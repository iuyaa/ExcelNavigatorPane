"""Upload the verified EXE/MSI release; promote signed latest.xml last.

Requires boto3, python-dotenv, requests, cryptography. Dry-run unless --apply.
Never removes old objects or overwrites differing immutable release files.
"""
import argparse
import base64
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote, urlsplit

import boto3
import requests
from botocore.config import Config
from botocore.exceptions import ClientError
from dotenv import dotenv_values
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding, rsa


def version_of(data):
    if len(data) > 1024 * 1024 or b'<!DOCTYPE' in data.upper():
        raise ValueError('Invalid update manifest')
    root = ET.fromstring(data)
    if root.tag != 'release' or len(root) or root.get('product') != 'ExcelNavigatorPane':
        raise ValueError('Wrong update manifest')
    version = root.get('version', '')
    if not re.fullmatch(r'(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.0', version):
        raise ValueError('Invalid version')
    parts = tuple(map(int, version.split('.')))
    if parts[0] > 255 or parts[1] > 255 or parts[2] > 65535 or parts < (1, 0, 1, 0):
        raise ValueError('Version exceeds MSI limits')
    if (root.get('file') != f'releases/{version}/ExcelNavigator-Setup-{version}.exe'
            or not re.fullmatch(r'[0-9a-f]{64}', root.get('sha256', ''))):
        raise ValueError('Invalid installer path or hash')
    return parts


def verify_release(manifest, public_key):
    version_of(manifest)
    root = ET.fromstring(manifest)
    key = ET.fromstring(public_key)
    number = lambda name: int.from_bytes(base64.b64decode(key.findtext(name), validate=True), 'big')
    public = rsa.RSAPublicNumbers(number('Exponent'), number('Modulus')).public_key()
    message = '\n'.join(root.attrib[name] for name in ('product', 'version', 'file', 'sha256'))
    public.verify(base64.b64decode(root.get('signature', ''), validate=True), message.encode(),
                  padding.PKCS1v15(), hashes.SHA256())
    return root


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--version', required=True)
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()
    if not re.fullmatch(r'\d+\.\d+\.\d+\.\d+', args.version):
        raise ValueError('Four-part version required')
    repo = Path(__file__).resolve().parents[1]
    values = dotenv_values(repo / '.env', interpolate=False)
    endpoint = values.get('TYAPP_S3_ENDPOINT', '').rstrip('/')
    origin = urlsplit(endpoint)
    if (origin.scheme != 'https' or not origin.hostname or origin.username or origin.password
            or origin.path or origin.query or origin.fragment):
        raise ValueError('HTTPS origin required')
    bucket = values.get('TYAPP_S3_BUCKET')
    if bucket != 'excel-navigation' or values.get('TYAPP_S3_FORCE_PATH_STYLE', '').lower() != 'true':
        raise ValueError('Wrong bucket or addressing mode')
    base = endpoint + '/' + bucket + '/'
    manifest = (args.directory / 'latest.xml').read_bytes()
    version = tuple(map(int, args.version.split('.')))
    if version_of(manifest) != version:
        raise ValueError('Manifest version mismatch')
    release = verify_release(manifest, (args.directory / 'update-public-key.xml').read_bytes())
    exe = 'ExcelNavigator-Setup-' + args.version + '.exe'
    installer = (args.directory / exe).read_bytes()
    checksum = (args.directory / (exe + '.sha256')).read_bytes()
    if checksum.decode('ascii').split()[0].lower() != hashlib.sha256(installer).hexdigest():
        raise ValueError('Installer checksum mismatch')
    if release.get('sha256') != hashlib.sha256(installer).hexdigest():
        raise ValueError('Signed installer hash mismatch')
    names = [exe, exe + '.sha256', '安装说明.txt', 'latest.xml', 'update-public-key.xml']
    names += [f'ExcelNavigator-{args.version}-{arch}.msi' for arch in ('x86', 'x64')]
    files = {'releases/' + args.version + '/' + name: (args.directory / name).read_bytes() for name in names}
    files['latest.xml'] = manifest
    # Keep old ClickOnce objects untouched; promote only the new MSI update entry.
    immutable = lambda key: key.startswith('releases/')
    ordered = sorted(files, key=lambda key: (key == 'latest.xml', key))
    print(('APPLY' if args.apply else 'DRY RUN') + ': ' + args.version + ', ' + str(len(files)) + ' objects -> ' + base)
    if not args.apply:
        return
    client = boto3.client('s3', endpoint_url=endpoint, region_name=values['TYAPP_S3_REGION'],
        aws_access_key_id=values['TYAPP_S3_ACCESS_KEY_ID'], aws_secret_access_key=values['TYAPP_S3_SECRET_ACCESS_KEY'],
        config=Config(signature_version='s3v4', s3={'addressing_style': 'path'},
                      connect_timeout=10, read_timeout=30, retries={'max_attempts': 0},
                      request_checksum_calculation='when_required', response_checksum_validation='when_required'))
    session = requests.Session()
    session.trust_env = False

    def existing(key):
        try:
            body = client.get_object(Bucket=bucket, Key=key)['Body']
            try:
                return body.read()
            finally:
                body.close()
        except ClientError as error:
            if error.response['Error']['Code'] == 'NoSuchKey':
                return None
            raise

    previous = existing('latest.xml')
    if previous is not None and (version_of(previous) > version or
            (version_of(previous) == version and previous != manifest)):
        raise ValueError('Refusing downgrade or different artifacts with the same version')
    # Preflight every immutable name before any remote mutation.
    present = {}
    for key in ordered:
        if immutable(key):
            present[key] = existing(key)
            if present[key] is not None and present[key] != files[key]:
                raise ValueError('Immutable remote artifact differs: ' + key)
    hashes = {}
    for key in ordered:
        if key == 'latest.xml' and existing(key) != previous:
            raise ValueError('Live manifest changed during upload; stop and inspect')
        data = files[key]
        if present.get(key) != data:
            client.put_object(Bucket=bucket, Key=key, Body=data, CacheControl='no-store',
                              ContentType='application/xml' if key.endswith('.xml') else 'application/octet-stream')
        response = session.get(base + quote(key, safe='/'), timeout=30, allow_redirects=False)
        if response.status_code != 200 or response.content != data:
            raise ValueError('Anonymous download verification failed: ' + key)
        hashes[key] = hashlib.sha256(data).hexdigest()
    report = args.directory / 'rustfs-publish.json'
    report.write_text(json.dumps({'version': args.version, 'base_url': base, 'sha256': hashes},
                                indent=2, ensure_ascii=False), encoding='utf-8')
    print('PASS: uploaded and anonymously verified every object; live manifest promoted last')


if __name__ == '__main__':
    try:
        main()
    except ClientError as error:
        print('FAIL: S3 HTTP', error.response.get('ResponseMetadata', {}).get('HTTPStatusCode'))
        raise SystemExit(1)
    except (ValueError, FileNotFoundError) as error:
        print('FAIL:', str(error))
        raise SystemExit(1)
    except Exception as error:
        print('FAIL:', type(error).__name__)
        raise SystemExit(1)
