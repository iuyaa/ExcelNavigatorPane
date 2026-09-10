"""Upload the verified release artifacts; promote the VSTO entry last.

Requires locally installed boto3, python-dotenv, requests. Dry-run unless --apply.
Never removes old objects or overwrites differing immutable release files.
"""
import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote, urlsplit
from zipfile import ZipFile

import boto3
import requests
from botocore.config import Config
from botocore.exceptions import ClientError
from dotenv import dotenv_values


def version_of(data):
    if len(data) > 1024 * 1024 or b'<!DOCTYPE' in data.upper():
        raise ValueError('Invalid deployment manifest')
    root = ET.fromstring(data)
    ns = '{urn:schemas-microsoft-com:asm.v1}'
    identity = root.find(ns + 'assemblyIdentity')
    if root.tag != ns + 'assembly' or identity is None or identity.get('name') != 'ExcelNavigatorPane.vsto':
        raise ValueError('Wrong deployment manifest')
    version = identity.get('version', '')
    if not re.fullmatch(r'\d+\.\d+\.\d+\.\d+', version):
        raise ValueError('Invalid version')
    return tuple(map(int, version.split('.')))


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
    archive = args.directory / ('ExcelNavigator-Publish-' + args.version + '.zip')
    files = {}
    with ZipFile(archive) as zipped:
        for entry in zipped.infolist():
            if entry.is_dir():
                continue
            name = entry.filename.replace('\\', '/')
            if (any(part in ('', '.', '..') for part in name.split('/'))
                    or ':' in name or name in files or entry.file_size > 128 * 1024 * 1024):
                raise ValueError('Invalid archive entry')
            files[name] = zipped.read(entry)
    manifest = files.get('ExcelNavigatorPane.vsto', b'')
    version = tuple(map(int, args.version.split('.')))
    if version_of(manifest) != version:
        raise ValueError('Archive version mismatch')
    if (base.encode('utf-16le') not in files.get('setup.exe', b'')):
        raise ValueError('Bootstrapper does not point to this update directory')
    exe = 'ExcelNavigator-Setup-' + args.version + '.exe'
    installer = (args.directory / exe).read_bytes()
    checksum = (args.directory / (exe + '.sha256')).read_bytes()
    if checksum.decode('ascii').split()[0].lower() != hashlib.sha256(installer).hexdigest():
        raise ValueError('Installer checksum mismatch')
    for name, content in {exe: installer, exe + '.sha256': checksum,
                          archive.name: archive.read_bytes(),
                          '安装说明.txt': (args.directory / '安装说明.txt').read_bytes()}.items():
        files['releases/' + args.version + '/' + name] = content
    # First all immutable assets, then bootstrapper resources, finally the live entry.
    immutable = lambda key: key.startswith(('Application Files/', 'releases/'))
    ordered = sorted(files, key=lambda key: (2 if key == 'ExcelNavigatorPane.vsto' else 0 if immutable(key) else 1, key))
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

    previous = existing('ExcelNavigatorPane.vsto')
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
        if key == 'ExcelNavigatorPane.vsto' and existing(key) != previous:
            raise ValueError('Live manifest changed during upload; stop and inspect')
        data = files[key]
        if present.get(key) != data:
            client.put_object(Bucket=bucket, Key=key, Body=data, CacheControl='no-store',
                              ContentType='application/x-ms-vsto' if key.endswith('.vsto') else 'application/octet-stream')
        response = session.get(base + quote(key, safe='/'), timeout=30, allow_redirects=False)
        if response.status_code != 200 or response.content != data:
            raise ValueError('Anonymous download verification failed: ' + key)
        if key == 'ExcelNavigatorPane.vsto' and response.headers.get('Content-Type', '').split(';')[0] != 'application/x-ms-vsto':
            raise ValueError('Incorrect public VSTO content type')
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
