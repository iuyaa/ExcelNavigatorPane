"""Live upload check. Writes only unique diagnostics objects; never deletes files.

Run from any directory: python tests/RustFSUploadCheck.py
Requires the local .env and installed boto3, python-dotenv, requests.
"""
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urlsplit
from uuid import uuid4

import boto3
import requests
from botocore.config import Config
from botocore.exceptions import ClientError
from dotenv import dotenv_values


def main():
    repo = Path(__file__).resolve().parents[1]
    config = dotenv_values(repo / '.env', interpolate=False)
    endpoint = config.get('TYAPP_S3_ENDPOINT', '').rstrip('/')
    url = urlsplit(endpoint)
    if (url.scheme != 'https' or not url.hostname or url.username or url.password
            or url.path or url.query or url.fragment):
        raise ValueError('Expected an HTTPS S3 origin without credentials or a path')
    bucket = config.get('TYAPP_S3_BUCKET')
    if bucket != 'excel-navigation' or config.get('TYAPP_S3_FORCE_PATH_STYLE', '').lower() != 'true':
        raise ValueError('Expected excel-navigation bucket and path-style addressing')
    for name in ('TYAPP_S3_REGION', 'TYAPP_S3_ACCESS_KEY_ID', 'TYAPP_S3_SECRET_ACCESS_KEY'):
        if not config.get(name):
            raise ValueError('Missing configuration: ' + name)
    client = boto3.client(
        's3', endpoint_url=endpoint, region_name=config['TYAPP_S3_REGION'],
        aws_access_key_id=config['TYAPP_S3_ACCESS_KEY_ID'],
        aws_secret_access_key=config['TYAPP_S3_SECRET_ACCESS_KEY'],
        config=Config(signature_version='s3v4', s3={'addressing_style': 'path'},
                      connect_timeout=10, read_timeout=20, retries={'max_attempts': 0},
                      request_checksum_calculation='when_required',
                      response_checksum_validation='when_required'))
    session = requests.Session()  # No browser session, cookies, or credentials.
    session.trust_env = False  # Avoid implicit netrc authentication on anonymous checks.
    prefix = 'diagnostics/upload-check-' + uuid4().hex
    key = prefix + '/Application Files/目录/校验 a+b%.txt'
    report = {'time': datetime.now(timezone.utc).isoformat(), 'endpoint': endpoint,
              'bucket': bucket, 'test_key': key, 'checks': [], 'complete': False}

    def passed(name):
        report['checks'].append(name)
        print('PASS:', name)

    def denied(name, operation):
        try:
            operation()
        except ClientError as error:
            if error.response['ResponseMetadata']['HTTPStatusCode'] == 403:
                passed(name)
                return
            raise
        raise RuntimeError('Expected access denied: ' + name)

    try:
        client.list_objects_v2(Bucket=bucket, MaxKeys=1)
        passed('authenticated bucket listing')
        denied('bucket policy administration denied', lambda: client.get_bucket_policy(Bucket=bucket))
        denied('unassigned bucket listing denied', lambda: client.list_objects_v2(
            Bucket='tyobj-test20260829', MaxKeys=1))
        # A missing key tests permission without deleting any existing object if misconfigured.
        denied('object deletion denied', lambda: client.delete_object(Bucket=bucket, Key=prefix + '/absent'))
        anonymous_list = session.get(endpoint + '/' + bucket,
                                     params={'list-type': '2', 'max-keys': '1'}, timeout=20,
                                     allow_redirects=False)
        if anonymous_list.status_code != 403:
            raise RuntimeError('Anonymous bucket listing was not denied')
        passed('anonymous bucket listing denied')

        public_url = endpoint + '/' + bucket + '/' + quote(key, safe='/')
        for revision in (1, 2):
            payload = ('Excel Navigator upload verification; revision=' + str(revision) + '\n').encode()
            client.put_object(Bucket=bucket, Key=key, Body=payload,
                              ContentType='text/plain', CacheControl='no-store')
            body = client.get_object(Bucket=bucket, Key=key)['Body']
            try:
                authenticated = body.read()
            finally:
                body.close()
            response = session.get(public_url, timeout=20, allow_redirects=False)
            expected = hashlib.sha256(payload).hexdigest()
            if (response.status_code != 200 or hashlib.sha256(authenticated).hexdigest() != expected
                    or hashlib.sha256(response.content).hexdigest() != expected):
                raise RuntimeError('Download status or SHA-256 mismatch')
            report['sha256'] = expected
            report['download_headers'] = {name: response.headers.get(name) for name in
                ('Content-Type', 'Content-Disposition', 'Cache-Control', 'X-Content-Type-Options')}
            passed('upload/overwrite and signed/anonymous SHA-256; revision=' + str(revision))

        anonymous_put = session.put(endpoint + '/' + bucket + '/' + prefix + '/anonymous.txt',
                                    data=b'Anonymous upload must be denied.\n', timeout=20,
                                    allow_redirects=False)
        if anonymous_put.status_code != 403:
            raise RuntimeError('Anonymous upload was not denied')
        passed('anonymous upload denied')
        report['complete'] = True
    finally:
        session.close()
        output = repo / 'work' / 'rustfs' / 'upload-check.json'
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print('Report:', output)


if __name__ == '__main__':
    try:
        main()
    except ClientError as error:
        # Do not print server messages, signed headers, or credential values.
        print('FAIL: S3 HTTP', error.response.get('ResponseMetadata', {}).get('HTTPStatusCode'))
        raise SystemExit(1)
    except Exception as error:
        print('FAIL:', type(error).__name__)
        raise SystemExit(1)
