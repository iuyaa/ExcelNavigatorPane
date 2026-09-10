"""Offline promotion-order and immutable-version regression check; no network."""
import hashlib
import importlib.util
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from urllib.parse import unquote
from zipfile import ZipFile

from botocore.exceptions import ClientError

spec = importlib.util.spec_from_file_location('publisher', Path(__file__).resolve().parents[1] / 'installer/Publish-RustFS.py')
publisher = importlib.util.module_from_spec(spec)
spec.loader.exec_module(publisher)


def main():
    base = 'https://updates.example.invalid/excel-navigation/'
    remote, writes = {}, []

    def get_object(Bucket, Key):
        import io
        if Key not in remote:
            raise ClientError({'Error': {'Code': 'NoSuchKey'}, 'ResponseMetadata': {'HTTPStatusCode': 404}}, 'GetObject')
        return {'Body': io.BytesIO(remote[Key])}

    def put_object(**args):
        writes.append(args['Key'])
        remote[args['Key']] = args['Body']

    def download(url, **kwargs):
        key = unquote(url.removeprefix(base))
        return SimpleNamespace(status_code=200, content=remote[key], headers={'Content-Type': 'application/x-ms-vsto'})

    settings = {'TYAPP_S3_ENDPOINT': base.split('/excel-navigation/')[0], 'TYAPP_S3_BUCKET': 'excel-navigation',
                'TYAPP_S3_FORCE_PATH_STYLE': 'true', 'TYAPP_S3_REGION': 'us-east-1',
                'TYAPP_S3_ACCESS_KEY_ID': 'fixture', 'TYAPP_S3_SECRET_ACCESS_KEY': 'fixture'}
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp)
        name = 'ExcelNavigator-Setup-1.0.0.2.exe'
        (path / name).write_bytes(b'fixture installer')
        (path / (name + '.sha256')).write_text(hashlib.sha256(b'fixture installer').hexdigest() + '  ' + name)
        (path / '安装说明.txt').write_text('fixture', encoding='utf-8')
        with ZipFile(path / 'ExcelNavigator-Publish-1.0.0.2.zip', 'w') as zipped:
            zipped.writestr('ExcelNavigatorPane.vsto', '<assembly xmlns="urn:schemas-microsoft-com:asm.v1"><assemblyIdentity name="ExcelNavigatorPane.vsto" version="1.0.0.2" /></assembly>')
            zipped.writestr('Application Files/1_0_0_2/fixture.dll', b'fixture dll')
            zipped.writestr('setup.exe', base.encode('utf-16le'))
        argv = ['Publish-RustFS.py', '--version', '1.0.0.2', '--directory', temp, '--apply']
        with patch.object(sys, 'argv', argv), patch.object(publisher, 'dotenv_values', return_value=settings), \
             patch.object(publisher.boto3, 'client', return_value=SimpleNamespace(get_object=get_object, put_object=put_object)), \
             patch.object(publisher.requests, 'Session', return_value=SimpleNamespace(get=download)):
            publisher.main()
            assert writes[-1] == 'ExcelNavigatorPane.vsto' and writes.index('setup.exe') > writes.index('Application Files/1_0_0_2/fixture.dll')
            writes.clear()
            publisher.main()
            assert all(not key.startswith(('Application Files/', 'releases/')) for key in writes)
            for key, replacement in [('Application Files/1_0_0_2/fixture.dll', b'conflict'),
                                     ('ExcelNavigatorPane.vsto', remote['ExcelNavigatorPane.vsto'].replace(b'1.0.0.2', b'1.0.0.3'))]:
                original = remote[key]
                remote[key] = replacement
                writes.clear()
                try:
                    publisher.main()
                except ValueError:
                    assert not writes
                else:
                    raise AssertionError('Conflict or downgrade accepted')
                remote[key] = original
    print('PASS: promotion order, exact retry, immutable conflict and downgrade rejection; no network')


if __name__ == '__main__':
    main()
