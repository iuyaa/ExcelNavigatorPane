"""Offline release promotion checks; all S3/HTTP calls are mocked."""
import base64
import hashlib
import importlib.util
import io
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from urllib.parse import unquote
import xml.etree.ElementTree as ET

from botocore.exceptions import ClientError
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding, rsa

spec = importlib.util.spec_from_file_location('publisher', Path(__file__).resolve().parents[1] / 'installer/Publish-RustFS.py')
publisher = importlib.util.module_from_spec(spec)
spec.loader.exec_module(publisher)


def main():
    base = 'https://updates.example.invalid/excel-navigation/'
    remote, writes = {'ExcelNavigatorPane.vsto': b'keep legacy entry'}, []
    bad_download, concurrent_publish = False, False

    def get_object(Bucket, Key):
        if Key not in remote:
            raise ClientError({'Error': {'Code': 'NoSuchKey'}}, 'GetObject')
        return {'Body': io.BytesIO(remote[Key])}

    def put_object(**args):
        writes.append(args['Key'])
        remote[args['Key']] = args['Body']

    def download(url, **kwargs):
        key = unquote(url.removeprefix(base))
        if concurrent_publish:
            remote['latest.xml'] = newer
        return SimpleNamespace(status_code=200, content=b'corrupted' if bad_download else remote[key])

    def reject():
        try:
            publisher.main()
        except ValueError:
            return
        raise AssertionError('Unsafe publication accepted')

    settings = {'TYAPP_S3_ENDPOINT': base.split('/excel-navigation/')[0], 'TYAPP_S3_BUCKET': 'excel-navigation',
                'TYAPP_S3_FORCE_PATH_STYLE': 'true', 'TYAPP_S3_REGION': 'us-east-1',
                'TYAPP_S3_ACCESS_KEY_ID': 'fixture', 'TYAPP_S3_SECRET_ACCESS_KEY': 'fixture'}
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    numbers = key.public_key().public_numbers()
    b64int = lambda x: base64.b64encode(x.to_bytes((x.bit_length() + 7) // 8, 'big')).decode()
    public = f'<RSAKeyValue><Modulus>{b64int(numbers.n)}</Modulus><Exponent>{b64int(numbers.e)}</Exponent></RSAKeyValue>'
    name = 'ExcelNavigator-Setup-1.0.1.0.exe'
    installer = b'fixture installer'
    digest = hashlib.sha256(installer).hexdigest()
    root = ET.Element('release', product='ExcelNavigatorPane', version='1.0.1.0',
                      file='releases/1.0.1.0/' + name, sha256=digest)
    message = '\n'.join(root.attrib[n] for n in ('product', 'version', 'file', 'sha256'))
    root.set('signature', base64.b64encode(key.sign(message.encode(), padding.PKCS1v15(), hashes.SHA256())).decode())
    manifest = ET.tostring(root)
    newer = manifest.replace(b'1.0.1.0', b'1.0.2.0')
    publisher.verify_release(manifest, public)
    notes = base64.b64encode('## 1.0.1.0\n### 新增功能\n- 功能说明'.encode()).decode()
    root.set('notes', notes)
    root.set('notesSignature', base64.b64encode(key.sign(
        ('ExcelNavigatorPane notes\n1.0.1.0\n' + notes).encode(), padding.PKCS1v15(), hashes.SHA256())).decode())
    manifest = ET.tostring(root)
    publisher.verify_release(manifest, public)
    from cryptography.exceptions import InvalidSignature
    try:
        publisher.verify_release(manifest.replace(notes.encode(), b'dGFtcGVyZWQ='), public)
        raise AssertionError('Tampered release notes accepted')
    except InvalidSignature:
        pass
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp)
        for filename, data in {name: installer, name + '.sha256': (digest + '  ' + name).encode(),
                               'latest.xml': manifest, 'update-public-key.xml': public.encode(),
                               '安装说明.txt': b'fixture', 'ExcelNavigator-1.0.1.0-x86.msi': b'x86',
                               'ExcelNavigator-1.0.1.0-x64.msi': b'x64'}.items():
            (path / filename).write_bytes(data)
        argv = ['Publish-RustFS.py', '--version', '1.0.1.0', '--directory', temp, '--apply']
        with patch.object(sys, 'argv', argv), patch.object(publisher, 'dotenv_values', return_value=settings), \
             patch.object(publisher.boto3, 'client', return_value=SimpleNamespace(get_object=get_object, put_object=put_object)), \
             patch.object(publisher.requests, 'Session', return_value=SimpleNamespace(get=download)):
            publisher.main()
            assert len(writes) == 8 and writes[-1] == 'latest.xml'
            assert remote['ExcelNavigatorPane.vsto'] == b'keep legacy entry'
            baseline = remote.copy()
            writes.clear()
            publisher.main()
            assert writes == ['latest.xml']
            for object_key, replacement in [('releases/1.0.1.0/' + name, b'conflict'), ('latest.xml', newer),
                                             ('latest.xml', manifest.replace(digest.encode(), b'0' * 64))]:
                remote[object_key] = replacement
                writes.clear()
                reject()
                assert not writes
                remote.clear()
                remote.update(baseline)
            writes.clear()
            bad_download = True
            reject()
            assert 'latest.xml' not in writes
            bad_download = False
            concurrent_publish = True
            writes.clear()
            reject()
            assert 'latest.xml' not in writes and remote['latest.xml'] == newer
            concurrent_publish = False
            remote.clear()
            remote.update(baseline)
            (path / name).write_bytes(b'tampered installer')
            writes.clear()
            reject()
            assert not writes
            try:
                publisher.verify_release(newer, public)
            except Exception as error:
                assert type(error).__name__ == 'InvalidSignature'
            else:
                raise AssertionError('Tampered signature accepted')
    print('PASS: signed metadata, promote last, preserve legacy, retry, conflict/downgrade/hash rejection, download failure and concurrent publication; no network')


if __name__ == '__main__':
    main()
