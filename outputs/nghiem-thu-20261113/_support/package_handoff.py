"""Đóng gói snapshot hỗ trợ để clone có thể tái xuất báo cáo đã nghiệm thu."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
archive = OUT / 'support-data-20261113.zip'
manifest = OUT / 'handoff-support-manifest.json'
inputs = [p for p in OUT.glob('*.json')
          if p != manifest and not p.name.endswith('-preview-cells.json')]
for directory in OUT.glob('backup-*'):
    if directory.is_dir():
        inputs.extend(p for p in directory.rglob('*') if p.is_file())
inputs.sort(key=lambda p: p.relative_to(OUT).as_posix())
records = []
temporary = archive.with_suffix('.zip.writing')
with ZipFile(temporary, 'w', compression=ZIP_DEFLATED, compresslevel=9) as package:
    for source in inputs:
        relative = source.relative_to(OUT).as_posix()
        package.write(source, relative)
        with source.open('rb') as stream:
            source_digest = hashlib.file_digest(stream, 'sha256').hexdigest()
        records.append({'file': relative, 'bytes': source.stat().st_size,
                        'sha256': source_digest})
with ZipFile(temporary) as package:
    assert package.testzip() is None
    assert package.namelist() == [record['file'] for record in records]
    for record in records:
        with package.open(record['file']) as stream:
            assert hashlib.file_digest(stream, 'sha256').hexdigest() == record['sha256']
assert temporary.stat().st_size < 100 * 1024 * 1024, 'Gói vượt giới hạn file GitHub'
os.replace(temporary, archive)
with archive.open('rb') as stream:
    digest = hashlib.file_digest(stream, 'sha256').hexdigest()
manifest.write_text(json.dumps({'archive': archive.name, 'bytes': archive.stat().st_size,
                               'sha256': digest, 'verified': True, 'files': records},
                              ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'archive': archive.name, 'files': len(records),
                  'MiB': round(archive.stat().st_size / 1024**2, 2), 'verified': True}, ensure_ascii=False))
