"""Replay untouched ZIP payload, metadata, and ordering preservation independently."""
import argparse
import hashlib
import json
from pathlib import Path
import struct
import zipfile

P = Path(__file__).resolve().parent
FIELDS = ('date_time', 'compress_type', 'flag_bits', 'create_system', 'create_version',
          'extract_version', 'reserved', 'volume', 'internal_attr', 'external_attr',
          'extra', 'comment', 'CRC', 'file_size', 'compress_size')


def digest(data):
    return hashlib.sha256(data).hexdigest()


def compressed(data, info):
    offset = info.header_offset
    assert data[offset:offset + 4] == b'PK\x03\x04'
    name_len, extra_len = struct.unpack_from('<HH', data, offset + 26)
    start = offset + 30 + name_len + extra_len
    return data[start:start + info.compress_size]


def analyze():
    manifest = json.loads((P / 'artifacts/manifest.json').read_text())
    rows = []
    for case in manifest['cases']:
        source_path = P / 'artifacts' / case['source_archive']['path']
        output_path = P / 'artifacts' / case['policy_outputs'][0]['output']['path']
        source_bytes, output_bytes = source_path.read_bytes(), output_path.read_bytes()
        with zipfile.ZipFile(source_path) as left, zipfile.ZipFile(output_path) as right:
            assert left.namelist() == right.namelist(), case['case_id']
            assert left.comment == right.comment, case['case_id']
            unchanged, changed = [], []
            for name in left.namelist():
                source, output = left.getinfo(name), right.getinfo(name)
                before, after = left.read(name), right.read(name)
                if before != after:
                    changed.append(name)
                    continue
                mismatched = [field for field in FIELDS if getattr(source, field) != getattr(output, field)]
                assert not mismatched, (case['case_id'], name, mismatched)
                raw_source, raw_output = compressed(source_bytes, source), compressed(output_bytes, output)
                assert raw_source == raw_output, (case['case_id'], name, 'compressed payload')
                unchanged.append({'name': name, 'decoded_sha256': digest(before), 'compressed_sha256': digest(raw_source), 'metadata_equal': True})
            if case['format'] == 'DOCX':
                assert changed == ['word/document.xml'], (case['case_id'], changed)
                if case['origin'] == 'caller-named-real-file':
                    assert any(row['name'] == 'word/_rels/document.xml.rels' for row in unchanged)
            rows.append({'case_id': case['case_id'], 'source_sha256': digest(source_bytes), 'output_sha256': digest(output_bytes), 'member_order_equal': True, 'archive_comment_equal': True, 'changed_members': changed, 'untouched_members': unchanged})
    return {'schema': 'litchi.performance.0818.zip-preservation.v1', 'cases': rows, 'metadata_fields': list(FIELDS), 'scope': 'default output per case; five-policy byte equality is checked independently by artifact admission'}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = analyze()
    target = P / 'zip-preservation.json'
    encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
    if args.check:
        assert target.read_text() == encoded
    else:
        assert not target.exists()
        target.write_text(encoded)
    print('0818 ZIP preservation PASS: six default outputs; untouched metadata and compressed payloads; archive order/comments')


if __name__ == '__main__':
    main()
