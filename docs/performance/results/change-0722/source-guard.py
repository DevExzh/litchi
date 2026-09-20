#!/usr/bin/env python3
"""Prove the writer pilot leaves existing read implementations unchanged."""
import hashlib
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
PREFIX = Path('crates/litchi-docx/src')


def read(lane, name):
    return (P / lane / PREFIX / name).read_bytes()


def sha(data):
    return hashlib.sha256(data).hexdigest()


namespace = read('baseline', 'namespace.rs')
assert namespace == read('candidate', 'namespace.rs')
codec = read('baseline', 'alt/codec.rs')
candidate_codec = read('candidate', 'alt/codec.rs')
declarations = b'mod document_scan;\npub(crate) use document_scan::scan_with_block_ranges;\n\n'
assert candidate_codec.count(declarations) == 1
assert candidate_codec.replace(declarations, b'', 1) == codec
document = read('baseline', 'parts/document_part.rs')
start = document.index(b'/// Select the active supported block ranges from original document XML.')
end = document.index(b'/// Select active direct children of the main document body in source order.', start)
expected_document = (document[:start] + document[end:]).replace(
    b'use crate::alt::{Chunk, active, scan};', b'use crate::alt::{Chunk, scan};'
).replace(b'use std::collections::BTreeSet;\n', b'')
assert expected_document == read('candidate', 'parts/document_part.rs')
child = read('candidate', 'alt/codec/document_scan.rs')
for public_function in [b'pub fn scan(', b'pub fn active(', b'pub fn is_relationship(']:
    assert public_function not in child
assert child.index(b'drop(reader);') < child.index(b'let active_offsets = active(xml, &offsets)?;')
assert child.index(b'drop(range_scanner);') < child.index(b'let active_offsets = active(xml, &offsets)?;')
result = {
    'status': 'pass',
    'script_sha256': sha(Path(__file__).read_bytes()),
    'namespace_whole_file_unchanged': sha(namespace),
    'codec_unchanged_except_private_child_declarations': sha(codec),
    'document_part_unchanged_except_dead_writer_helper_and_import_removal': sha(document),
    'removed_writer_helper_sha256': sha(document[start:end]),
    'candidate_files': {str(p.relative_to(P / 'candidate')): sha(p.read_bytes())
                        for p in sorted((P / 'candidate').rglob('*.rs'))},
    'parser_drops_precede_first_mce': True,
    'scope': 'Exact source transformation guard; release code generation and resource outcomes require measurement.',
}
encoded = json.dumps(result, indent=2) + '\n'
out = P / 'source-guard.json'
if sys.argv[1:] == ['--check']:
    assert out.read_text() == encoded
else:
    assert not sys.argv[1:]
    out.write_text(encoded)
print('PASS: existing public codec, namespace scanner, and read consumers preserved')
