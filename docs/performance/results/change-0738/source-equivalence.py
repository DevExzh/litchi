#!/usr/bin/env python3
"""Check fixed source blocks against the original retained probe."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
before = (P/'archive-probe/src/lib.rs').read_text()
after = (P/'probe/src/lib.rs').read_text()


def block(source, marker):
    start = source.index(marker)
    cursor = source.index('{', start)+1
    depth = 1
    # The selected fixed blocks have balanced braces in their string literals.
    # This witness is intentionally not a general Rust parser.
    while depth:
        if source[cursor] == '{':
            depth += 1
        elif source[cursor] == '}':
            depth -= 1
        cursor += 1
    return source[start:cursor]


proof = {}
for marker in ('#[derive(Clone, Debug, Serialize)]\nstruct Sample {',
               'fn oracle_for_output(', 'fn public_format_edit(',
               'fn measured_public_format(', 'fn timed_format(', 'fn output_sample('):
    original = block(before,marker)
    restored = block(after,marker)
    assert original == restored, marker
    proof[marker] = dict(bytes=len(original.encode()),
                         sha256=hashlib.sha256(original.encode()).hexdigest(),identical=True)
(P/'source-equivalence.json').write_text(json.dumps(dict(status='passed',blocks=proof,
    scope='Exact source fields and owner/oracle blocks; no ABI size, address or unique runtime cause claim.'),indent=2)+'\n')
print('PASS original Sample definition and five owner/oracle function blocks')
