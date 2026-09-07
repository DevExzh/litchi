#!/usr/bin/env python3
"""Check that the content-types oracle preserves the original closing whitespace."""
import importlib.util
import json
from pathlib import Path

path = Path(__file__).with_name("external-verifier.py")
spec = importlib.util.spec_from_file_location("external_oracle", path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
source = b'<Types><Default Extension="xml"/>\n</Types>'
insertion = b'<Override PartName="/added.xml"/>'
boundary = source.index(b'\n')
valid = source[:boundary] + insertion + source[boundary:]
module.prefix_suffix(source, valid, b'</Types>', 'valid', before_whitespace=True)
mutations = [
    valid.replace(b'\n</Types>', b'</Types>'),
    valid.replace(b'\n</Types>', b' </Types>'),
    valid.replace(b'Extension="xml"', b'Extension="bin"'),
]
for index, mutated in enumerate(mutations):
    try:
        module.prefix_suffix(source, mutated, b'</Types>', str(index), before_whitespace=True)
    except module.VerificationError:
        continue
    raise RuntimeError(f'preservation mutation {index} was accepted')
print(json.dumps({'status': 'pass', 'valid_insertions': 1, 'rejected_mutations': len(mutations)}))
