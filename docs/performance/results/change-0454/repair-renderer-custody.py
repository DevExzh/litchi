#!/usr/bin/env python3
"""Restore capture protocol identities and separate the later renderer correction."""
import datetime
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OLD_DERIVE = '316841a5e46e248c486bb2f7b4c16ae17eabc051539bdd2a67332ea2f4da4c3f'
NEW_DERIVE = '2be514cfb04c4e21d80881a605040943da608347555611c70e9b213ec1999797'
OLD_PROTOCOL = '6c3ab253d2e6680d1beffe975fc0c8fce472d10f04304eff986264ddfdaaaa3b'
INTERMEDIATE_PROTOCOL = '932d4ccb85820a460fe40ff5a43dd3eb0e1466182fe2cc4aaa0a84df8c868363'

def sha(raw):
    return hashlib.sha256(raw).hexdigest()

def encoded(value):
    return (json.dumps(value, indent=2) + '\n').encode()

corrected = (ROOT / 'derive.py').read_bytes()
if sha(corrected) != NEW_DERIVE:
    raise ValueError('unexpected corrected renderer identity')
old_block = '''        source_calls = work["source_reads"]["logical_calls"] if work["source_reads"] else 0
        destination_calls = work["destination_reads"]["logical_calls"] if work["destination_reads"] else 0'''
new_block = '''        source_reads = work.get("source_reads", work.get("source"))
        destination_reads = work.get("destination_reads", work.get("destination"))
        source_calls = source_reads["logical_calls"] if source_reads else 0
        destination_calls = destination_reads["logical_calls"] if destination_reads else 0'''
original = corrected.decode().replace(new_block, old_block).encode()
if sha(original) != OLD_DERIVE:
    raise ValueError('renderer correction is not the exact declared display-only change')
raw_protocol = (ROOT / 'protocol.json').read_bytes()
if sha(raw_protocol) != INTERMEDIATE_PROTOCOL:
    raise ValueError('unexpected intermediate protocol identity')
protocol = json.loads(raw_protocol)
protocol['bound_files']['derive.py'] = OLD_DERIVE
original_protocol = encoded(protocol)
if sha(original_protocol) != OLD_PROTOCOL:
    raise ValueError('capture protocol cannot be reconstructed exactly')
archive = ROOT / 'custody-corrections/renderer-r1'
archive.mkdir(parents=True)
for name in ['derive.py', 'protocol.json', 'measurements.json', 'measurements.md']:
    raw = (ROOT / name).read_bytes()
    (archive / (name + '.txt' if name.endswith('.py') else name)).write_bytes(raw)
rows = []
for kind, count in [('provider-runs', 16), ('external-runs', 2)]:
    for lane in range(count):
        relative = f'{kind}/{lane}/receipt.json'
        path = ROOT / relative
        raw = path.read_bytes()
        value = json.loads(raw)
        if value['protocol_sha256'] != INTERMEDIATE_PROTOCOL or value['status'] != 'pass':
            raise ValueError('unexpected intermediate receipt: ' + relative)
        target = archive / relative
        target.parent.mkdir(parents=True)
        target.write_bytes(raw)
        value['protocol_sha256'] = OLD_PROTOCOL
        restored = encoded(value)
        path.write_bytes(restored)
        rows.append({'path': relative, 'intermediate_sha256': sha(raw), 'restored_sha256': sha(restored)})
(ROOT / 'derive.py').write_bytes(original)
(ROOT / 'derive-final.py').write_bytes(corrected)
(ROOT / 'protocol.json').write_bytes(original_protocol)
amendment = {
    'status': 'pass', 'original_driver': 'derive.py', 'original_sha256': OLD_DERIVE,
    'corrected_driver': 'derive-final.py', 'corrected_sha256': NEW_DERIVE,
    'captured_protocol_sha256': OLD_PROTOCOL,
    'reason': 'Post-capture Markdown rendering used provider counter names for external rows. Only counter-key display lookup changed; sampling, statistics and timing inputs did not change.',
    'recorded_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
}
(ROOT / 'derivation-amendment.json').write_bytes(encoded(amendment))
(ROOT / 'custody-corrections/renderer-r1/restoration.json').write_bytes(encoded({
    'status': 'pass', 'intermediate_protocol_sha256': INTERMEDIATE_PROTOCOL,
    'restored_protocol_sha256': OLD_PROTOCOL, 'receipts': rows,
    'basis': 'The capture coordinator acknowledged changing only protocol_sha256 in the 18 receipts. Root reconstructed that original field from the captured pilot protocol hash and the exact inverse protocol edit; all intermediate bytes are retained here. This record discloses reconstruction, not uninterrupted receipt custody.',
    'recorded_utc': amendment['recorded_utc'],
}))
print('restored capture identities; renderer amendment retained separately')
