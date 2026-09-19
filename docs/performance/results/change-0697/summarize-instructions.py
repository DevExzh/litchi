#!/usr/bin/env python3
"""Extract instruction observations; these are not estimates of saved time."""
import json
import re
from pathlib import Path

P = Path(__file__).resolve().parent
sites = {
    'start': [
        ('parent Ctx namespace clone', '4311a', '4311e'),
        ('parent Ctx directive clone', '43131', '43135'),
        ('temporary Inherited namespace clone', '43168', '4316c'),
        ('temporary Inherited emitted-boundary clone', '432fb', '432ff'),
        ('ordinary-path child emitted-boundary owner', '45b2e', '45b32'),
    ],
    'inherited-drop': [
        ('temporary Inherited namespace release', '4767f', '47683'),
        ('temporary Inherited emitted-boundary release', '47697', '4769b'),
    ],
    'ctx-drop': [('Ctx namespace release', '4746f', '47473')],
}
rows = []
for name, selected in sites.items():
    text = (P / 'instructions' / (name + '.stdout')).read_text()
    count = int(re.search(r'\((\d+) samples,', text).group(1))
    parsed = {}
    for line in text.splitlines():
        match = re.match(r'\s*(\d+)\s*:\s*([0-9a-f]+):\s*(.*)', line)
        if match:
            samples, address, instruction = match.groups()
            assert address not in parsed
            parsed[address] = dict(samples=int(samples), instruction=instruction)
    assert sum(row['samples'] for row in parsed.values()) == count
    assembly = (P / 'instructions' / (name + '-assembly.stdout')).read_text()
    observations = []
    for label, atomic, following in selected:
        assert re.search(r'\b' + atomic + r':\s+lock (incq|decq)', assembly)
        assert parsed[atomic]['instruction'] == 'lock'
        observations.append(dict(label=label, atomic_address='0x' + atomic,
                                 sampled_following_address='0x' + following,
                                 following_instruction=parsed[following]['instruction'],
                                 following_samples=parsed[following]['samples']))
    rows.append(dict(symbol_case=name, symbol_samples=count, observations=observations))
(P / 'instruction-summary.json').write_text(json.dumps(dict(
    scope='Instruction samples in the isolated MCE profile, not atomic latency or predicted savings; samples may skid.',
    rows=rows,
), indent=2) + '\n')
print('Instruction observations extracted and sample totals checked.')
