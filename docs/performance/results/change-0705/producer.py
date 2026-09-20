#!/usr/bin/env python3
"""Separate producer control: numeric edit on planning sheet, SST elsewhere."""
from capture import BINARY, HERE, capture

for repeat in [1, 2]:
    name = f'producer-r{repeat}'
    capture(name, ['taskset', '-c', '12', str(BINARY), '--warmup', '20', '--samples', '100',
                   '--case', 'xlsx_producer_medium_source_one_edit_save',
                   '--producer-evidence', str(HERE / (name + '.producer.json')),
                   '--json', str(HERE / (name + '.json'))])
