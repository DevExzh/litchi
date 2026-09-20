#!/usr/bin/env python3
"""Whole-child RSS includes generator, lifecycle oracles and output verification."""
from capture import BINARY, HERE, capture

for shape in ['medium', 'dense-sparse']:
    for repeat in [1, 2]:
        name = f'rss-r{repeat}-{shape}'
        capture(name, ['/usr/bin/time', '-v', '-o', str(HERE / (name + '.rss.txt')),
                       'taskset', '-c', '12', str(BINARY), '--warmup', '2', '--samples', '3',
                       '--case', 'xlsx_source_backed_cell_values_one_percent_edit_save',
                       '--xlsx-cell-crud-shape', shape, '--json', str(HERE / (name + '.json'))])
