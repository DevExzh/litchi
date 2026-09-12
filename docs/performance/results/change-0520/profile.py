"""Retain separately dumped commit calls; untimed lifecycle calls stay visible."""
import json
from capture import HERE, BINARY, capture


def main():
    plan = json.loads((HERE / 'profile-plan.json').read_text())
    native = json.loads((HERE / 'plan.json').read_text())
    for repeat in range(1, plan['repeats'] + 1):
        for shape in plan['shapes']:
            name = f'profile-r{repeat}-{shape}'
            capture(name, ['taskset', '-c', str(native['cpu']), 'valgrind',
                           *plan['options'], '--callgrind-out-file=' + str(HERE / (name + '.callgrind')),
                           str(BINARY), '--warmup', str(plan['warmup']),
                           '--samples', str(plan['samples']), '--case', native['case'],
                           '--xlsx-cell-crud-shape', shape, '--json', str(HERE / (name + '.json'))])


if __name__ == '__main__':
    main()
