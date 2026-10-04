"""Admit exact setup/measured contexts before continuing the v3 schedule."""
import importlib.util
import json
from pathlib import Path

P = Path(__file__).resolve().parent


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def main():
    cap = module('capture_v3', P / 'capture_v3.py')
    reader = module('reader_0845', P.parent / 'change-0845/reader_v2.py')
    plan = cap.read(P / 'plan-v3.json')
    for item in [plan['binary'], plan['helper'], *plan['inputs'], *plan['tools']]:
        assert cap.desc(item['path']) == item, item['path']
    row = plan['schedule'][0]
    folder = P / 'capture-v3' / row['label']
    command = cap.read(folder / 'command.json')
    assert command['exit_code'] == 0 and command['row'] == row
    assert command['plan'] == cap.desc(P / 'plan-v3.json')
    assert command['binary'] == plan['binary'] and command['helper'] == plan['helper']
    assert command['cwd'] == plan['cwd'] and command['environment'] == plan['environment']
    expected_argv = ['/usr/bin/taskset', '-c', str(plan['cpu']), '/usr/bin/valgrind', *plan['options'], '--callgrind-out-file=' + str(folder / 'callgrind.out'), plan['binary']['path'], row['kind'], '0', '1', str(folder / 'artifacts')]
    assert command['argv'] == expected_argv
    for item in command['artifacts']:
        assert cap.desc(item['path']) == item
    assert cap.read(folder / 'semantics.json') == cap.validate_semantics(plan, row, folder)
    contexts = {}
    terminal = []
    for path in sorted(folder.glob('callgrind.out*')):
        raw = reader.parse_raw(path, allow_empty=True)
        header = raw['header']
        assert header['events'] == ['Ir'] and header['summary_ir'] == header['totals_ir']
        assert ' '.join(header['command'].split()) == ' '.join(expected_argv[expected_argv.index(plan['binary']['path']):])
        if header['summary_ir'] == 0:
            assert header['trigger'] == 'Program termination'
            assert sum(f['self_ir'] for f in raw['functions'].values()) == 0
            terminal.append(cap.desc(path))
            continue
        assert header['trigger'] == '--dump-after=' + plan['owner']
        ids = [key for key, name in raw['names'].items() if name == plan['owner']]
        assert len(ids) == 1
        owner = raw['functions'][ids[0]]
        edges = reader.incoming_edges(raw, ids[0])
        assert len(edges) == 1 and edges[0]['calls'] == 1
        edge = edges[0]
        labels = [label for label, caller in plan['expected_positive_callers'].items() if caller == edge['caller']]
        assert len(labels) == 1 and labels[0] not in contexts
        child_sum = sum(e['inclusive_ir'] for e in reader.aggregate_edges(owner, raw['names']))
        all_self = sum(f['self_ir'] for f in raw['functions'].values())
        assert owner['self_ir'] + child_sum == edge['inclusive_ir'] == all_self == header['summary_ir']
        contexts[labels[0]] = {'raw': cap.desc(path), 'caller': edge['caller'], 'calls': edge['calls'], 'owner_ir': edge['inclusive_ir'], 'owner_self_ir': owner['self_ir'], 'direct_children_ir': child_sum, 'all_function_self_ir': all_self}
    assert set(contexts) == {'setup', 'measured'} and len(terminal) == 1
    cap.write(P / 'first-child-admission-v3.json', {'status': 'pass', 'plan': cap.desc(P / 'plan-v3.json'), 'helper': cap.desc(__file__), 'reader': cap.desc(P.parent / 'change-0845/reader_v2.py'), 'command': cap.desc(folder / 'command.json'), 'semantics': cap.desc(folder / 'semantics.json'), 'contexts': contexts, 'terminal': terminal, 'classification': 'Exact incoming callers; no dump-order selection. Setup excluded from measured comparison.'})
    print('0911 v3 first-child admission PASS')


if __name__ == '__main__':
    main()
