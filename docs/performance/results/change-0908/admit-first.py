"""Admit the retained first child only when its dump is exactly the timed call."""
import importlib.util
from pathlib import Path

P = Path(__file__).resolve().parent

def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

c = load('capture0908v2', P / 'capture_v2.py')
r = load('reader0845', P.parent / 'change-0845/reader_v2.py')
plan = c.read(P / 'plan-v2.json')
folder = P / 'capture/eager-a'
command = c.read(folder / 'command.json')
assert command['exit_code'] == 0
for descriptor in command['artifacts']:
    assert c.desc(descriptor['path']) == descriptor
semantic = c.validate_semantics(plan, plan['schedule'][0], folder)
assert semantic['status'] == 'pass'
paths = sorted(folder.glob('callgrind.out.*'))
assert len(paths) == 1, paths
parsed = r.parse_raw(paths[0])
header = parsed['header']
assert header['part'] == 1 and header['events'] == ['Ir']
assert header['trigger'] == '--dump-after=' + plan['owner']
assert header['summary_ir'] == header['totals_ir'] and header['summary_ir'] > 0
assert header['command'].split() == command['argv'][command['argv'].index(plan['binary']['path']):]
edges = [edge for fid, name in parsed['names'].items() if name == plan['owner'] for edge in r.incoming_edges(parsed, fid)]
assert len(edges) == 1
edge = edges[0]
assert edge['calls'] == 1 and edge['inclusive_ir'] == header['summary_ir']
assert edge['caller'] == plan['parent'] and edge['position'] == '51174'
assert edge['file'] == '/tmp/litchi-0907-src/tools/perf-baseline/src/lib.rs'
owner = r.function_view(parsed['functions'][edge['callee_id']], parsed['names'])
assert owner['self_ir'] + owner['direct_children_ir'] == header['summary_ir']
assert sum(f['self_ir'] for f in parsed['functions'].values()) == header['summary_ir']
terminal = r._header_0844((folder / 'callgrind.out').read_text().splitlines(), str(folder / 'callgrind.out'))
assert terminal['part'] == 2 and terminal['events'] == ['Ir']
assert terminal['summary_ir'] == terminal['totals_ir'] == 0 and terminal['trigger'] == 'Program termination'
c.write(P / 'first-child-admission.json', {'status': 'pass', 'plan': c.desc(P / 'plan-v2.json'), 'original_command': c.desc(folder / 'command.json'), 'corrected_semantics': c.desc(folder / 'semantics.json'), 'raw_dump': c.desc(paths[0]), 'terminal': c.desc(folder / 'callgrind.out'), 'reader': c.desc(P.parent / 'change-0845/reader_v2.py'), 'helper': c.desc(__file__), 'header': header, 'measured_edge': edge, 'owner': owner, 'scope': 'Exactly one collected measured call at51174; expected-output callsite51143 excluded. Instruction counts only. Zero terminal remainder.'})
print('first child admitted: one measured call, exact Ir conservation, zero terminal')
