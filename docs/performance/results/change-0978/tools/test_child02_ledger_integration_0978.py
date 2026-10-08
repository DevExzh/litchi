import importlib.util, json, re, sys
from pathlib import Path

ROOT = Path('/home/zhuhe/litchi-performance-0978')
TOOLS = ROOT / 'tools'
CHILD = ROOT / 'cpu-profile-child02-v1' / 'capture' / 'child-02'
DUMP = CHILD / 'child-02.perf-mmap-dump.txt'
DATA = CHILD / 'child-02.perf.data'
BINARY = Path('/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile')
PERF = Path('/usr/bin/perf')
PID = int(json.loads((CHILD / 'child-02.launch.json').read_text())['exec_pid'])
CPU = 31
PERIOD = 1_000_000
BUILD = 'e5c457384eeebd057932c444efda77618f137399'
TARGET = str(BINARY)
MARKER = 'doc_retention_probe_0942_cfb_profile::profiled_measured_case'


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

life = load('life_0978_child02', TOOLS / 'lifecycle_ledger_0978.py')
scanner = load('scanner_0978_child02', TOOLS / 'raw_stack_0978.py')
decoder = load('decoder_0978_child02', TOOLS / 'focuseddecode_0978.py')

print(json.dumps({'stage':'lifecycle_start','dump_bytes':DUMP.stat().st_size}), flush=True)
lifecycle = life.parse_lifecycle_dump(DUMP, expected_pid=PID, expected_cpu=CPU, require_raw_metadata=True)
print(json.dumps({'stage':'lifecycle_done','records':lifecycle['record_count'],'samples':lifecycle['sample_descriptor_count'],'raw_metadata':lifecycle['raw_metadata_count']}), flush=True)
contract = scanner.parse_evlist_contract((CHILD / 'child-02.perf-evlist-v.txt').read_text(), expected_period=PERIOD, expected_stack_cap=32768)
print(json.dumps({'stage':'scanner_start','pid':PID}), flush=True)
raw = scanner.scan_dump(DUMP, perf_data=DATA, binary=BINARY, expected_pid=PID, expected_cpu=CPU, expected_period=PERIOD, stack_cap=32768, event_contract=contract, retain_samples=True)
print(json.dumps({'stage':'scanner_done','samples':raw.get('sample_count'),'loss':raw.get('raw_lost_sum'),'violations':raw.get('violations')}), flush=True)
corrected, corrected_lost = decoder._parse_corrected(CHILD / 'child-02.corrected.perf-script-no-inline.txt', expected_pid=PID, expected_cpu=CPU, expected_period=PERIOD, marker=MARKER, marker_dso=TARGET)
print(json.dumps({'stage':'corrected_done','samples':len(corrected),'lost':corrected_lost,'marker':sum(any(f.get('symbol_base') == MARKER and f.get('dso') == TARGET for f in row.get('frames',[])) for row in corrected)}), flush=True)
records = lifecycle['records']
comms = [e for e in records if e.get('kind') == 'PERF_RECORD_COMM' and e.get('pid') == PID and e.get('tid') == PID and e.get('exec') is True and str(e.get('comm')) == BINARY.name[:15]]
if len(comms) != 1:
    raise RuntimeError(f'target COMM_EXEC count={len(comms)}')
comm_seq = int(comms[0]['record_sequence'])
all_maps = [e for e in records if e.get('kind') == 'PERF_RECORD_MMAP2' and e.get('pid') == PID and e.get('tid') == PID]

def attribution(row, desc):
    seq = int(desc['record_sequence'])
    cutoff = comm_seq if seq >= comm_seq else None
    maps = all_maps
    attr = None
    frames = row.get('frames') or []
    if frames:
        leaf = frames[0]
        if leaf.get('resolved') is True and isinstance(leaf.get('dso'), str) and type(leaf.get('ip')) is int:
            active = life._interval_owner(maps, int(leaf['ip']), record_sequence=seq, comm_sequence=cutoff)
            if isinstance(active, dict) and active.get('path') == leaf.get('dso') and re.fullmatch(r'[0-9a-fA-F]{8,64}', str(active.get('build_id') or '')):
                attr = {
                    'physical_leaf': {
                        'ip': int(leaf['ip']), 'dso_path': str(leaf['dso']),
                        'build_id': str(active['build_id']).lower(),
                        'symbol': str(leaf.get('symbol') or ''),
                        'map_record_sequence': int(active['record_sequence']),
                    },
                    'target_source_eligible': str(leaf['dso']) == TARGET and str(active['build_id']).lower() == BUILD,
                }
            else:
                attr = {'physical_unknown': True, 'target_source_eligible': False}
    marker_frames = [f for f in frames if f.get('symbol_base') == MARKER and f.get('dso') == TARGET]
    if marker_frames:
        mf = marker_frames[0]
        if type(mf.get('ip')) is not int:
            raise RuntimeError(f'marker frame has no IP at sample {row.get("index")}')
        active_marker = life._interval_owner(maps, int(mf['ip']), record_sequence=seq, comm_sequence=cutoff)
        if not isinstance(active_marker, dict) or active_marker.get('path') != TARGET or str(active_marker.get('build_id') or '').lower() != BUILD:
            raise RuntimeError(f'marker active map mismatch at sample {row.get("index")}: {active_marker}')
        if attr is None:
            attr = {'physical_unknown': True, 'target_source_eligible': False}
        attr['marker_frame'] = {
            'present': True, 'resolved': bool(mf.get('resolved')),
            'dso_path': TARGET, 'build_id': BUILD, 'symbol': MARKER,
            'ip': int(mf['ip']), 'map_record_sequence': int(active_marker['record_sequence']),
        }
    return attr

attrs = {}
for row, desc in zip(corrected, lifecycle['sample_records']):
    a = attribution(row, desc)
    if a is not None:
        attrs[int(row['index'])] = a
print(json.dumps({'stage':'attribution_done','rows':len(attrs),'marker_attrs':sum('marker_frame' in a for a in attrs.values())}), flush=True)
terminal_text = (CHILD / 'child-02.terminal-report.txt').read_text()
losses = re.findall(r'Total Lost Samples:\s*(\d+)', terminal_text)
if len(losses) != 1:
    raise RuntimeError(f'terminal loss lines={losses}')
packet = life.build_ledger(
    lifecycle, raw['sample_ledger'], expected_pid=PID, expected_cpu=CPU,
    expected_period=PERIOD, target_path=TARGET, target_build_id=BUILD,
    target_comm=BINARY.name[:15], terminal_lost_samples=int(losses[0]),
    marker=MARKER, marker_dso=TARGET, attributions=attrs,
)
summary = {
    'schema':'litchi.performance.0978.child02_build_ledger_integration.v1',
    'status':'complete', 'child':'child-02', 'pid':PID,
    'inputs': {
      'dump': life.file_binding(DUMP), 'perf_data': life.file_binding(DATA),
      'corrected_script': life.file_binding(CHILD / 'child-02.corrected.perf-script-no-inline.txt'),
      'lifecycle_helper': life.file_binding(TOOLS / 'lifecycle_ledger_0978.py'),
    },
    'lifecycle': {k:lifecycle.get(k) for k in ('record_count','raw_event_count','raw_metadata_count','sample_descriptor_count','raw_lost_sum','unique_record_offsets','raw_metadata_crosscheck')},
    'raw_scanner': {k:raw.get(k) for k in ('status','sample_count','raw_lost_sum','sample_conservation','period_valid','pid_tid_cpu_valid','time_valid','bounds_valid','violations')},
    'corrected': {'sample_count':len(corrected),'marker_sample_count':sum(any(f.get('symbol_base') == MARKER and f.get('dso') == TARGET for f in row.get('frames',[])) for row in corrected)},
    'attribution': {'rows':len(attrs),'marker_rows':sum('marker_frame' in a for a in attrs.values())},
    'ledger': packet,
}
out = ROOT / 'analysis' / 'child-02-build-ledger-integration-v3.json'
out.parent.mkdir(parents=True, exist_ok=True)
if out.exists():
    raise RuntimeError(f'refusing to overwrite retained integration output: {out}')
out.write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n')
print(json.dumps({'stage':'ledger_done','output':str(out),'bytes':out.stat().st_size,'D':packet['ledger']['D'],'M':packet['ledger']['M'],'R':packet['ledger']['R'],'T':packet['ledger']['T'],'U_target':packet['ledger']['U_target'],'U_marker':packet['ledger']['U_marker'],'marker_count':packet['ledger']['marker_sample_count']}), flush=True)
