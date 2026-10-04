"""GDB-side owner-gated native SHA backend reachability; no timing inference."""
import gdb
import json
import os
from pathlib import Path

OUTPUT = Path(os.environ['LITCHI_0912_DISPATCH_LOG'])
SYMBOLS = json.loads(Path(os.environ['LITCHI_0912_SYMBOL_RANGES']).read_text())
RANGES = SYMBOLS['ranges']
OWNER = 'litchi_pptx::package::model::Package::opened_presentation_transaction'
PARENT = 'litchi_perf_baseline::pptx_owned_benchmark::timed_iteration'
SETUP = 'litchi_perf_baseline::build_pptx_source_edit_corpus'
BACKENDS = {'hardware': 'sha2::sha256::x86_sha::compress', 'software': 'sha2::sha256::soft::unroll::compress'}
LOG = {'status': 'running', 'events': [], 'backend_hits': [], 'owner_entries': [], 'owner_returns': [], 'errors': []}
ACTIVE = None
RUNTIME = {}
RETURNS = []
BPS = {}


def flush():
    OUTPUT.write_text(json.dumps(LOG, indent=2, sort_keys=True) + '\n')


def register(name):
    return int(gdb.parse_and_eval('$' + name))


def stack():
    frames = []
    frame = gdb.newest_frame()
    try:
        while frame is not None and len(frames) < 128:
            frames.append({'name': frame.name(), 'pc': int(frame.pc())})
            frame = frame.older()
    except gdb.error as error:
        return frames, str(error)
    return frames, None


def fail(error):
    LOG['errors'].append(str(error))
    LOG['status'] = 'fail'
    flush()
    return True


def inside(name, pc):
    return RUNTIME[name]['start'] <= pc < RUNTIME[name]['end']


class Backend(gdb.Breakpoint):
    def __init__(self, kind, name):
        super().__init__('*0x%x' % RUNTIME[name]['start'], internal=True)
        self.kind = kind
        self.name = name
        self.enabled = False

    def stop(self):
        try:
            assert ACTIVE is not None and ACTIVE['context'] == 'measured'
            assert gdb.selected_thread().global_num == ACTIVE['thread']
            pc = register('pc')
            assert inside(self.name, pc)
            frames, error = stack()
            names = [f['name'] for f in frames]
            assert SETUP not in names
            owner_positions = [i for i, n in enumerate(names) if n == OWNER]
            parent_positions = [i for i, n in enumerate(names) if n == PARENT]
            frame_scope = len(owner_positions) == len(parent_positions) == 1
            if frame_scope:
                assert owner_positions[0] < parent_positions[0]
                assert inside(OWNER, frames[owner_positions[0]]['pc'])
                assert inside(PARENT, frames[parent_positions[0]]['pc'])
            LOG['backend_hits'].append({'frame_scope_available': frame_scope, 'scope_proof': 'Exact owner-entry caller return address, same-thread gated breakpoint state, and exact matched owner return', 'kind': self.kind, 'symbol': self.name, 'pc': pc, 'owner_entry': ACTIVE['index'], 'frames': frames, 'unwind_error': error, 'backtrace': gdb.execute('bt', to_string=True)})
            self.enabled = False
            LOG['events'].append({'kind': 'backend', 'backend': self.kind, 'owner_entry': ACTIVE['index'], 'breakpoint_disabled_after_hit': True})
            flush()
            return False
        except Exception as error:
            return fail(repr(error))


class OwnerReturn(gdb.Breakpoint):
    def __init__(self, owner):
        super().__init__('*0x%x' % owner['return_address'], internal=True)
        self.owner = owner

    def stop(self):
        global ACTIVE
        try:
            assert ACTIVE is self.owner
            assert gdb.selected_thread().global_num == ACTIVE['thread']
            assert register('pc') == ACTIVE['return_address']
            assert register('rsp') == ACTIVE['entry_rsp'] + 8
            for bp in BPS.values():
                bp.enabled = False
            LOG['owner_returns'].append({'owner_entry': ACTIVE['index'], 'context': ACTIVE['context'], 'pc': register('pc'), 'rsp': register('rsp'), 'backend_breakpoints_disabled': all(not bp.enabled for bp in BPS.values())})
            LOG['events'].append({'kind': 'owner_return', 'owner_entry': ACTIVE['index'], 'context': ACTIVE['context'], 'backend_breakpoints_disabled': True})
            ACTIVE = None
            self.enabled = False
            flush()
            return False
        except Exception as error:
            return fail(repr(error))


class OwnerEntry(gdb.Breakpoint):
    def __init__(self):
        super().__init__('*0x%x' % RUNTIME[OWNER]['start'], internal=True)

    def stop(self):
        global ACTIVE
        try:
            assert ACTIVE is None and register('pc') == RUNTIME[OWNER]['start']
            rsp = register('rsp')
            address = int.from_bytes(bytes(gdb.selected_inferior().read_memory(rsp, 8)), 'little')
            contexts = [('measured', PARENT), ('setup', SETUP)]
            matches = [context for context, name in contexts if inside(name, address)]
            assert len(matches) == 1, hex(address)
            frames, error = stack()
            ACTIVE = {'index': len(LOG['owner_entries']), 'context': matches[0], 'entry_pc': register('pc'), 'entry_rsp': rsp, 'return_address': address, 'thread': gdb.selected_thread().global_num, 'frames': frames, 'unwind_error': error}
            LOG['owner_entries'].append(ACTIVE)
            RETURNS.append(OwnerReturn(ACTIVE))
            for kind, bp in BPS.items():
                bp.enabled = ACTIVE['context'] == 'measured' and not any(hit['kind'] == kind for hit in LOG['backend_hits'])
            LOG['events'].append({'kind': 'owner_entry', 'owner_entry': ACTIVE['index'], 'context': ACTIVE['context'], 'backend_breakpoints_enabled': {kind: bp.enabled for kind, bp in BPS.items()}})
            flush()
            return False
        except Exception as error:
            return fail(repr(error))


def exited(event):
    LOG['exit_code'] = getattr(event, 'exit_code', None)


gdb.events.exited.connect(exited)
try:
    gdb.execute('set pagination off')
    gdb.execute('set confirm off')
    gdb.execute('set disable-randomization off')
    gdb.execute('set language c')
    LOG['startup'] = gdb.execute('starti', to_string=True)
    LOG['mappings'] = gdb.execute('info proc mappings', to_string=True)
    for row in RANGES:
        start = int(gdb.parse_and_eval("(void*) &'" + row['name'] + "'"))
        RUNTIME[row['name']] = {'start': start, 'end': start + row['size'], 'file_address': row['address'], 'size': row['size'], 'load_bias': start - row['address']}
    biases = {row['load_bias'] for row in RUNTIME.values()}
    assert len(biases) == 1
    zero_maps = [line.split() for line in LOG['mappings'].splitlines() if line.strip().endswith(SYMBOLS['binary']['path']) and int(line.split()[3], 16) == 0]
    assert len(zero_maps) == 1 and int(zero_maps[0][0], 16) in biases
    LOG['mapping_load_bias'] = int(zero_maps[0][0], 16)
    LOG['runtime_symbols'] = RUNTIME
    BPS.update({kind: Backend(kind, name) for kind, name in BACKENDS.items()})
    owner_bp = OwnerEntry()
    flush()
    gdb.execute('continue')
    assert not LOG['errors'], LOG['errors']
    assert LOG.get('exit_code') == 0
    assert ACTIVE is None
    assert [r['context'] for r in LOG['owner_entries']] == ['setup', 'measured']
    assert [r['context'] for r in LOG['owner_returns']] == ['setup', 'measured']
    assert len(LOG['backend_hits']) == 1, LOG['backend_hits']
    assert all(not bp.enabled for bp in BPS.values())
    LOG['status'] = 'pass'
    LOG['interpretation'] = 'One exact backend reached under the measured owner; breakpoint sample is not timing and absence of another hit is not a general dispatch claim.'
    flush()
except Exception as error:
    fail(repr(error))
    gdb.execute('quit 1')
