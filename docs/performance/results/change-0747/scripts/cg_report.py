#!/usr/bin/env python3
"""Per-measured-iteration callgrind attribution from isolation pairs.

For each arm and shape, (Ir(samples=3) - Ir(samples=1)) / 2 is the cost of one
measured iteration, which cancels corpus construction, lifecycle gates and the
eager expected-output oracle. Inclusive Ir of a function is the sum of the
inclusive cost of every call into it."""
import json, sys
sys.path.insert(0, '/home/zhuhe/code/litchi-worktrees/scratch/0747/cg')
from cg_parse import parse

FUNCTIONS = {
    'publication': 'litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream',
    'topology_writer': 'litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream',
    'preservation_writer': 'litchi_opc::source_backed::SourceBackedPackage::write_changed_overlays_with_appended_inner',
    'planning': 'litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets',
    'commit': 'litchi_xlsx::cell_values::source::MultiSourceEdit::commit',
    'audit_slice_before': 'xml_minifier::audit::verify_with_policy',
    'audit_pair_after': 'xml_minifier::audit::verify_source_replacement',
}

def per_iteration(prefix, arm, shape):
    one = parse(f'{prefix}-{arm}-{shape}-s1.out')
    three = parse(f'{prefix}-{arm}-{shape}-s3.out')
    result = {'program_total': (three[5] - one[5]) / 2}
    for key, name in FUNCTIONS.items():
        incl = [(n, (three[1][n] - one[1][n]) / 2, (three[2][n] - one[2][n]) / 2)
                for n in set(one[1]) | set(three[1]) if n == name]
        if incl:
            result[key] = {'inclusive_ir': incl[0][1], 'calls': incl[0][2]}
    return result

if __name__ == '__main__':
    prefix, arms, shapes = sys.argv[1], sys.argv[2].split(','), sys.argv[3].split(',')
    out = {f'{arm}/{shape}': per_iteration(prefix, arm, shape) for arm in arms for shape in shapes}
    json.dump(out, open(sys.argv[4], 'w'), indent=2, sort_keys=True)
    for key, value in out.items():
        print(key, json.dumps(value))
