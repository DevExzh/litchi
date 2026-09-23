#!/usr/bin/env python3
"""Timed-region attribution of folded frame-pointer profiles (change 0753).

Keeps samples whose stack contains the harness's timed call
(run_fresh_writer > write_fresh_{doc,ppt,xls}), weights them by period, and
reports the share of kernel leaves (page faults and other kernel work), the
top leaf frames, and the inclusive share of named components. Stacks were
folded by profile-r2's fold.py (kernel frames collapsed to K).

usage: profile_summary.py FOLDED_DIR OUT_JSON
"""
import json, os, re, sys
from collections import Counter

GLUE = re.compile(r'^(branch<|map_err<|and_then<|map<|into<|from$|from<|unwrap_or_else<|ok_or_else<|call_once|\{closure|try_fold|try_for_each|for_each<|fold<)')
COMPONENTS = {
    'utf16_count_or_decode': r'^(count<|next_code_point|utf16_code_unit_len$|utf16_units$|encode_utf16|next$)',
    'text_widening_or_append': r'^(push_utf16le$|append_utf16le$|extend_from_slice<|append_elements<|needs_to_grow)',
    'field_character_scan': r'^contains_field_character$',
    'siphash': r'^(hash_one<|write_str|c_rounds)',
    'memmove': r'^__memmove',
    'memset': r'^__memset',
    'memcmp': r'^__memcmp',
    'realloc_growth': r'^(__libc_realloc|grow_amortized<|finish_grow<)',
    'cfb_write_to': r'^write_to<.*OleWriter|^write_sector_aligned',
    'harness_text_generation': r'^writer_payload_text$',
}


def name(frame):
    return frame.split('|R:')[0]


def summarize(path, runner=r'^run_fresh_writer', child=r'^write_fresh_'):
    runner_rx, child_rx = re.compile(runner), re.compile(child)
    rows, total = [], 0
    for line in open(path):
        period, count, comm, stack = line.rstrip('\n').split('\t')
        if not comm.startswith('lpb'):
            continue
        period = int(period)
        total += period
        frames = stack.split(' ;; ')
        index = next((i for i, f in enumerate(frames) if runner_rx.search(name(f))), None)
        if index is None:
            continue
        j = index + 1
        while j < len(frames) and GLUE.search(name(frames[j])):
            j += 1
        if j >= len(frames) or not child_rx.search(name(frames[j])):
            continue
        rows.append((period, frames[j:]))
    timed = sum(p for p, _ in rows)
    kernel = sum(p for p, f in rows if f[-1] == 'K')
    leaves = Counter()
    components = Counter()
    for period, frames in rows:
        k = 0
        while k < len(frames) and frames[-1 - k] == 'K':
            k += 1
        user = [re.sub(r'<.*', '<..>', name(f))[:60] for f in frames[:len(frames) - k]]
        leaves[('[K] ' if k else '') + ' < '.join(reversed(user[-2:]))] += period
        seen = set()
        for frame in frames:
            for component, rx in COMPONENTS.items():
                if component not in seen and re.search(rx, name(frame)):
                    seen.add(component)
                    components[component] += period
    return {
        'file': os.path.basename(path),
        'timed_share_of_process_cycles': round(timed / total, 4) if total else None,
        'kernel_leaf_share_of_timed': round(kernel / timed, 4) if timed else None,
        'top_leaves': [[key, round(v / timed, 4)] for key, v in leaves.most_common(12)],
        'component_inclusive_share': {k: round(components[k] / timed, 4) for k in COMPONENTS},
    }


def main():
    folder, out = sys.argv[1], sys.argv[2]
    result = {}
    for kind in ('doc', 'ppt', 'xls'):
        for leg in 'AB':
            path = os.path.join(folder, f'{kind}-{leg}.folded')
            result[f'{kind}_fresh_write_to/payload-heavy/{leg}'] = summarize(path)
    with open(out, 'w') as handle:
        json.dump(result, handle, indent=1)
    for key, value in result.items():
        comps = ' '.join(f"{k}={v:.3f}" for k, v in value['component_inclusive_share'].items() if v >= 0.01)
        print(f"{key}: kernel={value['kernel_leaf_share_of_timed']:.3f} {comps}")


if __name__ == '__main__':
    main()
