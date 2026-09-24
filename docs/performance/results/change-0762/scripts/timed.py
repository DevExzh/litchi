#!/usr/bin/env python3
"""timed.py FOLDED ANCHOR_REGEX [TOP]: attribute the samples whose stack
contains a frame matching ANCHOR_REGEX (the timed call) by leaf symbol and by
inclusive component. FOLDED is fold.py output (period, count, comm, frames
root->leaf joined by ' ;; ')."""
import re, sys
from collections import Counter
folded, anchor = sys.argv[1], re.compile(sys.argv[2])
top = int(sys.argv[3]) if len(sys.argv) > 3 else 30
COMPONENTS = [
    ('deflate (zlib-rs)', r'zlib_rs::|deflate_medium|deflate_quick|deflate_fast|longest_match|compress_block|build_tree|flush_block'),
    ('crc32', r'crc32|CrcStage'),
    ('budget consume', r'ExecutionContext::consume|Budget>::consume|budget::charge_chain|charge_chain'),
    ('budget reserve/commit', r'reserve_scoped|ExecutionContext::reserve|Reservation.*commit|ScopedReservation|release_chain'),
    ('escaping', r'for_each_escaped_chunk|emit_escaped_text|push_escaped|write_escaped'),
    ('text scan', r'scan_plain_text|encoded_text_size'),
    ('owned stage copy', r'write_owned'),
    ('memcpy/memmove', r'memmove|memcpy'),
    ('memset', r'memset'),
    ('sink hash (sha256)', r'sha2::|sha256'),
    ('text generation (format!)', r'core::fmt::|alloc::fmt::format'),
]
def name(frame):
    return frame.split('|R:')[0]
total = kept = 0
leaf = Counter(); incl = Counter()
for line in open(folded, errors='replace'):
    period, count, comm, stack = line.rstrip('\n').split('\t')
    period = int(period); total += period
    frames = stack.split(' ;; ')
    idx = next((i for i, f in enumerate(frames) if anchor.search(name(f))), None)
    if idx is None:
        continue
    kept += period
    below = [name(f) for f in frames[idx:]]
    leaf[re.sub(r'<.*', '<..>', below[-1])[:110]] += period
    joined = ' ;; '.join(below)
    for label, rx in COMPONENTS:
        if re.search(rx, joined):
            incl[label] += period
print(f'samples(period) total={total} anchored={kept} ({100*kept/max(total,1):.1f}%) anchor={anchor.pattern}')
print('== inclusive components, % of anchored ==')
for label, _ in COMPONENTS:
    print(f'{100*incl[label]/max(kept,1):6.2f}%  {label}')
print(f'== top {top} leaf symbols, % of anchored ==')
for sym, p in leaf.most_common(top):
    print(f'{100*p/max(kept,1):6.2f}%  {sym}')
