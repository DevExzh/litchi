import sys

def load(path):
    seq = []
    region = None
    for line in open(path):
        if not line.startswith('SSTTRACE'):
            continue
        _, rs, re_, off, span = line.rstrip('\n').split('\t')
        region = (int(rs), int(re_))
        seq.append((int(off), int(span)))
    return region, seq

def sim(region, seq, ceiling, first_fill, grow=True, region_if_fits=True):
    rs, re_ = region
    region_len = re_ - rs
    win_start, win_len = 0, 0
    nxt = first_fill
    reads = 0
    read_bytes = 0
    for off, span in seq:
        if span > ceiling or off < rs or off + span > re_:
            reads += 1
            read_bytes += span
            continue
        if win_len and win_start <= off and off + span <= win_start + win_len:
            continue
        stride = max(nxt if grow else first_fill, span)
        stride = min(stride, ceiling)
        if region_if_fits and region_len <= stride:
            fs, fl = rs, region_len
        else:
            fs, fl = off, min(stride, re_ - off)
        win_start, win_len = fs, fl
        reads += 1
        read_bytes += fl
        if grow:
            nxt = min(nxt * 2, ceiling)
    return reads, read_bytes

def baseline(seq):
    return len(seq), sum(s for _, s in seq)

for name in sys.argv[1:]:
    region, seq = load(name)
    if not seq:
        print(name, 'empty'); continue
    rl = region[1] - region[0]
    b = baseline(seq)
    print(f"\n=== {name.split('/')[-1]}  region={rl}  resolves={len(seq)}  baseline reads={b[0]} bytes={b[1]}")
    for ceiling in (4096, 8192, 16384, 32768, 65536, 131072, 262144, 524288):
        r = sim(region, seq, ceiling, 512)
        rx = sim(region, seq, ceiling, ceiling, grow=False)
        print(f"  ceiling={ceiling:>7}  grow512: reads={r[0]:>6} bytes={r[1]:>9}   flat: reads={rx[0]:>6} bytes={rx[1]:>9}")
    # exact-span window only (no read-ahead)
    e = sim(region, seq, 65536, 1, grow=False, region_if_fits=False)
    print(f"  exact-span window       : reads={e[0]:>6} bytes={e[1]:>9}")
    # region-only (enabled only when region fits ceiling)
    for ceiling in (65536, 262144, 524288):
        if rl <= ceiling:
            print(f"  region-only c={ceiling}: reads=1 bytes={rl}")
        else:
            print(f"  region-only c={ceiling}: reads={b[0]} bytes={b[1]} (disabled)")
