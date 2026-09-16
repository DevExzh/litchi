"""Run-length-encodes a change-under-read sweep's outcome classes and compares
the two legs, so a shifted ordinal is visible as a band length change rather
than as noise on every row."""
import re, sys, os, glob

def klass(text):
    text = re.sub(r'ArtifactFingerprint\(\[[^\]]*(\]\))?', 'FP', text)
    text = re.sub(r'\{[^}]*$', '{..}', text)
    text = re.sub(r'expected: FP, observed: FP \}', '{..}', text)
    return text.strip()

def load(path):
    rows, head = [], {}
    for line in open(path):
        if line.startswith('#'):
            k, _, v = line[2:].strip().partition('=')
            head[k] = v
            continue
        if line.startswith('trigger'):
            continue
        t, fired, reads, outcome = line.rstrip('\n').split('\t', 3)
        rows.append((int(t), klass(outcome)))
    return head, rows

def rle(rows):
    out = []
    for t, c in rows:
        if out and out[-1][2] == c:
            out[-1][1] = t
        else:
            out.append([t, t, c])
    return out

def fmt(bands):
    return '\n'.join(f"  {a}-{b}\t{c}" if a != b else f"  {a}\t{c}" for a, b, c in bands)

def refused(rows):
    return sum(1 for _, c in rows if c != 'OK')

def trailing_ok(rows):
    n = 0
    for _, c in reversed(rows):
        if c != 'OK':
            break
        n += 1
    return n

def leading_ok(rows):
    n = 0
    for _, c in rows:
        if c != 'OK':
            break
        n += 1
    return n

def interior_ok(rows):
    return [t for t, c in rows[leading_ok(rows):len(rows)-trailing_ok(rows)] if c == 'OK']

def main():
    bad = 0
    for before in sorted(glob.glob(os.path.join(sys.argv[1], '*.tsv'))):
        after = os.path.join(sys.argv[2], os.path.basename(before))
        hb, rb = load(before)
        ha, ra = load(after)
        bb, ba = rle(rb), rle(ra)
        same = [c for _, _, c in bb] == [c for _, _, c in ba] and \
               [b - a for a, b, _ in bb] == [b - a for a, b, _ in ba]
        flag = 'SAME' if same else 'DIFF'
        if interior_ok(ra):
            flag += ' INTERIOR-OK!'
            bad += 1
        if trailing_ok(ra) != 1:
            flag += f' TRAILING={trailing_ok(ra)}'
        if leading_ok(ra) != leading_ok(rb):
            flag += f' LEADING {leading_ok(rb)}->{leading_ok(ra)}'
            bad += 1
        print(f"{os.path.basename(before)}\t{flag}\tordinals {len(rb)}->{len(ra)}\trefused {refused(rb)}->{refused(ra)}")
        if not same and '--v' in sys.argv:
            print(" BEFORE:"); print(fmt(bb))
            print(" AFTER:");  print(fmt(ba))
    print(f"# problems={bad}")


if __name__ == "__main__":
    main()
