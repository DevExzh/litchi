# change-0570 evidence packet: contiguous CFB FAT run batching

Change record:
[`docs/performance/0570-cfb-fat-run-batching.md`](../../0570-cfb-fat-run-batching.md).
Disposition: retained. `performance_claim: none`; **no timing is measured and
none is claimed.**

## Contents

| Path | What it is |
| --- | --- |
| `corpus.before.txt`, `corpus.after.txt` | Reads and bytes for one `OleFile::open` of every container fixture under `test-data/ole`, counted through the crate's existing instrumented reader, before and after the change. |
| `make_corpus_probe.py` | The throwaway instrumentation that produced them. |

## Result

```
files 98   reads 493 -> 463 (saved 30)   bytes 276,992 -> 276,992 (identical)
  doc/picture.doc           26 -> 6   saved 20
  doc/testPictures.doc       8 -> 4   saved 4
  doc/FloatingPictures.doc  12 -> 9   saved 3
  xls/WithCustomViews.xls    5 -> 3   saved 2
  ppt/SampleShow.ppt         7 -> 6   saved 1
unchanged files: 93
```

Bytes read are byte-for-byte identical: the change removes calls, not work. The
five files that improve are exactly the five the corpus survey retained with
change 0565 predicted, each by exactly the predicted amount, which is a useful
check on that survey's model as well as on this change.

## Replay

Recompute the totals from the retained files:

```sh
python3 -B -c "
import re, pathlib
def load(p):
    return {m.group(1): (int(m.group(2)), int(m.group(3)))
            for m in (re.match(r'CORPUS (.+) reads=(\d+) bytes=(\d+)$', line)
                      for line in pathlib.Path(p).read_text().splitlines()) if m}
b = load('docs/performance/results/change-0570/corpus.before.txt')
a = load('docs/performance/results/change-0570/corpus.after.txt')
k = sorted(set(b) & set(a))
print(len(k), sum(b[x][0] for x in k), '->', sum(a[x][0] for x in k),
      'bytes identical:', sum(b[x][1] for x in k) == sum(a[x][1] for x in k))
print([(x, b[x][0], a[x][0]) for x in k if b[x][0] != a[x][0]])"
```

Note that four fixture names contain a space, so a parser that splits on
whitespace silently drops them and reports 85 files and 428 reads. The pattern
above anchors on the trailing fields instead.

## What is not here

No timing, allocation-profile, cold-cache, physical-device or cross-platform
capture. Thirty fewer positional reads across a 98-file corpus, twenty of them in
one file, is not separable from timing noise on this host, which is why no timing
was attempted rather than reported inconclusively.
