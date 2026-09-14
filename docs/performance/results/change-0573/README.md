# change-0573 evidence packet: one read per strict ZIP local header

Change record:
[`docs/performance/0573-zip-single-local-header-read.md`](../../0573-zip-single-local-header-read.md).
Disposition: retained. `performance_claim: none`; **no timing is measured and
none is claimed.**

## Contents

| Path | What it is |
| --- | --- |
| `corpus.before.txt`, `corpus.after.txt` | Reads and bytes for one whole-archive strict layout proof over every DOCX/PPTX/XLSX fixture under `test-data/ooxml`, counted through a positional reader that records every `read_at`, before and after the change. |
| `make_proof_probe.py` | The throwaway instrumentation that produced them. It inserts a counting `ReaderAt` and a corpus walk into `crates/soapberry-zip/src/office.rs`'s test module, because `IndexedArchive::build_strict_layout_proof` is private and is exactly the whole-archive proof being measured. |
| `survey_local_variable_regions.py` | The corpus survey that sizes `STRICT_LOCAL_HEADER_WINDOW`. |
| `local-variable-regions.ooxml.txt` | That survey over `test-data/ooxml`: 179 archives, 4,270 members. |
| `local-variable-regions.all.txt` | The same survey over all of `test-data`: 533 archives, 14,744 members. |
| `summary.txt` | The totals and per-fixture breakdown recomputed from the two corpus files. |

## Result

```
fixtures 168   members 4089   descriptor-bearing members 148
reads  8326 -> 4385   (-3941, -47.3%)
bytes  514578 -> 517371   (+2793, +0.54%)
model  descriptor-free members x1 + descriptor-bearing members x3 = 3941+444 = 4385

fixtures with byte-identical proof reads: 159
fixtures whose reads fell:                160
fixtures whose reads were unchanged:        8  (every member carries a descriptor)
fixtures whose reads rose:                  0

xlsx/ConditionalFormattingSamples.xlsx  132 members  reads 264 -> 132  bytes 10740 -> 10740
```

The measured count matches the model exactly: every descriptor-free member is
proved in one read, and every descriptor-bearing member keeps the three reads it
already cost, because the window reserves the widest possible descriptor and the
fixtures all carry a narrower one. Change 0572's plan predicted 264 requests for
`ConditionalFormattingSamples.xlsx`; that prediction is confirmed, and the
after figure is 132.

Change 0572 counted the same proof from the other side of the crate boundary,
through a counting `litchi_core::ReadAt` handed to `litchi-opc`. Its **before**
figures and these agree with no adjustment: 26 / 2,478 B for `sheet-names.xlsx`,
39 / 861 B for `universal-content.xlsx`, 264 / 10,740 B for
`ConditionalFormattingSamples.xlsx`.

## Window sizing

`local-variable-regions.ooxml.txt` is why the constant is 640 and not 512:

```
max local variable region = 539 bytes (one-read window 569)
  window=  512: 3969/4270 fit, 301 fall back
  window=  576: 4270/4270 fit, 0 fall back
  window=  640: 4270/4270 fit, 0 fall back
```

The distribution is bimodal because Microsoft Office writes a `0xa220` growth
hint whose payload is 260 or 516 bytes. Over all of `test-data`, exactly two
members — the `xl/revisions/userNames.xml` member of two revision-bearing XLSX
fixtures — carry a 2,056-byte extra field and would need a 2,112-byte window;
those are the corpus's demonstration that the heap fallback is reachable by real
files and must stay.

## Replay

Recompute the totals from the retained files:

```sh
python3 -B -c "
import re, pathlib
def load(p):
    return {m.group(1): (int(m.group(3)), int(m.group(4)))
            for m in (re.match(r'PROOF (.+) members=(\d+) reads=(\d+) bytes=(\d+) ', line)
                      for line in pathlib.Path(p).read_text().splitlines()) if m}
b = load('docs/performance/results/change-0573/corpus.before.txt')
a = load('docs/performance/results/change-0573/corpus.after.txt')
k = sorted(set(b) & set(a))
print(len(k), sum(b[x][0] for x in k), '->', sum(a[x][0] for x in k))
print(sum(b[x][1] for x in k), '->', sum(a[x][1] for x in k))
print('byte-identical fixtures:', sum(1 for x in k if b[x][1] == a[x][1]))"
```

The `PROOFTOTAL` line at the end of each file carries the same totals as
printed by the probe.

To reproduce the capture, insert the probe and run it:

```sh
python3 docs/performance/results/change-0573/make_proof_probe.py \
    crates/soapberry-zip/src/office.rs
cargo test -p soapberry-zip --lib zz_strict_layout_proof -- --nocapture --test-threads=1
git checkout -- crates/soapberry-zip/src/office.rs
```

Note that `cargo test --nocapture` prints the first line of test output on the
same line as the test name, so an extraction anchored at the start of a line
silently drops the first fixture and reports 167 instead of 168. The retained
files were extracted with `grep -oE "PROOF[A-Z]* .*"`.

## What is not here

No timing, allocation-profile, cold-cache, physical-device or cross-platform
capture. The counts are logical `ReaderAt` calls over warm fixtures. A 47%
reduction in positional reads is not separable from timing noise on a warm local
file, which is why no timing was attempted rather than reported inconclusively;
change 0572 is the record that prices these requests on a latency-bearing
source.
