# change-0572 evidence packet: OOXML source-backed range attribution

Change record:
[`docs/performance/0572-ooxml-range-source-attribution.md`](../../0572-ooxml-range-source-attribution.md).
Disposition: attribution only, no production change. `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | The frozen plan, `status: frozen-before-capture`. **Not edited after capture.** |
| `probe/` | The throwaway probe: a counting `litchi_core::ReadAt` with an optional simulated transport (`src/counting.rs`), the arm matrix and repeat loop (`src/harness.rs`), and the three scenario bodies (`src/main.rs`). Nothing under `crates/` was touched. |
| `classify_zip_requests.py` | Parses each fixture's ZIP central directory and local records independently of the library and assigns every recorded byte to exactly one region; locates the strict-layout proof sweep. |
| `summarize.py` | Produces the tables the change record cites and evaluates the determinism, attribution and control gates. |
| `assemble_results.py` | Joins capture, attribution, summary and host environment into `results.json`. |
| `results.json` | The retained machine-readable result: environment, corpus manifest with SHA-256, parsed ZIP layouts, gate verdicts, headline counts and timings, and per-arm per-region totals. |
| `timing-repeatability.json` | Two independent captures of the whole matrix compared arm for arm. |
| `xlsx-cell-survey.json` | Every `.xlsx` fixture under `test-data/ooxml/xlsx` driven through the open-then-`cell("A1")` scenario, to size how often that scenario refuses. |
| `summary.txt` | The human-readable tables, as printed by `summarize.py`. |

## Headline

Eleven fixtures, three scenarios, three policies, two construction routes, three
transports, five repeats: **132 arms, 660 scenario runs.**

```
scenario                    fixture                        mem  desc  exact  open  proof  fs4k  fs64k
docx_open_full_text         comment.docx                    10    10     15    12      -     4      4
docx_open_full_text         endnotes.docx                   16     0     13    11      -     5      3
docx_open_full_text         testComment.docx                17     0     13    11      -     5      3
pptx_open_middle_slide_text shape-glow-effect.pptx          17     0     19    17      -     5      3
pptx_open_middle_slide_text shapes.pptx                     48     0     49    47      -    14      3
pptx_open_middle_slide_text shape-soft-edges.pptx           65    65     87    84      -    30      4
xlsx_open_cell_a1           sheet-names.xlsx                13     0     40    11     26     6      3
xlsx_open_cell_a1           universal-content.xlsx          13    13     58    18     39     8      4
xlsx_open_cell_a1           ConditionalFormattingSamples    132    0    354    89    264    77     49
xlsx_open_cell_a1           SimpleNormal.xlsx               12     0     40    11     24     4      3
xlsx_open_cell_a1           ExcelPivotTableSample.xlsx      27     0     86    25     54    15      3

determinism_gate  pass   44 arms, 0 divergent across 3 transports and 5 repeats
attribution_gate  pass   0 requests in a gap, 0 bytes past end of file (see caveat)
control           pass   11/11 native-leaf and adopt-route sequences byte-identical
timing_gate       pass   medians over 5 repeats, load average 7.88 at capture
policy_gate       pass   4 policy-accepting entry points, read from source
```

## Seven corrections were made in review, before publication

An independent audit recomputed every table cell from `results.json` and found
the arithmetic clean but **six prose conclusions wrong**, plus one packet
defect. All are listed in `decision.json` under `corrections_made_in_review`.
The one that mattered: the draft asserted that *no styles payload is ever read*
and scored a plan prediction WRONG on that basis — contradicted by this packet's
own retained bytes, where `ExcelPivotTableSample.xlsx` reads `xl/styles.xml` at
requests 84 and 85. The prediction is now scored confirmed, and the error is
recorded in the change record rather than quietly removed.

The packet defect: `results.json` originally carried verdicts for three of the
five gates, so the snippet below printed three where the record asserts five.
`timing_gate` and `policy_gate` are now recorded there too.

## The build was pinned, deliberately

The first capture was discarded. Another agent held an 879-line uncommitted
change to `crates/soapberry-zip/src/archive.rs` and the probe linked it,
reporting the XLSX proof as one read per member instead of two. `results.json`
records `library_source` as a `git archive` of
`163ac1bd67f2a0d72c27bfec60e0a2620768cc9d` extracted to a scratch tree, which is
how the reported numbers are isolated from the working tree.

Anyone replaying this **must do the same** if `crates/` is dirty. Otherwise the
probe measures whatever is in flight.

## Replay

```sh
rev=163ac1bd67f2a0d72c27bfec60e0a2620768cc9d
work=$(mktemp -d)
git -C . archive "$rev" crates Cargo.toml rust-toolchain.toml | tar -x -C "$work"
cp -r docs/performance/results/change-0572/probe "$work/probe"
sed -i "s|\.\./\.\./\.\./\.\./\.\./crates|$work/crates|g" "$work/probe/Cargo.toml"

CARGO_TARGET_DIR="$work/target" cargo +1.95.0 build --release -j 8 \
  --manifest-path "$work/probe/Cargo.toml"
"$work/target/release/litchi-0572-probe" . "$work/capture.json" "$rev"

python3 docs/performance/results/change-0572/classify_zip_requests.py \
  --capture "$work/capture.json" --out "$work/attribution.json" --repo . \
  --raw-sequence-for '*:exact'
python3 docs/performance/results/change-0572/summarize.py \
  --attribution "$work/attribution.json" --capture "$work/capture.json" \
  --json-out "$work/summary.json"
python3 docs/performance/results/change-0572/assemble_results.py \
  --capture "$work/capture.json" --attribution "$work/attribution.json" \
  --summary "$work/summary.json" --repo . --out "$work/results.json"
```

**This was verified, not assumed.** The retained `probe/` source was copied out,
repointed by the `sed` above and rebuilt, and its capture reproduces the ordered
`(offset, length)` sequence of **all 132 arms** with zero mismatches. The
delayed-transport medians should reproduce to a couple of percent on an
otherwise comparable host.

Check the gate verdicts and the corpus hashes without rebuilding anything:

```sh
python3 -B -c "
import json
d = json.load(open('docs/performance/results/change-0572/results.json'))
for name, gate in d['gates'].items():
    print(name, gate['verdict'])
for f in d['corpus']:
    print(f['sha256'][:16], f['size'], f['path'])
print('arms', len(d['results']), 'raw requests retained', d['retained_raw_requests'])"
```

## What is not here

No filesystem, cold-cache, physical-device or cross-platform capture. No
allocation, peak-RSS, instruction-count or syscall measurement. No ABBA
comparison, because this record proposes no candidate. The counting source holds
each fixture in memory, so nothing here measures the page cache, and no
warm-local latency is claimed.

The intermediate `capture.json` and `attribution.json` are **not** retained: both
are regenerable from the probe and the scripts above, and `results.json` keeps
everything the change record cites, including the full ordered request sequence
for the `exact` policy on every fixture.
