# Change 0682 evidence: DOCX paragraph-index reuse

`performance_claim: none` — this packet retains scoped before/after evidence,
not registered claims. [The change record](../../0682-docx-paragraph-index-reuse.md)
describes the implementation, measured gains, resolved hot-query regression and retained
budget cost. iWork is outside this change.

The retained probe is in [probe/](probe/). The release binary was built from
the detached baseline worktree at revision
`35686023471e066d54a26737316e0d648f787f91` with an external target directory.
The baseline source tree was clean apart from the new `change-0682` evidence
directory; no crate source was edited.

The candidate was built from the final working tree after the DOCX memo and
its tests passed integration. [The after manifest](measurements/after/manifest.json)
binds it to 392 source files, the unchanged compiled probe source/lock, all raw
measurement streams, the build command and the executable hash. The temporary
executables and isolated build/checkouts are removed after integration; source,
lockfiles, raw samples and reproduction scripts remain.

Run `python3 docs/performance/results/change-0682/audit.py` from the final
revision to verify source/probe/raw hashes and regenerate
[`comparison.json`](comparison.json). It checks all 114 semantic matrix cases,
physical-cache counters and result digests across the paired windows. Summary
grouping uses input hashes so different checkout paths cannot split a case.

The 448 primary ABBA samples show lower fresh-view p50 latency in both legs:
58.45–63.44% on generated eager/source-backed cases, 77.86–79.75% on generated
managed cases, and 15.82–20.24% on the real fixture. The final longer managed
same-view control changes −1.32% to −0.04%, with no control speedup claim. These
results cover the named already-open-package loops on this host only.

Final quality logs are in [`integration/final-verified/`](integration/final-verified/):
format/check/Clippy, 1,523 DOCX tests, 104 facade tests and rustdoc all pass.
Earlier compiler failures remain in `integration/initial-check` and
`integration/final-quality`; the first successful full suite is in
`integration/quality`. Publication follow-up formatting and test-policy
failures remain under `publication-fmt` and `publication-budget`; the explicit
small-policy delayed-publication test subsequently passes. Seven
evidence/dependency checks are retained under `integration/evidence`,
`integration/boundaries` and `integration/final-evidence`.

`measurements/initial-candidate` and `integration/pre-publication-verified`
retain the first candidate before final review added operation-start cleanup
to publication paths. Its manifest paths describe the packet layout at capture;
the three moved measurement directories now live under `initial-candidate`.
`integration/pre-publication.diff` retains that production diff. The final
measurements and source binding use the top-level `after`/`abba` directories.

`measurements/second-candidate` and `integration/pre-layout-verified` retain
the complete pressure-safe implementation before the memo layout fix. Its
longer hot-query control regressed 3.02–6.26%. Retained disassembly identifies
the additional load through an inner `Arc<Index>`; `integration/pre-layout.diff`
records that source. The final memo owns the index directly, removing this
load and allocation. The final assembly slice is retained under
`measurements/after`. Neither archived candidate is the final measured source.

## Reproduction

Environment:

```text
Linux 7.0.0-1012-aws #12-Ubuntu SMP PREEMPT Tue Aug 11 2026 x86_64
AMD EPYC 9R45; 32 CPUs; one thread per core
rustc 1.95.0 (59807616 2026-04-14)
cargo 1.95.0 (f2d3ce0bd 2026-03-21)
perf 7.0.14
valgrind-3.26.0
```

Recreate the baseline checkout and copy this retained probe into it before
building. The baseline revision predates the probe itself:

```sh
git worktree add --detach /tmp/litchi-0682-before \
  35686023471e066d54a26737316e0d648f787f91
mkdir -p /tmp/litchi-0682-before/docs/performance/results/change-0682
cp -a docs/performance/results/change-0682/probe \
  /tmp/litchi-0682-before/docs/performance/results/change-0682/
```

Build the retained probe in an external target directory (substitute the
recreated checkout path for the original measurement path below):

```sh
cd /home/zhuhe/code/litchi-docx-cache-baseline
env CARGO_BUILD_JOBS=2 cargo build \
  --manifest-path docs/performance/results/change-0682/probe/Cargo.toml \
  --release --offline --locked \
  --target-dir /home/zhuhe/code/litchi-target-0682-before
```

Run the machine-readable matrix and same-binary A/A floor as follows. CPU 12
was used for the retained run; omit `CPU` for an unpinned reproduction.

```sh
BIN=/home/zhuhe/code/litchi-target-0682-before/release/probe0682 \
OUT=docs/performance/results/change-0682/measurements/before \
CPU=12 \
  ./docs/performance/results/change-0682/probe/run-before.sh

BIN=/home/zhuhe/code/litchi-target-0682-before/release/probe0682 \
OUT=docs/performance/results/change-0682/measurements/before \
CPU=12 SAMPLES=8 \
  ./docs/performance/results/change-0682/probe/run-aa.sh
```

`run-abba.sh` accepts `BEFORE_BIN` and `AFTER_BIN` and runs the same selected
cases in `A1 B1 B2 A2` order. It is retained for the candidate build lane:

```sh
BEFORE_BIN=/path/to/before/probe0682 \
AFTER_BIN=/path/to/after/probe0682 \
OUT=docs/performance/results/change-0682/measurements/abba \
CPU=12 SAMPLES=8 \
  ./docs/performance/results/change-0682/probe/run-abba.sh
```

The ABBA script resolves one `ROOT` and passes the same real-fixture path to
both binaries. The summarizer groups rows by route, mode, input SHA-256, and
repetition count, so an absolute checkout path cannot split a real-fixture
comparison.

Summaries are generated without discarding raw JSONL samples:

```sh
./docs/performance/results/change-0682/probe/summarize.py \
  docs/performance/results/change-0682/measurements/before/aa.jsonl \
  docs/performance/results/change-0682/measurements/before/aa-summary.json
```

Build the candidate from the final revision using the same `cargo build`
command and a separate `--target-dir`; run `run-before.sh` with `PHASE=after`
and the candidate binary to recreate the allocation/control matrix. For the
longer managed same-view follow-up, run `probe/run-managed-control.sh` with
`BEFORE_BIN` and `AFTER_BIN`, CPU 12 and eight samples. Use the baseline binary
for both variables to recreate `measurements/control-aa`, then the two binaries
for `measurements/control-abba`. Each invocation runs 100,000 queries on one
view; opening and eager managed index construction remain outside that loop.
Pass the resulting `control.jsonl` to `probe/summarize.py` as above.

## Corpus and measurement shape

The generated packages are deterministic and contain direct body paragraphs
with the same fixed WordprocessingML namespace and text pattern as the 0680
probe. The retained input hashes and dimensions are:

| corpus | paragraphs | visible XML | archive bytes | input SHA-256 |
| --- | ---: | ---: | ---: | --- |
| `generated-200` | 200 | 10,113 | 1,454 | `c8a34325e82a3239bb56b4fda5cb0cb132831ca368b1941ea9a1a68caa66b405` |
| `generated-10000` | 10,000 | 500,113 | 27,080 | `f8ff818837365e2004a9b057db83ae78459a99a15daac319d26fd78c987dc90f` |
| `ComplexNumberedLists.docx` | 17 | n/a | 14,458 | `297a085a7d433af2eeee7661e8db21539452cb585096484774a1e9f5f258b0b6` |

`fresh-count` creates and drops a document view for every measured iteration;
`same-count` repeats the query on one view; `first-count` measures one query on
that view. `fresh-document` isolates view construction. The selective modes
also cover `paragraph`, `paragraphs`, and source-backed `paragraph_text`.
Every counted row asserts the expected paragraph count, and every text row
requires the selected paragraph to exist. The result digest makes an accidental
elision of the query visible in the JSON output.

Package opening and the base source-cache allocation happen outside the timed
loop. `fresh-document` is the no-query control for view construction. In
`fresh-count/5`, the denominator is five fresh views: the first counted view
builds the paragraph memo and the next four exercise reuse when the route
admits it. `same-count` keeps one view and times its queries, including its
first query's memo build for eager/source routes. These boundaries price the
paragraph view/index work; they do not support a whole-package speedup or a
zero-allocation claim.

The source-backed rows retain one physical cold load followed by cache hits.
The source version is printed as the pair `(id, revision)`; the ID is
process-local by design. Eager rows report null source-version fields because
their owning package is an unmanaged in-memory package. Managed rows use a
finite explicit execution context and retain its memory/cache gauges.

The real fixture contains markup-compatibility content that the managed source
contract refuses before semantic ownership can be charged. It is therefore
included in eager/source rows, while managed repeated-view evidence remains on
the two generated corpora. This is an admission result, not a missing sample.

For orientation, the baseline matrix rows below show costs for
the timed loop (elapsed values are wall-clock samples, not a claim or a
distribution):

| route | corpus | mode/repetitions | elapsed ns | allocated bytes | source cache after loop |
| --- | --- | --- | ---: | ---: | --- |
| eager | generated-200 | `fresh-count/25` | 1,712,519 | 221,125 | n/a |
| source | generated-200 | `fresh-count/25` | 1,716,438 | 312,815 | 1 cold, 24 hits |
| managed | generated-200 | `fresh-count/25` | 4,090,909 | 8,546,805 | 1 cold, 24 hits |
| eager | generated-10000 | `fresh-count/25` | 81,001,407 | 8,632,325 | n/a |
| source | generated-10000 | `fresh-count/25` | 81,582,120 | 9,239,641 | 1 cold, 24 hits |
| managed | generated-10000 | `fresh-count/25` | 197,372,845 | 414,128,631 | 1 cold, 24 hits |

The corresponding `same-count/25` controls allocate 342,735 B for the
10,000-paragraph eager and source views because their first index build is
inside the timed query, while the managed route performs its documented eager
index build during `document()` setup and reports zero query-loop allocation.
These rows include the visible XML pass and, for fresh source views, the first
physical payload load; they must not be read as a zero-allocation whole-document
claim. The A/A summary is retained to show the same-binary floor in the same
window before any before/after comparison.

## Retained files and hashes

`measurements/before/matrix.jsonl` has 114 rows covering the bounded matrix;
`aa.jsonl` has 224 rows (eight A1/A2 samples for each selected case). Their
SHA-256 sidecars, summaries, and the machine-readable
[before manifest](measurements/before/manifest.json) are retained. The probe
source hashes used by both baseline and candidate builds are:

```text
Cargo.toml     b7d81facd01cc99a64a1fcf67b9d074f54125aaf7f807114c58c1173921abfe0
Cargo.lock     83ca2b1fee71bb0b677fa76b2464415b738a0f0aa4e7b2b592f2f51ae9eb6c62
src/main.rs    07940463cf9aab7cbb5e74005c65eff618544dc64265805a18b2017c92a7c24e
run-before.sh  e95a0a26b66f46a3394e1318668ea733dfdae3fb931933cd34cc348bdc37eee2
run-aa.sh      b85ced6ef263df5390d2333add97999a8d2d7465e37091bf9d07ef824f3a74db
run-abba.sh    d1eb89d578f016f14e8b431ea77ae5f0d2d17d15dd67450a909555534f9df701
summarize.py   d15c2ae95982249d3b93727462530764dba5c12f04852c286d78f6c86a71afc9
```

Executable hashes identify the measured builds; the executables and external
target directories are temporary, not retained replay artifacts. Reproduction
requires rebuilding the two revisions with the retained probe and lockfile.
