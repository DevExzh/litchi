# Evidence packet for change 0633

Record: [`docs/performance/0633-xls-commit-single-framing.md`](../../0633-xls-commit-single-framing.md).

Queue item **XLS-9** of change 0630, priced by change 0620. One implemented
value-identical removal (`commit_source_backed`'s second complete target parse),
one implemented value-identical reuse (`Snapshot::from_bytes`'s shared-string
property table), and one frozen design with the measurement that says why it is
not implemented (fusing the open's two framing passes).

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (branch `feat/office-format-completeness`) |
| branch | `perf/0633-xls-commit-single-framing` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680` (shared, read-only, detached at the base) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0 |
| pinning | `taskset -c 9` under `setarch x86_64 -R`, `RAYON_NUM_THREADS=1`, `OMP_NUM_THREADS=1` |

Binary sha256 (`binaries.sha256`):

| binary | leg | sha256 |
| --- | --- | --- |
| `litchi-perf-baseline` | before | `c8a72cdec51672dcd14643304ea550120da7d13de7ce8009ebfbcb254efcc6de` |
| `litchi-perf-baseline` | after | `34e7752fb927716e8d00130331678546b0cbcc76f48b754f3a0142825672d0b5` |
| `xls_edit_probe` | before | `ac9e73fbefa724184df5ad0f10f6683da8af6c4dc7bc0b7c1e11e90a9d1216ea` |
| `xls_edit_probe` | after | `a7b4f254da5d19175c857945c94b6f4bd41c584ace0666a421648a431afc709b` |
| `xls_edit_corpus` | before | `2d39c87ba729c8d4ee15a9951494e84a82a086a89e581e955f61bbdf0b31c142` |
| `xls_edit_corpus` | after | `abeb2afdf1586dde7c6f4dac7c6c58ea626394e44fb2c84b29e078f6000b7194` |
| `xls_error_matrix` | before | `9c79c36762e459849d1572b72737ea81c6373d382395efafacf43f33801bf882` |
| `xls_error_matrix` | after | `6b87ef07f0eda0da019272a1452b50ab634a1953c2b2623d9327407c6950712f` |

The harness binaries are `tools/perf-baseline` built `--release --locked` from
each leg's tree with its own `CARGO_TARGET_DIR`. The three scratch binaries are
standalone Cargo projects whose only dependencies are path dependencies on the
leg's crates; their complete sources are in `probe/`, and the `Cargo.toml` files
show the only difference between the legs (the path). Every binary was copied
out of its Cargo target directory into a staging directory before being
measured, so a concurrent build could not relink one mid-run (change 0627).

`xls_edit_probe`, `xls_edit_corpus` and the six `capture_*.sh` scripts are change
0620's, reused with two changes: the default CPU pin (14 → 9) and the scratch
path. `analyze.py` is 0620's with one fix, described below.
`xls_error_matrix`, `framing_attribution.py` and `make_analysis.py` are new here.

## Contents

| path | what it is |
| --- | --- |
| `probe/main.rs`, `probe/Cargo.toml` | the attribution probe (change 0620's): one open, one staged edit and one commit per iteration on a real fixture, six operations, each phase timed separately, with a counting global allocator armed for exactly one measured iteration |
| `probe/corpus-main.rs`, `probe/corpus-Cargo.toml` | the corpus differential (change 0620's): every `.xls` through all five publication paths, printing either the exact refusal text or the SHA-256 of the published artifact |
| `probe/matrix-main.rs`, `probe/matrix-Cargo.toml` | **new**: the 0541-style first-error matrix. 29 synthetic packages, each a small valid package whose Workbook stream carries one defect at one chosen position, through the public `Snapshot::from_bytes`, printing the exact first typed refusal |
| `capture_counters.sh` | deterministic counters and the per-fixture record census, both legs |
| `capture_callgrind.sh` | instruction attribution: isolation pairs with `--separate-callers=2` |
| `capture_perf.sh` | native `perf stat` cycles and instructions, same isolation method |
| `capture_latency.sh` | A1 B1 B2 A2 wall clock on real fixtures through the probe |
| `capture_selectors.sh` | A1 B1 B2 A2 over the twelve registered harness selectors |
| `capture_corpus.sh` | the corpus differential, three runs per leg |
| `corpus_oracle.py` | the three correctness oracles over the retained corpus runs |
| `analyze.py` | reads the raw outputs into `callgrind/*.json`, `perf/*.json`, `latency*/latency-summary.json`, `counters-*.json` |
| `framing_attribution.py` | **new**: reads one before-leg annotation pair down to its sub-1% rows and prices the two framing passes of `Snapshot::from_bytes` against the semantic work around them |
| `make_analysis.py` | renders every summary into `analysis.txt` |
| `analysis.txt` | **the tables the record cites**, in the record's order |
| `framing-attribution.txt` | the framing attribution for `54016.xls` and `WithCustomViews.xls`, which is the frozen design's evidence |
| `corpus-summary.txt` | the three oracles' results plus the 16-run digest sets |
| `binaries.sha256` | the eight measured binaries |
| `environment.txt` | host, toolchain, base commit, branch, pinning |
| `gates.txt` | tails of the six gates |
| `decision.json` | the decision, its reason codes, accepted evidence and costs, known gaps and what is withheld |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE (merged by the coordinator, not by this change) |

### Raw outputs

| path | what it is |
| --- | --- |
| `callgrind/ann-<leg>-<fixture>-<operation>-{small,large}.txt` | 64 `callgrind_annotate --inclusive=yes` function-total sections (the per-source auto-annotation is stripped; nothing reads it) |
| `callgrind/pairs-<leg>.txt` | the (fixture, operation, small, large) tuples the analyzers difference |
| `callgrind/callgrind-<leg>.json` | per-operation inclusive Ir per function-and-caller-chain |
| `perf/perf-<leg>-<fixture>-<operation>-s<N>.csv` | raw `perf stat -x,` counters |
| `perf/perf-<leg>.json` | per-operation cycles, instructions, branches, branch misses, task clock |
| `counters-<leg>.jsonl` | one line per (fixture, operation): allocations, allocated bytes, peak live bytes, published bytes, source-backed diagnostics, and the raw per-sample phase vectors |
| `inventory-<leg>.jsonl` | the per-worksheet record census of each fixture and the two edit targets the probe picked |
| `latency/{a1,b1,b2,a2}/<fixture>/<operation>.json` | round 1 raw per-sample phase vectors |
| `latency/latency-summary.json` | round 1 p50/mean/p95/p99 per round, paired deltas in both directions, and the same-binary A/A and B/B floors |
| `latency-round2/…`, `latency-round3/…` | rounds 2 and 3, identical shape and setup; round 3 is the one the record quotes |
| `owned-source-chase.txt` | the callgrind isolation pair that chased the one selector group above the 5% review trigger |
| `selectors/{a1,b1,b2,a2}/{numeric,semantic,owned}.json` | the harness's own reports, including `output_sha256`, the corpus manifest and the per-sample vectors |
| `selectors/selector-summary.txt` | per-case p50 with both directions and both floors |
| `corpus/corpus-<leg>-r<1,2,3>.jsonl` | 591 file/operation rows per run; three runs per leg |
| `corpus/nondeterminism-<leg>.jsonl` | 16 runs per leg over `59858.xls` and `pivottable_dates_grouping.xls`, the two fixtures change 0620 found nonreproducible |
| `matrix/matrix-<leg>.jsonl` | the 29 first-error rows per leg |

## Three rounds of timing are retained, not one

Seven other agents shared the host for the whole of this change's measurement
window (load average 9-22); it was never quiet and waiting for quiet was not an
option. Round 1's same-binary A/A floor on `54016.xls` reached +12.3% at p50 and
its `WithCustomViews.xls` tail is unusable, so the capture was repeated twice.
Round 3's floors on the two changed scenarios are −0.15% and +1.92%, and it is
the round the record quotes. Rounds 1 and 2 are retained unchanged and are not
discarded: the six independent p50 readings of the changed scenario across the
three rounds are −38.13%, −43.18%, −41.82%, −44.06%, −38.50% and −38.25% on
`54016.xls`, and the spread between them is the honest size of the host effect —
as is the ±15% range the *unchanged* `54016.xls` controls cover across the same
three rounds while their instruction counts stay flat to within 0.5%.

## One fix to the reused analyzer

Change 0620's `analyze.py` matched `^\s*([\d,]+) \(([\d.]+)%\)`.
`callgrind_annotate` right-aligns the percentage, so a row under 10% prints
`( 4.83%)` with a leading space and did not match: every retained row under 10%
was silently dropped. The copy here matches `\(\s*([\d.]+)%\)`. Nothing change
0620 reported depended on the dropped rows — its tables are built from
`PROGRAM TOTALS` and from rows above 10% — but the framing attribution this
record turns on is entirely below that line, and the packets from change 0620
onward can be re-read with this copy.

## How to reproduce

```sh
PKT=docs/performance/results/change-0633
ST=/home/zhuhe/code/litchi-worktrees/staged-0633       # binaries copied out of their target dirs
export REPO=/home/zhuhe/code/litchi CPU=9
bash $PKT/capture_counters.sh  $ST/xls_edit_probe-before before $PKT
bash $PKT/capture_counters.sh  $ST/xls_edit_probe-after  after  $PKT
$ST/xls_error_matrix-before > $PKT/matrix/matrix-before.jsonl
$ST/xls_error_matrix-after  > $PKT/matrix/matrix-after.jsonl
bash $PKT/capture_callgrind.sh $ST/xls_edit_probe-before before $PKT/callgrind
bash $PKT/capture_callgrind.sh $ST/xls_edit_probe-after  after  $PKT/callgrind
bash $PKT/capture_corpus.sh    $ST/xls_edit_corpus-before $ST/xls_edit_corpus-after $PKT/corpus
bash $PKT/capture_perf.sh      $ST/xls_edit_probe-before before $PKT/perf
bash $PKT/capture_perf.sh      $ST/xls_edit_probe-after  after  $PKT/perf
bash $PKT/capture_latency.sh   $ST/xls_edit_probe-before $ST/xls_edit_probe-after $PKT/latency
bash $PKT/capture_latency.sh   $ST/xls_edit_probe-before $ST/xls_edit_probe-after $PKT/latency-round2
bash $PKT/capture_latency.sh   $ST/xls_edit_probe-before $ST/xls_edit_probe-after $PKT/latency-round3
bash $PKT/capture_selectors.sh $ST/litchi-perf-baseline-before $ST/litchi-perf-baseline-after $PKT/selectors
python3 $PKT/analyze.py callgrind $PKT/callgrind before after
python3 $PKT/analyze.py perf      $PKT/perf      before after
python3 $PKT/analyze.py latency   $PKT/latency
python3 $PKT/analyze.py latency   $PKT/latency-round2
python3 $PKT/analyze.py latency   $PKT/latency-round3
python3 $PKT/framing_attribution.py $PKT before 54016 open >  $PKT/framing-attribution.txt
python3 $PKT/framing_attribution.py $PKT before cv    open >> $PKT/framing-attribution.txt
python3 $PKT/corpus_oracle.py $PKT > $PKT/corpus-summary.txt
python3 $PKT/make_analysis.py > $PKT/analysis.txt
```

Run the deterministic captures first and the two timing captures last, with as
little else on the host as possible — see the two-rounds note above.

## What is not here

- No callgrind raw `.out` files: they are regenerated by the script and are two
  orders of magnitude larger than the annotations.
- No published artifacts. The corpus differential hashes them in memory.
- No profile of the eager `litchi_xls::Workbook` open path outside
  `Snapshot::from_bytes`, of the facade, or of any other crate. This packet is
  scoped to `crates/litchi-xls`'s `cell_values` editor.
- No RSS, cold-cache, physical-I/O, throughput or producer measurement.
