# Evidence packet for change 0620

Record: [`docs/performance/0620-xls-edit-save-attribution.md`](../../0620-xls-edit-save-attribution.md).

Item **XLS-9** (rank 31) of change 0587. An attribution of the XLS
edit-and-save path, one implemented value-identical reuse in
`Transaction::commit_source_backed`, and one frozen design that is **not**
implemented.

## Provenance

| | |
| --- | --- |
| base commit | `1e4198321` (branch `feat/office-format-completeness`) |
| branch | `perf/0620-xls-edit-save-attribution` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-1e4198321` (shared, read-only, detached at the base) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0 |
| pinning | `taskset -c 14` under `setarch x86_64 -R`, `RAYON_NUM_THREADS=1` |

Binary sha256:

| binary | leg | sha256 |
| --- | --- | --- |
| `litchi-perf-baseline` | before | `0af006afa55262ee4b43179060d5e2ed78c0ecb8b66fc2286d9be97cd34859dc` |
| `litchi-perf-baseline` | after | `35291f3925aa583ef4339e9eecd061b37a38e47aa03e0b3be15b8121e2521aae` |
| `xls_edit_probe` | before | `a7130b304b47a4674db9be9a4af5f088facf8194ce1e6989d8894da20f3f3eb6` |
| `xls_edit_probe` | after | `bc7c2501e96a1ac133395cff5095bb370caff6f59f8078e8ada4ac6fa60cb3f4` |
| `xls_edit_corpus` | before | `1c2bd3d4846ccf8f964300db26876232f9a8c974e6b3b7e960d505aa59ce4ec9` |
| `xls_edit_corpus` | after | `3a6fca0e60802226e6ecdb3e51c56208d788675b07df36553a7a07606a1a4590` |

The harness binaries are `tools/perf-baseline` built `--release --locked` from
each leg's tree with its own `CARGO_TARGET_DIR`. The two scratch binaries are
standalone Cargo projects whose only dependency is a path dependency on the
leg's `crates/litchi-xls`; their complete sources are in `probe/` and the two
`Cargo.toml` files show the only difference between the legs (the path).

## Contents

| path | what it is |
| --- | --- |
| `probe/main.rs`, `probe/Cargo.toml` | the attribution probe: one open, one staged edit and one commit per iteration on a real fixture, six operations, each phase timed separately, with a counting global allocator armed for exactly one measured iteration |
| `probe/corpus-main.rs`, `probe/corpus-Cargo.toml` | the corpus differential: every `.xls` through all five publication paths, printing either the exact refusal text or the SHA-256 of the published artifact |
| `capture_counters.sh` | deterministic counters and the per-fixture record census, both legs |
| `capture_callgrind.sh` | instruction attribution: isolation pairs with `--separate-callers=2` |
| `capture_perf.sh` | native `perf stat` cycles and instructions, same isolation method |
| `capture_latency.sh` | A1 B1 B2 A2 wall clock on real fixtures through the probe |
| `capture_selectors.sh` | A1 B1 B2 A2 over the registered harness selectors |
| `capture_corpus.sh` | the corpus differential, three runs per leg |
| `corpus_oracle.py` | the three correctness oracles over the retained corpus runs |
| `analyze.py` | reads the raw outputs into `callgrind/*.json`, `perf/*.json`, `latency/latency-summary.json`, `counters-*.json` |
| `make_analysis.py` | renders those into `analysis.txt` |
| `analysis.txt` | **the attribution tables the record cites**, plus the native counter table |
| `corpus-summary.txt` | the three oracles' results, printed |
| `admission-census.txt` | which of the 126 fixtures the editor opens, which publish through each path, and the exact refusal texts with their counts |
| `nondeterminism-summary.txt` | 16 runs per leg over the two fixtures whose generic-commit output is not reproducible, with the digest sets |
| `environment.txt` | host, toolchain, base commit |
| `gates.txt` | tails of `cargo fmt --all --check`, `cargo clippy -p litchi-xls --all-targets`, `cargo test -p litchi-xls`, `cargo doc -p litchi-xls --no-deps` |
| `decision.json` | the decision, its reason codes, accepted evidence and costs, known gaps and what is withheld |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE (merged by the coordinator, not by this change) |

### Raw outputs

| path | what it is |
| --- | --- |
| `callgrind/ann-<leg>-<fixture>-<operation>-{small,large}.txt` | 64 `callgrind_annotate --inclusive=yes` function-total sections (the per-source auto-annotation is stripped; nothing reads it) |
| `callgrind/pairs-<leg>.txt` | the (fixture, operation, small, large) tuples the analyzer differences |
| `callgrind/callgrind-<leg>.json` | per-operation inclusive Ir per function-and-caller-chain |
| `perf/perf-<leg>-<fixture>-<operation>-s<N>.csv` | raw `perf stat -x,` counters |
| `perf/perf-<leg>.json` | per-operation cycles, instructions, branches, branch misses, task clock |
| `counters-<leg>.jsonl` | one line per (fixture, operation): allocations, allocated bytes, peak live bytes, published bytes, source-backed diagnostics, and the raw per-sample phase vectors |
| `inventory-<leg>.jsonl` | the per-worksheet record census of each fixture and the two edit targets the probe picked |
| `latency/{a1,b1,b2,a2}/<fixture>/<operation>.json` | raw per-sample phase vectors for each round |
| `latency/latency-summary.json` | p50, mean, p95, p99 per round; paired deltas in both directions; the same-binary A/A and B/B floors |
| `selectors/{a1,b1,b2,a2}/{numeric,semantic,owned}.json` | the harness's own reports, including `output_sha256`, the corpus manifest and the per-sample commit/publication vectors |
| `corpus/corpus-<leg>-r<1,2,3>.jsonl` | 591 file/operation rows per run; three runs per leg |
| `corpus/nondeterminism-<leg>.jsonl` | 16 runs per leg over `pivottable_dates_grouping.xls` and `59858.xls` |

## How to reproduce

```sh
PKT=docs/performance/results/change-0620
BEFORE=.../targets/0620-probe-before/release/xls_edit_probe
AFTER=.../targets/0620-probe-after/release/xls_edit_probe
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_counters.sh  "$BEFORE" before $PKT
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_counters.sh  "$AFTER"  after  $PKT
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_callgrind.sh "$BEFORE" before $PKT/callgrind
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_callgrind.sh "$AFTER"  after  $PKT/callgrind
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_perf.sh      "$BEFORE" before $PKT/perf
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_perf.sh      "$AFTER"  after  $PKT/perf
REPO=/home/zhuhe/code/litchi CPU=14 bash $PKT/capture_latency.sh   "$BEFORE" "$AFTER" $PKT/latency
CPU=14 bash $PKT/capture_selectors.sh <before-harness> <after-harness> $PKT/selectors
bash $PKT/capture_corpus.sh <before-corpus-bin> <after-corpus-bin> $PKT/corpus
python3 $PKT/analyze.py callgrind $PKT/callgrind before after
python3 $PKT/analyze.py perf      $PKT/perf      before after
python3 $PKT/analyze.py latency   $PKT/latency
python3 $PKT/analyze.py counters  $PKT before after
python3 $PKT/make_analysis.py > $PKT/analysis.txt
```

Run the deterministic captures first and the two timing captures last, with
nothing else on the host: the first latency run of this change was taken while
a corpus sweep occupied two other cores and its A/A floor reached 30% at p50 on
`54016.xls`, against under 2% when the host was quiet. Only the quiet run is
retained.

## What is not here

- No callgrind raw `.out` files: they are regenerated by the script and are two
  orders of magnitude larger than the annotations.
- No published artifacts. The corpus differential hashes them in memory; the
  `XLS_DUMP_DIR` environment variable in `probe/corpus-main.rs` writes them out
  and was used once, to locate the differing bytes of the CFB storage-order
  finding. Those dumps were deleted.
- No profile of the eager `litchi_xls::Workbook` open path, the facade, or any
  other crate. This packet is scoped to `crates/litchi-xls`'s `cell_values`
  editor.
