# Evidence packet for change 0748

Record: [`docs/performance/0748-cfb-overlay-fingerprint-reuse.md`](../../0748-cfb-overlay-fingerprint-reuse.md).

A plan over sealed owned bytes computes its source and target SHA-256 digests
once; its composed view, direct write and atomic save stop re-hashing bytes
that cannot change. The object editor's copy-through render and the XLS
sheet-visibility source-backed commit open their immutable sources sealed.

## Provenance

| | |
| --- | --- |
| base | `ab29ac6291` (`feat/office-format-completeness` tip; includes 0745, not 0746) |
| branch | `perf/0748-cfb-overlay-fingerprint-reuse` |
| commits | `3cdcead75c` CFB sealed contract, `604132d280` sealed ingress adoption, `d663c504f4` harness evidence v2, `0a3476edd2` test extension, `f45930efb5` review follow-up (emission hash derived from the seal, two doc fixes, multi-chunk test) |
| after leg (measured) | the branch at `d663c504f4`; `0a3476edd2` changes one test, and `f45930efb5` moves the no-hash decision from a caller argument to the plan's seal without changing which passes run for either provenance |
| before leg | detached worktree at `ab29ac6291` with the harness-only commit cherry-picked (`03cf32fbce`, the same change as `d663c504f4`) |
| host | `environment.txt` |
| pinning | `taskset -c 16` for every measured process |
| toolchain | rustc 1.95.0 for every binary: the harness through the workspace pin, the probe and census with `+1.95.0` |

`binaries.sha256` lists every measured binary (`aa/` is a byte-identical copy
of the before harness used for the A/A control; `*_fp` are the frame-pointer
profiling builds).

Because the before leg carries the harness commit, its XLS numeric
source-backed and plan-only reports label their operation evidence `v2` while
the base library still produces the v1 owned values; `tools/perf_abba_summary.py`
would refuse those rows. They are not used as evidence here: this packet reads
only `elapsed_ns`, and the operation shape of each leg is pinned by its own
unit tests.

## Contents

| path | what it is |
| --- | --- |
| `latency/summary.json`, `latency/tables.md` | the 29-case ABBA matrix (8 processes per case) and its rendered tables, flags and isolation table |
| `latency/raw/` | every raw report of that matrix, gzipped (`<case>-<process>-<leg>.json.gz`; `summary.json` names them without `.gz`) |
| `latency/cfb-generic-long/` | a second, 60-sample ABBA round of the generic `cfb_file` control |
| `latency/cfb-generic-aa/` | the A/A round of the same case: the before binary in both arms |
| `latency/ole-common-finish-render/` | the `ole_common_finish_render` control, run after the gates with the same method |
| `instructions/isolation.json` | per-iteration `instructions:u`/`cycles:u` from `perf stat` pairs (4 vs 24 iterations, 3 repeats) |
| `instructions/isolation-long-spread.json` | the two small RK/MulRK cases again with 10 vs 210 iterations and 5 repeats |
| `alloc/alloc.json` | allocation calls, bytes and peak for the first measured probe iteration, two runs per leg |
| `profiles/` | flat profiles and frame-pointer call trees of the `54016.xls` generic commit loop, before and after; `first-change-fractions.txt` |
| `correctness/` | the census fixture list, the after census output (gzip) and the SHA-256 of both legs' outputs |
| `probe/` | the census source (`corpus.rs`: 0746's census plus fingerprint, second-publication, composed-view and save digests) and the two manifests |
| `scripts/` | `abba.py` (ABBA driver; resumes from existing raw reports; `ABBA_AA=1` for the A/A control), `isolate.py`, `alloc.py`, `tables.py`, `first_change.py`, `gates.sh`, and 0746's `perftree.py` |
| `binaries.sha256`, `environment.txt`, `gates.txt`, `cleanup.json`, `log-sections.md` | identity, host, gates, cleanup and the coordinator's log paragraphs |

The timing probe is 0746's `probe/main.rs`, unchanged (SHA-256
`9bd72f55935194379e2bf7dba1fcaf92e6fb19d3f0667f99211bc80709da03d6`); it is not
copied here.

## Reproduction

1. Build the harness for each leg with
   `cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`
   (the before leg from `ab29ac6291` plus the harness commit).
2. Build the probe and census for each leg from `probe/Cargo-{before,after}.toml`
   (0746's `probe/main.rs` at `../probe-src/src/main.rs`, `corpus.rs` beside it,
   the workspace `Cargo.lock` copied in) with `cargo +1.95.0 build --release
   --offline`, once plain and once with `--features alloc-count`.
3. Fixtures: `54016.xls` is `test-data/poi/test-data/spreadsheet/54016.xls`
   (SHA-256 `2e050f1f…a911a`); `xls-large.xls` comes from
   `xls_edit_probe --generate-xls-large` (SHA-256 `228c6585…20fb`).
4. `python3 scripts/abba.py OUT 16`, `python3 scripts/isolate.py OUT.json 16`,
   `python3 scripts/alloc.py OUT.json 16`, then `python3 scripts/tables.py
   OUT/summary.json OUT.json`.
5. Census, from the repository root, for each leg:
   `CENSUS_SAVE_DIR=DIR xargs -d '\n' -a correctness/fixtures.txt xls_edit_corpus > leg.jsonl`.
   The two outputs must be byte-identical.
