# Evidence: change 0599, one XLSB workbook parse per cell-value commit

Change record: [`0599-xlsb-commit-single-parse.md`](../../0599-xlsb-commit-single-parse.md).

Disposition: retained, implemented. `performance_claim: none`. The paired
medians and isolation-pair counters here are evidence, not a registered claim.

## Contents

| Path | What it is |
| --- | --- |
| `callgrind/isolation-pairs.txt` | The deterministic table: `noop_transaction_commit_save` and `edit_one_existing_scalar_save`, both legs, on `testVarious.xlsb` and `synthetic-4x2000x12.xlsb`, profiled at 5 and 15 samples and differenced over 10. Columns are `Ir@5`, `Ir@15`, `Ir/op`. |
| `callgrind/callgrind.sh` | The script that produced it. `--cache-sim=no --branch-sim=no`, `taskset -c 17`. |
| `callgrind/symbol-diff-noop-poi.txt`, `symbol-diff-edit-poi.txt` | Per-symbol `after − before` over the 15-sample profiles on `testVarious.xlsb`, 14 largest reductions and 4 largest increases. The two files agree to the instruction on the removed symbols; that is the proof that exactly one workbook parse was removed in each case. |
| `callgrind/symbol-diff.py` | The script that produced those two. |
| `callgrind/annotate-*-poi.txt` | `callgrind_annotate --threshold=80` for each leg and case, for the self-cost profile behind the diffs. |
| `timing/{A1,B1,B2,A2,A3,A4}-{poi,cond,syn_small,syn_large}.json` | The 24 raw `xlsb_crud` reports. `A` legs are the base build, `B` legs the branch build; `A3`/`A4` are the A/A control pair. Each carries all 40 per-sample nanosecond readings, the corpus description, the output SHA-256 and every gate for all eight cases. |
| `timing/timing.sh` | The driver: order A1 B1 B2 A2 A3 A4, 3 warmups and 40 samples per leg, `taskset -c 17`, four fixtures. |
| `timing/analyse.py`, `timing/summary.txt` | The analysis and its output: per fixture and case, pooled before and after p50/mean/p95/p99, the delta, and both A/A control deltas. |
| `timing/harness-gates.txt` | Every `xlsb_crud` gate value across 6 legs × 4 fixtures × 8 cases, and the changed-package-member list per fixture and case. |
| `differential/differential.rs` | The differential probe. One source file, included by two binaries. |
| `differential/Cargo-before.toml`, `Cargo-after.toml` | Their manifests: identical but for the `litchi-xlsb` and `litchi-core` paths, one pointing at the base checkout and one at this branch's worktree. |
| `differential/before.tsv`, `after.tsv` | The two reports. **`diff` is empty.** |
| `fixture/synthetic-xlsb.md` | How the two synthetic fixtures were generated, their digests and shapes, and what they are not. The generator is checked in at `tools/perf-baseline/src/bin/xlsb_synthetic_fixture.rs`; the fixtures are not. |
| `gates.txt` | The tails of `cargo fmt --all --check`, `cargo clippy -p litchi-xlsb --all-targets`, `cargo test -p litchi-xlsb` and `cargo doc -p litchi-xlsb --no-deps`. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge. |
| `decision.json` | The `litchi-perf-change-decision` record. |

## Provenance

Base: `08d968f8ec39a4fbcfd7dbb4e59cdd44e4ed99c8` (change 0587), on branch
`feat/office-format-completeness`.
Branch: `perf/0599-xlsb-commit-single-parse`.

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0;
valgrind 3.26.0. Eight agents were building concurrently throughout; every
measured process was pinned to **CPU 17** with `taskset`.

Both legs are `cargo build --release --locked --features xlsb-crud --bin xlsb_crud`:

| binary | sha256 |
| --- | --- |
| `xlsb_crud`, base leg (read-only checkout of `08d968f8e`, own `CARGO_TARGET_DIR`) | `2694b2b95ba2bd7f70403d875096d5199e78f666e8ece319b4b21a8db94a8da5` |
| `xlsb_crud`, branch leg | `b1f9d9d2c450607a07a883edd6f2442810b25c96dd84dccf914a6a7151d945e1` |
| `xlsb_commit_differential`, base leg | `f25d38f5a3495b6e23a0f24c3cc449ef8536b347e075935a65d9365f44b0ecdf` |
| `xlsb_commit_differential`, branch leg | `23a0870e73fb81a52d89ccd69c120474956b195c354d471ba10de63788bd1020` |

`xlsb_crud`'s own source is byte-identical in both legs; only `litchi-xlsb`
differs. The branch leg's manifest gains one `[[bin]]` entry for the fixture
generator, which does not participate in `xlsb_crud`'s compilation and does not
move `Cargo.lock`.

## What this packet does not establish

* No claim is registered; nothing here is a claim-registry entry.
* Timing is CPU only, from a warm in-memory source, on one pinned core of a
  contended host. No allocation, RSS, syscall, cold-cache, physical-device or
  cross-platform measurement was taken.
* This window's A/A floor is **worse** than the host's standing figure: |p50|
  median 0.84%, p90 3.22%, max 10.04% over 64 control pairs. The two synthetic
  fixtures' timing deltas fall inside it and are **not** separable by timing;
  only the instruction counts separate the legs there. The two real-producer
  fixtures' commit deltas are outside it.
* The corpus still has no large real-producer `.xlsb`. The synthetic fixtures
  are producer-free and deliberately carry none of the features whose parsing
  this change removes, so they bound the saving from below, not from above.
