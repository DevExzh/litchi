# Evidence packet — change 0603

Record: [`../../0603-xlsx-fused-traversal-marker-admission.md`](../../0603-xlsx-fused-traversal-marker-admission.md)

Admission of declaration-only markup-compatibility worksheets to change 0546's
fused validate-and-parse traversal, with the event-by-event proof that the MCE
preprocessor's rewrite is invisible to the worksheet parser.

## Provenance

| | |
| --- | --- |
| base commit | `6c4c1469b` (branch `feat/office-format-completeness`) |
| branch | `perf/0603-xlsx-fused-traversal-marker-admission` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-6c4c1469b` (shared, read-only, detached at the base) |
| after checkout | `/home/zhuhe/code/litchi-worktrees/0603` (this branch, own `target/`) |
| probe | change 0602's `results/change-0602/probe`, unmodified; only the three path dependencies were repointed at the two checkouts |
| build | `cargo build --release`, `CARGO_TARGET_DIR` outside both checkouts, rustc 1.95.0 (59807616e 2026-04-14) |
| binaries | `binary.sha256` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, valgrind 3.26.0 |
| CPU pin | 24, for every measured process; seven other agents were active on the host |

The shared working copy at `/home/zhuhe/code/litchi` was neither built in nor
modified.

## Contents

| path | what it is |
| --- | --- |
| `census/markers.py` | classifies every `xl/worksheets/sheet*.xml` part of every `.xlsx` under the given roots exactly as `source_stream_admission` does, and reports which marker gate each part trips |
| `census/markers.tsv` | its output over `test-data/ooxml/xlsx` and `test-data/office-interop` — 95 files, 207 worksheet parts |
| `census/admission_census.rs` | the scratch in-crate test that counts, per worksheet part, which traversal is taken and whether it completes; appended to `shared_traversal_tests.rs`, run, and removed again (it is not part of the committed suite) |
| `census/admission-census.txt` | its output on the real corpus: 71 `Borrowed`, 31 `Rewritten`, 105 refused, and **0 parts on which the traversal completes, before or after** |
| `census/admission-census-derived.txt` | its output on the derived projections: 7 of 16 worksheet parts newly complete the traversal, 0 before |
| `fixtures/derive.sh` | regenerates the five measurement fixtures by calling change 0602's `degate.py` and `project.py` with the `plain` variant; the scripts themselves are not duplicated here |
| `fixtures/sha256.txt` | hashes of the five derived packages actually measured |
| `cg/run.sh` | the callgrind driver: isolation pairs at N=1 and N=4 against one retained editor, so the residue of the difference is M=3 complete plan-and-commit operations |
| `cg/diff.py` | change 0602's differencing script, unmodified, used for `attribution.txt` |
| `cg/attribution.txt` | per-operation inclusive Ir by function, both legs, all five fixtures |
| `cg/codec-share.py` | extracts the single `mce/codec.rs` row of `process_markup_compatibility'…process_ooxml` — the whole preprocessing pass — and differences the pair exactly, without `diff.py`'s largest-row selection |
| `cg/codec-share.txt` | its output: the preprocessing share per fixture per leg, including the control's byte-identical 166,834,244 Ir |
| `cg/inclusive-tops-before.txt`, `cg/inclusive-tops-after.txt` | the rows at or above 1.00% of the N=4 profile for each fixture and leg, retained in place of the 64 MiB of full annotations |
| `bench/measure.sh` | the timing driver: A1 B1 B2 A2 then S1..S4, 5 warmup and 40 samples per leg (30 on the largest), pinned to CPU 24 |
| `bench/measure-control-repeat.sh` | the second interleaved pass over the marker-free control (A3 B3 B4 A4), run immediately after the main window |
| `bench/stats.py` | change 0602's quantile and paired-delta script, retitled for these legs |
| `bench/legs/*.txt` | all 44 legs, one nanosecond duration per line (the 40 base legs plus the four-leg `ndp-[AB][34]` repeat of the control) |
| `bench/summary.txt` | per-leg p50/mean/p95/p99, the paired deltas in both directions, and the A/A floor |
| `tests/negative-control.txt` | the same committed test with `MceRewriteEquivalence::observe` forced to accept: 5 of the 12 adversarial synthetics then lose a typed refusal |
| `gates.txt` | tails of `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx --all-targets`, `cargo doc -p litchi-xlsx --no-deps`, `cargo test -p litchi-xlsx` (1,297 passed, 0 failed) and the focused shared-traversal run |
| `decision.json` | the machine-readable decision, in the shape of `results/change-0587/decision.json` |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge |

## Reproducing

```sh
# fixtures (into a scratch directory, not the repository)
docs/performance/results/change-0603/fixtures/derive.sh "$PWD" "$SCRATCH/derived"

# both probe legs
for leg in before after; do
  cp -r docs/performance/results/change-0602/probe "$SCRATCH/probe-$leg"
  sed -i "s#/home/zhuhe/code/litchi/crates#<checkout-for-$leg>/crates#g" "$SCRATCH/probe-$leg/Cargo.toml"
  ( cd "$SCRATCH/probe-$leg" && CARGO_TARGET_DIR=<target-for-$leg> cargo build --release )
done

# counts, then instructions, then timing
bash docs/performance/results/change-0603/cg/run.sh before
bash docs/performance/results/change-0603/cg/run.sh after
bash docs/performance/results/change-0603/bench/measure.sh
python3 docs/performance/results/change-0603/bench/stats.py "$SCRATCH/bench/legs"
```

`cg/run.sh` and `bench/measure.sh` carry the scratch paths used for this batch;
point `S` at your own scratch directory before running them.

## What is not here

* No full callgrind annotation: the twenty `*.incl` files were 64 MiB and were
  deleted after `attribution.txt`, `codec-share.txt` and the two
  `inclusive-tops-*.txt` extracts were taken from them.
* No derived `.xlsx`: `fixtures/derive.sh` and `fixtures/sha256.txt` regenerate
  and verify them from fixtures already in the repository.
* No allocation, RSS, cold-cache, publication or concurrency measurement, and no
  measurement of a real producer file as shipped — the record's Limitations
  section says why the last one does not exist.
