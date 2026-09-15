# Evidence: change 0607, the PPTX authored-deck slide regeneration

Change record:
[`0607-pptx-authored-slide-regeneration-design.md`](../../0607-pptx-authored-slide-regeneration-design.md).

Disposition: design, retained. `performance_claim: none`, no claim-registry
entry, no production code changed. This packet holds the censuses, counts,
oracles and timing the record cites, and the probe that produced them.

Everything here was measured on the **read-only before checkout**
`/home/zhuhe/code/litchi-worktrees/before-6c4c1469b`. There is no "after" leg,
because no crate under `crates/` was modified.

## Contents

| Path | What it is |
| --- | --- |
| `probe/src/main.rs` | The probe. Subcommands: `admission <dir>` (open every `.pptx` under a directory and call `presentation_mut()`), `opened <dir>` (open every `.pptx` and `to_bytes()` it unedited, comparing with the source), `census <slides>` (deterministic per-materialization counts for an authored deck), `stability <slides> <out>` (save, re-assert the identical title, save again), `editdiff <slides> <out>` (save, change one title, save again), `resave <slides> <samples>` / `nopsave <slides> <samples>` / `create <slides> <samples>` (the timed loops). |
| `probe/Cargo.toml.example` | Its manifest, with the path dependency written as `<checkout>`; point it at the checkout being measured. |
| `admission/admission-test-data.tsv` | Every `.pptx` under `test-data/` (78), its size, whether `Package::open` succeeded, whether `presentation_mut()` was admitted, `is_modified()` after open, and the refusal text. Trailer line carries the totals. |
| `admission/admission-authored-reopened.tsv` | The same census over two decks litchi itself authored and saved. |
| `admission/opened-noop-save.tsv` | Every `.pptx` under `test-data/` opened with `from_vec` and saved with `to_bytes()` with no edit, source bytes against saved bytes. |
| `counts/census.txt` | Deterministic counts per materialization for authored decks of 1, 10 and 50 slides: slide and notes parts rebuilt, slide XML bytes serialized, how many of those slides are unmodified, and the two saves' sizes. |
| `counts/isolation-pair.txt` | The four callgrind whole-child `Ir` totals (`resave`/`nopsave` at `--samples 1` and `6`), the per-iteration differences, and the cross-check from the same profile's call graph. |
| `counts/cg-{resave,nopsave}-{1,6}.stderr` | The valgrind run logs behind those four totals. |
| `counts/whole-save-inclusive.txt` | `callgrind_annotate --inclusive=yes` lines for `Package::to_bytes`, `PackageWriter::to_bytes`, `PhysPkgWriter::write`, `DeflateEncoder::new`, `zlib_rs::deflate::{init,deflate}`, `PublicationPlan::from_package`, `verify_authored`, `flush_presentation` and `generate_slide_xml_with`, from the `resave 50 6` profile (7 saves). |
| `counts/flush-presentation-breakdown.txt` | The auto-annotated source of `flush_presentation` and `materialize_presentation` from the same profile, with each callee's `Ir` and call count. Divide by 7 for one materialization. |
| `oracle/no-op-rematerialization.txt` | Authored decks of 3 and 50 slides: save, re-assert the identical title, save again. Both archives byte-identical, same member count, names and order. |
| `oracle/one-slide-edit-member-diff.txt` | The 50-slide deck: save, change one title, save again. 160 of 161 members byte-identical; only `ppt/slides/slide1.xml` differs. |
| `timing/timing.sh` | The paired-timing driver: three warmup runs then 30 measured samples per leg, legs in A1 B1 B2 A2 order, every process pinned. |
| `timing/stats.py` | The summary generator: per-leg p50/mean/p95/p99, pooled legs, deltas, and the A/A and B/B floors from the repeated legs. |
| `timing/timing-50/`, `timing/timing-200/` | `A1.txt`, `B1.txt`, `B2.txt`, `A2.txt` (one nanosecond-per-iteration value per sample) and `summary.txt` for the 50-slide and 200-slide decks. |
| `timing/scaling-single-leg.txt` | The single-leg scaling table across 1, 10, 25, 50, 100 and 200 slides, with the member count of each deck. |
| `gates.txt` | The tail of every gate run in the worktree, and the design oracles' verdicts. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `6c4c1469beda47c2d44b502024a8dc680af35bfe` (`feat/office-format-completeness`) |
| Branch | `perf/0607-pptx-eager-save-slide-regeneration` |
| Measured leg | the probe built `--release` against the read-only checkout `/home/zhuhe/code/litchi-worktrees/before-6c4c1469b`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0607-probe-before` |
| Probe binary sha256 | `c5bfebc68e2bc7326937df30702ab7ac67eac6a90f109b7f6add1459cea42458` |
| After leg | none; no crate under `crates/` was modified |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 27` |
| Host load | other agents of this wave were active on the other 31 cores for the first pass; the paired legs retained here were taken in a quiet window, and the in-window A/A and B/B floors (+0.43% and +0.27% p50 at 50 slides, +0.06% and +0.30% p50 at 200) are the only statement about that |

## What this packet does not establish

No speedup, regression, allocation, RSS, cold-cache, real-producer or
cross-platform result, and no claim of what an implementation would save — none
was built. The deck is synthetic and generated by the probe; there is no
real-producer measurement of this path and there cannot be one, because a real
producer's file cannot enter the mutable writer at all (record §1). The
instruction counts rank work and not latency: `DeflateEncoder::new`'s 44.5%
instruction share is dominated by state zeroing, which callgrind counts per
byte, and its native cost is far smaller.
