# change-0589 evidence packet: an empty overlay hashes the artifact once

Change record:
[`docs/performance/0589-ole2-snapshot-fingerprint-passes.md`](../../0589-ole2-snapshot-fingerprint-passes.md).

Disposition: **retained, partially implemented.** `performance_claim: none`.
Production change is three edits in `crates/litchi-cfb/src/overlay.rs`; the
`litchi-doc` and `litchi-ppt` changes in this branch are tests only.

## Provenance

| | |
| --- | --- |
| Base commit (both legs' source) | `08d968f8e` (`docs(perf): survey what remains across the OLE2 and OOXML path (0587)`) |
| Branch | `perf/0589-ole2-snapshot-fingerprint-passes` |
| Before leg checkout | `/home/zhuhe/code/litchi-worktrees/before-08d968f8e` (read-only, detached at `08d968f8e`) |
| After leg checkout | `/home/zhuhe/code/litchi-worktrees/0589` (this branch) |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), `--release --locked`, default release profile |
| valgrind | 3.26.0 |
| CPU pin | **9**, via `taskset -c 9`, on every measured process |
| Concurrency | eight measurement agents active on the host throughout |

### Probe binaries (sha256)

The probe was built twice per leg. **Pair A** produced every number in the
record's "Measured" section (counts, `perf stat`, callgrind, paired timing,
censuses). **Pair B** is pair A plus the `digest` subcommand and a `sha2`
dependency used only by it; it produced the corpus differential. No mode
measured in the record is reachable from the added code, and no timing figure
came from pair B.

| pair | leg | sha256 |
| --- | --- | --- |
| A | before | `9e693c753dfe64d4d0dc21caf630cf82dc18a0231f33d2e7473abf2f193d0023` |
| A | after | `96e62c0918f15f5e27a6301e4aaaf0e47483953cb5192239ff4eb0c1caae4d98` |
| B | before | `32cc96c146b0e6285ba977c4d224bc3906db3b91d61000b8819b6231635472e6` |
| B | after | `c2f01a129a55665e4eaab065232439cd0254ef1b88d33cf51f67406b0831fca5` |

Build directories (deleted after this packet was assembled):
`/home/zhuhe/code/litchi-worktrees/targets/0589-before` for the before leg,
`/home/zhuhe/code/litchi-worktrees/0589/target` for the after leg.

## Contents

| Path | What it is |
| --- | --- |
| `probe/main.rs` | The scratch probe, one source for both legs. Derived from change 0587's `docppt_survey` driver (`results/change-0587/doc-ppt/docppt-survey/src/main.rs`). Subcommands: `census-doc` / `census-ppt` (which fixtures each snapshot admits), `counts` (deterministic per-open source-read counts through a counting `ReadAt` adapter), `profile` (a loop for callgrind and `perf stat` isolation pairs), `bench` (per-sample nanoseconds), `digest` (the corpus differential). |
| `probe/Cargo.toml.tmpl` | The manifest template; `@ROOT@` is substituted with each leg's checkout so the two builds differ only in the crates under test. |
| `counts/census-doc-before.tsv` | All 57 `.doc` under `test-data/` and whether `litchi_doc::body_text::source::SourceSnapshot::open` admits them. 8 do. |
| `counts/census-ppt-before.tsv` | All 30 `.ppt` and whether `litchi_ppt::text_edit::SourceSnapshot::open` admits them. All 30 do. |
| `counts/doc-fixtures.txt`, `counts/ppt-fixtures.txt` | The admitted fixture lists, in the order every sweep uses. |
| `counts/counts-before.jsonl`, `counts/counts-after.jsonl` | Per-open `read_calls`, `read_bytes`, `full_artifact_reads`, `len_calls` and `version_calls` for `doc-snapshot-open` (8), `doc-snapshot-resolve` (4 — the other 4 fixtures refuse a paragraph read with `Refused(StructuralContent)`) and `ppt-textedit-open` (30). **`diff` of the two files is empty**: the change removes hashing, not reading. |
| `perf/perfstat-before.jsonl`, `perf/perfstat-after.jsonl` | `perf stat -r 5 -e cycles,instructions` isolation pairs, one line per fixture per mode: run `low` and `high` operations in one process, difference and divide by `high - low`. `low/high` is 100/600 under 500 KB and 20/120 above. |
| `perf/perfstat-beforeAA.jsonl` | A second `before` leg run in the same window, for the A/A floor. |
| `perf/perfstat-extra.jsonl` | The DOC readback (`doc-snapshot-resolve`, open plus first paragraph) on both legs, and `cfb-index-open` — one `SharedOleFile::open`, the unit the design section prices for the duplicate index parse. That path is untouched by this change, so it was measured on the after leg only. The `picture.doc` readback row is noise: that fixture refuses a paragraph read, so the probe exits before the loop. |
| `perf/perfstat-summary.txt` | The 38-fixture table with per-fixture deltas, the A/A column, and the medians. |
| `perf/native-sha-cost.txt` | The derived native per-pass SHA-256 cost. Dividing each fixture's before/after cycle delta by the removed pass count and the artifact length gives one constant — 2.048 cycles and 2.579 instructions per byte, within 4% across all 38 fixtures and a 149x size range — which is the evidence that the removed work is whole-artifact hashing and nothing else. Also carries the hashing share of a native open before and after, and the 20.2x factor by which callgrind overprices SHA-256. |
| `perf/bench-summary.txt` | Paired timing: p50/mean/p95/p99 per leg, deltas in both directions, and the A/A floor. |
| `perf/bench-samples/bench-<stem>-{A1,B1,B2,A2,A3,A4}.txt` | The raw per-operation nanosecond samples. `A1 B1 B2 A2` is the paired order (before, after, after, before); `A3` and `A4` are two further before legs, and the floor is `(A2+A4)` against `(A1+A3)`. |
| `callgrind/isolation-pairs.txt` | Total and `sha2` instructions per operation for both legs on three fixtures, and the derived **hash-pass count** per open (12 → 6 for the DOC generic open, 4 → 2 for the PPT text-edit open). Includes the reconciliation with change 0587's survey figures. |
| `callgrind/*.stderr` | The per-run valgrind banners. The `.out` files were deleted after extraction. |
| `differential/digest-before.tsv`, `differential/digest-after.tsv` | One line per `.doc` and `.ppt` artifact under `test-data/` (87 of them) carrying every identity this change can reach: the empty-splice plan's source and target fingerprints, its `is_noop` flag, the SHA-256 of what it publishes and the publish report's fingerprints; an exact-byte no-op splice plan's fingerprints, span count and published digest; an effective one-byte splice plan's fingerprints, span count and published digest; and the DOC and PPT snapshot fingerprints or their exact typed refusals. **The two files are byte-identical.** |
| `scripts/counts.sh` | Runs the `counts` sweep for one leg. |
| `scripts/perfstat.sh`, `scripts/perfall.sh` | One isolation pair, and the sweep over every fixture. |
| `scripts/bench.sh` | The `A1 B1 B2 A2 A3 A4` paired-timing run for one scenario. |
| `scripts/cgrun.sh` | The callgrind isolation pairs for one leg. |
| `scripts/summarize.py`, `scripts/bench_summary.py`, `scripts/native_sha_cost.py` | The three summary generators. `native_sha_cost.py` takes this packet's directory as its argument and regenerates `perf/native-sha-cost.txt` from the retained JSONL. |
| `scripts/gates.sh` | The gate runner. |
| `gates.txt` | The tail of every gate. |
| `decision.json` | The decision record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Reproducing

```sh
# both legs
git -C <repo> worktree add --detach <before-dir> 08d968f8e
git -C <repo> worktree add -b perf/0589-... <after-dir> <this branch>

# one probe per leg
for leg in before after; do
  mkdir -p /tmp/probe-$leg/src
  cp probe/main.rs /tmp/probe-$leg/src/
  sed "s#@ROOT@#<$leg-dir>#g" probe/Cargo.toml.tmpl > /tmp/probe-$leg/Cargo.toml
  (cd /tmp/probe-$leg && CARGO_TARGET_DIR=<target-$leg> cargo build --release --locked)
done

# then scripts/counts.sh, scripts/perfall.sh, scripts/bench.sh, scripts/cgrun.sh,
# and `snapfence_probe digest <repo>/test-data` on each leg.
```

Paths inside the scripts are absolute to this host's session scratch directory
and must be edited to rerun elsewhere.

## What this packet does not contain

No allocation, RSS, syscall, cold-cache, physical-device or range-source
measurement was taken. No `perf-baseline` selector opens either measured path, so
no harness scenario moved. The commit and save paths get the same halving on a
no-op publication; that is covered by `litchi-cfb`'s unit tests, not measured
here.
