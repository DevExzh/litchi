# Evidence packet — change 0647

A relationship call that establishes nothing — `Relationships::get_or_add`
reusing a relationship the collection already carries, or `add_relationship`
naming an identifier that is already taken — now keeps change 0593's open-time
source capture, so the `.rels` member is published from the source instead of
being reserialized and audited into the same decision. Record:
[`docs/performance/0647-opc-get-or-add-noop-reuse-design.md`](../../0647-opc-get-or-add-noop-reuse-design.md).
`performance_claim: none`.

The change was written as a frozen design for the byte question change 0628
raised — keeping the capture would publish the source spelling where the source
spelling differs from the canonical one — and was implemented because the
measurement showed that question is already answered: the serialize-and-compare
route the capture replaces compares against the same open-time canonical capture
and, on equality, copies the **source** member. Both routes end at
`PreservationAction::Copy` of the same archive entry.

## Provenance

| field | value |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (`feat/office-format-completeness`) |
| branch | `perf/0647-opc-get-or-add-noop-reuse-design` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0647` (removed after commit) |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680` (shared, read-only) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21), valgrind 3.26.0 |
| build | `cargo build --release` for every probe, both legs, separate `CARGO_TARGET_DIR`s |
| CPU affinity | every measured process pinned to CPU 22 with `taskset` |
| staging | every measured binary copied out of its Cargo target directory before use (change 0627) |
| host state | eight agents building and testing concurrently; load average 13–39 across the window |

The two legs are the shared read-only base checkout (before) and this branch's
worktree (after); the only source difference is
`crates/litchi-opc/src/rel.rs` plus the tests the record lists. Probe binary
SHA-256:

| binary | before | after |
| --- | --- | --- |
| `rels_spelling` | `4cca923758ef418c…acf543a160d9e9c7` | `105418688723dd2a…84544e16567838f9` |
| `reuse_publish` | `5e8b5892cb31a566…3bb70927ef2235c6` | `c8cfd63e47f55220…952a38eb3a4a3b76` |
| `reuse_counts` | `e497bf1ce456ffe9…bb42ead7b18aac03` | `ff23978f498b248d…28ec93bc7d5554b6` |
| `reuse_iters` | `1d90ec93a719477c…972c396b594631d6` | `1b3d14984d80199b…4bcc8c97ba555605` |
| `docx_route` | `740d9416d7fdca2f…886dbb89b1df7a53` | `6f6eb6876c8fd8bb…71ca4ca7b228ec3e` |

## Contents

| path | what it is |
| --- | --- |
| `gates.txt` | every gate with its tail and exit status, plus a summary of the differential checks |
| `differential/spelling-before.txt`, `differential/spelling-after.txt` | `rels_spelling` over all 336 OOXML fixtures: per relationships member, the source byte length, the canonical byte length, the `same`/`differs` verdict, the internal relationship count and the distinct internal (type, target) count, plus per-fixture and corpus totals. `diff` of the two is **empty**. |
| `differential/reuse-before.txt`, `differential/reuse-after.txt` | `reuse_publish` over all 336 fixtures: 2,190 collections, 6,281 reuse calls, each publishing twice (seam-only baseline and reusing call) and comparing, with per-fixture rolling SHA-256 digests over every baseline and every reuse outcome. `diff` of the two is **empty**. |
| `differential/docx-route-before.txt`, `differential/docx-route-after.txt` | `docx_route` over all 63 `.docx`-family fixtures: the documented `Package::open` → `document_mut().add_paragraph_with_text` → `to_stream` route, classifying every source relationships member of the publication as `kept` or `moved`, with a SHA-256 of each publication. `diff` of the two is **empty**. |
| `counts/counts-before.txt`, `counts/counts-after.txt` | `reuse_counts` on five (fixture, owner) pairs: allocations and allocated bytes for the open, the seam, the reusing call and the save, the canonical `.rels` length of the owner, and the published byte count and SHA-256. Published bytes and digests are identical on both legs. |
| `callgrind/cg-summary.txt`, `callgrind/cg-summary2.txt` | the twelve plus four isolation-pair runs with their whole-run Ir totals |
| `callgrind/call-counts.txt` | per-scenario call counts for `try_to_xml_bytes`, `verify_authored`, `rels_uri`, `add_relationship`, `get_or_add` and `reuse_candidate`, from the N = 1 / N = 11 difference, for all four scenarios on both legs |
| `callgrind/annotate-{package,drawing}-{before,after}-11.txt` | `callgrind_annotate --threshold=99.0` tables for the N = 11 runs of the two reuse scenarios |
| `timing/summary.txt` | p50, mean, p95 and p99 per leg, the paired deltas in both directions, and the A/A floor, for all three timing windows |
| `timing/timing-package/`, `timing/timing-drawing/`, `timing/timing-drawing-2/` | the raw per-sample nanosecond blocks: `a1 b1 b2 a2` for the paired legs and `aa1 aa2 aa3 aa4` for the A/A floor, 500 samples each |
| `probe/` | the complete source of all five probe binaries with `Cargo.toml`, plus `scripts/callgrind-pairs.sh`, `scripts/timing.sh`, `scripts/stats.py` and `scripts/call-counts.py`. Path dependencies point at the 0647 worktree; retarget them to reproduce. |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge |

## Headline numbers

| measurement | before | after | tier |
| --- | --- | --- | --- |
| relationships members in the corpus whose source spelling differs from canonical | 2,189 of 2,216 (98.78%) | same | measured |
| reuse calls over the corpus, all returning an established identifier | 6,281 | 6,281 | measured |
| reuse publications byte-identical to the seam-only baseline | 6,281 / 6,281 | 6,281 / 6,281 | measured |
| cross-leg `diff` of the 6,281-call oracle | — | **empty** | measured |
| cross-leg `diff` of the DOCX ordinary route over 63 fixtures | — | **empty** | measured |
| `try_to_xml_bytes` per save with one reusing call (CFS) | 42 | **41** | measured |
| `verify_authored` per save with one reusing call (CFS) | 1 | **0** | measured |
| canonical `.rels` bytes built and discarded per save (CFS package / drawing1) | 733 / 2,260 | **0 / 0** | measured |
| save allocations (CFS package owner / drawing1 owner) | 402 / 523 | **363 / 363** | measured |
| instructions per open+reuse+save, CFS package member | 26,402,859.7 | 26,382,164.8 (−0.078%) | measured |
| instructions per open+reuse+save, CFS `drawing1.xml` | 26,408,332.9 | 26,289,565.8 (−0.450%) | measured |
| marginal instructions of the reusing call over the bare seam (package / drawing) | +23,176 / +109,387 | **+10,294 / −7,363** | measured |
| paired p50, CFS package | 1,552.21 µs | 1,541.53 µs (−0.69%) | measured, not claimed |
| paired p50, CFS `drawing1.xml`, second window | 1,552.23 µs | 1,530.51 µs (−1.40%) | measured, not claimed |
| A/A floor in the same windows (p50 / p99) | — | ≤0.09% / ≤1.77% | measured |
| the reported regression: drawing window 1, after leg, mean / p99 | — | **+11.70% / +230.79%** | measured, attributed to one contaminated block |

## Reproducing

```sh
# 1. Point the probe manifest at the two checkouts and build each leg.
cargo build --release   # from probe/, with CARGO_TARGET_DIR per leg

# 2. The corpus sweeps (deterministic; the leg only matters for the cross-leg diff).
taskset -c 22 ./rels_spelling  /home/zhuhe/code/litchi/test-data
taskset -c 22 ./reuse_publish  /home/zhuhe/code/litchi/test-data
taskset -c 22 ./docx_route     /home/zhuhe/code/litchi/test-data

# 3. Counts.
taskset -c 22 ./reuse_counts --package  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
taskset -c 22 ./reuse_counts --part /xl/drawings/drawing1.xml  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx

# 4. Instructions and timing.
probe/scripts/callgrind-pairs.sh <staged-dir> <fixture> <out>
probe/scripts/timing.sh          <staged-dir> <fixture> <out> 500 --package
python3 probe/scripts/stats.py       <out>
python3 probe/scripts/call-counts.py <callgrind-out-dir> package before
```
