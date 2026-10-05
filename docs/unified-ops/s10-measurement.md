# S10 — the M0a before/after performance gate

- Status: measured. The verdict is in §1; owner decisions are needed on §9
  F1 and F2.
- Date: 2026-10-05
- Branch: `uo/m0a-s10-measure`, from `uo/m0a` at `b906691f3b` (S1–S9 and
  records 0011–0014 merged).
- Scope:
  - plan [0003](0003-m0a-migration-plan.md) §4 step S10, §5.3 and risk R5;
  - goal [`UNIFIED_OPS_GOAL.md`](../UNIFIED_OPS_GOAL.md), "PERFORMANCE BAR".
- Inputs:
  - the S1 baseline ([`results/0003/before`](results/0003/before/README.md));
  - the S8b route map ([`s8b-perf-routes.md`](s8b-perf-routes.md),
    [`results/0003/s8b/route-map.json`](results/0003/s8b/route-map.json));
  - the S9 notes ([`s9-deletion.md`](s9-deletion.md) §6);
  - the S8a, S5d and S6 notes in [`m0a-log.md`](m0a-log.md).
- Evidence: [`results/0003/after/`](results/0003/after/README.md), with the
  commands, environment, per-route tables, the pruned raw data and every
  follow-up.

## 1. Verdict

**M0a passes the gate on its stated terms ("blocks the merge until explained
or fixed"), not under the raw rule.**

Applied mechanically, S1's rule flags 8 of the 99 facade rows. Every one is
explained by a measured mechanism that is not a cost of M0a's code, and none
needs a fix in M0a's code:

- **E1, first call in a fresh process (7 rows).** The XLSX open and open
  lifecycle, the two PPTX `slide_count` rows, PPTX selected slide, and DOCX
  eager paragraph count and full text all time one call per fresh child
  process. Small probes built from the same sources with the same recipe show
  the same first-call time at both revisions (XLSX open 187.6 vs 187.9 µs,
  PPTX `slide_count` 0.54 vs 0.47 µs). Callgrind shows the same work (XLSX
  open +0.75% instructions). The gap belongs to the harness executables.
- **E2, the RTF sample loop's allocator state (1 row).**
  `rtf_file_unified_open` on the 540 KB corpus: with malloc trimming off, the
  unified open is 0.98× the pre-M0a one, and in a bare loop it equals its
  leaf twin.
- **The S8a separator.** The DOCX eager full-text row also carries S8a's
  paragraph separator (B3, an accepted behaviour change). It costs +2.77%
  instructions on the first call (76.4 vs 75.7 µs in the probe).

All 47 kept predecessors pass (0.87–1.03), and so do the other 91 facade
rows. R5 is confirmed and bounded (§5). The DOC routes and the PPT byte
route hold about one input size more after open; `retain_source(false)`
removes that; peak memory during open is unchanged.

The owner should confirm that E1 and E2 are accepted as explanations (§9
F1, F2). Without that, the raw rule would block the merge on artifacts of
the measuring harness.

## 2. What ran

- **Legs.**
  - **A:** the pre-M0a binary (S1's, SHA-256 `9c887ee5…2c79c04`, from
    `a8f45144c2`, formerly `da4c207478`) on S1's selectors.
  - **U:** the post-S9 binary (`dbe581e7…6985761`, from `d10dd4c747`) on the
    `unified_` successors, plus the 19 selectors S9 kept, which run under one
    name in both binaries.
  - **L:** the post-S9 binary on the `leaf_` twins.

  Both binaries use S1's build recipe, toolchain 1.95.0, identical
  third-party versions, and corpora with identical archive hashes.
- **Order.** `A U L L U A` per group and block, with each group's block
  contiguous in time:
  - 8 blocks for the in-process groups;
  - 4 for ODP and ODS, 3 for XLSX, 2 for PPTX and 1 for DOCX.

  That made 420 children, every one pinned to CPU 20, with S1's warm-ups,
  samples and warm cache.
- **Statistic.** The median of the per-child p50s.
- **Rule.** S1's: trigger = max(5%, S1 spread); +0.5 µs floor under 5 µs;
  full separation above 10% spread; the bimodal DOC rows compared mode by
  mode.
- **Overhead check.** The goal's open budget (≤ max(5%, 20 µs)) on S1 "open"
  rows; S1's rule as a review trigger elsewhere.
- **Host.** Shared. The 1-minute load was 2.55–21.1; 11 children waited at
  the gate for 600 s in total, against the 2-hour cap.
- **Re-run.** The PPTX and DOCX groups were re-run after an interruption,
  with the harness's corpus copies on disk (§8).

## 3. Rows over their trigger

| Scenario | Old selector → successor | Corpus | n A/U | before µs | after µs | ratio | trigger | separated | mechanism |
|---|---|---|---|---:|---:|---:|---:|---|---|
| open | `xlsx_file_open` → `xlsx_file_unified_open` | xlsx-cell-values-medium | 6/6 | 253.2 | 376.8 | 1.488 | 19.7% | yes | E1 |
| open | `xlsx_file_open_lifecycle` → `xlsx_file_unified_open_lifecycle` | xlsx-cell-values-medium | 6/6 | 2,648 | 2,801 | 1.058 | 5.0% | yes | E1 |
| list | `pptx_file_eager_slide_count` → `pptx_file_unified_eager_slide_count` | pptx-source-backed-media | 4/4 | 1.71 | 4.39 | 2.564 | 22.0%, 0.5 µs floor | yes | E1 |
| list | `pptx_file_source_slide_count` → `pptx_file_unified_source_slide_count` | pptx-source-backed-media | 4/4 | 1.61 | 4.45 | 2.765 | 55.4%, 0.5 µs floor | yes | E1 |
| read one | `pptx_file_source_selected_slide` → `pptx_file_unified_source_selected_slide` | pptx-source-backed-media | 4/4 | 71.1 | 77.6 | 1.091 | 5.9% | yes | E1 |
| list | `docx_file_eager_paragraph_count` → `docx_file_unified_eager_paragraph_count` | docx-source-backed-media | 2/2 | 99.3 | 107.4 | 1.081 | 5.0% | yes | E1 |
| full text | `docx_file_eager_full_text` → `docx_file_unified_eager_full_text` | docx-source-backed-media | 2/2 | 58.7 | 68.9 | 1.173 | 5.0% | yes | S8a B3 + E1 |
| open | `rtf_file_open` → `rtf_file_unified_open` | rtf-semantic-plain-large | 16/16 | 1,892 | 3,179 | 1.680 | 8.0% | yes | E2 |

### 3.1 E1: the first call in a fresh process

**Mechanism.** These routes spawn one fresh child process per sample and
time a single open or a single query in it. One-time costs therefore land
inside the timer: first execution of the code path, its page faults and
cache misses. In every row the old route and the successor make the same
leaf calls; for example, both `slide_count` paths are `check_source()` then
the leaf count.

In the harness:

- the leaf twin of the open rows is slower than the old route too (XLSX open:
  leaf 352 µs, old 253 µs);
- the post-S9 children take more page faults inside the timer (median per
  sample: XLSX open A 99, L 132, U 144).

**Test.** `cold_open_probe.rs` makes the same calls in a small program built
at both revisions with S1's recipe. It opens the harness's own corpus files
and runs as 30 fresh processes per variant, pinned and alternating. Median
first call:

- XLSX facade open: 187.6 µs before, 187.9 µs after.
- XLSX leaf open: 182.4 → 179.0 µs. The OPC index step is 136.5 → 136.3 µs.
- DOCX source open: 181.1 → 180.1 µs.
- DOCX eager `paragraph_count`: 105.1 → 104.1 µs.
- PPTX `slide_count`: 0.54 → 0.47 µs.
- PPTX `slide(100)` with text and name: 54.0 → 53.9 µs.

Callgrind on the first XLSX open: 1,130,713 → 1,139,175 instructions
(+0.75%); leaf +0.38%.

**Disposition.** Explained, not an M0a regression. The successor's code does
the same work at the same first-call cost; the gap comes from the harness
executables (265 MB; S8b added 118 routes and S9 removed 47 between the two
builds). The probes measure the result, not the mechanism. Code placement
and first-touch faults in the harness binaries are the likely cause, but
this record does not prove it. Owner decision on the harness: §9 F1.

### 3.2 E2: the RTF sample loop's allocator state

**Mechanism.** After each timed open, the in-process RTF loop builds an
untimed text and paragraph projection. That leaves glibc's heap such that the
next open re-faults about 2,700 pages (10.5 MiB): 2,684 minor faults per
open in the probe's harness-shaped loop and 0 in a bare loop. Which leg falls
into that mode depends on the leg's projection and the process layout.

**Evidence** (`rtf-allocator-check.json`, `rtf-open-probe.json`):

- Re-running the `rtf` group without the GNU `time` wrapper made the *leaf*
  lifecycle the slow one (3,193 µs).
- With `GLIBC_TUNABLES` that disable trimming and pin the mmap threshold,
  the slow mode disappears in every leg. Unified open is 1,899 µs and
  pre-M0a 1,932 µs (0.98×); open + text is 1,890 vs 1,904 µs.
- In a bare loop, unified and leaf opens cost 1,684 and 1,680 µs.
- The unified route's own pre-parse work is small: the flat-ODF probe costs
  3.86 µs on this input.

**Disposition.** Explained, not an M0a regression. Owner decision on the
harness: §9 F2.

### 3.3 DOCX full text and the S8a separator

Since S8a (B3), `File::text()` writes `\n` between DOCX paragraphs: 10,199
bytes instead of 10,000 on this corpus. The old route timed the
unseparated text, so the facade comparison times different work. S8a's log
already said earlier full-text timings were not comparable.

- On the first call, the separated text costs +2.77% instructions (1,249,832
  → 1,284,430) and 76.4 vs 75.7 µs in the probe.
- Unified and leaf now time the same text, at 1.020.

The remainder of the harness's +17% is E1.

**Disposition.** Explained: an accepted behaviour change plus E1.

## 4. Unified vs leaf: the goal's budgets

- **Open budget, max(5%, 20 µs): 21 of 25 rows within.** The four over:
  - `xlsx_file_unified_open`: +24.6 µs (7.0%). In the standalone fresh-process
    probe it is +8.9 µs (187.9 vs 179.0), within budget; the rest is E1.
  - `docx_file_unified_source_open`: +35.0 µs (10.8%). The probe gives
    +10.0 µs (180.1 vs 170.1), within budget; the rest is E1.
  - `rtf_file_unified_open` and `rtf_file_unified_open_lifecycle`, large:
    +1,355 µs and +102 µs. Both are E2; a bare loop gives +0.3%.

  The DOC and PPT opens are within budget: +0.3 to +3.3 µs. That includes
  the PPT/DOC source retention S5d/S6 added (S6 measured +2.8 µs on the PPT
  byte route).
- **Review items (23), all explained.**
  - **Detection, 16 rows: +0.5–0.6 µs (bytes), +4.1–5.4 µs (reader).** The
    unified detector is the coordinator `litchi::open` runs. It also decides
    whether a ZIP carries an OPC catalog (central-directory work; for readers
    it prepares the ODF package), while the ODF-only leaf reads just the
    `mimetype` member. It costs what the pre-M0a public detector cost: all 36
    facade detection rows are 0.99–1.04, and all 36 kept `detect_*` rows
    (12 selectors × 3 corpora) pass.
  - **Small DOC fixtures, 4 rows: +1.2 to +1.6 µs on 9–16 µs reads.** This is
    the fixed per-open cost of the unified DOC route: detection,
    compound-file classification, sealing the retained source into
    `SharedOleFile`, and one `Arc`. The facade comparison on the same files
    is 0.92–0.97.
  - **DOCX `list_paragraphs`, 2 rows: 1.31× and 1.30× the leaf.** The view
    returns `compat::word::Paragraph` wrappers, which M1 replaces with the
    node view. The facade comparison is 1.00, because the old facade built
    equivalent wrappers.
  - **DOCX source open + full text, 1 row: +30 µs (6.3%).** This is the
    unified path open (+10 µs standalone) plus E1. The facade comparison is
    1.005.
- **Within budget but not isolating.** The ODS unified opens are
  0.56–0.62× their leaf twins. The unified handle adopts the package
  detection prepared, while the leaf twin's `Spreadsheet::from_bytes`
  re-indexes. That pair does not measure the unified layer's cost.

## 5. Memory (R5)

The default harness and its allocator binary report no per-route memory for
these routes, because neither harness opens an allocation region on them. A
counting-allocator probe (`mem_probe.rs`) was built against each revision
with the harness's dependencies and lockfile. It was deterministic across
rounds and runs, with no leak on either revision. Per fixture and route:

- **DOC, path and bytes routes.** Held after open: +1.01× to +1.94× the input
  size.
  - `FloatingPictures.doc`: +349,838 B for a 335,360 B input.
  - A 1.62 MB DOC: +1,640,879 B.

  That is the input plus its CFB index, sealed into `SharedOleFile`.
- **PPT, byte route.** +1.01× to +1.14× the input: SampleShow +132,094 B for
  125,440 B; an 842 KB PPT +854,098 B.
- **PPT, path route.** +1.3 to +7.3 KB; that route keeps the source-backed
  package, not the bytes.
- **Peak during open.** Unchanged on every DOC/PPT route (within ±0.6 KB).
  Retention keeps memory alive longer; it does not add a copy.
- **`OpenOptions::retain_source(false)`.** It removes the retention: DOC
  +4,768 B and PPT +1,256 B over pre-M0a, the handle's own state.
- **Other formats.** +0.6 to +4.3 KB, except ODP by path, which holds
  −13.2 KB because the unified path route uses the source-backed ODP owner.
- **Fresh-child peak RSS** (filesystem routes, process lifetime high-water):
  unified − leaf is +0.4 to +0.8 MiB on every route.
- **Whole-child peak RSS** (GNU `time`) does not measure the routes. It is
  the `git` process the harness runs to record the worktree state, about
  199 MiB here. That is why S1 saw 199.4–199.7 MiB in every group. The
  pre-M0a binary's `git` lookup no longer resolves, so its column shows the
  harness alone.

This is the trade-off the owner override of 2026-10-04 accepted (plan 0003
§2.3); the opt-out works as designed.

## 6. Correctness

- **Harness checks.** Every child in the tables exited 0, so every leg passed
  the harness's untimed checks: projections, digests, oracles, frozen OLE2
  outcomes and source replays.
- **Outputs.**
  - XLSX semantic and source digests are identical on all legs.
  - The 12 successful DOC and PPT observations are identical before and
    after.
  - All legs refuse `duplicate-style-names.doc`. The unified leg words the
    refusal at code level ("corrupt_file: the file is malformed [Doc]") and
    keeps the leaf error in `Error::leaf()`, as S8b recorded.
  - DOCX full text is 10,199 bytes after and 10,000 before: the S8a
    separator.
  - No PPT or DOC text difference was expected (those fixes are on `uo/m1`),
    and none was seen.
- **Harness defect.** The harness itself was wrong at `b906691f3b`; see §7.

## 7. Harness fix (committed)

S8a changed `File::text()` for DOCX after S8b had merged, and nothing re-ran
the DOCX filesystem routes. At `b906691f3b`:

- the DOCX leaf full-text legs timed the unseparated text;
- the leaf signature and the fixed oracle used the unseparated text too.

Every DOCX filesystem route failed its untimed verification ("DOCX unified
and leaf semantic signatures differ"). A one-sample run of a binary built at
`b906691f3b` reproduced this; the other eleven groups passed.

`d10dd4c747` fixes the harness, not the library:

- the leaf legs call `text_with_paragraph_separator("\n")` (eager) and
  `extract_text_with_paragraph_separator("\n")` (source-backed);
- the leaf signature uses the separated text;
- the oracle joins the paragraphs with `\n`;
- the harness README records it.

`c475a527df` adds a test that all four full-text legs return the oracle.
Harness fmt and clippy (`-D warnings`) are clean, and six targeted harness
tests pass. The measured binary was built at `d10dd4c747`.

## 8. Problems and deviations

The README gives details.

- **Quota failures.** The shared per-user tmpfs quota on `/tmp` (49.38 GiB,
  held almost entirely by other agents' build directories) failed four PPTX
  children with `Disk quota exceeded`. The harness writes 12 × 17 MB corpus
  copies per PPTX child.
- **My error.** One PPTX child failed because I replaced the after binary's
  scratch copy with a link to the same file while its parent was running.
- **Re-run.** The session was then stopped, and the partial PPTX blocks on
  tmpfs were set aside. PPTX (2 blocks) and DOCX (1 block) were re-run with
  `--filesystem-root` on disk (ext4) for all legs. The other groups ran on
  tmpfs, as in S1.
- **Slow groups.** PPTX and DOCX have 4 and 2 children per leg, the minimum
  of 2 or more.
- **Raw data.** Pruned bundles (every field the analysis reads) are
  committed. The full raw set (8.8 MB tar, SHA-256 `70e44835…c0eb9676b6`)
  is kept outside the repository, with the probes, scripts and logs, at
  `/home/zhuhe/code/litchi-worktrees/uo-evidence/m0a-s10/s10-capture-raw.tar`.
- **No allocator binary.** It would not have reported these routes, so the
  allocation probe replaced it.

## 9. Follow-ups and owner decisions

- **F1 (owner decision): fresh-child routes measure the harness executable.**
  The `*_file_*` routes time one call in a fresh child, so their result
  depends on the harness binary's code placement, which every harness change
  moves.

  Options:
  1. accept E1 as explained (recommended for this merge);
  2. give those routes an untimed warm-up call in the child, which needs new
     route names (S8b naming rule);
  3. gate cold first calls with small standalone binaries, as `cold_runs.py`
     does here.
- **F2 (owner decision): the RTF in-process loop's allocator mode.** The
  untimed projection decides whether the next timed open re-faults about
  10 MiB.

  Options:
  1. accept E2 (recommended);
  2. run RTF children with fixed malloc tunables;
  3. add bare-loop variants under new names.
- **F3: drop S1's whole-child RSS tripwire.** It measures the `git` child,
  not the harness. Use the allocation probe or the fresh-child procfs peak
  instead.
- **F4 (not gating): DOCX `list_paragraphs`.** It costs 1.3× the leaf for the
  `compat` wrappers; M1's node view replaces them.
- **F5 (not gating): unified XLSX open.** It costs +8.9 µs over the leaf
  calls in a fresh process (within the 20 µs budget); DOCX source open costs
  +10.0 µs.
- **F6: future captures.** Set `--filesystem-root` (or `TMPDIR`) off the
  shared tmpfs; the per-user quota is shared by every agent.

## 10. Not claimed

- **No optimization claim.** Ratios below 1 are reported, not claimed: ODP
  opens 0.42–0.82×, DOC read-one 0.85×, and the bimodal DOC full text, whose
  pooled 0.70 comes from the pre-M0a binary's high mode; the low-mode ratio
  is 0.999.
- **Only S1's routes, corpora, host and warm cache.** Nothing about
  `cold-requested` caches, other corpora or hosts, concurrency, or the
  steady-state cost of the M1 surfaces.
- **E1's mechanism is inferred, not proved.** The probes show the code costs
  the same; that the harness gap comes from executable layout is inferred
  from that and from the fault counts, not measured. The XLSX lifecycle's
  text part and the eager DOCX open were not probed separately.
- **The allocation probe counts requested heap bytes.** It does not measure
  RSS, allocator overhead or page cache.
- **Budget-only successors and unpaired routes.** The two DOCX read-one
  successors with a different timed scope have no facade verdict. ODS
  selected cell/media and the untimed `docx_source_selected_paragraph` have
  no successor in M0a. The four ODS routes ran in both binaries as kept
  predecessors and pass.
