# Evidence, harness and coverage gaps against docs/GOAL.md (OLE2/OOXML)

Area: program-level documentation/tooling audit, not a code-path survey. HEAD `2fc5fc657`
(after change 0586). All citations are read-only; nothing was changed.

## 1. Ranked gap table

Ranked by how much each gap blocks `docs/GOAL.md`'s DEFINITION OF DONE (lines 796-821) for
OLE2/OOXML specifically. DoD clause letters: (a) reproducible baselines, (b) Amdahl/ROI order,
(c) work proportional to accessed content, (d) avoid parsing/recompressing unrelated content,
(e) unchanged media flow without decompression/copies, (f) bounded streaming memory,
(g) explicit bounded parallelism with real scaling, (h) hardware-counter-justified layout,
(i) SIMD only on proven hot loops, (j) ADR/correctness/preservation/security pass, (k) no
claim without reproducible evidence.

| # | Gap | GOAL requires | Exists | Missing | DoD | Effort |
|---|---|---|---|---|---|---|
| 1 | Per-PR/push CI check cannot detect a real regression | Deliverable 8: CI smoke check | `smoke` job builds `baseline` as a byte-copy of `current` with only the revision label changed, then asserts the comparator says "pass" against itself | Any comparison to actual prior history on push/PR | b, k | Medium |
| 2 | Most real DOC/PPT and a chunk of real XLS/XLSX/DOCX/PPTX fixtures can't reach the measured path | Corpus classes incl. multi-producer; DoD (a) | 8-of-57 `.doc`, 10-of-93 XLSX-cell, 1 DOCX full-text (0 bytes), 1 PPTX genuine-middle-slide fixtures are the whole measurable population per category | A tracked, first-class refusal census (reason histogram) instead of narrative repetition per record | a, b | High (admission scope) / Low (tracking) |
| 3 | Zero OLE2 range-source coverage vs. 3 independent OOXML mechanisms | "caller-supplied remote/range sources" as a benchmarked dimension, both priority formats | `SimulatedRangeSource`/`ReadAt` wraps OPC+XLSX; a dedicated `pptx_range_source` module; provider-lifecycle pacing (0447/0448) | Any `Cfb*RangeSource`/`Xls*RangeSource`/`Doc*`/`Ppt*` case | c | Low-Medium |
| 4 | L1/LLC hardware counters unavailable on this guest | Record L1 data misses, LLC misses | Zero-valued/unusable L1 alias, LLC absent from event set (standing since 0408/0409, restated 0417/0435/0437/0584) | A host where these events are real | h | Not fixable in-repo |
| 5 | Lock-wait/contention evidence scoped to exactly one producer | "lock wait time or a reliable contention proxy" for explicit bounded parallelism | One opt-in direct-mutex boundary: the OPC managed source-backed cache | Any equivalent for `litchi-cfb`'s own bounded rayon session (`src/shared_bulk.rs`) | g | Medium |
| 6 | Fuzz targets exist but never execute, in CI or locally | Continuous malformed/adversarial-input coverage | 13 `fuzz/` targets incl. litchi-cfb, litchi-doc, litchi-opc, litchi-docx, soapberry-zip | `cargo-fuzz`/nightly on any CI runner (`rust-ci.yml`, `perf-baseline.yml` have zero "fuzz" occurrences); local host also lacks both (0582) | j | Medium (add a nightly-toolchain CI job) |
| 7 | No DIFAT-bearing OLE2 fixture anywhere, real or synthetic | Very-large / structurally complete OLE2 corpus | Largest real `.doc`/`.xls` are ~1.45-1.4 MB; DIFAT needs the FAT array past 109 sectors, roughly 7 MB+ | A synthetic CFB shape built past that threshold | a, c | Low |
| 8 | The CI-regression-gated corpus (43 corpora/213 bindings) is almost entirely synthetic | "files from multiple real-world producers where licensing permits" as a corpus class | Real-producer breadth exists ad hoc per one-off record (0584's 57 `.doc`, 0570's 98 fixtures) | Any real-producer fixture in the schema-2 default/regression-gated catalog (only the RTF watermark is real) | a, k | Medium (licensing review) |
| 9 | `crud-coverage-index-v1.json` is a thin, self-declared "representative" sample | Scenario x format matrix | 33 rows over 15 categories; several categories show one format only (creation-from-scratch: DOC/PPT/XLS only, no OOXML; security row: XLS+XLSX only) | XLSB in the index at all — 0 of its 8 dedicated `xlsb_crud` selectors are mapped, despite XLSB being a named priority format | a | Low-Medium |
| 10 | Claim registry stalled at change 0467; HEAD is 0586 | Deliverable 6: machine-readable before/after results | 10 claims (7 landed/2 rejected/1 held), none after 0467 | Verified **not** a silent gap: every `performance_claim:` line in 046[8-9]/05xx records reads `none` — the last ~119 records are count/instruction-level or rejections by design | k | None (descriptive finding) |
| 11 | No consolidated final performance report | Deliverable 9: before/after tables, geomeans, scaling curves, remaining serial fraction, spanning HEAD | Per-change geomeans throughout REPORT.md; one appendix table explicitly labeled "descriptive; not current claims", stale at early changes | A synthesis pass across all landed changes as of current HEAD | k | Medium-High (likely end-of-program work, not overdue yet) |
| 12 | Verification breadth: perf is Linux-only; no Miri/sanitizers/loom anywhere | Cross-platform confirmation; memory/concurrency safety verification | `rust-ci.yml` runs a 3-OS correctness matrix (ubuntu/macos/windows); zero Miri/ASan/TSan/loom in any Cargo.toml or workflow | Perf numbers on macOS/Windows (perf-baseline.yml pins `ubuntu-latest` in all 3 jobs); any Miri/sanitizer/loom job | h, j | Medium |

## 2. Measurement blockers

- **DOC**: `SourceSnapshot::open` "admits only ordinary Unicode main-story documents"; 49 of 57
  real `.doc` fixtures refuse outright, leaving 8 measurable. `docs/performance/0586-doc-paragraph-hint-rejected.md:38-39,78-82`;
  restated at `docs/performance/HOTSPOTS.md:15`.
- **XLS**: `ConditionalFormattingSamples.xls` (the largest real `.xls` fixture, 1.4 MB) cannot
  complete full-text extraction through the public text API — "the library refuses its text
  extraction" — so change 0584's 703,937-link prediction for it is unreachable in practice.
  `docs/performance/GOAL_AUDIT.md` (0585-0586 entry, "What this batch does not discharge"
  paragraph, ~line 44).
- **XLSX**: 10 of 93 fixtures refuse the cell scenario outright, two of them named in a frozen
  plan itself. `docs/performance/GOAL_AUDIT.md:213-215`.
- **DOCX**: one fixture extracts zero bytes under the "full text" scenario.
  `docs/performance/GOAL_AUDIT.md:212-213`.
- **PPTX**: only `shapes.pptx` in the whole corpus has enough slides to exercise a genuine
  middle slide; every other PPTX read scenario is first/last-slide only.
  `docs/performance/0572-ooxml-range-source-attribution.md:129`.
- **CFB density gate** (change 0568's skip-uninterpreted-globals-bytes gate) "cannot fire on any
  input in this repository" — synthetic-only coverage, recorded as a limitation rather than
  tested. `docs/performance/0568-xls-worksheet-window.md:68-90,209`; restated
  `docs/performance/GOAL_AUDIT.md:268`.
- **DIFAT**: no corpus fixture, real or synthetic, has a DIFAT sector at all; the loop is
  provably untested. `docs/performance/0570-cfb-fat-run-batching.md:156-157`.
- **Fuzzing**: `cargo-fuzz` not installed, no nightly toolchain, on this host — and CI never
  installs either. `docs/performance/0582-zip-strict-scope-differential-fuzz.md:19-20,31-32`;
  `.github/workflows/rust-ci.yml`, `.github/workflows/perf-baseline.yml` (no "fuzz" string).
- **Hardware counters**: L1 alias returns validated zeroes, LLC events absent from the set,
  standing since 0408/0409. `docs/performance/changes/0408-opc-materialization-evidence.md:90`;
  `docs/performance/HOTSPOTS.md:2779,2797`.
- **No correctness defect newly observed this pass.** The one previously-known tooling defect
  in this territory (0567's discarded-classification bug in detect-then-open) was already
  repaired by change 0571 per `docs/performance/GOAL_AUDIT.md` (0565-0567 entry) — reporting it
  again would be stale, not new.

## 3. Stale or contradictory documentation

- `docs/performance/HOTSPOTS.md:6359` **"Ranked work queue"** — its own text says "provisional
  until baseline measurements are recorded" and its 15 rows cite work through roughly changes
  0001-0151/0190, with no rank for anything in the 300-586 range (all the recent OLE2/OOXML
  selective-read work). This directly contradicts the file's own newest-first convention: line 3
  is change 0586, yet line 6359 reads as if it were still current. No "superseded" marker exists
  between them. The adjacent **"Evidence still missing"** section at `HOTSPOTS.md:6381` has the
  same problem.
- `docs/performance/CRUD_COVERAGE.md:1187` **"Highest-return next cases"** — same fossil
  pattern: item 1 discusses change 0165 as the top "highest-return" candidate, deep inside a
  file whose newest entry (line 3) is change 0508. Both stale sections predate roughly 400 and
  340 subsequent numbered records respectively and could mislead a reader (or a fresh agent
  wave) into treating superseded priorities as current.
- `docs/performance/GOAL_AUDIT.md` mid-file entries (e.g. the "0572-0576" entry, lines 167-226)
  restate blockers ("A new OOXML hotspot is recorded and not yet addressed: the open reads every
  `*/_rels/*.rels` part…") that a *later* entry in the same file (0577, lines 100-166) already
  addresses with a design (still blocked on a missing accessor, per the standing-state brief) —
  not contradictory once read in order, but easy to quote out of context since both entries use
  present tense.
- No hard contradiction was found between `docs/GOAL.md` and the harness; the harness's own
  README repeatedly self-discloses its limits in the same paragraph as the feature (e.g.
  `tools/perf-baseline/README.md:4485-4488`, "Timing is only a first baseline… does not claim
  peak RSS, allocation counts, CPU utilization, lock contention, or cache misses").

## 4. Cheapest high-value evidence additions

1. **Wrap `litchi-cfb`'s existing `ReadAt` sources in the harness's `SimulatedRangeSource`.**
   CFB already reads through the generic `ReadAt` trait (`crates/litchi-core/src/source.rs:62`,
   implemented at `crates/litchi-cfb/src/shared.rs:2690` and `overlay.rs:595`), and
   `SimulatedRangeSource` already implements the same trait
   (`tools/perf-baseline/src/lib.rs:7914`). Adding `Cfb`/`Xls` range-source cases mirroring
   `XlsxRangeSourceOpen` et al. (`lib.rs:1970-1976`) looks like harness plumbing, not new
   production code — closes gap 3 for a fraction of the effort of any OOXML range work already
   done.
2. **Fix the `smoke` job's self-comparison** (`.github/workflows/perf-baseline.yml:259-267`) by
   fetching the most recent successful `full` run's artifact as `baseline` instead of copying
   `current`. The comparator, policy, and artifact plumbing (`reference-regression` job,
   `perf_compare.py`, `perf-regression-policy-v1.json`) already exist end to end; this is a
   matter of wiring an automatic trigger instead of requiring an operator to supply
   `reference_run_id` by hand. Closes gap 1, the highest-ranked item.
3. **Add one synthetic CFB "very-large" shape past the DIFAT threshold** (~7 MB, past 109 FAT
   sectors) to the existing `litchi-cfb-synthetic-v1` generator family
   (`docs/performance/CORPUS_MANIFEST_V2.md:69`). Closes gap 7 with a deterministic, licensable
   fixture — no real-world file needed.
4. **Map the 8 existing `xlsb_crud` selectors** (`tools/perf-baseline/src/bin/xlsb_crud.rs:168`)
   into `crud-coverage-index-v1.json`. The harness work is already done; only the index rows are
   missing, and `tools/validate_crud_coverage_index.py` already exists to keep it honest.
5. **Turn the two stale "ranked/highest-return" sections into pointers.** Replacing
   `HOTSPOTS.md:6359-6427` and `CRUD_COVERAGE.md:1187` with a one-line "superseded; see the
   newest entries at the top of this file" costs nothing and removes a standing
   documentation-staleness risk for every future reader or agent wave.
6. **Add a refusal census case** to the harness: for each format, report admitted/refused counts
   and refusal-reason histogram over `test-data/` as a first-class JSON field, rather than each
   OLE2 record re-deriving "49 of 57" or "10 of 93" by hand. Cheap relative to the value: every
   future record currently repeats this count from scratch.

## Notes on scope and confidence

Everything above is grounded in the current tree (`grep -n`/`sed -n` reads with file:line cited)
or in explicit self-declared statements inside the numbered records (quoted where the phrasing
matters, e.g. "no corpus fixture has any DIFAT sector at all"). Two items are explicitly
*verified-absent-of-defect* rather than gaps to fix: the claim registry's stall (row 10) and the
smoke job's own honesty about what it measures (it is named "smoke", not "regression gate", and
its self-comparison is a real, working plumbing check — just not the regression detector its
presence in a "Performance baseline" workflow might suggest to a new reader). No code was run
beyond `--help` invocations and read-only `grep`/`python3 -c` analysis of already-checked-in
JSON; no fixture, binary, or profiler was executed, so no new performance number is claimed by
this survey itself.
