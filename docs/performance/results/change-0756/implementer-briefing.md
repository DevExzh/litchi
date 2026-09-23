# Wave briefing for implementers (2026-09-22, base 009d515bef)

You are an implementer in the Litchi performance program (Rust workspace, lossless
Office read/write library). A coordinator dispatched you with one task, a record
number NNNN, a git worktree and a measurement core. Read this whole file first.

## Authority and hard constraints

- `docs/GOAL.md` is the program charter. Read its "NON-NEGOTIABLE ADR AND
  SEMANTIC CONSTRAINTS" and "DECISION AND REGRESSION RULES" sections.
- Accepted ADRs in `docs/adr/` are hard constraints (read `docs/adr/README.md`
  and at least ADR 0003, 0005, 0006 and any ADR owning the code you touch).
- Owner decisions of 2026-09-16 (`docs/performance/0652-owner-decisions-for-the-third-wave.md`)
  set three standing trade-offs:
  1. The library is early alpha: **breaking public API / durable-format changes are
     acceptable** — make the change, document it, refuse old formats with a typed error.
  2. **Correctness and safety first, performance second.** No `unsafe`, no weakened
     malformed-input defences, no weakened limits/budgets, typed refusals stay typed,
     preservation by default, validation never mutates, no silent repair.
  3. Most inputs are benign: optimize the common benign path; the malicious minority
     must still be refused correctly (it may pay more).
- Scope: OLE2 (DOC/XLS/PPT/CFB) and OOXML (DOCX/XLSX/PPTX/XLSB/OPC/ZIP). Do **not**
  touch ODF crates or any iWork crate (`litchi-iwa*`, `litchi-numbers*`,
  `litchi-pages`, `litchi-keynote`).
- Do not add external dependencies. Do not add ambient threads, Rayon, filesystem,
  clock or network behaviour. Keep crate dependency direction (tools/check_crate_boundaries.py).
- Do not claim more than you measured. Records use `performance_claim: none`
  (numbers are reported as evidence, not registered as claims).
- Exact semantic no-ops must remain exact no-ops. Deterministic output must stay
  deterministic. If your change alters output bytes anywhere, say so explicitly,
  justify it, and prove determinism and semantic equality.

## Where you work

- Your worktree: `/home/zhuhe/code/litchi-worktrees/NNNN-<slug>` on branch
  `perf/NNNN-<slug>` (already created at base `009d515bef`, root `Cargo.lock` copied
  in — it is gitignored; never commit it).
- Your Cargo target dirs: `/home/zhuhe/code/litchi-worktrees/targets/NNNN` (after)
  and `/home/zhuhe/code/litchi-worktrees/targets/NNNN-before` (before). Always set
  `CARGO_TARGET_DIR`; always pass `--locked` (and `--offline` works). Use
  `CARGO_BUILD_JOBS=6`.
- Your scratch dir: `/home/zhuhe/code/litchi-worktrees/scratch/NNNN/`. Do **not** use
  `/tmp` (it is a RAM-backed tmpfs shared with other agents).
- Read-only shared base checkout: `/home/zhuhe/code/litchi-worktrees/base-009d515bef`
  (detached at the base). Never modify it. A prebuilt base harness binary (native,
  no allocator metrics) is at
  `/home/zhuhe/code/litchi-worktrees/targets/base-009d515bef/release/litchi-perf-baseline`.
  You may use it as the "before" leg if your branch does not change the harness.
  If you change the harness, build the before leg from a separate detached worktree
  of the base with only your harness commit(s) applied (create it under
  `/home/zhuhe/code/litchi-worktrees/NNNN-before-src`, copy the root Cargo.lock in,
  remove it when done).
- Never modify the main checkout `/home/zhuhe/code/litchi` or other agents' worktrees.
  The untracked `docs/performance/results/change-0741/` in the main checkout is an
  unfinished packet from a previous agent: you may read it, never modify it.
- Other agents build and measure concurrently on this 32-core host.

## Lessons from this wave (apply them)

- Build the before leg with the IDENTICAL cargo command, flags and features as the
  after leg (from a detached worktree at your base). The prebuilt base binary
  shifted untouched paths by 2.7–3.4%; builds are not bit-reproducible either.
- Code layout alone moved wall clock by up to ±40% on this host with identical
  instructions (0746), and argv[0]/path length moved PPT timings by about ±15% (0745).
  Report instructions and cycles (`perf stat -e instructions,cycles`) with every timing,
  keep path and argv lengths equal across legs, and prefer instruction/allocation
  evidence for small effects.
- Set `TMPDIR` to a folder under your scratch dir, so tests do not write to the
  RAM-backed /tmp.
- Four agents in this wave stalled for hours at their final cleanup step, apparently
  waiting on a background command's completion notification that never arrived.
  Run long commands (large `rm -rf`, builds, test suites) in the FOREGROUND with an
  explicit Bash `timeout` of up to 600000 ms. If you must background something,
  poll its log/exit file yourself; never end your turn waiting for a notification.
  Commit your record and packet BEFORE the final deletions, and send your final
  message promptly after them.
- Reviewers found real problems in every branch of this wave (undo/redo breakage,
  CPU DoS in a new check, admission on the wrong element). Before you finish, ask
  yourself adversarially: what input or call sequence makes the new path differ from
  the old one? Test that.

## Measurement method (before/after)

- Harness: `tools/perf-baseline` (`litchi-perf-baseline --help` lists ~520 cases;
  `--case`, `--samples`, `--warmup`, shape flags, `--json PATH`). Allocation counts:
  build `--features allocator-metrics --bin litchi-perf-baseline-alloc`.
- Pin every measured process to your assigned core with `taskset -c <core>`.
- Interleave before/after processes in ABBA order for at least 4 rounds (8 processes
  per case), each process >=15 samples after >=3 warmups (more for sub-ms cases).
  Report per-process p50/p95/mean, the median of process p50s per arm, the paired
  ratio after/before, and a simple bootstrap CI over paired processes. Keep all raw
  JSON reports in the packet. Do not measure while your own build is running.
- Include every regression >5% in any measured case in the record (do not hide it in
  an average). Measure at least one control case your change should not affect.
- Before optimizing, profile the timed region (perf is available: `perf record`,
  `perf report --stdio`; valgrind/callgrind and heaptrack exist). Profiles of the
  whole process include untimed corpus construction — isolate the timed work
  (e.g. a small probe binary, `--call-graph dwarf`, or symbol filtering) before
  drawing conclusions.

## Gates (run in your worktree with your target dir; record commands + exit codes)

- `cargo fmt --all --check`
- `cargo check -p <every touched crate and its in-scope dependents> --all-targets --locked`
- `cargo clippy -p <touched crates> --lib --no-deps --locked -- -D warnings`
  (and `--all-targets` for touched crates if it passed on the base; note pre-existing
  failures rather than fixing unrelated code)
- `cargo test -p <touched crates and in-scope dependents> --locked`
  (the facade: `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked`
  if a facade-visible behaviour/API changed)
- `RUSTDOCFLAGS="-D warnings" cargo doc -p <touched crates> --no-deps --locked`
- If the harness changed: `cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked`
  and the coverage validator:
  `python3 tools/validate_crud_coverage_index.py --index docs/performance/crud-coverage-index-v1.json --catalog docs/performance/results/perf-corpus-manifest-v2.json --selector-source tools/perf-baseline/src/lib.rs --checklist docs/CRUD_Scenario_Checklist.md --repo-root .`
- `python3 tools/check_crate_boundaries.py`, `python3 tools/non_iwork_gate.py verify`,
  `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural`

## Deliverables (commit them on your branch)

1. Production change + focused tests (preservation, refusal, limits, determinism,
   no-op, and the specific proof obligations your task names).
2. `docs/performance/NNNN-<slug>.md` in the house style of recent records (see e.g.
   0656, 0660, 0734): title stating the result; `Status:` line (retained / rejected /
   design-only) with `performance_claim: none`; what changed (files, functions);
   authority (ADRs, 0652 decision numbers); the evidence that motivated it; the
   before/after tables with corpus, case names, samples, core, binaries' SHA-256;
   every regression flag; what is NOT claimed; verification (gates); cleanup.
   Rejected results are valuable too: if the measured result does not justify the
   change, revert the production change, and commit the record + evidence as a
   rejection.
3. Evidence packet `docs/performance/results/change-NNNN/`: `README.md`, raw JSON
   reports, the scripts you used (small Python is fine), `gates.txt`, `cleanup.json`,
   and `log-sections.md` containing ready-to-paste short sections (one paragraph each,
   house style `## NNNN — <title>` + `[NNNN](NNNN-<slug>.md) ...`) for
   `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md`. Keep the packet small: no
   binaries, no perf.data, no corpora; summarized profiles only.
4. Do **not** edit the shared logs yourself: HOTSPOTS.md, REPORT.md, GOAL_AUDIT.md,
   BASELINE.md, ADR_COMPLIANCE.md, CRUD_COVERAGE.md, claim registries, coverage
   index, report-claim classification. (Exception: a coverage-index change strictly
   required by a harness selector you add — flag it.)
5. Commit messages: conventional style (`perf(pptx): ...`), ending with
   `Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`.
   Do not push. Do not merge into `feat/office-format-completeness`.
6. Cleanup before you finish: record binary SHA-256s, then delete your target dirs
   and scratch contents (`/home/zhuhe/code/litchi-worktrees/targets/NNNN*`,
   `/home/zhuhe/code/litchi-worktrees/scratch/NNNN/*`, any before-src worktree via
   `git worktree remove --force`). Keep your worktree and branch.
7. Final message to the coordinator (under ~600 words): result (retained/rejected),
   headline before/after numbers with case names, files changed, commits, gate
   results, regressions, anything you could not finish, and cleanup status.
