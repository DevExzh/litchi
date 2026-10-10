# 0053 — Capability reports and generated OfficeCLI comparison

## Scope

Publish `File::capabilities()` and `capabilities::Capabilities::of(format)`
through operations and the facade. This advances M1 R5; it does not close M1,
OfficeCLI parity, the property registry or the full goal.

The local base tip `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an
ancestor of this branch. No rebase is needed against that local tip.
Code-only portions of candidate `b9c6fd0bf4` supplied the report/generator
foundation; no candidate history or measurement evidence was merged.

## Behavior

Reports expose static format/state, retained-source availability, verbs, view
modes, element paths/relationships, property access flags and registered gaps.
DOCX is the available adapter; enabled pending adapters and disabled features
have distinct refusal codes/follow-ups. Locked files refuse every verb; signed
files permit reads and report protected edits. Property flags remain read-only
until M2 registers real setters. Element flags describe intrinsic format
support; whole-file verbs carry owner-state restrictions.

Report construction and iteration do no document I/O or allocation. A stale
file can still report its format capabilities, while actual reads validate
its retained source and fail. This is static capability discovery, not a
promise that arbitrary document payloads are valid or readable.

ADR 0034 ordinary reports omit native identity fields and native path forms.
Property filtering shares the query compiler's native-key classifier.
Positional and semantic paths remain visible; explicit raw selection stays in
path/query raw wrappers. No public native identity introspection door is added.

Pending File read errors now use the same capability refusal source.
`view::Mode::ALL` provides exhaustive mode enumeration. Public capability gap
subjects and property access flags provide a seam for later registry work.

Three JSON snapshots and `tools/officecli_capability_delta.py` generate
`officecli-capability-delta.md` from the base checkout's OfficeCLI schemas.
The table records schema and snapshot hashes. Generator unit tests and the
snapshot-hash drift check run in Rust CI without requiring OfficeCLI there.
Native paragraph/content-control selection metadata is explicitly classified
as an ADR policy difference. Missing rows remain visible and untriaged; the
table does not claim complete parity or complete gap registration. The current
DOCX comparison exposes nine missing element rows and 314 missing property
rows. Run subtypes sharing an ordinal path are not claimed as supported kinds;
raw and set-only schemas remain out of scope even for pending adapters.
View gaps have unique subjects and use the mode dispatcher as their source.
Semantic identity metadata has no separate public accessor in this step;
semantic path templates remain available.

## Verification

- All-feature checkpoint: 616 library tests, 27 integration tests and 141
  doctests passed. Review added a gap-uniqueness regression and strengthened
  pending/protection assertions; final targeted results are recorded below.
- DOCX-only/test-support: 351 library and 26 integration tests passed before
  the final review changes. Its first doctest run failed to load dependencies
  during overlapping target-directory builds; rerun results are below.
- No-default-feature library checkpoint: 227 tests passed.
- Facade DOCX integration test passed, including capability discovery.
- Warning-denied all-target operations Clippy, operations/facade rustdoc and
  all-feature workspace library check passed (including Python binding build).
- API glossary: 11 profiles, 144 banned names, zero transaction families;
  passed. Crate boundaries: 66 packages, 267 internal edges, 11 existing debt
  items; passed.
- Python checker/generator checkpoint: 1,159 tests passed; the final delta
  suite with two additional review regressions passed all 10 tests.
- Generated delta reproduces byte-for-byte with `--check` against
  `/home/zhuhe/code/litchi/3rdparty/OfficeCLI`. CI checks snapshot hashes without
  that external checkout.
- Existing allocation harness measured zero allocations for capability
  construction and iteration. DOCX fixture checks cover declared ordinary
  paths, queryable kinds, emitted-property membership and mode refusals;
  they do not prove that every declared property occurs in the corpus.

Final review verification passed: nine all-feature capability-related library
tests, four public snapshot/state tests, the warning-denied Clippy rerun, the
serial DOCX-only doctest rerun and the minimal-build capability tests. The
serial rerun resolved the dependency-loading failure. Changed-source formatting,
`git diff --check`, generated-table reproduction and artifact-limit checks also
passed. No Python runtime API test or paired performance acceptance is claimed.

## Remaining work

All non-DOCX adapters, complete property liveness and delta triage, typed public
view options, unified edits/batches, Python/CLI integration and paired read/write
performance acceptance remain outstanding. The allocator check is not a paired
performance benchmark. No iWork sources or base-checkout files are changed.

## Commit and hygiene

Staging again failed because worktree Git metadata is read-only:
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`.
This is the third consecutive goal turn with that restriction, but meaningful
implementation and verification remain possible. The goal remains active.
Records 0051–0053 are uncommitted; the partial old index must be refreshed before
committing. No commit or successful staging is claimed.

No worktree is created. Build output, generated lockfiles and scratch logs are
removed after verification. Repository evidence-size and binary checks print
nothing; no compiled artifacts are added.
