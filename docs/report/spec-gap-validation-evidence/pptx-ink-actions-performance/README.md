# PPTX existing InkAction performance scaffold

This directory is a review-gated bounded harness scaffold for absolute Litchi
host observations about the existing-target PowerPoint 2014 `inkAction` owner
at semantic owner `cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`. The production
transitive source and workspace/build inputs are pinned separately to baseline
`2a2ffa1cae4e6b7070082768ce84483e5d411dc8`; the exact capture checkout HEAD
is recorded at runtime. The normative semantic-owner design input is
`docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md` at
semantic-owner Git blob `597400950b1027c47cd6e4cbbedd23915bc0980e` and
SHA-256 `30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d`.
The public adapter, executable
host probe, matrix preflight, verifier, full-source manifest tool, and
isolated lockfile are present. No release profile or timing receipt has been
captured, and there is no native PowerPoint acceptance result or speedup claim.
The committed [`source-contract.json`](source-contract.json) records the
reproducible profile contract; the workspace-only untracked `docs/GOAL.md` is
explicitly excluded from the source input set.

The initial profile is deliberately finite: exactly 23 recipes and 42 lanes,
with no Cartesian expansion. [`PLAN.md`](PLAN.md) defines the public
`Package`/`Presentation` boundaries, fresh mutable setup,
`from_vec_with_limits` versus owner limits, exact graph variants, and
allocator/drop accounting. [`requirements.md`](requirements.md) defines the
receipt and later run gate. [`corpus-manifest.json`](corpus-manifest.json)
records the complete recipe/lane mapping and retained helper hash. Its matrix
preflight requires every recipe to be used by at least one lane and every lane
to resolve to a real recipe; orphaned or unresolved entries fail with a
nonzero status. The dual-pin guard derives every transitive local path package
from Cargo metadata, compares semantic contract paths with the semantic owner,
and compares the complete production package/build closure with the production
source baseline.

The fixture authority is the synthetic complete OPC helper at
[`crates/litchi-pptx/tests/pptx_ink_actions.rs`](../../../../crates/litchi-pptx/tests/pptx_ink_actions.rs),
committed in the pinned semantic owner commit, with SHA-256
`bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e`. The
`source_commit` field and the retained `owner_commit` field are compatibility
aliases for `semantic_owner_commit`; neither names the production baseline or
the runtime capture HEAD. Accepted ADRs are current capture context and are
hashed against the runtime capture HEAD. The checked corpus has no native
`inkAction` package. Synthetic owner graphs
establish only the named Litchi host behavior; they cannot establish
PowerPoint acceptance, rendering, playback, recognition, or producer-specific
path/MIME behavior.

Each fixture records exact member, content-type, package-relationship, and
part-relationship manifests. Successful save/reopen and apply lanes verify
those manifests, while refusal lanes require the complete serialized source to
remain unchanged. The opaque fixture contains an unknown MCE choice, an
inactive fallback, retained opaque profile bytes, and both internal and
external outbound diagnostics. Its `opaque_mce_scalar_edit` lane publishes a
prepared scalar patch through `Package::apply_ink_actions_patch()` and checks
the exact opaque target slice after serialization and reopen.

The adapter owns the isolated `harness/Cargo.lock`; the repository root
lockfile is not the profile lock. The lockfile is materialized and included in
the source/provenance closure before any later release profile. A missing or
root-lockfile substitution fails preflight. The later authorized run requires
a clean isolated checkout whose capture HEAD descends from both explicit pins,
full source/toolchain/host/binary/fixture hashes, three fresh processes, two
warm-ups, and twenty samples per process: at most 126 process launches, 252
warm-ups, 2,520 measured calls, and 2,772 operation calls. Until that review
gate passes,
[`report.md`](report.md), [`root-review.md`](root-review.md), and
[`results/`](results/) remain free of timing receipts.

The runner delegates target cleanup to [`cleanup_target.sh`](cleanup_target.sh).
The executable timing-free scaffold tests run that helper through valid,
missing, malformed, mismatching, symlink, and failed-run cases, and verify that
capture-gate and verifier failures do not create or delete diagnostic paths.
The later verification report retains host/load metadata, command start/exit
provenance and output paths, and `/usr/bin/time -v` user/system/elapsed fields.
Status-zero cleanup requires an exact verification-success sentinel in addition
to the ownership sentinel; a missing or mismatching receipt retains the target.

The timing-free correctness operator is [`operator_capture.py`](operator_capture.py).
It requires an explicit full `--expected-head`, independently checks that HEAD,
uses fresh external results and target directories with symlink-free lineage,
and stops at the first failed preflight, build, host probe, or matrix gate while
retaining final diagnostics. Its pure checks and mocked orchestration regressions
are in [`operator_capture_tests.py`](operator_capture_tests.py); these tests do
not build or invoke the executable.
