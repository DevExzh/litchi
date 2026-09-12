# PPTX existing InkAction performance scaffold

This directory is a review-gated bounded harness scaffold for absolute Litchi
host observations about the existing-target PowerPoint 2014 `inkAction` owner
at `cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`. The public adapter, executable
host probe, matrix preflight, verifier, full-source manifest tool, and
isolated lockfile are present. No release profile or timing receipt has been
captured, and there is no native PowerPoint acceptance result or speedup claim.

The initial profile is deliberately finite: exactly 23 recipes and 42 lanes,
with no Cartesian expansion. [`PLAN.md`](PLAN.md) defines the public
`Package`/`Presentation` boundaries, fresh mutable setup,
`from_vec_with_limits` versus owner limits, exact graph variants, and
allocator/drop accounting. [`requirements.md`](requirements.md) defines the
receipt and later run gate. [`corpus-manifest.json`](corpus-manifest.json)
records the complete recipe/lane mapping and retained helper hash. Its matrix
preflight requires every recipe to be used by at least one lane and every lane
to resolve to a real recipe; orphaned or unresolved entries fail review.

The fixture authority is the synthetic complete OPC helper at
[`crates/litchi-pptx/tests/pptx_ink_actions.rs`](../../../../crates/litchi-pptx/tests/pptx_ink_actions.rs),
committed in the pinned owner commit, with SHA-256
`bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e`. The
checked corpus has no native `inkAction` package. Synthetic owner graphs
establish only the named Litchi host behavior; they cannot establish
PowerPoint acceptance, rendering, playback, recognition, or producer-specific
path/MIME behavior.

The adapter owns the isolated `harness/Cargo.lock`; the repository root
lockfile is not the profile lock. The lockfile is materialized and included in
the source/provenance closure before any later release profile. A missing or
root-lockfile substitution fails preflight. The later authorized run requires
a clean isolated checkout,
full source/toolchain/host/binary/fixture hashes, three fresh processes, two
warm-ups, and twenty samples per process: at most 126 process launches, 252
warm-ups, 2,520 measured calls, and 2,772 operation calls. Until that review
gate passes,
[`report.md`](report.md), [`root-review.md`](root-review.md), and
[`results/`](results/) remain free of timing receipts.
