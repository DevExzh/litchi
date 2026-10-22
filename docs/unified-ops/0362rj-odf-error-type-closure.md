# 0362rj — ODF public error boundaries and metadata output admission

Status: source repair complete; native performance and full-workspace integration remain open.

The initial unmerged source receipt for this slice was `6ad7dfd8e8`; the first
review return was `90bcb2d07c`; this continuation amends that receipt with the
remaining workspace consumers and typed replay boundaries found in review.

The public ODF owner, projection, edit, and output boundaries now use the
format-owned `litchi_odf_common::Error`. ODS package/document and flat owners,
source-backed spreadsheet/catalog operations, and source feature snapshots
translate their private core results at the public boundary. ODT generic
package and flat owners, specialized flat owners, and the root transaction
package now do the same. The ODT transaction edit verbs, history and joined
edits, patch and transfer plans, durable patches, sealed patches, and paragraph
transfer application carry that family error through their public staging,
commit, parse, serialize, apply, and transfer methods. Existing typed limit,
refusal, I/O, and validation causes are preserved by the format error
conversion. Private codecs, replay helpers, and other internal core helpers
remain on `litchi_core::Result` where they are not format API boundaries.

The metadata patcher now admits replacement values, including their escaped XML
length, before cloning the caller patch or allocating escaped values. Its output
writer also enforces the same bounded output budget while preserving metadata
XML fidelity and typed output-limit errors. A tight-output regression verifies
that an oversized replacement is refused as `Resource::OutputBytes`.

The review follow-up narrows the replay classification to public stale-base
routes (`PatchConflict`), the fixed transaction ceilings to semantic `Units` or
`Edits` limits, and joined/transfer isolation or dependency failures to typed
`UnsupportedOperation` refusals. Malformed artifacts and internal replay
integrity failures remain core validation errors.

The second review pass also migrates the ODT leniency, parser-hardening,
streaming, real-corpus, and crate-level documentation consumers to the family
error at the package/document boundary. Semantic document queries continue to
expose their core query errors and are converted only at the test boundary.

Validation completed before the coordinator's disk-space hold:

* `cargo check --offline --locked -p litchi-odt --lib --all-features`
* `cargo check --offline --locked -p litchi-ods --lib --all-features`
* `cargo test --offline --locked -p litchi-odf-common --lib metadata_patch_admits_output_before_cloning_values` — 1 passed, 366 filtered
* `cargo fmt --all`
* `git diff --check`

The affected ODT and ODS checks were run with the existing capped lane target.
Strict Clippy, broader tests, native timing, and full-workspace validation were
held when host free space fell below the coordinator's stop threshold; this
record makes no performance acceptance claim.

Coordinator integration also retains the independent 0362rh replay fixes.
Overlapping Formula history, ODM durable/history, ODI semantic replay and ODT
resource-GC paths now return the format-owned error with the same typed
conflict/refusal classification. The corresponding replay tests were migrated
together; the duplicate ODI artifact-kind helper from the parallel branches
was removed. Combined pinned-toolchain validation is pending on this merge.
