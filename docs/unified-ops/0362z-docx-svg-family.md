# 0362z — DOCX source-backed SVG family canonicalization

This slice starts at `3c1d25de24` on `uo2/w8-docx-svg`. It removes the
remaining public SVG batch record surface from the DOCX source-backed root and
makes `source_backed::svg` the only public owner family. The implementation
keeps the existing SVG graph scanner, source admission, deferred splice and
bounded writer; this record makes the API boundary consistent with 0351 and
0355a.

## API boundary

The public edit route is now:

```text
source_backed::Package::edit_svg(selectors) -> source_backed::svg::Edit
source_backed::svg::Edit::{attach_svg, attach, detach_svg}(&mut self, ..)
source_backed::svg::Edit::commit(self)
    -> Result<source_backed::svg::Commit, litchi_docx::Error>
source_backed::Package::apply_svg(&self, &source_backed::svg::Patch)
    -> Result<source_backed::svg::Commit, litchi_docx::Error>
```

`svg::Commit` is the shared `litchi_core::edit::Commit` triad containing the
new source-backed `Package`, the family `Patch` and `Diagnostics`. Its output
is the committed package's `write_to`/`save` route. The package remains
immutable, the edit is the only staging owner, and dropping an edit discards
its operations.

The readonly and staging capability map is:

| Previous public spelling or route | Canonical public surface |
| --- | --- |
| root `SourceBackedSvgAttachmentBatchEdit` | `source_backed::svg::Edit` |
| root `SourceBackedSvgAttachmentBatchCommit` | `source_backed::svg::Commit` |
| root `SourceBackedSvgAttachmentBatchPatch` | `source_backed::svg::Patch` |
| root `SourceBackedSvgAttachmentBatchSnapshot` | `source_backed::svg::Snapshot` |
| root `SourceBackedSvgAttachmentSnapshot` | `source_backed::svg::Picture` |
| root `SourceBackedSvgAttachmentReplacement` | `source_backed::svg::AttachmentReplacement` |
| root `SourceSvgAttachment` | `source_backed::svg::Attachment` |
| root `SourceSvgPictureSourceView` | `source_backed::svg::PictureSource` |
| root `SourceSvgPictureView` / `SvgPictureView` | `source_backed::svg::PictureView` |
| root `SvgResourceView` / `RasterResourceView` | `source_backed::svg::Resource` / `Raster` |
| root `SvgResourceSourceView` / `RasterResourceSourceView` | `source_backed::svg::ResourceSource` / `RasterSource` |
| root `SvgPictureOwnerState` | `source_backed::svg::OwnerState` |
| root `SvgInput` | `source_backed::svg::Input` |
| root `SvgAttachmentCommitDiagnostics` | `source_backed::svg::Diagnostics` |
| `Package::edit_svg_attachments` | `Package::edit_svg` |

`svg::Edit::source` and `projected` retain the exact selected-batch
projections. `Snapshot`, `Picture`, `Attachment`, `PictureSource`,
`PictureView`, `Resource`, `Raster`, `ResourceSource`, `RasterSource`,
`OwnerState`, `Input`, `AttachmentReplacement` and `PictureSelector` retain
the existing read accessors and identity semantics. The advanced attach
request and the atomic multi-selector attach/detach verbs remain available
through the canonical edit.

The former lifecycle records are now private engine types. They still carry
the detached before/after state used by the canonical wrapper, but no public
engine commit publishes a snapshot, exposes a target snapshot, or provides a
second commit spelling. All consumers, including the drawing lifecycle
integration tests, use `edit_svg` and `source_backed::svg` names in this
commit.

## Source and replay behavior

The source-backed family shape from 0355a is unchanged. Commit validates the
full selected XML and relationship closure and returns a deferred owner;
`write_to` and `save` perform the bounded source splice. The package's
`OpenOptions` source limits remain the limit policy; SVG has the existing
empty family `LimitKind`, and no `*_with_limits` entry was added.

The family retains no-op diagnostics, forward replay, inverse replay,
content-bound patch checks, source revision and lexical-freshness checks,
opaque relationship preservation, raster fallback preservation and exact
source-artifact restoration for the committed package pair. The tests cover
single and multi-picture edits, rejected intermediate operations, stale
sources, foreign bases, inverse restoration, managed reservations, lazy
resource reads and bounded limits.

Both DOCX glossary profiles already contained the `source_backed::svg` family
and `Package::apply_svg -> svg::Commit` result rule. This slice keeps those
rows and adds bans for every removed root spelling and the old edit route, so
the canonical names cannot regress.

## Validation

`cargo fmt --all` and `git diff --check` passed. With the capped W8
environment, the following checks passed:

```text
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 cargo check --locked -p litchi-docx --all-features --tests
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 CARGO_TEST_THREADS=4 cargo test --locked -p litchi-docx --test drawing_svg_lifecycle --all-features
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 cargo clippy --locked -p litchi-docx --lib --all-features -- -D warnings
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-docx --lib --all-features --no-deps
RUSTC_BOOTSTRAP=1 cargo rustdoc --locked -p litchi-docx --lib --all-features -- -Zunstable-options --output-format json --document-hidden-items
python3 tools/check_api_glossary.py --profile docx-minimal --json target/doc/litchi_docx.json
python3 tools/check_api_glossary.py --profile docx-full --json target/doc/litchi_docx.json
python3 tools/api_debt_inventory.py target/doc/litchi_docx.json
```

The focused suite passed all 24 tests. The rustdoc capture contains no legacy
SVG spelling; the fresh inventory has 24 residual rows outside this SVG
family and no SVG row. Logs and the 78,285,079-byte rustdoc JSON remain
outside the checkout; their hashes and regeneration commands are recorded in
`results/0362z/validation.json`. No benchmark, native timing or performance
claim is made here.

## Coordinator integration

The merge onto `bf529cb322` needs no textual conflict resolution. Its first
all-target Clippy run found twelve stale test conversions from the previous
record patch to the now-canonical patch type. Those consumers now borrow
`commit.patch()` directly, avoiding both the conversion and an unnecessary
clone. The repaired assembly passes all 24 lifecycle tests and strict
all-target/all-feature Clippy.

The worker checked both policy names against one all-features capture. The
coordinator additionally generated distinct minimal and full feature captures;
both pass, with 42 and 43 explicit result contracts respectively. The failed
Clippy log, passing checks and both capture hashes are retained in
[`results/0362z/coordinator-integration.json`](results/0362z/coordinator-integration.json).
This does not replace the final combined workspace gate or remaining
performance qualification.
