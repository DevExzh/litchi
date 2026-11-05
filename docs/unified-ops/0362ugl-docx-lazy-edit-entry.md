# 0362ugl — DOCX lazy edit entry

This correction closes the DOCX rows left open by the edit-construction cost
audit in 0362ug. `Package::edit_with_options` and the source-backed primary
edit now retain the package, options and source policy and admit the native
document transaction on the first operation, projection or commit. The raw
OPC edit retains its source graph and copies the candidate only on the first
mutation. Ink retains its validated edit options and defers the package scan,
story inventory and binding capture until the first inventory,
selector-dependent operation or commit.

The same deferred boundary is used by source-backed story text, glossary text,
section layout, plain-paragraph copy/removal and SVG families. Selectors and
typed limits are retained before any source projection is kept. A failed
materialization leaves the edit without a staged state, while a failed native
operation leaves the existing staged state unchanged. Existing commit,
replay, inverse, signature, source-freshness and bounded-publication routes
remain the owners of validation and output.

Because lazy materialization can fail after construction, the affected
read-only projections are fallible `&self` accessors returning owned,
cheaply-clonable snapshots. All tracked DOCX and facade consumers were migrated
in this correction; no compatibility alias or eager fallback was added.

The follow-up cold-path proof also keeps SVG selector admission on the
deferred owner's retained read policy: constructing the next SVG edit does not
materialize the pending output from the preceding commit. The focused lifecycle
test checks that the selected raster and an untouched raster retain their read
counts across that constructor. The source-backed budget and plain-paragraph
copy refusal tests now force their first projection explicitly, so their
limit/refusal assertions still cover first use rather than the lazy constructor.

The source checks in this worktree are `rustfmt --edition 2024` and
`git diff --check`. Runtime tests, strict Clippy and the full gate remain
pending the coordinator's isolated Cargo slot. No performance qualification is
claimed by this source-only record.
