# PPT picture-bullet validation

The candidate adds bounded `BlipCollection9Container` and `BlipEntityAtom`
readers and source-checked snapshot/edit/commit/patch APIs for replacing,
appending, and removing picture bullets. OfficeArt BLIP/FBSE payloads retain
unknown source bytes; delayed FBSE payloads resolve only through a supplied
store. Snapshots share admitted collection/source state, and exact semantic
no-ops return the original source allocation.

Destructive removal validates references in document PP9 records and drawing
shape PP9 records. Existing malformed references remain readable and survive
exact no-ops. Nonidentity reorder requires a reference rewrite and is refused.
Changed commits enforce protected-source and stream-only CFB admission; these
restrictions do not claim arbitrary nested-storage rewriting. Rendering and
ambient resource resolution are outside this owner.

Root validation used Rust 1.95.0 on the isolated candidate recorded in
`receipt.json`: 1,229 unit/integration tests and fourteen doctests passed,
with three unit/integration and eight doctest ignores. Strict all-target,
all-feature Clippy, formatting, and whitespace checks passed. The receipt
binds the three source/documentation files and compressed/raw test logs.
The matrix includes the previously committed native metadata coverage rows.

Coverage is synthetic, including malformed admission, quotas, source sharing,
shape references, signed/nested-source refusal, exact no-ops, changed save and
reopen, and exact inverse patches. This receipt makes no native-producer
acceptance or measured performance claim.
