# 0362ul42 — DOCX/OOXML raw record limits

This leaf is based on `30aaafbf98e097aeec8cf6fe3df6caaae62a37d6` and covers
the DOCX mail-merge and modern-comment codecs plus OOXML-common Custom XML
properties.  It also carries the required validation-name cleanup requested by
the integration naming audit.

## API and consumer changes

* `litchi-docx::mail_merge` now requires the caller's MCE policy on the public
  settings/recipient parse and write functions and on detached snapshots.
  Mail-merge transactions retain that policy, package settings/recipient
  readers derive it from the retained `SourceLimits`, and package edits bound
  the complete settings output before publication.
* All eight public modern-comment parse/write functions now require a caller
  MCE policy.  Package load/store, eager document reads, source-backed reads,
  and threaded-comment authoring derive that policy from the owning package's
  retained source limits.  Typed `modern_comments::LimitKind` values cover
  input/output bytes, depth, items, and string bytes; output-limit failures
  remain typed through package and transaction boundaries.
* `litchi_ooxml_common::custom_xml::{read_props,write_props}` now require a
  `&SourceLimits`.  Discovery, add/edit/snapshot/validation, DOCX data-store
  replacement, and all tests pass the package's retained policy.  Properties
  MCE processing and serialized output are refused before they exceed the
  retained per-part bound; typed Custom XML limit kinds remain intact.
* Validation names are module-local: DOCX exposes
  `validation::{DEFAULT_LIMITS, Limits, Error, validate_read_at}` and PPTX
  exposes `validation::{DEFAULT_LIMITS, Limits, Error, validate_source,
  validate_source_backed}`.  The old root-prefixed validation names and root
  re-exports were removed.  `litchi-ops` raw validation, the performance
  harness, and validation tests use the module paths.  The public PPTX format
  projection is now `format::pptx::Slide`; its tests and presentation view
  were migrated in the same change.

## Focused checks added

The source-only test additions cover typed output refusal for mail merge and
modern comments, and retained per-part refusal for Custom XML properties.
`git diff --check` passes.  `rustfmt --edition 2024 --check` parsed all
changed Rust files; it reports pre-existing repository formatting differences
(primarily trailing-comma style) but no parser error.  The initial source-only
review had no Cargo run because the coordinator withheld a bounded build slot;
the later bounded validation is recorded below.

## Bounded validation follow-up

The final source tip for this leaf is
`0bc42e8b0f79d5c74d68b214585e5d7226571f48`, based on
`30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`. The final source-only change
after the library run is a test initializer cleanup required by `clippy`; it
does not alter production code. The approved root lock was `Cargo.lock` SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
(139,606 bytes), with the pinned 1.95.0 toolchain and the bounded one-job,
20 GiB, offline/locked environment. No binaries or diagnostic logs are
committed.

The hydrated, sparse-safe validation results are retained outside the
repository:

* DOCX library tests: `cargo test --offline --locked -q -p litchi-docx
  --all-features --lib`, 1,525/1,525 passed, exit 0. Receipt
  `.../0362ul42-docx-ooxml-562-validation-20261025e/01-docx-library-tests.receipt.json`,
  15,836 bytes, SHA-256
  `cd862b559bb7253a465663b2267671bbd642c0e47dd2514789bacdf0b63a1ef3`.
  This run is bound to the production-equivalent `562602ed` tip; the later
  `0bc` change only fixes a test initializer for clippy.
* Affected all-target check: `cargo check --offline --locked -q -p
  litchi-docx -p litchi-ooxml-common -p litchi-pptx -p litchi-ops -p
  litchi --all-features --all-targets`, exit 0. Receipt
  `.../0362ul42-docx-ooxml-562-static-20261025g/01-affected-all-target-check.receipt.json`,
  14,144 bytes, SHA-256
  `7c0401beb37d1ad517b5ee7b568a0a28792d9e6c7a58817aa96dd10758989e82`.
* Affected all-target clippy at the final `0bc` tip: `cargo clippy
  --offline --locked -q -p litchi-docx -p litchi-ooxml-common -p litchi-pptx
  -p litchi-ops -p litchi --all-features --all-targets -- -D warnings`, exit
  0. Receipt
  `.../0362ul42-docx-ooxml-0bc-clippy-20261025h/01-affected-all-target-clippy.receipt.json`,
  19,542 bytes, SHA-256
  `bd592bc777eeb621e1b52747411c8f01c8f5eb2ed8ab3c28afcb2325e89f802a`.

The exact compile-time fixture inventory for the affected crates was hydrated
from tracked `test-data/` only (no `3rdparty/` path): 156 literal
`include_bytes!`/`include_str!` inputs, including the DOCX/PPTX corpus, Pages
merge fixtures, and PPTX view-properties fixture. This closes the earlier
sparse-checkout setup failures. An earlier safe-chunked `litchi-ops` runtime
test attempt still reports missing runtime corpus fixtures; it is not claimed
as a pass or as a production failure. The all-target check and clippy results
above are the accepted affected-consumer gates. No native benchmark, rustdoc,
or performance qualification is claimed by this record.

## Remaining PPTX raw-codec audit

The validation-name migration does not silently claim closure of the separate
PPTX configurable raw-codec sweep.  Concrete remaining public default/no-policy
boundaries in the current source are:

* `crates/litchi-pptx/src/laser/codec.rs`: `read` delegates to a mutable
  default `Limits`, while `write` has no caller policy.
* `crates/litchi-pptx/src/transition/reader.rs`: `read` selects
  `Limits::DEFAULT` beside the explicit `read_with` route.
* `crates/litchi-pptx/src/shape/{reader.rs,mod.rs}` and
  `shape/owned.rs`: public default `read` routes sit beside explicit
  `read_with` routes.
* `crates/litchi-pptx/src/view_properties/codec.rs`: public `parse` and
  `parse_strict` use fixed/default tolerance and `to_xml` has no output
  policy.
* `crates/litchi-pptx/src/media_parts/codec.rs`, `tag/codec.rs`,
  `notes/codec.rs`, `animations/codec/semantic.rs`, `comments/codec.rs`, and
  the presentation-properties codecs expose variable XML parse/write paths
  without a required module policy.
* `crates/litchi-pptx/src/model3d/package.rs` selects
  `codec::Limits::DEFAULT` for package read/write rather than deriving the
  retained source policy.

Root confirms these boundaries on current integration, not merely the old
leaf base. They are required engineering returns, assigned to 0362ul61
(laser/transition/shape), 0362ul62 (content codecs), and 0362ul63
(view/presentation properties and model3d). They remain open until the raw
APIs, all consumers and selected owner profiles conform; they are not accepted
exceptions or part of the redaction deferral.

## Coordinator composition review

Atomic `f9e3bff3b30f2cda083b3b1eac2a0cf0c6c0a58e` has the same tree as
record tip `4f05c153c0683ede4091a809da22f1d34f838a28`. The modern-comment
sizing correction at `562602ed` passes independent review. Its delta to
tested `0bc42e8b` only replaces a default-then-field-assignment in a test with
an equivalent struct initializer. Root verifies all three final receipt
identities, command logs, Cargo traces, unchanged locks, clean source and
empty terminal owned-process lists. The final Clippy target is 1,116,403,882
bytes. Earlier failed fixtures and output-length assertions remain recorded.

Composition finds one further standalone consumer: the performance harness
must call `litchi_pptx::validation::validate_source` after removal of the root
re-export. That correction and the explicit historical DrawingML adapter
migrate in this same merge (0362ul57b). The adapter preserves the historical
full library through six exact RTF/XLS/DOCX/PPTX spelling reversals. The
composed standalone harness check is still queued; scoped Cargo results above
do not establish that separate consumer's build or native performance.
