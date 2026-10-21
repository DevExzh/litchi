# 0362k — Retain Pages footnote resource policies

- Date: 2026-10-21.
- Status: source slice validated; integration with the final W4/W4b workspace
  remains required.
- Contract: 0351 §2.5 and 0355a §1 require owner-derived limits at every raw
  parse, rewrite and output boundary.

The Pages footnote graph projection and body-footnote edit paths now derive
wire, text-rewrite and Snappy profiles from the admitted package limits. The
projection keeps the package archive's input, field, nesting and rewrite-work
ceilings while constructing codec options; it no longer silently falls back to
its fixed codec defaults. Body-footnote insertion, removal, metadata rewrites,
attachment-owner validation, storage-candidate discovery and text-component
rewrites use the retained package profile for every wire parse and Snappy
bound/compression call. The Pages text helper maps an unsupported text-wire
nesting profile to its typed limit error rather than treating it as malformed
content.

The migration preserves the existing content policy. Malformed or semantically
invalid wire remains `InvalidSource`; input/output bytes, fields, nesting,
rewrite work, allocation and archive/Snappy ceilings retain typed limit or
allocation errors. Unknown storage fields are still copied through the rewrite
path. A package whose archive nesting limit is below the text-wire schema's
minimum is refused as `WireNesting` before attempting to decode. Source-derived
limits are checked before the corresponding parse, output allocation or
compression operation; there is no default-policy retry after a limit error.

Focused tests cover exact and one-over input/output byte ceilings, the text-wire
minimum nesting requirement, unknown wire-field preservation and typed failure
from a too-small field budget. Existing Pages footnote unit tests and IWA
footnote graph tests continue to exercise malformed, unknown, reference,
attachment and body-marker behavior.

Validation on this branch:

- `cargo check -p litchi-iwa -p litchi-pages --lib` — passed.
- `cargo test -p litchi-pages --lib` — 189 passed, 0 failed.
- `cargo test -p litchi-iwa --lib pages::editor::footnotes::tests:: -- --nocapture`
  — 10 passed, 0 failed.
- `cargo clippy -p litchi-iwa -p litchi-pages --lib --all-features -- -D warnings`
  — passed.
- `cargo fmt --all -- --check` — passed.

The attempted `cargo check --workspace --all-targets` reached the changed IWA
and Pages crates, then stopped at the pre-existing W8 integration mismatch in
`crates/litchi-ops/src/error.rs:1850`: the current `litchi_ppt::slide_order::Error`
is non-exhaustive and that downstream match has no wildcard. This slice does not
modify that classifier; the final workspace check must rerun after the
coordinator's error-contract integration. No native timing is claimed here.
