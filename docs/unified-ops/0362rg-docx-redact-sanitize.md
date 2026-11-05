# 0362rg — DOCX redaction, sanitization, and story hyperlink products

This slice is based on `532cd8634b` on branch `uo2/0362rg-docx-redact-sanitize`.
It removes the remaining public raw `Commit` products from the three DOCX
record families and makes their source-backed consumers use the canonical
`edit_*()` → `Edit` → `commit()` shape.

## Capability map

| Record family | Raw product after this slice | Source-backed owner product |
| --- | --- | --- |
| `sanitize` | `SanitizePlan::apply() -> Result<Patch>`; reversible `Patch` retains exact no-op, forward, inverse, and source-conflict behavior | `source_backed::hyperlink_detachment::Commit` with the shared Package/Patch/Diagnostics triad |
| `redact` | `Plan::apply() -> Result<IrreversibleRedactionPatch>`; the patch is forward-only and carries no public source/XML payload access | `source_backed::hyperlink_redaction::Commit` with deferred bounded publication |
| `story_hyperlinks` | `Plan::apply() -> Result<ForwardOnlyPatch>`; the patch retains effect and completeness reports without inverse/publication methods | `source_backed::story_hyperlink_redaction::Commit` with deferred story overlays |

The raw families retain genuine standalone bounded mutation through
`sanitize::raw::Codec`, `redact::raw::Codec`, and
`story_hyperlinks::raw::Codec`. XML and relationship ingress requires the
family limits, and the codecs own mutable output state under those limits.
Redaction codecs accept one forward application and expose only the redacted
output and removed relationship identifiers. Public redaction plans, patches,
reports, errors, and debug output do not retain or print the removed target
payload.

All DOCX redaction, sanitization, and story-hyperlink tests and the two
story-hyperlink perf consumers now publish through canonical owner edits.
Read-only inventory and raw plan timing remain available where they are the
capability under test. Replay accepts a matching content artifact; changed or
stale source state remains refused. Signed changes fail at canonical commit,
while exact no-ops preserve source bytes and signatures.

## Validation

The lane used the memory-capped W8 environment, one build/test worker, and
`CARGO_TARGET_DIR=(deleted-coord-artifact)`.

- `cargo test -p litchi-docx --all-features` passed, including 1,490 unit tests,
  all DOCX integration suites, and 168 passing doctests (31 ignored).
- Focused redaction, sanitization, and story-hyperlink suites passed 5, 6, and
  23 tests respectively.
- `cargo clippy -p litchi-docx --all-features --all-targets -- -D warnings`
  passed.
- `cargo check --manifest-path tools/perf-baseline/Cargo.toml` passed.
- Perf-baseline all-target Clippy passed with the existing unrelated
  `while_let_on_iterator` and `needless_borrows_for_generic_args` lints
  explicitly allowed; the strict command reports those five pre-existing
  locations outside this slice.
- Both `docx-minimal` and `docx-full` rustdoc JSON glossary checks passed;
  the affected DOCX inventory rows are absent.

No benchmark or performance result is claimed by this API cleanup.
