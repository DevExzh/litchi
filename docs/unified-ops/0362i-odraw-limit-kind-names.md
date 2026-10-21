# 0362i — Canonical ODraw resource-kind names

- Date: 2026-10-21.
- Base: `916fc215ff`.
- Status: naming slice complete; remaining raw/owner/error work is required.

ODraw traversal failures now use `litchi_odraw::LimitKind`; image failures use
`litchi_odraw::image::LimitKind`. The former root exports `Limit` and
`ImageLimit` are removed, with no public compatibility aliases. All internal
producers, pattern matches and tests migrate together. Workspace source
inspection found no consumers outside ODraw explicitly naming either old
type, and all five direct workspace consumers compile with every target and
feature enabled. The glossary bans the two old root names without banning
the unrelated picture family's `Error::Limit` variant.

This slice preserves resource dimensions, enum variants, error fields,
diagnostics, parsing, admission checks, allocations and output algorithms.
The image enum moves unchanged into its focused module. The
[source comparison](results/0362i/source-parity.json) covers all 12 changed
Rust files, retaining complete source hashes and normalized hashes. Reproduce
it with `python3 docs/unified-ops/results/0362i/verify_sources.py`; it checks
the relocated enum separately and normalizes only imports, exports, the
declared type-name changes and whitespace. No new native timing claim is
needed for this naming-only slice.

The [validation receipt](results/0362i/validation.json) records 95 unit tests
and the doctest, all-target/all-feature ODraw Clippy, a fresh rustdoc glossary
check, the five-consumer compilation, glossary tooling tests and workspace
formatting. Independent source review found no behavioral or fidelity
change. The target remains below the 20 GB limit.

This does not complete the ODraw raw ingress audit. `Record`, `Parser`,
`Children`, `Container` and the image file APIs still need the remaining
explicit-policy, bounded-allocation and observed-count work already tracked
by 0362c/0362d. Shared raw authoring and owner-policy propagation also remain
required. W4/W4b/W5/W8 integration, native performance repairs and the final
workspace gate remain open; item 22 alone is deferred.
