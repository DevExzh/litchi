# 0361b — Shared iWork errors follow the non-exhaustive contract

- Date: 2026-10-21.
- Base: `5f82848509`.
- Contract: 0351 §2.4, final paragraph; ADR 0034 §6.
- Status: shared-error slice validated; final integrated gate remains required.

Twenty-five public `Error` enums in `litchi-iwa-common` still permitted
exhaustive downstream matching. They cover the shared wire error, chart
values, comments, media playback, shape values, and table coordinates and
models. Each now has `#[non_exhaustive]`. The variants, payloads, messages,
validation rules and wire representations are unchanged.

All affected consumers migrate with these declarations: the physical IWA
archive, text rewrite, detection, Pages, Keynote, Numbers and the legacy IWA
adapter. Existing resource and allocation errors keep their precise mappings.
Existing invalid-source and invalid-format cases keep their categories.
Consumers that retain a common typed error continue to do so for an unknown
future variant; other fallbacks refuse through their existing domain error.
New fallback messages are fixed strings or include only an existing static
operation context. No new branch formats authored content.

The API debt inventory now emits `exhaustive_error` for a reachable public
`Error` enum without the attribute. A real-rustdoc regression distinguishes
exhaustive and non-exhaustive declarations. This extends the inventory's
coverage; it does not certify other crates or rename their errors. Remaining
rows outside this shared iWork slice must be resolved in the final sweep.

An independent coordinator peer review found no change to the mapping of any
current variant. This slice changes downstream matching requirements and
otherwise unreachable future fallbacks; it does not alter a codec, output
algorithm or allocation on a currently reachable path. No native performance
claim is made.

## Validation

`cargo check --workspace --all-targets --all-features` passes after migrating
every exhaustive consumer reported by the compiler. Intermediate failures
identified those consumers; none is accepted or suppressed.

The glossary/inventory tooling and iWork public-surface tests pass all 47
tests (`tools.test_check_api_glossary`, 31;
`tools.test_check_iwork_leaf_public_api`, 16).

All-target/all-feature Clippy with `-D warnings` passes for the eight affected
crates. The common, physical archive, text-wire and detection suites pass 417
tests across 11 targets, including one compile-fail doctest, with no ignored
tests. Workspace formatting, common rustdoc with warnings denied, and the
`iwa-common` glossary profile pass. Both independently locked tool manifests
(`perf-baseline` and `native-resave`) pass all-target checks; native-resave
reports its four existing unused-import/dead-code warnings.

The fresh common-crate inventory has no `exhaustive_error` rows. Its sole row
is `wire::NestedFieldEdit`: this is a `Copy` descriptor containing a borrowed
field-number path, expected presence and replacement value. It has only a
constructor and value getters. The bounded raw function
`patch_nested_fields_batched(data, edits, WireLimits)` consumes a slice of
these descriptors; the descriptor holds no package or staged mutable state
and has no edit/commit lifecycle. It is a naming-heuristic false positive,
like the raw `SubEdit` descriptor distinguished by 0351 §1. The row remains
in the committed inventory for review; no runtime exception or compatibility
alias is introduced.

The compact [validation receipt](results/0361b/validation.json) records the
commands and counts, and the [inventory](results/0361b/inventory.tsv) records
that reviewed row. The full integrated 53-check gate still follows the
pending wave merges. W4/W4b/W5/W8, remaining raw/shared API work and performance
repairs remain required. Nothing is pushed.
