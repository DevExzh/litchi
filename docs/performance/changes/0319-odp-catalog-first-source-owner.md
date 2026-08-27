# Change 0319: ODP catalog-first source owner

Date: 2026-08-27
Status: additive source/catalog implementation; no performance claim

## Decision

Add `litchi_odp::SourceBackedPresentationCatalog` as a small catalog-first
owner alongside the existing `SourceBackedPresentation`. The owner retains
the bounded positional ZIP package, captured `SourceVersion`, and ordered
`SlideCatalogEntry` values only. It does not retain `content.xml`,
`styles.xml`, metadata, slide models, or media payloads.

## Mechanism

Opening validates the ODP MIME type and performs shared namespace-aware,
document-level XML validation followed by one catalog scan of a temporary
`content.xml` string. The scan requires the
`office:document-content` -> `office:body` -> `office:presentation` hierarchy,
records direct `draw:page` positions and optional `draw:name` values, bounds
the catalog at 65,536 pages, bounds names at 1 MiB, and uses fallible vector
reservation. Namespace aliases are resolved by URI rather than a required
producer prefix; malformed prolog/epilog content and unresolved references are
rejected. The temporary XML is dropped before the owner is returned.

`slide_at` rereads `content.xml` and optional `styles.xml` only for an in-range
selection, then reuses `Parser::parse_slide_with_styles_at`. That parser keeps
the established full-document validation and transition-style error
precedence; this change does not claim byte-isolated unselected-slide parsing.
Out-of-range and missing-name selectors return after freshness checks without
reading payload members. Media and arbitrary package members remain explicit
on-demand reads. `materialize` delegates to the exact source-backed common
owner and the existing mutable `Presentation` boundary.

All constructors use `SourcePackageLimits`, including password variants. Every
operation checks the captured source revision before and after work, with a
concurrent `SourceChanged` result taking precedence through the same
`prefer_current` pattern used by the ODS catalog owner.

## Verification scope

Focused integration coverage exercises catalog order/count, valid custom
namespace prefixes and decoded names, selected-slide readback parity, fresh
content reads, cold media, out-of-range no-read behavior, stale-source errors,
malformed prologue/epilogue/CDATA/entity input, and source-size limits. Exact
validation is recorded below.

## Claims deliberately not made

This change makes no claim about streaming, RSS, OOM behavior, allocations,
latency, throughput, physical I/O, cold-cache behavior, or broad ODF
performance. It also does not alter the unified `litchi::Presentation` facade.

## Validation

Validation ran serially with `CARGO_BUILD_JOBS=1` in one isolated target
directory:

- Full `litchi-odp` library and integration suite: passed, including 157
  library tests and all 8 catalog integration tests.
- Strict `litchi-odp` library/test Clippy with `-D warnings`: passed.
- `litchi-odp` rustdoc with `RUSTDOCFLAGS="-D warnings"`: passed.
- `rustfmt` and `git diff --check` for the changed sources: passed.
