# Change 0320: ODT source-backed text-block catalog

## Scope

This change adds `litchi_odt::SourceBackedDocumentCatalog` as an additive,
read-only lifecycle. It retains the validated positional package and a bounded
catalog of visible paragraphs and headings, while dropping the temporary
`content.xml` projection after catalog construction. The established
`SourceBackedDocument` owner is unchanged.

The catalog uses the existing ODT `Block` and `Kind` vocabulary. Its selected
block operation rereads `content.xml` and retains only the selected semantic
block. It does not expose physical XML offsets because the current ODF package
owner decompresses complete members for reads.

## Retention and validation

Opening retains no `content.xml`, `styles.xml`, `meta.xml`, media payload,
style registry, or semantic model. Styles, metadata, and media remain cold
unless the caller explicitly requests a package member or media payload.

The catalog follows the existing text parser's visible block order and
suppresses tracked-change definitions, note bodies, and ruby pronunciation
runs. Namespace aliases are resolved through the same XML namespace machinery.
The package source and archive limits remain authoritative; the existing ODT
content, text-block, and nesting ceilings continue to bound input and selected
semantic work.

Every operation checks the captured source version. A source revision observed
during selection or materialization returns `SourceChanged`, and an out-of-range
selection returns `None` without reading `content.xml`.

## Evidence

The deterministic integration coverage checks mixed/nested ordering, namespace
aliases, suppression, fresh selection reads, cold optional members, source
freshness, malformed input, package limits, encrypted sources, and exact
materialization. The new `source_catalog` target contains 11 tests; the
existing `source_backed` target contains 14 tests.

The serialized validation batch passed:

- `litchi-odf-common --lib`: 274 tests passed.
- `litchi-odt --lib --tests`: 554 library tests passed, all integration
  targets passed, including `source_backed` (14) and `source_catalog` (11).
- Strict Clippy with `-D warnings` passed for `litchi-odf-common --lib` and
  `litchi-odt --lib --tests`.
- Rustdoc with `-D warnings` passed for both crates.
- `rustfmt` and `git diff --check` passed cleanly.

Materialization now preflights both the physical ZIP source size and encrypted
manifest plaintext sizes. The same preflight is used by this catalog and by
the existing `SourceBackedDocument` owner, keeping the explicit materialization
boundary consistently bounded.

No streaming, RSS, OOM, throughput, latency, or other performance claim is
made here; those require a separately measured corpus and process-level
benchmark.
