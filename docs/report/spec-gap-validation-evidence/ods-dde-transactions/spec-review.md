# ODS DDE transaction specification review

Checked 2026-09-13 UTC against the vendored ODF 1.4 schema and Part 3
§§9.1.1–9.1.6, 9.8, and 14.7.1–14.7.5. This is a bounded review of the
current DDE lexical validation, typed cache renderer, and source authoring
boundary; it is not a full-content ODF validator review.

## Clear areas

`crates/litchi-ods/src/model/dde/lexical.rs` follows the XML Schema 1.0
`date`/`dateTime` lexical rules used by the RNG `dateOrDateTime` definition:
year-zero and leading-zero rejection, calendar day bounds, Gregorian leap
years, the `T` time form, fractional seconds, the `24:00:00` boundary, and
timezone offsets through `+14:00`/`-14:00`. `CachedValue::Time` delegates to
the repository's exact XML Schema duration parser. No date/dateTime or time
datatype blocker was found.

ODF Part 3 §9.1.1 says columns are implied by cells at the same row position,
and the RNG gives `table:number-columns-repeated` a default of one. The one
`table:table-column` emitted by the cache renderer is therefore sufficient for
the grammar even when rows contain multiple cells; no equal declared/effective
column count requirement was found.

Source construction validates required tuple attributes for nonempty XML text,
and newly authored/replaced sources require `office:name` to implement the
§14.7.1 prose requirement even though the RNG makes the attribute optional.
Unchanged unnamed legacy declarations remain a deliberate preservation case.

## Findings and disposition

The three previously recorded blockers are resolved in the reviewed candidate:

* An empty authored `CachedTable` is materialized as one
  `<table:table-row><table:table-cell/></table:table-row>`. This satisfies the
  RNG's one-or-more row and one-or-more cell requirements while retaining an
  empty logical cache value.
* Non-whitespace text or CDATA in a retained cache, including `text:p`, marks
  the cache opaque. Typed whole-cache replacement then refuses the edit before
  publication, while a source-only edit preserves the original cache bytes.
* Predefined and numeric character references in retained cache markup are
  admitted and retained byte-for-byte. The scanner marks that cache opaque, so
  typed replacement still fails closed; custom entity references remain
  refused because DTD/entity declarations are outside the accepted boundary.
* Scalar replacement of a named cache preserves its modeled `table:name`.
  Cache styles, spans, covered cells, row/column groups or headers, nested
  tables, unknown attributes, and unsupported reference vocabulary remain
  outside the typed projection and refuse replacement rather than being
  discarded.

No remaining blocker was found in this bounded review under the documented
canonical-refusal and authoring policies.

## Source and validation receipt

The final reviewed engine snapshot is identified by SHA-256
`18683f6ebac6cfdb2c2f48d1faa388f8253d8b2f9a172a7b3b361148afd8b891` for
`crates/litchi-ods/src/model/dde/transaction.rs`. Relative to the initial
semantic review snapshot, this final freeze contains the BOM source-range
coordinate correction, resource-estimate corrections, and modeled cache-name
preservation with conservative refusal of unmodeled cache structures; these
changes do not alter the earlier dispositions. The complete source hashes at
final review time were:

* `crates/litchi-ods/src/model/dde.rs` —
  `98c15134cd0a89fd16e4aa8e2b7fc22aa4f1e10c5189a125f0a9186e3976f4d8`
* `crates/litchi-ods/src/model/dde/lexical.rs` —
  `cc67145d64775ea993797c927639b4c98c100ada47e23df06e6429139686e427`
* `crates/litchi-ods/src/model/dde/transaction.rs` —
  `18683f6ebac6cfdb2c2f48d1faa388f8253d8b2f9a172a7b3b361148afd8b891`

`cargo check --locked --offline -p litchi-ods --tests` completed successfully
with no warnings. The final focused gate passed 38 tests (26 transaction and
12 facade), including reference-preservation and named-cache replacement/refusal cases.

The final integrated transaction source is
`bce59b24064e017d0a3f3cf945594fdd9870e84cdeda513d076ea9c86fc1bc4c`.
Subsequent changes enforce candidate cleanup order and capacity-based admission
and release of owned draft buffers, remove deep edit cloning, and roll back
failed lazy staging. Repeated source-only edits retain the staged declaration.
They do not change the supported XML vocabulary or the specification
dispositions above. The focused resource review covers those changes. Final root
gates pass 770 tests across 45 targets, including 33 transaction and 12 facade
tests; the public authoring replay retains its exact XML hash and passes the
whole-content schema and independent value-decoding checks.
