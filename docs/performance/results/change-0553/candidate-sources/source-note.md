# 0553 XLSX commit-local compact layout candidate

This isolated source snapshot starts from committed `8aa0c5baf0616d16c79eba0c6c28dc1716338ad6` and explores the commit-local route selected by the 0552 deferred-proof feasibility audit. It has no performance claim and is not a change to the live checkout. The root coordinator owns all builds, tests, captures, and any application of this snapshot.

## Ownership and scope

The candidate changes only the source-backed `MultiSourceEdit::commit` owner and the private worksheet edit codec. `Snapshot::from_source_selected` keeps the restored baseline source validation/parser path and retains only the source bytes and authoritative `Store`; no proof or layout is retained in a snapshot, patch, or publication object. The single-snapshot `SourceEdit` path is unchanged.

The commit path first computes effective actions. A no-op skips collection. Collection is attempted only for an action map that can be handled by the existing compact writer: existing cell owners and `Action::Update` payload/style effects, including the existing plain `SetFormula` representation when its target source is otherwise eligible. Insertions, removals, shared-string/shared-formula dependencies, source formulas, unknown cell payloads, and other uncertain structures fall back to the complete scanner/writer.

## Direct collector

`raw::worksheet::edit::compact` walks the exact borrowed worksheet bytes with one bounded `NsReader` immediately before the compact rewrite. It retains only u32 source spans for `sheetData`, rows, cells, and an optional dimension, plus the minimum row/cell envelope metadata consumed by the compact writer. It does not construct the complete scanner `Layout`, a semantic `Store`, owned tags for unchanged cells, or a retained snapshot cache.

The walker resolves explicit and inferred row/cell addresses with the same grid limits and requires each source cell to equal the next authoritative `Store` entry in sorted address order. It binds one transitional or strict SpreadsheetML dialect for every core event, checks strict row/column order, source/store cardinality, root and `sheetData` closure, and span bounds before returning a layout. It rejects formulas, shared/foreign/uncertain cell structures, unsupported worksheet constructs, leaf CDATA where the ordinary parser rejects it, and any malformed or undecodable event; a refusal is `None` and invokes the existing complete rewrite, preserving its diagnostic precedence and exact error text.

All element names and attributes are UTF-8 checked, duplicate-attribute checked by the XML reader, and XML 1.0 decoded/normalized before their values are discarded. Relevant row, dimension, and cell coordinates are parsed with existing checked helpers. This retains the complete scanner's late malformed-attribute boundary: the optional route never publishes a result from an uncertain source. Unsupported namespaces, MCE/x14ac markers, merges, formulas, extension payloads, and non-scalar cell forms remain on the established fallback path.

The source eligibility boundary remains the existing 8 MiB and 131,072 provisional-event limits. Collector metadata is capped at 2 MiB, with fixed representation charged up front and geometric row/cell/stack growth checked and fallible. Every refusal drops scratch state before fallback. The returned layout carries source and `Store` slice pointer/length identities; the writer verifies both before dereferencing spans. It lives only on the commit stack and is dropped before candidate snapshot publication.

The compact writer and reduced-readback provenance are inherited from the prior bounded proof work, including lexical byte copying for unchanged spans, changed-cell lazy tag decoding, dimension union, full output validation, and independent readback. No planner timing, performance gate, or acceptance threshold is changed by this candidate; only the frozen 0553 baseline/candidate comparison can decide whether it is useful.
