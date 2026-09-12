# 0528 XLSX source-backed publication path review

status: bounded read-only source map; no optimization decision

production_change: none

This review maps the multi-worksheet XLSX publication path from
`SourceBackedEditor::publish_multi_commit_to_stream` through
`SourceBackedPackage::write_topology_to_stream` and
`write_changed_overlays_with_omissions_and_appended`. It records the source,
provenance, semantic validation, resource, and physical-preservation fences
that a publication optimization must retain. OLE2/OOXML remains the active
priority; ODF is deferred and iWork is excluded. No Rust source was edited and
no build, test, benchmark, profile, allocator run, or capture was performed.

## Identity and method

| item | identity |
| --- | --- |
| baseline revision | `11a14d2043f65b1a6a935beaa1b9c16d924b749c` (`11a14d204`) |
| 0528 plan SHA-256 | `259454bb139493e82da98ad28c09332e65bcecfb0e1fea165a6b1ee6370f9b1e` |
| 0528 source-binding SHA-256 | `8a133620217daf78a9ed91500ebf551a7db56d341df3ab02127588430935361f` |
| 0528 baseline source manifest SHA-256 | `9af673c4c13f2abb3aeaf5c4e31df6a613de6297abc104bf472931ff7733529f` |
| 0528 baseline source.patch SHA-256 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| workspace `Cargo.toml` SHA-256 | `911a52cf6932b81550bc9ffd6e522c327dec178297ee9bc46d2ccedb693d4885` |
| perf harness `Cargo.toml` SHA-256 | `a04de024b9cbe9683cdb7307c3d7199daab7171bbdcb6bb8c30b361766857aca` |
| perf harness `Cargo.lock` SHA-256 | `13333f511914d8146c60282b5d6385693cf89c0f939fa52b12e4db821c9b8b36` |

The source files that own this route are pinned as follows:

| source file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `2f3f839adc91f0da204aefc83ebe2bb605cc02d269346759155b7bda54abc0e9` |
| `crates/litchi-xlsx/src/cell_values/patch.rs` | `762c222c7fc3580a66c4c1f11f9de44d4a43d4203210597cacdfe34fc17e4648` |
| `crates/litchi-opc/src/source_backed.rs` | `6f37f9a1e2ecd0435cc811b56a3e92d4bbdb0b41bcdf755478d04f2455f8ea48` |
| `crates/litchi-opc/src/lib.rs` | `fd9841c5a48911d750f37910422f45ac57f1a9d5cc2bd32791e40d0859320f0d` |
| `docs/adr/0011-ooxml-physical-package-ownership.md` | inspected; OPC is the physical package owner |
| `docs/adr/0005-io-memory-and-performance.md` | inspected; profiles and preservation require measured evidence |

The perf lock selects `quick-xml` `0.41.0`, registry checksum
`e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1`. That
dependency is relevant to the separate scanner review; this publication route
does not authorize changing the full validator or its namespace/error path.

## Ownership and call path

`SourceBackedPackage` is the physical OPC owner. Its documentation at
`source_backed.rs:5234-5243` says that the ordinary view is immutable and that
the package owns narrow sequential publication and raw copying of unchanged
ZIP members. `SourceTopologyPlan` at `source_backed.rs:159-164` is deliberately
opaque: the semantic caller describes logical Part and relationship changes,
while OPC retains ZIP preservation, content-types lexical preservation, and
relationship member placement. This is the boundary required by ADR 0011.

The route is:

| stage | source and line | owner | result and required fence |
| --- | --- | --- | --- |
| editor entry | `cell_values/source.rs:599-650` | `litchi-xlsx` | Consumes the editor, checks execution, checks exact source provenance, chooses the target snapshot, creates an empty or changed topology plan, and delegates once to OPC. |
| target planning | `cell_values/snapshot.rs:1146-1197` | `litchi-xlsx` | Verifies workbook/worksheet owner identity and selected-owner count, then records changed workbook and worksheet bytes plus calculation-chain relationship/Part operations. |
| topology authorization | `litchi-opc/src/source_backed.rs:6573-7761` | `litchi-opc` | Resolves plan names against the immutable source catalog, validates physical namespace, relationships, content types, signatures, encryption, limits, and source freshness. |
| ZIP preservation | `litchi-opc/src/source_backed.rs:9514-9884` | `litchi-opc` | Builds a preservation plan that regenerates selected members, omits requested members, copies every other entry, appends new entries, and writes through freshness, cancellation, context, budget, and partial-output sinks. |

### Editor entry and provenance

At `source.rs:605-637`, the multi-editor path calls
`self.package.check_execution()` before examining the commit. It asks
`before.matches_source_backed(&self.package)` for `Matched`, `Mismatched`, or
`Unavailable`:

* `Matched` avoids a reload, while the later topology writer still owns source
  freshness during output.
* `Mismatched` returns the typed `Error::PatchConflict` naming the first
  selected worksheet Part.
* `Unavailable` loads the selected worksheet positions with
  `MultiSnapshot::load_source_backed`; the subsequent `same_source` check at
  `source.rs:627-637` still refuses a changed source.

The target at `source.rs:639-643` is the reloaded/current or original snapshot
for an empty patch and the frozen `after` snapshot for a changed patch. An
empty patch receives `SourceTopologyPlan::new()` at `source.rs:644-646`.
Otherwise `MultiSnapshot::topology_plan_from` constructs the plan. Both paths
then call `write_topology_to_stream` exactly once at `source.rs:649`.

### Semantic commit before publication

`MultiSourceEdit::commit` at `source.rs:1125-1232` is part of the contract even
though its output is passed to the publisher later. It bounds the expanded
action count, rewrites each touched worksheet, validates the complete rewritten
XML in `Snapshot::from_rewritten_value_source` (`snapshot.rs:711-729`), and
performs scalar, insertion, clear/remove, and shared-formula readback checks at
`source.rs:1187-1213`. It invalidates the workbook calculation state for a
changed batch at `source.rs:1216-1222`, then freezes the ordered
`MultiSnapshot`/`MultiPatch` pair.

The rewrite path can use the reduced value-only readback only when source spans
and stored entries satisfy its existing proof; otherwise it falls back to the
complete worksheet parser (`snapshot.rs:715-755`). This preserves the current
refusal and fallback behavior for unsupported or opaque cell content. The
semantic snapshot therefore remains the authority for cell values and formulas;
OPC remains the authority for physical package preservation.

### XLSX topology plan

`MultiSnapshot::topology_plan_from` first rejects a changed selected-owner count,
then calls `append_owner_topology` once for the workbook owner and
`append_worksheet_replacement` for every selected worksheet
(`snapshot.rs:1177-1197`). The helpers enforce:

* workbook and worksheet URI/content-type identity;
* copied changed workbook or worksheet bytes through fallible reservations;
* calculation-chain removal as a relationship plus Part removal;
* calculation-chain addition as a Part plus an internal workbook relationship;
* refusal of calculation-chain replacement outside the cell transaction.

The changed worksheet/workbook payloads enter `SourceTopologyPlan` through
`try_replace_part`, so the plan carries ordinary payload bytes rather than an
OPC `SourceXmlPart` proof. That distinction matters below: the semantic layer
has already validated a candidate worksheet, but OPC still applies its own
generic XML, Part-byte, and archive-limit policy before output.

## OPC topology authorization

`write_topology_to_stream` begins by disabling publication read-ahead. At
`source_backed.rs:6578-6589`, an empty plan immediately uses
`write_exact_source`, preserving signed packages and physical details that the
rewrite primitive cannot support. A nonempty plan checks source freshness,
execution context, and encryption before destructuring and sorting all plan
operations (`source_backed.rs:6591-6628`). Source-authorized XML and
precompressed additions are checked before output.

The changed path then performs the following ownership checks:

1. It builds the physical member lookup and requires exactly one content-types
   member (`source_backed.rs:6654-6667`). The lookup folds names for collision
   detection while retaining source member identity.
2. It resolves replacement and removal Part names against the immutable source
   catalog and checks any source-XML replacement proof
   (`source_backed.rs:6669-6725`). Replacement targets cannot be Parts added by
   the same plan.
3. It builds the surviving/new Part namespace and canonicalizes relationship
   owners and internal targets (`source_backed.rs:6727-6855`). Inbound edges to
   removed Parts must be explicitly detached; external targets remain URI
   references.
4. It parses or lexically splices relationship members, validates generated XML,
   counts relationship events, and checks relationship ownership and limits
   (`source_backed.rs:6919-7417`).
5. It reads the content-types member only when additions/removals require it,
   computes required and removed overrides, and validates their generated
   replacement (`source_backed.rs:7495-7596`).
6. It calls `validate_topology_limits` before authored XML is audited at
   `source_backed.rs:7598-7606`. The pass accounts for archive, Part,
   relationship, member-name, event, and replacement-byte ceilings.

Changed existing Parts are compared with their source bytes before XML audit at
`source_backed.rs:7664-7697`. Exact no-op replacements are omitted from the
changed list and therefore remain raw copies. A plain replacement for an XML
Part runs the OPC `validate_overlay_xml` audit on both original and replacement
bytes at `source_backed.rs:7675-7687`. Source-proven replacements use the
separate source-proof contract and retain a late freshness token instead of
repeating the authored compactness audit.

If all operations collapse to exact no-ops, `source_backed.rs:7741-7747`
returns the exact source. A changed publication refuses non-Part/opaque
physical members, encrypted entries, and signed-source mutation at
`source_backed.rs:7749-7761`. Added decoded, source-XML, and precompressed
members are validated and retained through the transfer boundary
(`source_backed.rs:7763-7867`). The late checks are freshness, context, and
cancellation checks; they do not authorize skipping the earlier destination
and payload proof.

## Physical preservation writer

`write_changed_overlays_with_omissions_and_appended` is a thin entry to
`write_changed_overlays_with_appended_inner` (`source_backed.rs:9514-9562`).
The inner owner:

* disables read-ahead, monitors the source, and checks source freshness;
* builds a bounded ZIP preservation index with
  `preservation_index_with_limits` (`source_backed.rs:9563-9600`), refusing
  unsupported preservation layouts and trailing bytes;
* creates sorted replacement and omission lookups, rejects collisions, counts
  physical matches, and requires exactly one canonical member for every
  replacement/removal (`source_backed.rs:9602-9704`);
* checks the conservative output bound, then emits one
  `PreservationAction` per source entry: `Regenerate` for changed members,
  `Omit` for removed members, and `Copy` for all other entries. New entries are
  appended after that source order (`source_backed.rs:9720-9769`); and
* wraps the sequential sink with source checks, execution/cancellation checks,
  output-budget reservations, byte accounting, and ZIP error mapping. The final
  source/context decision and incomplete-output precedence are applied at
  `source_backed.rs:9771-9883`.

This is the physical preservation contract: untouched local/central records
and compressed payloads are copied through the ZIP preservation primitive where
possible; selected replacements are regenerated; omissions and additions are
explicit. The source package remains immutable and output errors are reported
with the existing typed byte-count/freshness precedence.

## Mandatory contracts for any candidate

The path map found no source-level correctness blocker in the existing route.
Any candidate that moves or combines work must preserve all of these observable
contracts:

* exact source lineage/version checks, including the `Unavailable` reload path,
  typed `PatchConflict`, and source-change detection during output;
* execution/cancellation checks before planning, at bounded loops, around
  source reads, during transfer, and before final success;
* semantic worksheet validation, scalar/formula/shared-formula readback,
  calculation-chain invalidation, unsupported-cell refusal, and the current
  unknown/opaque-content preservation or refusal behavior;
* Part URI/content-type identity, workbook graph ownership, relationship
  owner/target validity, content-types overrides, duplicate/case-folded member
  checks, archive and aggregate resource limits, and signed/encrypted policy;
* exact no-op output, including signed sources and physical layouts that changed
  publication cannot rewrite;
* raw preservation of unselected members and unchanged selected members,
  deterministic replacement/omission/addition ordering, sink failure and
  incomplete-output classification, and managed reservation release.

The separate 0528 scanner review remains separate from this path. In particular,
the full value-only validator's End-namespace check is not covered by a
publication change, and a scanner namespace optimization must not be folded
into this physical writer map.

## Possible repeated work: hypotheses only

Static reading identifies overlap that may explain publication profiles, but it
does not establish removable cost or authorize an optimization:

1. A changed XLSX worksheet is validated and semantically parsed while
   `from_rewritten_value_source` creates the candidate and performs readback
   (`snapshot.rs:711-755`, `source.rs:1187-1213`). Because the plan uses a plain
   `try_replace_part`, OPC later reads the source Part and runs its generic XML
   audit on the original and replacement (`source_backed.rs:7664-7687`). These
   validators have different owners and limits; proof reuse would need an
   explicit equal-policy, source-bound contract and must retain error ordering.
2. `write_topology_to_stream` first calls `validate_overlay_limits` for pending
   replacement bytes and current totals (`source_backed.rs:6721-6725`,
   `8704-8782`), then calls `validate_topology_limits`, which rechecks
   replacement limits while adding the full topology totals and relationship
   events (`source_backed.rs:7598-7606`, `8991-9225`). The second pass is
   topology-aware; merging or narrowing either pass could change typed resource
   precedence, so no removal is approved by this audit.
3. The physical writer scans source metadata more than once: the physical lookup
   sizes names and then populates its folded map (`source_backed.rs:9394-9442`),
   topology limits walk archive names and Part metadata
   (`source_backed.rs:9009-9072`), and the preservation writer indexes and
   emits source entries (`source_backed.rs:9574-9761`). Within the final writer,
   one entry loop counts replacement/omission matches before a second loop emits
   preservation actions (`source_backed.rs:9666-9688`, `9731-9761`). Combining
   these scans would require preserving fail-closed collision checks, limit
   order, source fences, and the no-op fast path; it is an unmeasured hypothesis.
4. When provenance is unavailable, the editor reloads selected snapshots before
   plan construction (`source.rs:619-625`), after which OPC still reads source
   metadata and changed Parts during authorization. That reload is conditional
   correctness work, not evidence that the source check can be removed.

These observations are path ownership and review targets only. They are not a
performance claim, a recommendation to remove validation, or approval for a
candidate patch. The rejected 0527 row-arena candidate remains outside this
review.

## Bounded follow-up: slice XML attribute scratch

The completed publication profile puts `validate_overlay_xml` on the
`xml_minifier::audit::verify_with_policy` edge. Its `inspect_attributes` helper
is a separate bounded follow-up target; this review does not authorize skipping
the OPC audit of either the original or replacement XML.

The exact follow-up source identity is:

| source/dependency input | SHA-256 |
| --- | --- |
| `crates/xml-minifier/src/audit.rs` | `40c7dfea199aa8f2e1951d0d75cd9a47e84f13bf49f2abcd1232230b64b355ac` |
| `crates/xml-minifier/Cargo.toml` | `7cdf4fac21a50ce8c39fc20d0af89c3ca24f869b74d95c6cd449249f0e28b797` |
| `Cargo.lock` | `9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a` |
| quick-xml `src/events/attributes.rs` | `c46448f11d7dba312e6ad2177dc31deab2ce51282e9f1e4069d3b35175c18518` |
| quick-xml `src/events/mod.rs` | `b5e38bbfc0d87b2fa49b2dedcdb1e14d42064e7d41049c1b568d6f86f7731abb` |

The pinned dependency is quick-xml `0.41.0`, registry checksum
`e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1`.

At `audit.rs:1300-1440`, the slice auditor creates `Reader::from_str`, reads a
borrowed event, and calls `inspect_attributes` only for `Start` and `Empty`.
At `audit.rs:738-956`, the streaming auditor uses a reusable parser event
buffer and guarded source capture, but calls the same `inspect_attributes`
function for its `Start` and `Empty` events. The shared helper at
`audit.rs:1504-1540` calls `tag.attributes()` for every attribute-bearing tag.

The inspected quick-xml source does not currently expose a reusable attribute
checker. `BytesStart::attributes` at `events/mod.rs:286-289` constructs a fresh
`Attributes` through `Attributes::wrap`. That creates a fresh `IterState` at
`events/attributes.rs:517-525` and `1040-1048`, with an empty `keys` vector and
no hash set. Its default duplicate check records each key range at
`events/attributes.rs:1135-1153`; once the tag crosses the 32-attribute
threshold, it allocates and seeds `key_hashes` at `events/attributes.rs:1156-1191`.
Therefore, the current streaming path reuses its event/capture buffers, but it
does not reuse quick-xml's per-tag duplicate-key scratch. The profile's
`quick_xml::events::attributes::IterState` edge is consistent with this source
shape. Any claimed reusable checker would be a new implementation or an
upstream API change, not reuse of an existing `audit.rs` object.

A feasible shape for a bounded candidate is to give the slice auditor a
state-owned per-tag key/range and hash scratch, call quick-xml's iterator with
duplicate checks disabled, and perform an equivalent check at the same lexical
point. That could retain capacity between tags and avoid the per-tag key-vector
allocation. It has two blockers that prevent approval in this audit:

1. `IterState::check_for_duplicates` runs after parsing the key but before
   parsing that attribute's value (`attributes.rs:1287-1293`). Simply calling
   `with_checks(false)` and checking after `Attribute` is yielded changes error
   precedence: a duplicate whose value is unquoted, unterminated, or otherwise
   malformed can report the value error before the duplicate. A replacement
   checker must detect the duplicate before value parsing, retain raw QName byte
   equality, and preserve the first duplicate's relative positions in the
   bounded `Malformed` detail.
2. `inspect_attributes` increments the aggregate `Attributes` resource only
   after quick-xml yields a valid attribute (`audit.rs:1513-1522`), then decodes
   `xml:space` and maps normalization/invalid-value failures
   (`audit.rs:1523-1537`). The candidate must preserve that ordering, the
   `check_start` lexical-layout check before it, and the existing
   `Error::Malformed`, `Error::Limit`, `Error::Allocation`, and authored/slice
   versus streaming error boundaries. Any retained scratch must be bounded by
   the per-event token/attribute policy, use the correct fallible reservation
   behavior, and not turn aggregate attribute acceptance into a per-event cap.

An upstream resettable `Attributes`/`IterState` API that retains the exact
quick-xml duplicate-check state would lower the first risk, but changing that
dependency is outside this source-only review. A raw pre-scan or a custom
attribute lexer could reproduce the timing, but it would duplicate quick-xml's
attribute grammar and malformed-error precedence. Neither approach is a
concrete accepted optimization here. The current XML audit remains the
correctness oracle.

Required focused evidence for a future candidate includes valid and authored
slice/stream report parity; duplicate attributes below and above the 32-name
hash threshold; duplicate-before-malformed-value precedence and error detail;
missing equals, quote, and separator errors; exact aggregate and token
attribute-limit boundaries; `xml:space` decoding, normalization, and invalid
values; arbitrary stream chunking and BOM/error offsets; and the existing
streaming memory envelope. An allocation/profile comparison must separate the
slice and streaming entry points and report retained scratch capacity. Until
that evidence exists, the profile edge identifies a hypothesis only, with no
removable-CPU, allocation-count, or throughput claim.

### Candidate shape: zero/one-attribute probe

The least invasive candidate is a common-helper fast path, still unaccepted:
create `tag.attributes().with_checks(false)`, inspect at most two iterator
results without touching `State` or decoding `xml:space`, and process the one
attribute with the existing attribute-processing logic only when the first
result is `Ok` and the second is `None`. For no attributes, return the
inherited `Space`. If the first result is an error, or the second result is
either `Ok` or an error, explicitly drop the probe and replay a fresh default
checked `tag.attributes()` iterator from the beginning.

The pinned quick-xml implementation supports the safety proof. `with_checks`
only changes `IterState::check_duplicates` (`events/attributes.rs:580-589`),
while key/value syntax parsing remains in `IterState::next`
(`events/attributes.rs:1222-1358`). With one successfully parsed attribute,
there is no possible duplicate; the unchecked first/second probe has the same
syntax result as the checked iterator. A second valid result identifies a
possible duplicate, and a second error may be a malformed value after a first
valid key. Replaying from the beginning is therefore required: the checked
iterator tests duplicate names before parsing the duplicate value
(`events/attributes.rs:1287-1293`), preserving duplicate-before-malformed-value
precedence and the existing bounded `Malformed` detail. The replay also keeps
the current aggregate attribute increment before `xml:space` decoding
(`audit.rs:1513-1537`) and leaves `check_start`'s earlier layout errors intact.

The probe must be dropped before checked replay or single-attribute decoding so
its iterator state cannot coexist with the checked scratch. The existing
`Attributes` aggregate/token limits, fallible audit-state reservations, typed
error mapping, and streaming `StreamError` boundary remain unchanged. Because
`inspect_attributes` is shared by slice and streaming verification
(`audit.rs:831-860`, `1326-1357`), a helper-level patch affects both entry
points; slice-only attribution would require a separate helper and its own
parity proof. This shape removes the quick-xml key-vector allocation only for
the proven zero/one-attribute case and intentionally pays the probe plus
replay cost for tags with two or more results. It is a fresh measurement
candidate, with no current adoption or ROI claim.

## Required tests before any publication candidate

The existing route should remain the oracle for a candidate comparison. At
minimum, focused tests and full all-feature checks would need to cover:

* changed and exact-no-op multi-sheet scalar, numeric insertion, clear/remove,
  formula, and shared-formula edits, including workbook calculation-chain
  invalidation and readback;
* unknown/interleaved worksheet XML, unsupported cell metadata, unselected
  sheets, worksheet relationships, and complete graph/content-type retention;
* stale, foreign, and unavailable provenance; source mutation during planning,
  transfer, and flush; cancellation; failing sinks; partial-output byte counts;
* malformed source/replacement XML, duplicate or case-equivalent physical
  members, relationship collisions/dangling targets, content-types conflicts,
  archive/Part/relationship/event limits, encrypted input, signed no-op, and
  signed changed-source refusals; and
* byte-for-byte no-op and untouched-member local/central preservation plus
  candidate reopen and semantic readback.

No correctness or ROI conclusion beyond this bounded source map is made until a
candidate has exact source binding, focused tests, and representative OLE2/
OOXML measurements under the accepted performance protocol.
