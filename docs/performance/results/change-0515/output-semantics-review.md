# XLSX changed-output validation and compaction review

This is a bounded, read-only design review for the OLE2/OOXML priority lane. It
records source evidence and a possible work-elimination seam; it makes no
performance claim and does not propose an ODF change. The measured numbers below
are prior evidence, not measurements from this review.

## Current output path

For an effective worksheet edit, `WorkbookEdit::commit` rewrites the source
worksheet, then compacts the rewritten XML, then validates the compacted bytes.
The relevant sequence is in
`crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs:1560-1840`:

1. `raw::worksheet::edit::rewrite` and the merge/web/metadata rewrite helpers
   construct `after` while retaining source spans where possible.
2. `raw::compact::changed_worksheet` produces the actual compacted bytes.
3. When the edit changes grid content, `raw::worksheet::parse` parses those
   exact compacted bytes.
4. `WorksheetOutput::into_bytes_and_web` finishes the web proof against those
   same bytes. An ineligible proof calls the authoritative
   `raw::web::read` on the output.
5. If a store was parsed, `validate_styles` checks its style indexes against
   the style catalog.
6. Every requested `Change` is checked against the parsed output state. The
   bounded parsed-store handoff is only used at 4,096 cells and 1 MiB.
7. The package is cloned and reopened through
   `Workbook::from_package_with_styles`, which performs the OPC/workbook graph
   checks before publication. This is the package-level validation boundary,
   not a ZIP-signature check.

Metadata-only worksheet changes deliberately set
`requires_store_verification` false, but their output still goes through the
output web check and the package reopen. That distinction limits the scope of a
grid-parser optimization candidate.

The direct attribution profile in
`docs/performance/changes/0512-xlsx-commit-attribution.md` placed source
worksheet storage at 25.80%, worksheet rewrite at 27.18%, changed worksheet
validation parse at 25.77%, and changed XML compaction at 20.38% of the
commit-only sample. These four **direct commit children are disjoint** and may
be added for that sample (rounding leaves the total near 99.13%). Nested
diagnostic rows, such as the inclusive `raw::worksheet::parse` and scanner
shares, overlap and must not be added. The follow-up in
`docs/performance/changes/0514-xlsx-fusion-rejection.md` rejected a broader
semantic-parser/layout fusion: it passed the available correctness tests but
regressed cold same-value no-op p50 by 64.43--84.64% and increased incremental
peak memory by 18.64--27.04%. The current source contains no such fusion.

## Semantic invariants that the output proof must preserve

The final artifact is the compacted byte stream, so the proof has to be about
that stream and its publication path. The following invariants are binding:

- **No-op identity.** An empty or semantically ineffective edit returns the
  original source/patch identity, preserves source bytes, touches no component,
  and does not pay changed-output validation. This is implemented at
  `transaction.rs:1023-1035` and `1485-1558`. Any candidate must retain the
  cold same-value guard that rejected the 0514 fusion.
- **XML and SpreadsheetML structure.** The output must have one valid
  SpreadsheetML worksheet root, legal namespace resolution, matching and
  bounded nesting, valid child order, no duplicate structural records, and
  legal dimensions, rows, cells, columns, merges, formulas, and metadata. The
  complete structural/state checks are in
  `crates/litchi-xlsx/src/raw/worksheet/codec.rs:265-403`,
  `405-557`, `736-841`, and `900-1052`.
- **Typed cell and style semantics.** Coordinates, row monotonicity, formula
  groups, value/formula/inline-text exclusivity, encoded-size limits, metadata
  indexes, and style indexes must retain the parser's error behavior. Style
  validation is in `transaction.rs:1669-1681` and
  `crates/litchi-xlsx/src/workbook/model.rs:1512-1559`.
- **XML text semantics.** `xml:space`, text-bearing element policy, whitespace
  removal, CDATA, entity/general-reference decoding, and escaped text must
  result in the same typed values and limits. The compactor's policy is in
  `crates/litchi-xlsx/src/raw/compact.rs:41-137`; the parser's text and general
  reference handling is in `codec.rs:900-945`.
- **Namespaces, MCE, and unknown content.** Qualified names and namespace
  aliases must resolve identically; supported markup-compatibility and web
  extension cases must keep their required fallback behavior; unknown content
  and lexical source spans that are preserved by the rewrite must remain
  preserved. The snapshot scanner's conservative boundaries are in
  `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs:200-317`,
  `1197-1329`, and `1333-1495`.
- **Web binding behavior.** The empty-binding shortcut is valid only when the
  `Probe` proves the exact output. Otherwise the full web reader remains the
  authority, including its 16 MiB bound, extension URI/name rules, required
  attributes, formula cardinality, and malformed-input errors. See
  `crates/litchi-xlsx/src/raw/web/check.rs:18-23`, `51-85`, and `88-228`, and
  `crates/litchi-xlsx/src/raw/web.rs:25-238`.
- **Error ordering and publication.** Compaction errors precede grid parsing
  errors, which precede web validation errors; style validation follows web
  finishing, and requested-change checks follow style validation. The tests
  encode the first three phases at `compact.rs:488-562`, while the transaction
  sequence is explicit at `transaction.rs:1669-1823`. Requested changes are
  checked against parsed output before package publication, and the final OPC/
  workbook reopen is retained. A candidate cannot move a check across those
  boundaries merely because valid inputs produce equivalent values.
- **Exact preservation and ownership.** Source-backed untouched parts and
  unknown markup remain the preservation authority. A generated output is
  validated as a fresh artifact and is reopened through `litchi-opc`, consistent
  with ADR 0003, ADR 0005, ADR 0006, ADR 0011, and ADR 0024. An optimization
  must not turn an output semantic check into a source snapshot check.

## Why the compacted bytes need their own proof

`raw::compact::changed_observed` is a transforming pass, not a transparent
event tap (`compact.rs:41-118`). It can drop formatting-only text according to
the inherited `xml:space` state and the local-name text-bearing list, rebuild
start and empty tags with `BytesStart`, regenerate end tags, and let the XML
writer normalize quoting. The `BytesStart` attributes are passed as raw values
to the writer (`compact.rs:104-118`), so a proof must also account for the
writer's escaping/serialization behavior and reject an unsafe output. It writes
borrowed events where possible, but the event passed to the observer is still
the source `Event` after the write; it is not an object produced by re-reading
or byte-inspecting the writer's output. The output therefore can differ
lexically and in its event sequence even when its intended SpreadsheetML value
is unchanged.

The current single-quote test makes this boundary concrete. In
`crates/litchi-xlsx/src/raw/web/check.rs:88-92`, `Probe::observe_start` marks a
start event ineligible if `attributes_raw()` contains an apostrophe. The test
`crates/litchi-xlsx/src/raw/compact.rs:592-601` feeds an attribute written with
single quotes and confirms that `changed_worksheet` marks the output
ineligible, so `into_bytes_and_web` falls back to the full web reader. The
writer can wrap the raw value in double quotes, while the source value may still
contain an unescaped literal double quote; the resulting bytes must be treated
as potentially malformed until an output-aware check proves otherwise. A
source-backed observer has only seen the original raw quote bytes. The fallback
is a conservative false negative and is correct; treating the source event as
an emitted event would be an unsound proof of the output. The same distinction
applies to dropped whitespace, escaped entities,
CDATA/general references, namespace spelling, and malformed tails.

This is also why parsing `after` before compaction is insufficient. The parser
would be checking a different event stream from the one that is eventually
published. Even if valid inputs usually decode to the same cells, compaction can
fail, alter whitespace, normalize attributes, or expose a different malformed
construct. Error phase/order, original-offset reporting, and limits would then
become observable regressions.

## Recommended seam and proposed parser-feed design

A coherent candidate is an **output-aware parser feed at the compactor's writer
boundary**, started only after an edit is known to be effective. For each event
that the writer actually emits, it would feed the existing worksheet parser an
equivalent normalized event; dropped whitespace would produce no parser event.
The parser result would remain provisional until compaction reaches EOF. Then
the unchanged x14ac capture and MCE preprocessing would run on the exact output
bytes. Reuse is allowed only if those checks succeed, the MCE result is
`Cow::Borrowed`, output UTF-8 holds, each existing limit remains enforced at its
current owner/boundary, and the output-aware event feed has proved
serialization safety and event equivalence. On any failed gate, discard the
provisional result and run the current parser on the exact compacted bytes.

This closes the source/output quote mismatch only if the callback is driven by
the writer's actual normalized name, attribute, escaping, and omission
decisions. A callback that merely copies decoded source values is insufficient:
the current writer stores raw attribute values, and the single-quoted
`marker='contains "quotes"'` fixture can become malformed when wrapped in
double quotes. The callback must reject that case (or the writer must escape it)
before a provisional parser result can be reused. A safer first slice is
MCE-free, x14ac-free ordinary output with proven-safe attributes; all other
outputs fall back to the exact-byte parser.

The post-compaction MCE gate is necessary but not sufficient. The common
`process_ooxml` fast path returns borrowed input when no MCE namespace is found
without parsing the XML (`crates/litchi-ooxml-common/src/mce/codec.rs:530-548`),
so it cannot by itself detect malformed normalized attributes. The output event
feed must preserve the current XML well-formedness and UTF-8 checks, and retain
each existing limit at its owner (MCE input/output/event limits when that layer
is active, worksheet depth/value/cell limits in the parser, and the web 16 MiB
bound), or a reader over the exact bytes must remain in the path. The ordinary
worksheet parser has a depth bound and payload/cell limits but no general event
counter; the candidate must not invent one as a parity requirement. x14ac is
another ordering constraint:
`raw::worksheet::parse` receives captured extension values before parsing and
consumes them while opening defaults and rows (`codec.rs:265-275`, `559-657`).
Feeding the parser before a later output capture therefore requires either a
concurrent x14ac observer plus deferred materialization, or a narrowly gated
path with no x14ac values. It cannot be claimed as an unchanged parser API by
just attaching the extension map after parsing.

The seam does not justify removing style validation, requested-change checks,
or the package reopen. It also does not automatically replace the exact-byte
web fallback. The current 0512 attribution already profiles the available
rewrite, compaction, grid-parse, and source-store boundaries; there is no
provisional parser path to profile in this batch. If a candidate is implemented,
its own run should compare those current boundaries with its output audit,
post-output x14ac/MCE checks, fallback parse, and package reopen. Retain it only
if it removes measurable changed-output work without regressing cold no-ops,
peak RSS, output bytes, or validation behavior. The 0512 percentages identify
the lane; they do not establish that parser/event fusion is profitable.

## Required proof and regression matrix

Any implementation should differentially compare the current reference path
against the candidate for exact compacted bytes, parsed worksheet state, web
bindings, requested-change verification, phase/error identity, and publication
success. At minimum, include:

- single-quoted attributes, double-quoted attributes containing apostrophes,
  escaped quotes, and attribute normalization;
- `xml:space` inheritance, dropped formatting whitespace, text-bearing tags,
  CDATA, entities/general references, comments, processing instructions, and
  namespace aliases;
- supported and unsupported MCE/web extension forms, unknown qualified
  attributes, malformed names/tails, duplicate/ordering violations, and the
  16 MiB web limit;
- styles, row/column metadata, shared formulas, merges, inline strings,
  formulas/values, and the 4,096-cell/1 MiB validated-store handoff;
- competing compaction/grid/web failures to preserve `compact > grid > web`
  precedence; and
- ordinary edits, metadata-only edits, exact no-ops, and semantically
  ineffective actions to preserve source identity and untouched-part bytes.

No output-fusion performance claim is authorized until those comparisons pass
and operation-local measurements show a net benefit. ODF work remains deferred
until the OLE2/OOXML optimization goal is complete.

## Source and decision references

- `docs/GOAL.md` — work-elimination-first ordering; changed-output validation;
  correctness/preservation/safety gates; no claim without reproducible evidence.
- `docs/adr/0003-snapshots-edits-and-patches.md` — exact no-op identity,
  changed dependency-closure validation, source-backed preservation, and
  publication rules.
- `docs/adr/0005-io-memory-and-performance.md` — lazy semantic payloads,
  output validation/reopen, bounded caches, and memory/performance gates.
- `docs/adr/0006-validation-security-and-compatibility.md` — preservation
  default, deterministic validation, malformed-known-payload refusal, and
  structured error behavior.
- `docs/adr/0011-ooxml-physical-package-ownership.md` — OPC package ownership
  and reopen through `litchi-opc`.
- `docs/adr/0017-ooxml-producer-template-ownership.md`,
  `docs/adr/0018-xlsx-calculation-chain-ownership.md`, and
  `docs/adr/0024-current-topology.md` — format-owned output/validation and
  current XLSX topology.
- `docs/performance/changes/0469-xlsx-borrowed-compaction-events.md` and
  `docs/performance/changes/0470-xlsx-empty-web-proof.md` — existing compact
  event borrowing and exact-output web proof boundary.
- `docs/performance/changes/0512-xlsx-commit-attribution.md`,
  `docs/performance/changes/0514-xlsx-fusion-rejection.md`, and
  `docs/performance/results/change-0514/follow-up-options.md` — prior
  attribution, rejected broad fusion, and the pending changed-output lane.
