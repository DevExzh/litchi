# XLSX cell-values validation/parser boundary review

`change-0540` is a read-only source audit against revision
`49ad30bcaf8d24fce024896ff355b04c786fd061`. It makes no production change,
build, or performance claim. The active optimization lane is OLE2/OOXML;
ODF remains deferred.

The target is the selected worksheet path used by
`SourceBackedEditor::edit_sheets`:

```text
edit_sheets
  -> MultiSnapshot::load_source_backed
  -> load_source_catalog
  -> Snapshot::from_source_selected (once per selected worksheet)
       -> validation::worksheet_xml
       -> raw::worksheet::parse
```

The relevant source is [`source.rs`](../../../../crates/litchi-xlsx/src/cell_values/source.rs#L469),
[`snapshot.rs`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L179),
and [`snapshot.rs`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L426).
The boundary is a good attribution target because the validator and parser
both consume a complete worksheet, but they do different jobs and have
different error policies.

## Current execution and error order

`MultiSnapshot::load_source_backed` first checks execution and source version,
collects and bounds selectors, loads the workbook/catalog capture, resolves
sheet positions, and then loads each selected worksheet. For each selected
sheet, `from_source_selected` checks execution, worksheet kind, and worksheet
relationships; obtains the source-backed part data; charges the aggregate
worksheet-byte bound; and then runs the following sequence:

```text
source/package and aggregate-bound checks
  -> value-only worksheet XML validation
  -> raw worksheet parse and semantic materialization
  -> execution check
  -> shared-style reference validation
  -> scalar/metadata/unknown-cell closure validation
  -> Snapshot construction and later transaction work
```

The source payload is retained for the snapshot, so a candidate must preserve
both the exact bytes and the source identity/lifetime represented by
`SourcePayload`. The final source-version fence in `load_source_backed` also
remains outside any worksheet reader fusion.

The first-error ownership is as follows.

| Boundary | Current owner and first errors | Constraint on reuse |
| --- | --- | --- |
| Package/catalog | Execution cancellation, source/version, OPC part and relationship, workbook content type/XML, sheet graph, and selector errors | These checks stay before selected worksheet validation. A worksheet reader must not hide a catalog or source error. |
| Worksheet validation | XML read errors; DTD refusal; unbound, foreign, or mixed SpreadsheetML namespaces; root/depth/parent/closing errors; unsupported worksheet elements; disallowed, qualified, duplicate, or malformed attributes; text/CDATA/general references outside `f`, `v`, or `t` | This complete validation pass owns the error boundary before raw parsing. The error text and order are observable. |
| Raw worksheet parse | x14ac capture, MCE processing, UTF-8 conversion, row/column/cell structure, coordinate/style/metadata/type, formula/value and duplicate-cell errors, shared-string/formula resolution, and semantic materialization | Raw errors are reachable only after validation succeeds. Parser state is temporary until the whole parse and materialization complete. |
| Post-parse closure | Style indexes outside the captured styles table, then `Unknown` cells or cell/value metadata rejected by the value-only closure | These checks currently run after `Store` construction in a fixed style-before-scalar order. They cannot be moved earlier without preserving that order. |
| Snapshot/transaction | Allocation, aggregate limits, source fences, rewrite/readback/publication and action-projection errors | A deferred scanner or parser diagnostic cannot cross these existing phase boundaries. |

The validator is intentionally conservative. It permits only the supported
SpreadsheetML worksheet grammar (`worksheet`, dimensions/views, format/cols,
sheet data, rows, cells, formulas, values, and inline text), binds every
element to one transitional or strict SpreadsheetML dialect, checks the
parent stack, and rejects dependency-bearing markup such as `mergeCells`.
This is why [`merge_markup_is_refused_at_source_open`](../../../../crates/litchi-xlsx/tests/source_backed_cell_values.rs#L891)
reports the value-only validation error even though the general raw worksheet
model can parse merges. Validation does not establish coordinates, row/column
ordering, formula semantics, cell-type/value consistency, style-table bounds,
metadata closure, or package relationships; those remain raw or post-parse
owners.

Some details matter when comparing error order. `bind_dialect` runs before
worksheet grammar and attribute checks for each start/empty element. End tags
must match both the local name and the initially bound dialect. DTDs are
rejected, while declarations, comments, and processing instructions are
currently ignored by this validator (`Event::Decl`, `Event::Comment`, and
`Event::PI`); a review or candidate must not describe PI as a current
validation rejection. Whitespace text is accepted generally, and non-whitespace
text or references are accepted only in scalar contexts. These choices must
remain exact.

After validation, `raw::worksheet::parse` runs an x14ac descent preflight. If
the byte marker `dyDescent` is present, x14ac capture happens before MCE
processing and can fail before the semantic parser. On the ordinary no-marker
path, a parser failure triggers the historical x14ac capture retry; if that
retry fails, its x14ac error replaces the parser error. The parser then runs
`process_ooxml`, converts the selected bytes to UTF-8, and materializes the
worksheet. MCE may return an owned transformed buffer, while a no-MCE result
can borrow the original bytes. This distinction is an input identity and
source-offset boundary, not merely a content-equality optimization. A fused
driver must also defer any x14ac or MCE preprocessing failure until the
complete validation pass has finished, because validation currently owns the
whole preceding error boundary. Once validation succeeds, it must retain the
existing raw order: x14ac before MCE, MCE before UTF-8/parser work, and the
plain-path x14ac retry after a parser error. A candidate fallback or retry must
discard candidate-local diagnostics and preserve the original raw error
surface, including the current rule that a failing retry replaces the parser
error while a successful retry leaves the original parser error intact.

Within the raw parser, `scan_cell_attributes` performs one checked iteration
over the complete attribute list, decodes `r` at encounter time, and then
`start_cell` applies coordinate, style, cell-metadata, value-metadata, and
cell-type checks in that order. The fixed error behavior is covered by the
0539 tests in [`raw/worksheet/tests.rs`](../../../../crates/litchi-xlsx/src/raw/worksheet/tests.rs#L513),
including duplicate unknown attributes, entity-decoding failures, invalid
references, style/index errors, and metadata bounds. A shared event driver
must not replace these distinct attribute policies with a common decoded map.

## Attribution and duplicated work

The historical 0530 planning profiles are the available cost reference. For
the selected `edit_sheets` owner, the medium profile attributed 75,518,721 Ir
(60.1%) to raw worksheet parsing and 46,622,053 Ir (37.1%) to value-only XML
validation; dense-sparse attributed 141,675,512 Ir (59.8%) and 89,730,098 Ir
(37.9%) respectively. The 0530 aggregate table gives 434,424,105 Ir for raw
parsing and 272,706,313 Ir for validation across the profiles, with
`PartView::data` at 15,467,269 Ir (2.13%) and catalog loading at 1,814,338 Ir
(0.25%). MCE preprocessing, 37,572,982 Ir (5.18%), is nested inside raw
parsing and must not be added to it as a disjoint stage.

The direct edges in the historical medium validation/parser loops also show
the same XML machinery being paid twice: approximately 13.7 million Ir in
`read_event_impl`, 5.3 million in event processing, and 5.4 million in
namespace `resolve_event` per loop. The fresh 0540 capture is attribution-only
and must establish the exact current symbol edges before any candidate is
accepted. These counts demonstrate repeated traversal work; they do not prove
that the work can be removed or that native latency will improve.

## Why a simple one-pass fusion is unsafe

The central obstacle is whole-pass validation precedence. The current code
finishes `validation::worksheet_xml` before it starts x14ac/MCE preprocessing
or the raw parser. If a raw parser-invalid cell occurs near the beginning and a
validator-invalid element occurs near the end, the validator error wins today.
A parser or preprocessing observer that returns its error immediately would
change that result. A one-reader design would therefore need to keep parser
and preprocessing state provisional, record any x14ac/MCE/parser error without
exposing it, finish validation to EOF, and select a validation error whenever
one exists. If the validator fails, all parser/preprocessing state and any
derived layout must be discarded. If validation succeeds, the saved
preprocessing error keeps its existing precedence over parser work; a saved
parser error can be returned only after validation and the existing x14ac retry
rule. This deferred state is the minimum needed for validation-first
semantics.

There are further independent mismatches:

* The validator rejects foreign or mixed element namespaces and unsupported
  grammar. The raw parser has broader namespace/unknown-markup behavior (the
  raw tests include unknown namespace cases), so raw acceptance cannot serve
  as proof that the value-only closure is safe.
* The validator reads the original source. The raw parser reads the output of
  x14ac/MCE preprocessing, which can be an owned transformed `Cow`; event
  lifetimes, source spans, namespace scope, and error locations then refer to
  different buffers. A source event cannot be reused for an MCE-transformed
  parser input without an explicit identity proof.
* x14ac observes raw source events before MCE selection. Its stream join point,
  `capture_stream_with_active`, intentionally gives the raw observer and the
  active observer different views and typed error precedence. Running it in a
  fused worksheet parser would need to defer a raw x14ac error long enough for
  the complete validator to establish whether a validation error should win.
  On the plain no-marker path, the historical capture retry still has to run
  after a parser failure and can replace that error.
* The two consumers do not have one attribute contract. The validator checks
  all permitted XML attributes and namespace forms; the raw parser ignores
  some prefixed attributes, decodes unqualified `r` immediately, and defers
  fixed fields in the exact coordinate/style/metadata order. Shared borrowed
  events are viable only if each observer keeps its own attribute loop and
  decoder behavior.
* The parser mutates pending cells, rows, formulas, shared-formula maps, and
  materialization buffers before it returns a `Store`. Validation has no
  publication side effect. Any fused state must be transaction-local and
  unpublished until validation, semantic parsing, style checks, and scalar
  closure all succeed. Partial cells or layout spans cannot be retained on a
  failed validation or parser path.
* Shared strings and shared formulas need complete worksheet knowledge and
  may allocate after the event loop. A parser observer that reaches EOF is not
  yet equivalent to `raw::worksheet::parse`; deferred resolution,
  materialization, checked reservations, and their errors still need to finish
  before a validation or scanner error can escape.
* Style validation depends on the catalog's `style_count`, and scalar closure
  depends on the completed `Store`. Neither result is available from the
  validator's XML grammar state alone. Execution checks, aggregate worksheet
  limits, source-version fences, and allocation handling also remain required.
* A token buffer can preserve validation-first order by reading once and
  replaying owned events to the parser, but it adds an allocation for nearly
  the complete worksheet and must preserve raw names, namespace declarations,
  attribute spelling, entity text, event kinds, and source identity. It is
  unlikely to beat two direct passes without fresh evidence and has a larger
  peak-memory risk.

The existing selected stream scanner is useful evidence that a shared driver
can be built: [`raw/worksheet/selected.rs`](../../../../crates/litchi-xlsx/src/raw/worksheet/selected.rs#L195)
uses `capture_stream_with_active` and publishes only after EOF. It is a
first-slice eligibility scanner, however. It returns `NotEligible` for
unsupported semantic structures and falls back to the materialized parser;
it does not produce the complete `Store` or the cell-values validation proof.

## Feasible candidates

The candidates below are scoped as designs for a later measured pilot. None is
approved by this audit.

### A. Private proof-bearing one-reader path for the strict ordinary subset

Of the designs considered here, this candidate is the one that could remove
the duplicate XML tokenization in the selected load path. Gate it
conservatively to a source for which
`process_ooxml` is proven to return the original bytes (pointer and length,
not just equal contents), with no MCE transformation, no x14ac descent/branch,
and no unsupported namespace or dependency-bearing markup. Keep the existing
path as the fallback for every other input.

This is a future design sketch only. It is unadmitted, and this audit does not
authorize a production implementation.

The private driver would use one borrowed `NsReader` and two observers. The
validation observer runs first and retains its dialect/grammar/attribute stack.
The raw parser observer receives the same borrowed event only after the
validation observer accepts that event. Raw parser errors are saved and stop
the parser observer, but do not stop validation; validation must reach its
normal EOF/root checks so a later validation error still wins. A validator
error makes the provisional parser unusable. At EOF, the driver selects a
validator error first, then a saved parser error; only when neither exists does
it run the existing post-parse materialization/style/scalar phases. No `Store`,
source payload handoff, or layout is published until all existing phases
succeed.

This design must also preserve `read_event`/resolver scope, event kinds,
checked closing names, decoder errors, and the raw parser's complete attribute
iteration. It cannot use the validator's boxed local-name stack as a parser
name cache unless the lifetime and resolver proof is explicit. Parser-owned
cell strings and formula state must remain owned before the event is read
again.

The likely benefit is removing one `NsReader` event loop and one namespace
resolution pass for validator-safe, MCE-free worksheets. It will not remove
cell materialization, shared-string/formula work, x14ac processing on eligible
inputs, style/scalar closure, or source/budget fences. Invalid validator input
may get extra speculative parser work before the validator reaches its error;
that cost needs a separate gate. A parser error that occurs early also leaves
the reader running only for validation, so the candidate needs a bounded
failure-state representation and must avoid growing parser collections after
failure.

### B. Token or normalized-event handoff

Read and validate the original source once, retaining an owned event/token
sequence for a later raw parser invocation. This makes validation-first order
straightforward and can share XML tokenization, but the token sequence must
retain event kinds, raw attribute forms, namespace declarations/rebindings,
text/reference distinctions, and source bytes needed by the lossless source
payload. It adds full-input buffering and likely overlaps the parser's own
cell/formula allocations. It is a fallback design for cases where the
two-observer state machine cannot preserve errors, not a preferred candidate.
It should be rejected unless fresh RSS/allocation evidence shows that the
replay buffer is cheaper than the current two-reader path.

### C. Validated proof envelope without fusion

Keep the two passes but have validation return a private proof containing the
dialect, complete-root/closing result, and the exact source identity. The raw
parser may consume that proof to skip only checks that are demonstrably
identical. This is low risk if it changes no parser acceptance or errors, but
it does not remove the validation traversal and therefore has a small ceiling.
It is useful as a preparatory refactor for A, not as a performance candidate
on historical percentages alone. A proof must be invalidated for an owned MCE
buffer, a different worksheet `Arc`, source-version changes, or any rewrite.

### D. Fold the final style/scalar scans into raw materialization

This is adjacent to, rather than inside, the validation/raw boundary. During
materialization, the parser already visits cell, row, and column records. It
could collect the same offending-style and forbidden-scalar witnesses and
report them after parsing, preserving the current `style references` before
`scalar cells` order. The implementation must reproduce the current
`Store`-iteration ordering, including row/column style checks, and must not
surface either witness before parser, shared-string, shared-formula, or
execution errors. A compact witness or ordered flag may remove a final Store
scan with less semantic risk than XML-reader fusion, but it still needs a
fresh stage-local profile and differential tests. It must not weaken the
style-count or scalar closure.

The small catalog and `PartView::data` shares should not be the first target.
The MCE namespace substring-search candidate belongs to rejected 0531
([evidence](../change-0531/README.md)); its production change was reverted
after the final native gates failed. It does not solve the
validation/parser duplicate pass, and this audit does not recommend
reevaluating the unchanged helper. Any separate future work would require a
new frozen plan and new native evidence.

## Required differential coverage

The current tests are strong but split by owner. `validation_borrow_tests.rs`
covers transitional/strict aliases, default/prefix rebinding, foreign and
unbound namespaces, malformed/duplicate attributes, DTDs, text/CDATA/general
references, mismatched ends, truncation, invalid UTF-8, and a broad truncation
differential against the pre-0521 validator. `raw/worksheet/tests.rs` covers
semantic values, formulas/shared formulas, rows/columns/dimensions, merges,
styles/metadata, x14ac/MCE behavior, and the exact 0539 cell-attribute error
precedence. Snapshot tests cover candidate-output validation/readback and
store equivalence. Source-backed integration tests cover atomic publication,
budgets, cancellation, source freshness, and refusal of dependency-bearing
markup.

What is missing is a combined first-error matrix through
`Snapshot::from_source_selected` or the public source editor. The future
candidate must compare the unchanged path and the fused path for the typed
error domain, exact relevant message, stage, and whether any state was
published. At minimum, pair failures in these classes:

* package/catalog, source/version, worksheet relationship, aggregate-byte,
  and cancellation boundaries;
* transitional/strict namespaces, aliases, default/prefix rebinding,
  foreign/unbound/mixed scopes, root/depth/truncation/mismatched ends;
* unknown, qualified, duplicate, malformed, and entity-encoded attributes,
  including the raw parser's invalid-`r`/style/metadata/type precedence;
* DTD, PI, comments/declarations, invalid UTF-8, text/CDATA and general
  references in and outside scalar contexts;
* MCE `AlternateContent`/ignored branches and x14ac default/row descent,
  duplicate, malformed, and parser-error-retry cases;
* dimensions, rows, columns, merges, formulas/shared formulas, inline/shared
  strings, scalar type/value errors, duplicate cells, styles, metadata, and
  unknown-cell closure; and
* exact source bytes/provenance, parsed `Store` fields, no partial snapshot,
  source retention, aggregate limits, execution fences, and multi-sheet
  publication/readback.

For fusion A, add fixtures where a parser error is early and a validator error
is late, and the inverse, to prove that full validation precedence is retained.
Also test validator-only failure after substantial parser allocations, early
parser failure with a successful validator, x14ac failure competing with a
validator failure, and plain parser failure that triggers an x14ac retry.
Measure cold effective edits, cold same-value cell/row/column no-ops, hot
no-ops, dense-sparse worksheets, and fallback-heavy worksheets. Record
stage-local and whole-plan p50/p95/p99, mean, RSS/peak live bytes, and
allocation counts/bytes under the frozen 0540 protocol. Instruction shares
must not be converted into a latency or allocation claim.

## Decision

The fresh 0540 attribution should first establish exact direct edges for
`validation::worksheet_xml`, `validate_xml`, `raw::worksheet::parse`,
`Parser::parse`, `process_ooxml`, and x14ac. If the duplicate reader machinery
remains material in native measurements, candidate A is technically feasible
only as a private, identity-gated, transactional state machine with deferred
preprocessing/parser errors and an unchanged fallback. Candidate A remains a
future unadmitted design. Candidate B is a higher-memory alternative. Candidate
C is preparatory and has limited savings. Candidate D is the lower-risk
adjacent option if the final Store scans show measurable cost.

The 0514/0516 parser/layout fusion reviews are prior matches for the required
constraints: x14ac must precede MCE/parser work; MCE-owned buffers require a
separate path; source identity and derived-buffer guards matter; scanner or
parser errors must not replace earlier semantic, web, style, or publication
errors; partial state cannot be published; and cold no-ops must not pay
unbounded speculative work. The rejected 0539 attribute `Cow` candidate also
shows that source-level allocation plausibility is insufficient without native
gates. No boundary-fusion implementation should be admitted from the 0530
percentages alone.
