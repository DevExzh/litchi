# 0542 shared worksheet validation/parser traversal review

This review is read-only. It audits the proposed shared validation/parser
traversal against the restored XLSX source and the 0540 boundary review. It
does not admit a runtime change, run a build, or make a performance claim.

## Current boundary

The selected source-backed path remains:

```text
MultiSnapshot::load_source_backed
  -> Snapshot::from_source_selected
       -> validation::worksheet_xml
       -> raw::worksheet::parse
       -> style-reference validation
       -> scalar-cell validation
```

`from_source_selected` obtains a `SourcePayload`, checks the aggregate
worksheet-byte bound for the multi-sheet route, validates the complete
worksheet XML, parses it, checks execution, validates styles and scalar
closure, and only then constructs the immutable `Snapshot`. The outer
multi-sheet loader checks execution and source version before and after all
selected worksheets. A shared driver must leave these package, source,
cancellation, publication, and source-retention boundaries in place.

The validator and raw parser have different contracts. The validator accepts
only the value-only SpreadsheetML grammar and rejects foreign or mixed
namespaces, unsupported elements, disallowed/qualified/duplicate/malformed
attributes, dependency markup, DTDs, and text or references outside `f`, `v`,
or `t`. It ignores comments, declarations, and processing instructions when
there is no MCE preprocessing. Its persistent element names are owned because
the pinned quick-xml 0.41 API does not expose an input-lifetime name slice
after consuming a `BytesStart` or `BytesEnd` event. A candidate must not use a
lifetime extension, `unsafe`, or a resolver-internal buffer as a stack entry.

The raw parser performs its own complete attribute iteration and preserves
the established check order: cell coordinate, style, cell metadata, value
metadata, and cell type. It owns pending cell text, formulas, rows, columns,
shared-formula state, and materialized values; shared formulas resolve after
EOF, and shared strings can trigger a deferred package read. Style and scalar
closure checks happen after parsing. These states are provisional until all
existing phases pass and must never be published on a failed validation or
parse.

## Hard candidate constraints

### Callback lifetime and event scope

The MCE streaming callbacks are explicitly higher-ranked over their event
lifetime. `RawElement`, `SemanticElement`, `SemanticEnd`, and text values are
valid only during one callback. Their `Cow` fields do not authorize retaining
borrowed names, attributes, decoded text, or event objects in parser state.
Every value needed after the next XML read must be copied into an owned,
checked, bounded field. Parser strings and formula state already satisfy this
shape; a new state machine must retain it.

The existing MCE stream also creates its own owned context and event-level
allocations. It delivers only the selected semantic branch to its active
observer, while its raw observer sees every source start/empty element. A
candidate must not assume that a semantic callback is equivalent to the
original `NsReader` event or that its filtered attributes can replace the raw
parser's attribute loop.

### Validation-first errors

The current operation finishes the validator before any raw preprocessing or
parser work. A fused ordinary-source path may parse speculatively, but it must
retain only provisional parser state and defer every parser, x14ac, and MCE
diagnostic until validation has reached its normal EOF/root checks. A parser
error near the beginning followed by a validator error near the end must
return the validator error. The inverse ordering must return the validator
error as well. If validation fails, all parser state, preprocessing state,
shared-string handles, and derived layout must be dropped.

When validation succeeds, the raw error surface remains unchanged. In
particular, `raw::worksheet::parse` captures x14ac before MCE processing when
`may_contain_descent` says it is needed. On the plain path, a parser failure
causes the historical x14ac capture retry; a failed retry replaces the parser
error and a successful retry leaves the original parser error. A candidate
must keep this ordering and mapping even when the parser observer has stopped
after its first error.

### MCE and x14ac eligibility

The one-reader candidate is safe only where preprocessing is proven to return
the exact original bytes by allocation identity and length. A content-equal
owned `Cow` is not sufficient: parser event lifetimes, source offsets,
namespace scope, and retained `SourcePayload` identity then refer to different
buffers. Any MCE namespace/directive, `AlternateContent`, ignorable or
preserved extension branch, `MustUnderstand`, transformed output, or other
MCE-sensitive input must use the existing two-pass path unless a separate
identity proof covers it.

Likewise, any x14ac descent marker or eligible `dyDescent` capture must remain
on the existing path. x14ac observes raw source events before MCE selection,
and its typed errors and duplicate/default/row semantics are independent of
the value-only validator. A cheap byte substring test is not a sufficient
eligibility proof if it can miss a qualified attribute or reserved marker.

The callback MCE stream is not a drop-in validator replacement. It rejects
DTD and processing instructions, while the current value-only validator
ignores those events and the ordinary no-MCE parser path can accept a PI. It
also filters hidden branches and rejects hidden custom references according to
its own conformance contract. Therefore a candidate that routes all
worksheets through `process_markup_compatibility_stream*` changes observable
error and acceptance behavior. Keep the ordinary one-reader path MCE-free,
or retain the authoritative fallback for all such sources.

### Source, cancellation, and publication

The selected source payload must remain the exact `PartData`/`Arc` represented
by `SourcePayload`; managed payloads cannot be detached with `into_arc`.
`package.check_execution()` and `package.source_version()` must retain their
current placement and error mapping. The final source-version check is outside
the worksheet reader and cannot be replaced by a reader-local check. A
candidate must not cache a provisional store, expose it through a snapshot, or
retain a speculative shared-string read after any error.

### Resource bound and invalid-late behavior

The proposed invalid guard is useful evidence but is not itself a resource
invariant. A candidate invalid-input p50 at most twice the same-shape valid
planning p50 and incremental peak at most 1.10 times the valid planning peak
would bound the measured corpus. It cannot bound hostile input: the existing
OPC `ReadLimits` cap package/part bytes, but the raw worksheet parser does not
enforce its declared `MAX_XML_EVENTS`, and its `cells`, rows, formulas,
shared-formula maps, and decoded strings grow with the accepted worksheet.
The single-sheet source path also does not apply the multi-sheet aggregate
64 MiB check at this point; the part policy permits substantially larger
payloads.

Accordingly, the guard envelope is adequate as a frozen admission check only
when paired with the existing parser's checked reservations and an explicit
candidate rule that keeps provisional state bounded by the same configured
source/event/text/depth policies. If the implementation adds a retained event
or token log, a second full source buffer, or an unbounded diagnostic list,
the 1.10x measurement does not make it safe and the candidate should be
rejected. For late validator faults, the driver must stop both the reader and
parser at the first observer failure, release the provisional parser and
validator state before authoritative fallback, and retain the validator's
error order. A separate invalid-late peak and allocation guard is required;
valid-planning results cannot be used to claim adversarial safety.

## Proposed bounded public guards

The existing `source_backed_cell_values` integration helpers can exercise the
public `SourceBackedEditor::edit_sheets` path without a new public API. Keep
the package graph valid and replace only worksheet XML bytes in a deterministic
fixture. For each medium and dense-sparse shape, use the same release normal
and allocator binaries and source/cancellation policy as the ordinary primary
case.

Use these independent rows:

1. A valid source with an empty edit or same-value edit, to establish the
   ordinary no-op envelope and exact source/recovery behavior.
2. A parser-invalid cell near the beginning followed by a validator-invalid
   element near the end, to force parser allocations before a late validator
   error.
3. A valid prefix with a validator-invalid element after substantial cell
   records, to test validator-only late failure without parser-state growth
   being mistaken for success.
4. A valid worksheet with MCE/x14ac markers and the same fault, to verify
   authoritative fallback and error ownership. This row is a semantic guard,
   not an eligible fused speedup row.

For each row record operation p50/p95/p99 and mean, peak RSS or the available
peak-live proxy, allocation count/bytes, source/output identity, typed error,
and publication/recovery result. Compare each invalid row with its same-shape
valid baseline; retain every >5% adverse result for individual review rather
than averaging it away. The invalid p50 and 1.10x incremental-peak thresholds
are reasonable initial *admission gates* for this bounded corpus, but a
failure means fallback or rejection, not a relaxed limit. No allocation or
resource-safety claim should be made from normal binaries where counters are
unavailable.

The required differential matrix should also cover strict/transitional
aliases and namespace rebinding, foreign/unbound/mixed scopes, qualified and
duplicate attributes, DTD/PI/comment/declaration/reference event kinds,
formula/value/type errors, shared and inline strings, x14ac default/row and
retry errors, and a second selected worksheet. Every row should assert the
typed first error, stable message where appropriate, no partial snapshot, and
source usability after failure.

### Guard implementation scope audit

The current guard source calls `allocation_metrics::begin()` immediately
before `edit_sheets`, clocks that call, and finishes the Region while retaining
the returned transaction/result. Fixture construction, package opening,
selector setup, retry, commit, and inspection are outside the interval. The
Region reports an observer-ordered operation-local absolute live-byte peak;
the analyzer subtracts the same sample's before-live value for the
incremental guard. This supports the planned 1.10x invalid peak comparison.
A missing, overflowed, or observer-invalid Region is reported unavailable or
failed closed, never converted to zero. The report still labels the metric as
allocator evidence rather than RSS or an adversarial memory proof.

The guard's empty `edit_sheets` transaction followed by an out-of-clock
`commit` is a valid no-op oracle. The `semantic-noop` spelling aliases that
same operation and emits the canonical `valid` case with the same fixture; it
does not exercise a same-value staged edit. Any admission report should call
it an empty-edit row unless a real same-value setter is added, and should
preserve the source-version/source-byte check after both accepted and refused
runs.

## Candidate review

The frozen production patch changes only the selected multi-sheet
source-backed worksheet path. It keeps `SourcePayload` and the surrounding
execution/version checks intact. The eligibility predicate is conservative:
it rejects non-UTF-8 input, sources above 8 MiB, literal MCE or
`AlternateContent` markers, and x14ac/`dyDescent` markers. A false positive
only selects the established two-pass path; for a marker-free source,
`process_ooxml` is the borrowed identity path and the raw extension capture is
the empty default. I found no eligibility route that bypasses the established
MCE or x14ac diagnostics.

`Validator::observe` is a callback-scoped, owned-name copy of the existing
worksheet validator. It observes the event before `Parser::transition`, and
the parser transition is the old raw event match extracted into a shared
function. Provisional parser errors are discarded. On every provisional
failure the caller reruns the authoritative validator and then the ordinary
raw parser, so a parser error before a validator error cannot outrank the
validator and raw x14ac retry behavior remains in one place. No speculative
`Store`, shared-string read, source buffer, event queue, or published snapshot
survives that fallback. The outer cancellation and source-version fences are
unchanged.

The applied driver now returns immediately with `ProvisionalFailed` when the
observer rejects an event, the event cap is exceeded, the parser transition
fails, or the reader fails. The authoritative full validator in the fallback
therefore supplies the first diagnostic, including when an early raw parser
failure is followed by a later validator error. Immediate return prevents the
quick-xml namespace resolver from draining malformed tails and bounds the
provisional parser work by the failure prefix. This closes the earlier drain
concern.

One resource close condition remains: the caller keeps its `Validator` local
alive while fallback `worksheet_xml` and `raw::worksheet::parse` run. The
grammar-derived element stack is small for current worksheet rules, but
`first_error` owns an `Error::Invalid(String)`, and diagnostics can format a
source-sized element name, attribute name, or text value within the 8 MiB
eligibility cap. An explicit `drop(validator)` before both fallback calls is
needed to release that provisional diagnostic and avoid overlapping it with
authoritative allocations. If the local is retained, the invalid-late guard
must measure and admit that overlap explicitly; the frozen resource rule
should not assume it is negligible.

The explicit 8 MiB and 131,072-event gates make the provisional raw record
count and copied text finite, and parser stack growth is checked against the
existing 256-level transition limit. The validator stack has checked reserve
and arithmetic; its bound comes from the finite accepted worksheet grammar,
independently of `MAX_XML_DEPTH`. The source and event caps bound the accepted
worksheet corpus, but they do not make every allocation fallible or provide
an OOM guarantee. In particular,
`resolve_shared_formulas` still builds `members: Vec<_>` and
`masters: HashMap<_, _>` without `try_reserve`; attribute decoding and the
quick-xml namespace resolver also use ordinary allocations. These are existing
parser behavior and can be accepted only with a bounded-state claim. The
candidate documentation correctly makes no blanket checked-collection or
hostile-input memory-safety claim.

The final applied patch, cap-test patch, source manifest, and test-module
wiring are hash-consistent. The focused tests cover size/event-cap fallback
and late validator/raw cases, while the retained 0541 matrix covers an early
raw parser failure followed by a later validator failure. This closes the
earlier freeze-consistency and precedence-fixture concerns.

The implementation is therefore technically plausible for a private,
identity-gated, transactional ordinary-source path. Explicit validator
release before fallback remains the resource admission condition. The normal
and allocator guard captures remain required before a speedup claim:
ordinary planning bytes and incremental peak must stay within the frozen 1%
gate, and invalid p50/peak must satisfy their same-shape thresholds. Those
measurements are admission evidence, not a substitute for the explicit
8 MiB/event bound or a claim about the unrestricted baseline parser.
