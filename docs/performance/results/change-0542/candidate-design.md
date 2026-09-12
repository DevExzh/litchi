# Change 0542 candidate: shared source worksheet traversal

This is an isolated candidate for the selected multi-sheet source path. It
changes only `Snapshot::from_source_selected`: eligible worksheets use one
borrowed `quick_xml::NsReader` and otherwise retain the existing
`validation::worksheet_xml` followed by `raw::worksheet::parse` calls. No
single-sheet path, public API, dependency, source payload, or test is changed.

The raw worksheet codec owns the event driver and parser transition function.
For each borrowed event it calls the caller-owned validation observer first,
then advances the ordinary parser with that same event. The validator owns its
existing local-name stack; event payloads and the source remain borrowed. No
event buffer, rewritten XML buffer, or second source is created. The driver returns immediately and drops the parser and stack as soon as the
observer reports a validation failure, the parser transition fails, or a
provisional bound is reached; the fallback then performs full validation first. It
never materializes shared strings, formulas, or a `Store` until the observer
has accepted the EOF event and the validator has therefore checked the root
and depth.

Eligibility is conservative and byte based. The source must be valid UTF-8,
fit within `MAX_SHARED_SOURCE_BYTES` (8 MiB), and contain neither the MCE
namespace/`AlternateContent` marker nor x14ac/`dyDescent` markers. The MCE
default input and output limits are checked explicitly as well. With no MCE
namespace, `process_ooxml` would return the original borrowed bytes after
those same limit checks; with no x14ac marker, the raw extension capture is
the default empty value. Any source outside this proof uses the unchanged
two-pass path.

The provisional parser counts every reader event, including EOF. It drops
state after `MAX_SHARED_PROVISIONAL_EVENTS` (131,072), and also enforces the
worksheet model's `MAX_XML_EVENTS` (1,000,000) guard. Its element stack is
limited by the ordinary `MAX_XML_DEPTH` (256) transition. The new parser stack/row and validator-stack reservations are checked. Existing
attribute decoding, namespace resolution and some post-validation collections
still use ordinary allocations; this is not an OOM guarantee. The validator
stack is bounded by its accepted worksheet grammar, independently of the raw
parser depth limit. A retained
local-name copy corresponds to bytes in the bounded source; parser text has
the existing per-cell/formula limits, and the event cap bounds record count.
Thus a late validator error can retain only one bounded provisional parser,
which is released before authoritative parsing; there is no unbounded
failure-only store or event queue. A parser error returns immediately, so no parser state survives into the
authoritative fallback validation.

`ReaderFailed` and every other provisional failure run the authoritative
`worksheet_xml` and then `raw::worksheet::parse`. This preserves validation
first precedence, preprocessing/x14ac diagnostics, UTF-8 and MCE limit
errors, parser/materialization wording, and source-byte fences. The observer must accept EOF and complete root/depth checks before materialization;
all failure branches obtain the final diagnostic from the established
validator/parser calls.

The remaining risk is the extra temporary parser allocation on a valid or
late-invalid worksheet within the conservative 8 MiB/131,072-event window.
The fixed event/depth/byte bounds make that cost finite, while malformed
reader input, parser failure, or the guard immediately releases it and pays
the historical fallback cost. This candidate intentionally does not claim a
process-wide memory budget for the baseline parser or alter ODF/single-sheet
behavior.

The focused integration checks should cover a normal eligible source, a
comment-padded source above 8 MiB, a comment-heavy source above 131,072 events,
and late validator/parser errors for both in-window and fallback inputs. Each
cap is an admission fallback, so the oversized cases must retain the baseline
outcome rather than become a new rejection. MCE, x14ac, invalid UTF-8, and
reader-error fixtures should likewise exercise the unchanged authoritative
path.

Changed paths in this candidate patch:

- `crates/litchi-xlsx/src/cell_values/snapshot.rs`
- `crates/litchi-xlsx/src/cell_values/validation.rs`
- `crates/litchi-xlsx/src/raw/worksheet/mod.rs`
- `crates/litchi-xlsx/src/raw/worksheet/codec.rs`

The final applied source also includes three private test functions in
`shared_traversal_tests.rs`, wired through `validation.rs`. The original proposal,
cap-tests patch, and early-fallback supplement are retained;
`applied-candidate.patch` and the candidate source manifest bind the formatted
implementation actually built and measured. The retained 0541 public matrix
covers an early raw parser error followed by a later validator error.
