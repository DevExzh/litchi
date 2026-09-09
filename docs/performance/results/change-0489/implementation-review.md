# 0489 implementation review

The change stays within the private OPC splice owner. A prepared XML splice
now carries a scalar capability containing the full source/replay/candidate
proof, target index, and frozen splice limits. Preparation still runs the
standalone source audit first, then the initial candidate audit. Only success
mints the capability; the subsequent source-freshness and execution-context
checks must also succeed before a plan is returned. Binary and exact-no-op
paths retain their existing behavior.

Publication validates the capability against the immutable plan before opening
the replay provider. Each pass checks the proof and limits again. The private
expected-artifact preview copies the same capability. There is no global cache,
public proof-bypass switch, document-sized allocation, new dependency, or unsafe
code. A capability from another plan cannot authorize different proof scalars,
target or limits. The independent applied-diff review found no semantic blocker.

## Authentication and failure behavior

Later measurement, preview and emission passes drain the existing bounded
splice adapter without rerunning the XML parser. Every pass retains the
verified decoded source reader, complete source/replay/candidate digest and
length checks, authenticated EOF, source-first freshness checks, cancellation,
ZIP verification, and exact accepted-output accounting. A changed provider can
fail byte authentication instead of XML grammar; no failing pass yields a
successful publication. Final DOCX semantic reopen remains mandatory.

Existing fixed-plan and per-replay-pass XML workspace reservations remain even
when later passes do not allocate a parser. Concurrent publications therefore
keep independently charged replay windows and conservative XML admission.
Successful Work accounting remains byte-based. The raw drain consumes whole
bounded payload windows rather than XML token ranges, so the private prefix
reached at a Work limit may change. A rejected charge is atomic: a budget of
four units cannot partially accept an eleven-byte payload charge. The added
unit test verifies typed refusal, zero charged payload Work, and the exact
source prefix already accepted by the sink. Source Work is charged upstream
by the verified reader.

## Focused coverage

Five added adapter tests cover malformed source/candidate and XML token limits,
capability binding and rejection of initial audit mode at publication, exact
raw fixed-candidate output, atomic Work refusal, and retained XML workspace
admission. Together with existing callback, EOF, hash, cancellation and sink
failure cases, the adapter suite contains 21 tests.

A public replay test checks 36 combinations: Stored/Deflate, ordinary two-pass
or expected-artifact four-pass publication, and truncated/extra/same-length
changed replay data at each later provider open. It checks the failing open
count, typed errors, empty caller output for private passes, and exact
IncompleteOutput counts for final emission. Existing preservation, source
mutation, Work, preview, concurrent reservation and reopen tests remain.

The first all-feature and no-default suites passed. A retained Clippy attempt
then rejected the grouping of a new hexadecimal test literal. The grouping
was corrected without changing its value, and one comment was clarified to
distinguish source and payload Work accounting. Final validation and builds
bind source manifest
`0f5d9a455d25146823a7a8eb3620712f6e8f199abe0d40c32678d100c3994afa`;
terminal results are reported in the results review.
