# OPC ownership checkpoint for DOCX tail append

This checkpoint completes the shared OPC prerequisites discovered while
integrating the bounded DOCX scanner with the
[0482 decoded splice](../../changes/0482-bounded-xml-opc-splice.md).
The DOCX integration, settings admission, benchmark captures and full goal
remain in progress. This checkpoint makes no latency or RSS claim.

`SourcePartSpliceFragment` provides fixed-length storage whose package memory
reservation precedes allocation and survives the handoff into the splice
plan. A different package instance cannot substitute another budget. Storage
initialization checks cancellation and source freshness in bounded chunks;
failure and destruction release the lease. The existing `Arc<Vec<u8>>` entry
point retains its behavior. The reservation describes requested buffer
capacity, excluding allocator bookkeeping and the caller's encoder state.

The XML auditor's per-event attribute scratch now depends on the smaller of
the aggregate attribute policy and token window, including the first rejected
attribute. Its aggregate counter still rejects too many attributes spread
across events. This lets a small parser window coexist with a large document
event count without reserving document-sized duplicate-attribute scratch.

OPC positional reads now retry `Interrupted` while retaining the input-byte
reservation. Each managed retry consumes one Work unit and checks source
freshness; cancellation is observed before retrying. Only accepted bytes
consume input budget. Public `ExecutionError` values carried by format-owned
I/O guards retain their typed OPC cancellation/budget outcome, including the
existing partial-output wrapper.

The preservation oracle obtains compressed lengths from the central directory,
so deferred local sizes cannot omit payload or descriptor bytes. A separate
raw central-record comparison checks all directory bytes except the relocated
local-header offset. Negative oracle tests distinguish a directory-only
attribute mutation from permitted relocation. These independent parsers are
restricted to the generated non-ZIP64 fixtures; broader ZIP64 coverage belongs
to the existing archive and external-corpus gates.

The source review is recorded in [opc-foundation-review.md](opc-foundation-review.md),
and the attribute bound review in [xml-workspace-review.md](xml-workspace-review.md).
Validation receipts retain failed development attempts separately from passing
checks. The default OPC suite passed 515 tests with one external ZIP64 corpus
test ignored, followed by 29 passing splice tests after the central oracle
was added. The all-feature OPC suite passed 538 tests with the same external
corpus test ignored. Warning-denied OPC Clippy and rustdoc passed with all
features. The XML suite passed 46 tests with one explicit asset-regeneration
test ignored; its warning-denied Clippy and the scoped format check passed.
[foundation-validation.json](foundation-validation.json) binds the exact
commands, raw logs and source hashes. Raw test logs retain their original
trailing blank lines. These checks do not certify the unfinished DOCX path.
