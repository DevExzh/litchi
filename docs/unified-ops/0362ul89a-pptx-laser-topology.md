# 0362ul89a — PPTX laser operation scope and topology admission

This return closes two composition findings in the PPTX raw shape and motion
lane, on top of integration `126647b3f4d285e4367b7033fdf6b0aa1805af02`.

## Laser edit aggregate

`laser::Edit` now retains one `Usage` ledger for its staged operation.  Each
successful `set` charges its completed slide read-back once; the source
preflight remains local because a slide that already owns a trace is refused
before it can enter the affected-slide scope.  A failed verb stages both the
candidate graph and ledger, so the prior edit remains retryable.  `reset`
discards the ledger with the graph, and commit/replay continue to use one
ledger for their changed-part walk.  This preserves the selected aggregate
slide-byte, trace, and point bounds across multiple slides without a package
wide rescan or per-slide counter reset.

The raw writer call site consequently takes the caller's mutable `Usage`
ledger.  Existing dialect, exact output, read-back, and failure-atomicity
checks remain unchanged.

## Prepared topology admission

Prepared source snapshots now propagate the result of
`transition::admit_source_with_limits` instead of converting every admission
error to a plain slide.  Supported enclosed markup-compatibility shape
removal still receives `SlideMce::Enclosed`; malformed XML, selected shape or
transition limit failures, and unsupported MCE forms refuse before the owned
payload is published.

## Focused coverage and validation

The laser edit tests derive an exact one-slide XML bound, admit the first
changed slide, refuse the second aggregate charge, verify the retained usage
after failure and retry, then verify reset/no-op and a fresh retry.  The
topology change is source-only and preserves the existing whole-shape-removal
capability.  Cargo, rustdoc, native tools, and performance capture were not
run in this return; the coordinator must compile and qualify the composed
tree after the typed-error lane is merged.
