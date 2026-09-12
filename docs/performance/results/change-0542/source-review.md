# Applied source and error boundaries

`applied-candidate.patch` is the exact formatted baseline-to-candidate change.
It combines the original proposal, cap tests, and immediate-fallback supplement.
The five changed files are independently bound by `candidate-frozen-inputs.json`
and `candidate/source-manifest.json`; the three harness enabler files are
identical in both measured sources.

Only selected source-backed worksheet loading chooses the new traversal. Other
loaders keep their existing entry paths. The ordinary raw parser uses the same
extracted event transition function, so the eager read controls remain required
before retention. Source payload ownership and surrounding execution, version,
style, scalar and publication checks retain their existing positions.

Eligibility requires at most 8 MiB of original UTF-8 source, the existing MCE
input/output limits, and absence of the MCE and x14ac markers used by the
authoritative preprocessing paths. Ineligible bytes use the existing validator
and raw parser. An eligible reader delivers each borrowed event to the validator
before the raw parser. Root/depth checks occur on EOF before raw materialization.

Observer rejection, parser failure, the 131,072-event cap, and reader failure
return immediately and release provisional parser state. The caller then runs
the established full validator and raw parser. This preserves validation-first
errors even when a raw failure precedes a later validation failure, and retains
historical preprocessing and x14ac retry precedence. The 0541 public error matrix
tests that combined ordering, retries and selected-worksheet order. The three
new test functions cover accepted and refused cap fallback plus an eligible edit.

The candidate bounds source bytes, event count, record count and provisional
text; raw stack depth retains its existing limit and validator nesting is
restricted by its grammar. These are finite-state bounds, not a process memory
budget or an OOM guarantee. Existing namespace/attribute decoding and some
post-validation formula collections still allocate through ordinary APIs. The
implementation adds no event log, second XML buffer, unsafe block or runtime
dependency. Measured invalid-input memory is reported separately from these
source-level constraints.

The new standalone guard uses the existing allocator Region. A process lifetime
high-water mark was rejected during harness drafting, before source freeze or
capture. Every measured incremental peak subtracts that sample's before-live
bytes from its operation-region peak. Normal binaries explicitly report
allocation measurements unavailable. The valid guard is an empty-edit/no-op
operation; it is not a same-value setter benchmark.

The final static review identified a remaining overlap: `Validator` survives
into the measured fallback and can retain an error string while the authoritative
validator constructs another. Together with the failed late-raw latency gate,
this prevents retention. `next-candidate.patch` is an unmeasured follow-up that
drops that state and forwards only errors produced after validated EOF through
the existing x14ac retry. It has no build or performance admission in this batch.
