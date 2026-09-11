# Next-scope review after 0508

Two independent source reviews considered text export and ordered OPC Part
batches. ODT is selected for fresh profiling: the existing large export has
10,000 blocks, each currently starting a new String and dropping its completed
buffer after emission. The measured 0508 medians are about 1.9 ms. A single
bounded reusable buffer could remove repeated allocation without changing
writes, decoding, output progress or nested start-order publication. This is
a candidate, not yet an admitted optimization or causal attribution.

ODS performs a row-size walk plus the core joined writer's preflight and
emission walks. Eliminating one safely may need a core preflight API or bounded
staging, so it has a wider proof surface. ODT decoded-length scanning is also
a later candidate; UTF-8, CRLF and decoder error boundaries must remain.

The OPC reviewer found repeated first-wave preparation/fences and redundant
active-ordinal clearing. These are small local candidates, not evidence that
queue overhead dominates the remaining local latency. Removing fences also
needs explicit cancellation/source-error precedence review around worker
startup. The 0499 whole-child counters cannot attribute operation-local channel,
startup, cache or ZIP cost; they do not justify a larger scheduler rewrite.
That provider-profile follow-up remains open. No global pool or automatic
provider-based admission shortcut is proposed here.

Admission for ODT requires the fresh 0509 control profile, a fixed retention
bound, unchanged nested frontier and output-call order, exact bytes/progress
and failure checks, matched before/after samples, and retained adverse flags.
