# Measurement scope refinements

The frozen historical controls call `PartView::data_with_accounting`, while
the production batch API uses ordinary cache-backed reads without an operation
accounting report. This is an instrumentation difference. Historical
before/after comparisons remain descriptive and must not attribute their whole
delta to the scheduler. The final harness supports serial `PartView::data` and
`read_parts_ordered` in the same executable with accounting disabled; those
fixed-work controls are the primary evidence for batch scaling.

Both routes retain the selected handles through cache/budget observation and
verify bytes after the timer. The serial caller's vector is not a managed
library owner; the batch wrapper additionally charges its structural storage.
That deliberate accounting difference is not an unexplained retained-memory
regression. Zero memory/object usage after both result and package drop remains
a mandatory oracle.

The few-large corpus selects four Parts, so requested widths above four cannot
create additional independent Part work. Efficiency computed against requested
width must state that cap. The many-small corpus selects sixty-four Parts.
The synthetic delay source models an explicitly supplied positional provider;
it does not demonstrate a production network service or real remote latency.
FileSource runs use warm files, not a controlled cold-cache experiment.

Timing CSV source counters record calls and requested/returned bytes. The
independent integration gates, rather than these CSV counters, prove actual
overlap and independent declared-byte/task admission limits. Whole-child RSS
and perf counters include fixture setup, repeated package opens, verification,
and reporting; they are not operation-local memory or CPU attribution.

Accepted timing rows use three warmups and thirty measurements per repeat,
with two repeats. In current CSV files both are tagged `sample`; exclude
indices 0–2 before calculating measurement statistics. Retain every adverse
latency/RSS flag above five percent. Shared-host results and thirty-sample tail
percentiles are descriptive; no isolated-host or broad end-to-end claim follows.
