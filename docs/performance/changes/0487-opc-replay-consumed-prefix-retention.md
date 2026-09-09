# 0487: retain consumed OPC replay prefixes across short reads

Short replay reads previously flushed each consumed slice, repeating sink
freshness checks even when the adapter allocation had unused space. The
private adapter now appends new replay bytes after the consumed prefix and
flushes at capacity, fragment completion, or termination. Provider callback
freshness checks, error precedence, Work, hashes, EOF proofs and accepted
output counts remain enforced.

The [144-process comparison](../results/change-0487/results-review.md) shows
35.61–35.99% lower authored-heavy file median latency, with unchanged candidate
archive bytes and effectively unchanged operation heap. Small-input latency
and three RSS regressions remain explicitly recorded. Separate diagnostics
show 59.90% fewer authored-heavy file statx calls.

OPC/DOCX tests, format, Clippy, rustdoc, boundaries, Python helpers, and two
10,000-run ASan fuzz campaigns pass. A parallel allocator harness assertion
failed; all five harness tests pass in the retained serial retry. See the
review for the unresolved process-global counter race and complete evidence.

The implementation adds six adapter tests and a public Stored/Deflate reader
failure test. It changes no public API or dependency. The full non-iWork goal
remains open.
