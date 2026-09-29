# 0843 selection and evidence boundary

Previous turn: concrete progress, committed 0842 at 03fa423d73 with verified
runtime-dispatch results and owned-root cleanup. Current production and all 35
normative file hashes match that recorded state. The three unrelated local
files remain untouched.

The next shared-substrate question is whether one-Part publication spends
material work rebuilding preservation metadata when a package has many small
parts. SourceBackedPackage::write_changed_overlays_with_appended_inner creates
scratch/index, replacement and omission lookups/counts, scans the entries, then
creates a preservation action plan and scans again before emission. These
operations may be necessary and are not optimized merely because they exist.
The 0836 profile was dominated by Deflate on a few-large-member save and does
not establish their cost for many-small-member packages.

A fresh ordinary-route adoption trial of the already qualified PPT borrowed-text
candidate is also promising. The 0841 diagnostic wrapper and missing current
allocation lane prevent treating its favorable results as adoption evidence;
0782's historical regressions remain real. That trial is queued, not replaced
by a claim that this OPC diagnostic completes the broader goal.

The owner-approved durability policy already exists through 0761/0773, with
Full as the default. 0821 measured weaker-mode effects. The DOCX replayable-tail
path remains deliberately Full-only, as 0773 documents. No durability API or
policy is reimplemented in this batch.

## Fixed diagnostic

Four deterministic generic OPC packages contain 4/64/1024/4096 opaque 256-byte
parts plus content types and package relationships. Only the last Part changes,
by XORing its first byte. The independent reader regenerates every payload,
checks ZIP metadata/CRC, verifies exact outputs across all captures, and proves
untouched compressed local spans remain exact. These are not semantic Office
fixtures or native Office interoperability evidence.

Publish prepares source-backed open and the replacement input before its clock;
lifecycle includes open and plan construction. Both time the same public
write_topology_to_stream route, including prepare_topology validation and effective
graph construction before preservation emission, with result destruction and full output checking
outside the clock. The two latency populations are not subtracted to invent
phase times. Six native blocks alternate case order; thirty samples follow
three warmups. Independent processes, same CPU 12, one immutable binary.

Callgrind separately records one publication per process, two repetitions per
part count, using only the admitted profile_publish symbol as collection scope.
Guest instructions are not native cycles or elapsed time. Heaptrack separately
records the same one-call owner; only allocation events with that owner in
the decoded ancestry support operation attribution. Whole-process events remain
separate. Requested bytes are not physical copies or peak live memory. No
observed native elapsed/RSS under either profiler enters native summaries.

## Contracts

| Authority | Boundary |
| --- | --- |
| ADR 0001/0006 | Exact bytes, typed refusal and preservation outrank speed. |
| ADR 0002/0011/0024 | Probe uses public OPC/ZIP owners; no production grammar/API moves. |
| ADR 0003/0030 | No snapshot, patch, source-freshness or lazy-decode changes. |
| ADR 0005/0031 | Serial explicit owned-source operations; no hidden workers or ambient provider. |
| ADR 0032 | Profiling does not authorize an unbounded parsed-index memo. |
| 0758 | OLE2/OOXML priority; ODF deferred and iWork excluded. |

The current production inventory is exactly 0842 source-final, so its passing
owner/workspace-library/boundary gates are reused explicitly, with its two
baseline-confirmed excluded iWork gate failures. Fresh probe formatting,
warning-denied Clippy, tests, root-lock alignment and independent readers precede
measurement. Diagnostics may identify a next candidate, but this batch admits
no production optimization or broad coverage promotion.
