# 0420: share validated payload storage during owned OPC reopen

> Correction (0421): allocator `peak_live_bytes_*` values in this record
> may under-report peaks and must not support high-water claims.
> [Counter correction](0421-allocator-peak-counter.md) explains the defect.
> Live-byte totals, allocation requests, normal timing, RSS and Heaptrack
> metrics are unaffected by this specific bug. Raw reports are unchanged.

Owned PPTX cross-copy retains a freshly validated candidate archive and graph.
Its eager reopen also decoded payloads already held by the source and staged
destination. Candidate `81c70058a` lets OPC select an existing immutable payload
allocation after decoding the new archive. Reuse requires exact URI/content
type, agreement with the donor's visible bytes, equality with the new decoded
bytes, and no larger vector capacity. New archive metadata, relationships,
preservation records and source authorization remain independently established.

PPTX uses this advanced OPC constructor only for clean owned destinations.
The existing custom/dirty/save-option fallback and all detached candidate,
patch, physical and semantic publication checks remain. No dependencies,
ordinary CRUD signatures or read/output limits change.

Against clean control `68e4b0dfa`, matched warm generated-media and plain
lifecycle diagnostics show:

| Metric | Media-rich lifecycle | Plain lifecycle |
|---|---:|---:|
| Live-after process snapshot | 287,885,879 → 237,452,746 B (−17.518%) | 884,426 → 796,963 B (−9.889%) |
| High-water process snapshot | 896,603,242 → 812,687,527 B (−9.359%) | 3,565,760 → 3,559,495 B (−0.176%) |
| Whole-process RSS, all pairs | −9.15% to −9.27% | −0.08% to +0.15% |
| Requested allocation volume and calls | effectively unchanged | unchanged |
| Normal p50, paired directions | **+0.656%** / −1.554% | **+1.688% / +1.224%** |
| Normal p99, paired directions | **+0.887%** / −1.413% | **+2.384% / +5.215%** |

The plain A2/B2 p99 rises from 9.432934 to 9.924835 ms, crossing the declared
5% review trigger. That adverse observation remains in the result and receives
an explicit [resource review](../results/change-0420/resource-review.md).
All same-revision drift checks remain below 5%. The candidate is retained for
lower measured live bytes and media-rich RSS, accepting the small plain
latency cost on this corpus. Timing is mixed and establishes no speedup.

The [bundle](../results/change-0420/README.md) retains 16 fresh-process reports:
800 normal observations (100 samples/10 warmups per leg) and 240 separate
allocator observations (30/3). The normal count is below the 500-sample
release latency threshold; no latency claim is registered. CPU 2, one worker,
Rust 1.98.1, identical release flags and harness, and all source/binary/corpus
identities are recorded. Sources are warm and generated in memory on a shared
AMD EPYC 9R45 KVM host. Every output, topology and semantic/refusal gate matches.

Full target decompression and its allocation requests still occur before
reuse. Process live/high-water snapshots and RSS are distinct observations;
they do not measure an operation-local peak or a bounded aggregate budget.
The prior 0419 candidate heap trace is the unchanged control's profile basis;
0420 captures no new heap trace. The [ownership review](../results/change-0420/source-review.md)
distinguishes per-reopen payload accounting from lifecycle observations.

Validation: 1,286 OPC/PPTX all-feature tests passed, three ignored; seven final
focused tests passed. Tests cover shared preservation holders, exact target
output, donor metadata isolation, mismatched bytes/capacity, inconsistent
custom payloads, copy-on-write isolation, input limits and malformed ZIP errors.
Clippy passes with the same three established command-local exemptions;
rustdoc passes with warnings denied. Strict nine-claim, crate-boundary and
CRUD-index gates pass. Retained-report replay and semantic corruption probes
provide independent evidence checks.

Near-limit and low-memory behavior, operation-local retention/peak attribution,
source-backed lifecycles, native producers, physical cold/range I/O and scaling
remain open. This batch does not complete the non-iWork performance program.
