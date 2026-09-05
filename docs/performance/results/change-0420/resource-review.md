# 0420 resource and acceptance review

The frozen 0420 capture contains 16 fresh-process runs: normal and allocator
ABBA lanes for the media-rich and plain PPTX cross-copy lifecycles. The
reports bind the same corpus, output digest, plan topology, source-immutability
checks, and refusal gates to control `68e4b0dfa` and candidate `81c70058a`.
The resource acceptance rule is satisfied: both selectors show lower matched
`live_bytes_after` in both ABBA directions, with no material allocation-volume
change and no RSS increase beyond the review threshold.

## Allocator observations

The allocator lane has 30 retained samples per leg. Values below are the
means retained in `summary.json`; requested allocation bytes are allocator
request volume, not physical copying.

| Selector | Metric | Control | Candidate | Candidate delta |
|---|---:|---:|---:|---:|
| media-rich | live after | 287,885,879 B | 237,452,746 B | −50,433,133 B (−17.52%) |
| media-rich | process high-water after | 896,603,242 B | 812,687,527 B | −83,915,715 B (−9.36%) |
| media-rich | RSS, A1/B1 | 884,676 KiB | 803,760 KiB | −9.15% |
| media-rich | RSS, A2/B2 | 884,940 KiB | 803,508 KiB | −9.20% |
| plain | live after | 884,426 B | 796,963 B | −87,463 B (−9.89%) |
| plain | process high-water after | 3,565,760 B | 3,559,495 B | −6,265 B (−0.18%) |
| plain | RSS, A1/B1 | 82,740 KiB | 82,672 KiB | −0.08% |
| plain | RSS, A2/B2 | 82,608 KiB | 82,736 KiB | +0.15% |

Requested allocated bytes, allocation calls, and reallocation calls are
unchanged within the paired measurements. The largest difference is media-rich
A1/B1: requested bytes +0.00011%, allocation calls +0.00022%, and
reallocations unchanged. This is consistent with full target decompression
remaining in the candidate path.

The [lifecycle ownership review](source-review.md) distinguishes the first
plan's retained patch from the application candidate. The patch retains eight
copied media payloads (16,777,216 bytes); the application candidate retains
16 output media payloads (33,554,432 bytes). Sharing these accounts for
50,331,648 bytes of the 50,433,133-byte live-after difference. The remaining
101,485 bytes are an unassigned non-media/allocator residual. This is source
ownership accounting, not a per-allocation profile. The
83,915,715-byte high-water difference likewise is not an operation-local peak
or a proof of an aggregate memory bound.

## Normal-lane review trigger

The normal lane has 100 samples and is diagnostic only; no latency claim is
authorized. Media-rich timing moves by +0.661% in A1/B1 and −1.508% in A2/B2
for the mean, with p99 movement of +0.887% and −1.413%. Plain timing moves by
+1.590% and +1.571% for the mean. Plain p99 is +2.385% in A1/B1 and
+5.215% in A2/B2 (9.432934 ms to 9.924835 ms), which crosses the protocol's
5% review trigger by 0.215 percentage points. Plain p95 remains below 4% and
RSS is flat. This is a disclosed diagnostic trigger, not grounds for a
latency claim or a hidden replacement of the original result.

Same-revision drift is small for the allocator vectors: requested bytes,
allocation calls, reallocations, and live snapshots are effectively identical
within each revision. Normal timing drift is +2.45% for the media-rich control
mean and +0.25% for the candidate mean; plain control and candidate drift are
both below 0.5% for the mean. RSS drift stays below 0.16% in the retained
ABBA legs.

## Recommendation and limits

Retain the candidate implementation in its current guarded scope. It meets
the measured resource objective on both retained selectors: matching target
payload allocations are no longer retained after reopen, while complete
validation, decompression, metadata construction, source authorization, and
publication checks remain unchanged. The favorable media-rich RSS and
process-snapshot direction supports the ownership hypothesis for this corpus;
it does not establish a general RSS or peak-memory improvement.

Keep the result classified as a resource diagnostic with no latency claim.
The plain p99 trigger should remain visible and receive a larger latency
sample or operation-local attribution before any latency or broad workload
claim is considered. The capture does not measure cold input, range I/O,
scaling, native-producer packages, near-limit memory behavior, or an
operation-local peak. Full target decompression still occurs, donor custom
parts remain outside the PPTX opt-in gate, and raw source/output ZIP
allocations remain required.
