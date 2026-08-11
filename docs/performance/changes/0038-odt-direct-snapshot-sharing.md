# Change 0038: ODT direct snapshot byte sharing

## Scope

Direct `transaction::Snapshot::from_bytes` ingress now validates the consumed
package once and adopts the validated document's existing private
`Arc<Vec<u8>>`. Reopening a snapshot for transaction staging clones that Arc
into the existing shared-package constructor instead of allocating and copying
the complete archive again.

This removes two archive-sized copies from the common direct non-noop
transaction path: one during snapshot validation and one before staging the
edit. It changes no public type or signature and retains the 64 MiB transaction
bound, exact source bytes, immutable snapshot ownership, complete ODT parsing,
signed/encrypted refusal, compact-XML audit, final ODT reopen and semantic
readback, deterministic effects, patch replay, inverse restoration, and stale
source rejection. Existing-document `Document -> Snapshot` sharing remains the
separate accepted path from change 0014.

## Matched media-rich measurement

The existing opt-in `odt_media_paragraph_edit_save` case fixes a
16,786,287-byte ODT with 200 paragraphs and eight deterministic incompressible
2 MiB media members. Its input SHA-256 is
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
The timed region consumes direct snapshot bytes, replaces paragraph 100,
commits, and materializes the candidate. Untimed checks reopen the complete
ODT, verify every media member and manifest entry, replay the patch, apply its
inverse byte-exactly, reject a stale base, and require deterministic output.

Matched ABBA release runs used CPU 2, 10 warmups, and 100 samples per leg. The
pooled distributions contain 200 samples per state.

| State | p50 | mean | p95 | p99 |
| --- | ---: | ---: | ---: | ---: |
| Before | 32.270 ms | 30.125 ms | 34.706 ms | 36.723 ms |
| After | 7.798 ms | 7.879 ms | 8.534 ms | 9.045 ms |

The pooled p50 delta is **-75.84%**, mean **-73.84%**, p95 **-75.41%**,
and p99 **-75.37%**. Both matched comparisons improve independently: the A
leg p50 moves 32.289 -> 8.063 ms and the B leg 32.151 -> 7.657 ms. The before-B
distribution contains 23 low outliers below 20 ms; this lowers its mean, but
does not change the directional result or its 32.151 ms median.

Heaptrack over ten samples reports 129,341 -> 129,261 allocation calls
(-0.0619%, exactly eight removed calls per full harness iteration), unchanged
22,817 temporary allocations, and flat 106.03 MiB peak heap. Heaptrack RSS is
117.72 -> 117.92 MiB (+0.17%) and uninstrumented maximum RSS is
109,720 -> 109,592 KiB (-0.12%); both are treated as flat. The whole-process
counter run, which includes corpus construction and all untimed verification,
moves task clock -0.40%, cycles -0.38%, instructions -1.03%, cache references
-12.89%, and cache misses -18.53%. The latency claim is therefore scoped to
the harness's predeclared edit/save interval rather than generalized to all
process work.

Raw ABBA samples are retained as [`before A`](../results/abba-odt-shared-snapshot-before-a.json),
[`after A`](../results/abba-odt-shared-snapshot-after-a.json),
[`after B`](../results/abba-odt-shared-snapshot-after-b.json), and
[`before B`](../results/abba-odt-shared-snapshot-before-b.json). Heaptrack,
GNU Time, counter summaries, frozen binary hashes, and evidence hashes are
indexed by [`odt-shared-snapshot-sha256.txt`](../results/odt-shared-snapshot-sha256.txt).

## Preservation and correctness gates

A private regression proves both ownership boundaries: the input `Vec` buffer
survives direct snapshot construction at the same address, and a semantically
reopened `Document` shares the snapshot's exact package Arc. The existing
package transaction suite continues to cover exact no-op identity, ordinary
changed commit and readback, patch/inverse/stale behavior, malformed and
oversized input, signed/encrypted preservation/refusal, raw media preservation,
and the over-limit publication fallback.

Executed gates:

- all 525 library tests, every integration test, and all 55 doctests in the
  all-feature ODT suite passed;
- warning-denied all-feature ODT library Clippy passed; and
- package formatting and diff-hygiene checks passed.

## Remaining work

- ODT still parses required XML parts on each semantic rehydration; only the
  redundant complete-archive copies are removed.
- This result covers direct byte snapshots on the existing media-rich
  paragraph transaction. It is not evidence for structural/resource-adding
  operations, cold filesystems, other ODF owners, or source-backed ODF APIs.
- The predecessor-byte copy retained for per-operation reversible patch
  semantics is deliberately unchanged and requires separate attribution.
