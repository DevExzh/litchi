# Bounded splice-consumption batching

The retained change-0484 metadata2 traces identify 3,735,927 `statx` calls
on the prepared source descriptor in the authored-heavy diagnostic child.
`SpliceAuditReader::consume` checks the source twice and forwards output for
each parser-consumed authored slice. XML tokenization therefore amplifies
source metadata calls even though source bytes are not being read.

The candidate optimization batches only already-consumed bytes within the
existing adapter window. It must retain source-first errors, bounded work,
resource accounting, exact candidate/source/replay hashes, authenticated EOF,
and caller-visible partial-output accounting. The standalone source XML audit
remains a separate proof: a candidate can repair malformed source XML, so a
candidate audit cannot authorize deleting source validation.

| Constraint | Implementation obligation |
|---|---|
| ADR 0001 / 0006 correctness and preservation | Original bytes, grammar checks and typed refusals remain authoritative. |
| ADR 0002 / 0010 / 0011 / 0024 ownership | Optimization stays private to the OPC splice adapter; no dependency changes. |
| ADR 0003 snapshot and publication | Freshness checks guard external callbacks and final publication; immutable source policy remains intact. |
| ADR 0005 resources and I/O | Reuse the admitted buffer; no complete fragment allocation; work and accepted output remain counted. |
| ADR 0008 evidence | Focused adversarial, publication, replay and DOCX tests plus matched before/after measurements are required. |

All accepted ADR and README hashes match the previously read set in
`adr-refresh.json`. No ADR exception or new ambient provider is proposed.
The completed before/after capture and analysis are reviewed in
[results-review.md](results-review.md), including adverse tails, RSS, and
source-heavy metadata-call counts.

The finalized adapter retains per-parser-fragment authored Work charging and
checks source freshness before and after external replay/source/sink calls.
Consumed bytes stay in the admitted adapter buffer until a full window or
parser-return boundary. Their exact source, candidate, and replay digests
are then updated once; accepted sink bytes remain counted per successful
write. A pending typed failure prevents finalization callbacks. The separate
source XML audit and all authenticated replay passes remain present.

Ten focused unit tests cover bounded writes, mutation during sink output,
replay digest mismatch, short-sink accepted prefixes, Work-limit prefix
flushing, cancellation, source-window and authored-window version counts,
finalization after failure, and an overreported replay read count. The
authored-only test consumes a 128-byte payload one byte at a time through a
16-byte window and checks both exact output/digests and fewer than 64 source
version probes. This directly covers the measured authored-token hotspot.

The final source diff is confined to the private OPC splice implementation
and its adjacent unit tests. The unrelated Keynote formatting change is
already present in both historical and candidate source inventories and is
outside this batch. No iWork code is staged by this work.
