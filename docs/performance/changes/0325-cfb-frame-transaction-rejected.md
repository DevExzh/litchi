# Change 0325: Reject CFB frame transaction for XLS source-backed scans

Status: rejected before implementation.

## Scope

The candidate was an exact-range CFB frame transaction (or bounded frame
prefetch) for the source-backed XLS `WorksheetScan` path. The intent was to
coalesce adjacent BIFF-frame work and reduce the number of CFB reads and
freshness probes while preserving the existing logical result.

An exact-range transaction cannot coalesce reads beyond the requested range.
Any overread used to obtain that coalescing would need an explicit contract
for distinguishing bytes that are payload from bytes that may be skipped. It
would also need to account for the combined materialization and input limits,
including limits that apply before the extra bytes can be accepted. No such
contract was established, so the candidate cannot claim a read reduction
while preserving exact-range behavior.

## Correctness boundary

The current source-backed path deliberately keeps freshness, I/O error,
cancellation, and deferred-error boundaries around individual operations.
Reducing probes or coalescing operations into a longer transaction weakens
the point at which a source mutation is detected and changes which operation
receives a read or version error. It would also have to preserve whether an
error discovered while preparing another frame is surfaced immediately,
attached to that frame, or deferred, without changing observable behavior.
Finally, it extends the amount of work performed between cancellation
checks. No source-level argument justified changing those boundaries for
this optimization.

The measured perf selectors `xls_source_backed_open` and
`xls_source_backed_open_list_worksheets` have no `WorksheetScan` opportunity;
they are therefore zero-opportunity cases for this candidate. The production
`text` and `write_text_to` scan paths are outside this decision and are not
covered by that claim. The one-cell path is the only measured scan case in
scope, and its possible reduction is bounded by the frames that one query
actually touches. That ceiling is not an achievable performance claim: the
exact-overread constraint and the preserved per-operation boundaries provide
no measured or defensible gain meeting the optimization acceptance bar.

## Prior diagnostic evidence

The following numbers are scoped to the earlier A1 diagnostic control only;
they are not a benchmark of this rejected candidate and are not generalized
to other XLS files. On
`test-data/ole/xls/ConditionalFormattingSamples.xls`, the prior
source-backed control observed:

| Selector | Logical reads | Logical bytes | Version calls |
| --- | ---: | ---: | ---: |
| `open` | 655 | 567,685 | 1,266 |
| `list` | 655 | 567,685 | 1,266 |
| one-cell | 921 | 569,398 | 1,813 |

Those counters only establish the shape of the earlier control path. They do
not show that a frame transaction can safely remove any of those reads or
probes, and no candidate binary or before/after run was produced for 0325.

## Decision

No production implementation, test change, or benchmark was made. The XLS
micro-optimization line is closed pending a materially different contract
that explicitly defines safe range ownership, freshness semantics, error
precedence, and cancellation boundaries. This record makes no claim about
speedup, RSS, I/O reduction, or cold-start behavior.
