# 0586: the DOC paragraph hint is measured at exactly zero, and reverted

Status: **rejected and reverted.** `performance_claim: none` — this record
carries exact chain-link counts showing no change whatsoever, on every fixture
the library can read.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was tried

Change [0579](0579-cfb-resumable-chain-walk.md) built a resumable chain walk and
applied it to exactly one caller, `GlobalsBuffer::ensure` in `litchi-xls`. It
named `litchi-doc` as its own follow-up, and on inspection that looked like the
cleanest remaining fit in the repository:

- `resolve_paragraph` (`crates/litchi-doc/src/body_text/source.rs:1662`) reads
  8 KiB at a time through two nested loops, calling the **unhinted**
  `read_stream_range` each time.
- The offsets it reads are provably non-decreasing. `parse_clx` refuses any piece
  whose `fc` precedes the previous piece's `fc_end` — the
  `index != 0 && fc < previous_fc_end` guard returning
  `Error::Refused(Refusal::AmbiguousTopology)` — and `cp` only advances within a
  piece.
- It needed no change to `litchi-cfb` at all; the hinted API already existed.

The change was twenty lines: create one `StreamChainHint` before the piece loop
and pass it to `read_stream_range_hinted`.

## What was measured

Same method as change [0585](0585-cfb-resumable-cursor-construction.md): one
instrumented build with a `COLD_LEG` switch that makes every hint discard,
reproducing the unhinted walk. For the DOC path this isolates the change cleanly,
because `resolve_paragraph` is the **only** hinted caller in `litchi-doc` —
`GlobalsBuffer` is XLS-only, so no other hint is in play.

Scenario: open through `SourceSnapshot::open`, then walk every paragraph by
position until the snapshot refuses, over every `.doc` fixture under
`test-data/`.

| | unhinted | hinted | removed |
| --- | ---: | ---: | ---: |
| paragraph-walk chain links, 8 fixtures | **3,991** | **3,991** | **0.0%** |
| fixtures with zero saving | — | — | **8 of 8** |
| text digest mismatches | — | — | 0 |

Every fixture, individually:

| fixture | paragraphs | unhinted | hinted | removed |
| --- | ---: | ---: | ---: | ---: |
| `duplicate-style-names.doc` | 15 | 2,388 | 2,388 | 0.0% |
| `lists-margins.doc` | 4 | 796 | 796 | 0.0% |
| `noheadfoot-litchi.doc` | 1 | 261 | 261 | 0.0% |
| `documentProperties.doc` | 1 | 233 | 233 | 0.0% |
| `footnote.doc` | 0 | 103 | 103 | 0.0% |
| `endingnote.doc` | 0 | 90 | 90 | 0.0% |
| `table-merged-cells.doc` | 0 | 70 | 70 | 0.0% |
| `picture.doc` | 0 | 50 | 50 | 0.0% |

Not a single link was removed anywhere.

## Why it is zero, and why that is a design finding rather than a corpus accident

The hint is created **inside** `resolve_paragraph` and dies with it. It can only
shorten walks between reads issued by *one* call. A paragraph whose text is
reached in a single 8 KiB read has no prefix for it to resume, and neither does
the first read of any call.

So the saving is not merely small on this corpus — it is **structurally
unreachable at this scope**. Three of the eight fixtures resolve zero paragraphs
and four resolve one; the one fixture that resolves fifteen still issues too few
reads per call for a prefix to exist. Making this pay would require the hint to
span *across* `resolve_paragraph` calls, which means holding it in the snapshot
next to the `SharedOleFile` it borrows — a self-referential borrow, and a
materially different design from the twenty lines tried here.

**Most of the DOC corpus cannot exercise this path at all.** 49 of the 57 `.doc`
fixtures are refused by `SourceSnapshot::open`, which admits only ordinary
Unicode main-story documents. The eight above are the whole measurable
population, and they are small: the largest `.doc` files in the repository are
1,619,457 and 1,448,448 bytes, then the distribution drops steeply to 335,360 and
below.

## Disposition

**Reverted.** `docs/GOAL.md`'s decision rules are explicit — "revert speculative
complexity that does not improve representative workloads" — and a measured zero
on every measurable fixture is the clearest possible case. Keeping it would add a
hint, a mutable borrow threaded through two loops, and a paragraph of
justification, in exchange for nothing observable. `crates/litchi-doc/` is
byte-identical to `e927e139c`.

This is the fourth retained rejection in the recent record set, after changes
0524, 0534, 0548 and 0549, and it is recorded rather than discarded because the
*reason* is reusable: 0579's mechanism pays only where one scan issues many
ascending reads through one hint. Change 0585 pays because a text extraction
resolves hundreds or thousands of shared strings through a document-scoped
resolver. This did not, because the hint's scope was one call.

## What a future attempt would need

1. A hint whose lifetime spans many `resolve_paragraph` calls — which means
   solving the self-referential borrow against the snapshot's `SharedOleFile`,
   not merely moving the `let` statement outward.
2. A corpus that can show it. The eight measurable fixtures here cannot, and a
   deterministically generated large DOC would need to be labelled synthetic and
   could not by itself justify landing the change.
3. Evidence that the ordering premise still holds at the wider scope. Within one
   call it is guaranteed by `parse_clx`'s refusal; across calls, the caller
   chooses the positions, and nothing constrains them to ascend. Correctness
   would be unaffected — a hint past the requested offset is discarded and the
   walk restarts cold — but the saving would depend on access order in a way it
   does not today.

## Limitations

No timing, cycle, allocation or RSS measurement was taken; with a delta of
exactly zero links there was nothing for them to attribute. The measurement
covers the paragraph-read path only, not `edit_paragraph`, publication, or any
other `litchi-doc` scenario. The `perf stat` and callgrind captures originally
planned for this change were not run, because a zero-link delta makes them
uninformative about the mechanism.

The counter proves the hint removed no *chain links*. It does not prove the
twenty lines had no other cost — the discarded-hint bookkeeping is a few
instructions per call, unmeasured, and reverting removes it either way.

## Retained evidence

[`results/change-0586/`](results/change-0586/): `doc-chain-legs.tsv`, the raw
two-leg output for every `.doc` fixture including the 49 refusals, and
`doc-chain-summary.txt` folded from it. The instrumentation is documented in
change 0585's packet, at
[`results/change-0585/instrumentation.patch.md`](results/change-0585/instrumentation.patch.md),
and was removed from the tree after both measurements.
