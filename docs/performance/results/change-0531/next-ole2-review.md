# Next OLE2 action after the 0531 OOXML pilot

`scope: bounded read-only OLE2/CFB backlog review`

`performance_claim: none`

## Decision

After the current 0531 OOXML pilot, the highest-justified OLE2 action is a
fresh operation-local attribution of the CFB ownership and reconciliation
work in `OleFile::open`, led by `OleFile::claim_sector`. Keep
`validate_stream_allocations` and `validate_physical_sector_layout` as
separate measured owners. This is a measurement and candidate-selection step;
a runtime change still needs measured benefit and proof that every required
validation remains intact.

The latest CFB constructor profile gives these disjoint exclusive shares in
the XLS source-open scope ([0523 profile analysis](../change-0523/analysis.json)
and [chain review](../change-0523/chain-review.md)):

| Owner | Exclusive instruction share | Why it is next |
|---|---:|---|
| `OleFile::claim_sector` | 17.81–17.82% | largest residual owner after the rejected collector experiment |
| `validate_physical_sector_layout` | 14.25–14.26% | large mandatory reconciliation pass that has not had a focused candidate |
| `validate_stream_allocations` | 14.16–14.17% | large chain/marker validation pass that has not had a focused candidate |

`SectorChainScratch::collect_exact` remains the largest individual owner at
40.08–40.09%, but its checked visited-bit fusion is closed: the matched 0524
candidate reduced constructor instructions only 1.17–1.19%, left allocation
vectors unchanged, and did not reach the required 3% lower p50 in both
repeats. The few-large CFB guard also regressed 2.35–2.41% in p50
([0524 decision](../../changes/0524-cfb-visited-bit-evaluation.md)). The next
campaign must therefore measure the residual owners instead of reopening that
small collector change.

## Required measurement boundary

Use a fresh current-head release profile with separate owner toggles and
positive incoming caller checks. Cover the existing CFB tiny, many-small and
few-large shapes, plus the XLS source-backed open/list/one-cell consumers. The
few-large shape is the sensitivity case: `collect_exact` was 42.62% of its
direct CFB profile, versus 2.12% for tiny and 5.45% for many-small in 0523.
Retain the smaller shapes as guards rather than extrapolating from the large
case.

Pair the attribution with operation-local allocation calls, allocated bytes,
reallocations and incremental live peak. Then use the same serial ABBA native
workflow gate as 0524, with exact CFB/XLS semantic, source, physical-layout,
cycle, marker, overlap, limit and error-oracle checks. A candidate can be
considered only if the profile identifies a specific duplicate operation and
the source review proves unchanged error order, fallible allocation behavior,
sector ownership, and physical reconciliation. A lower Callgrind count alone
is diagnostic. Retain a candidate only under the declared end-to-end native
and drift gates; do not promote a source-open or CFB-guard result into a broad
OLE2 claim.

The proof must keep collection and claiming in their current order. The 0523
chain review records why fusing them can report an overlap before a late cycle
or marker error and can leave partial role mutations. The physical pass must
also continue to detect unclaimed non-free sectors; moving it into FAT loading
would require a new bounded state and first-error proof. Those are review
constraints, not an implementation design.

## Evidence that keeps this priority bounded

The accepted 0511 FAT-sector extension already reduced `load_fat` to 1.79% of
the final XLS constructor profile; the larger remaining shares are the checks
above. Its eager XLS p50 gains of 14.87–17.25% and CFB few-large gains of
45.78–46.30% do not transfer to plain owned-source workflows, whose results
were small and mixed. Repeating FAT batching or inferring a general CFB gain
from eager opening is not justified ([0511 record](../../changes/0511-cfb-fat-entry-reservation.md)).

The older XLS source attribution found a different, provider-specific lead:
`FileSource::version` accounted for 46.83–49.68% of the atomic-to-FileSource
mean gap, with 1,266 version calls for open/list and 1,802 for one-cell
([0278 attribution](../../changes/0278-xls-source-attribution.md)). The
operation-scoped freshness session then cut those calls by 97.9463–98.1132%
and improved descriptive p50 by 52.54–56.31%, but it was rejected because
four same-side p95/p99 drift cells exceeded the predeclared 5% limit
([0279 rejection](../../changes/0279-cfb-operation-freshness-session-rejected.md)).
Do not reinstate that candidate or treat its descriptive reductions as an
accepted OLE2 result. A separate provider investigation can return only after
the CFB residual-owner measurement and a new, predeclared tail-stability
policy; it is not the next implementation here.

Recent DOCX and PPTX records do not change this OLE2 ranking. The 0517–0519
DOCX results concern OPC XML/source-snapshot publication, while 0474–0478
concern PPTX streaming preflight, compressor state and ZIP directory metadata
([DOCX 0517](../../changes/0517-opc-source-xml-validation.md),
[DOCX 0518](../../changes/0518-docx-source-snapshot-reuse.md),
[DOCX 0519](../../changes/0519-opc-publication-xml-proof-reuse.md),
[PPTX 0475](../../changes/0475-pptx-streaming-attribution.md)). Their
owners and clocks are separate from CFB construction and must not be pooled
with this measurement.

OLE2 and OOXML remain ahead of ODF. ODF work stays deferred until that
optimization goal is complete, and iWork remains outside this review.
