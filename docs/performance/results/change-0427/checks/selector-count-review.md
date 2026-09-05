# Change 0427 selector-count review

This is a source-only review of the selector-count failure recorded in
`baseline-selector-count.log`. I did not run Cargo, tests, capture commands, or
profilers. The retained failure reports the mechanically counted `Case` enum
at 429 and the stale assertion at 427.

## Finding

The two-count increase predates the 0427 retention module. Commit
`39e8dca8a82099f7e3af14d174fc9a35eb985ace` added the two owned lifecycle
selectors and brought the enum from 425 to 427. Commit
`f8f9e6667284ae28fdbf9b9313b2ff9583ec74fc` then updated only the test constant
from 425 to 427. Commit
`6ca9962c7818e173538a40ed45f3a3c32cc1aa6a` subsequently added two matched
source-backed lifecycle selectors without updating that constant, bringing the
enum to 429.

The 0427 retention module adds no `Case` variants. It dispatches through the
four already registered lifecycle names:

| API and corpus | Existing selector |
| --- | --- |
| owned, plain | `pptx_cross_copy_plain_lifecycle` |
| owned, media-rich | `pptx_cross_copy_media_rich_lifecycle` |
| source-backed, plain | `pptx_source_backed_cross_copy_plain_lifecycle` |
| source-backed, media-rich | `pptx_source_backed_cross_copy_media_rich_lifecycle` |

The current enum contains those four lifecycle variants plus the existing
owned plain/media-rich and source-backed plain phase selectors, for seven
cross-copy entries in total. Each
new lifecycle variant has a `Case::name` arm and a `parse_case` arm. The
`Case::DEFAULT.len() == 36` invariant is independent of this selectable enum
count.

## Count progression

The enum body was counted from each revision with the same uppercase-variant
rule used by `selectable_case_count_matches_current_enumeration`:

| Revision | Change | Selectable variants |
| --- | --- | ---: |
| parent of `39e8dca8a` (`147b2f6dae`) | baseline before owned lifecycle selectors | 425 |
| `39e8dca8a` | add `PptxCrossCopyPlainLifecycle` and `PptxCrossCopyMediaRichLifecycle` | 427 |
| `f8f9e6667` | update the expected literal only | 427 |
| `6ca9962c7` | add `PptxSourceBackedCrossCopyPlainLifecycle` and `PptxSourceBackedCrossCopyMediaRichLifecycle` | 429 |
| current source | no later `Case` additions from 0427 | 429 |

The appropriate source correction is therefore to keep the mechanically
verified expected count at 429 (or replace the literal with a generated
registry count later). No selector-count change should be attributed to the
0427 retention diagnostic.

## Review result

This is a stale test constant, not a retention-dispatch regression. The
existing name/parse coverage for the two source-backed lifecycle selectors
should remain part of the selector guard when the corrected count is recorded.
