# 0894 — selected worksheet metadata route

This note records the candidate design, completed quality and corpus evidence,
and its measured qualification. The candidate is rejected under the frozen
primary gate for this source, build, corpus, and protocol; the result is not a
general Office performance conclusion. The [0892 producer profile](0892-current-producer-full-cell-profile.md) identified the authoritative
worksheet parse after selected-cell scanning as the next high-value work
elimination target. [0893](0893-forced-worksheet-tolerance-guard.md) then
strengthened the benchmark so that every worksheet tolerance report is forced
before an iteration is accepted.

## Baseline and hypothesis

The baseline source is commit
`506aa2edae6666d9c76ae6827c2f5941a8e47896`. Its selected worksheet scanner
reads the XML stream, records selected cells and their dependencies, and
refuses producer metadata that it cannot prove inert. For example, the current
producer refusal occurs at `sheetViews`, before selected cells are visited;
the source-backed route then falls back to the authoritative worksheet parser,
which performs the full worksheet parse and materializes the store.

The candidate admits a small, value-checked metadata subset to the selected
route. For a clean eligible worksheet, the route can retain the selected-cell
records, dependency closure, source and execution fences, EOF verification,
resource limits, cancellation behavior, and callback ordering without doing
that second authoritative parse. The paired trial tested this work-elimination
hypothesis and failed its primary gate. No causal profile was collected; the
measured result and the next investigation are recorded below.

The benchmark does not instrument route selection. `runtime_route_observed`
must remain `false` in both roles. Raw-scanner tests provide eligibility
evidence separately from the timed source-backed workload.

## Bounded candidate contract

The candidate validates metadata while preserving the existing semantic-event
stream, MCE dispatch, x14ac observation, XML depth and input/output ceilings,
and typed error precedence. The admitted profile is deliberately narrower than
the set of valid SpreadsheetML:

| Metadata | Admitted subset | Required fallback or refusal |
| --- | --- | --- |
| `sheetViews` | Direct container; bounded `sheetView` count; checked `workbookViewId`; at most one checked `selection` projection per view with bounded A1 and `sqref` values | Unknown children or attributes, duplicate selections, empty containers, and branches rejected by the existing MCE dispatcher |
| `sheetFormatPr` | Finite, bounded row/column defaults and boolean/outline fields; recognized x14ac `dyDescent` remains under the extension observer | Unknown fields, invalid values, duplicate or repairable metadata |
| `cols` / `col` | Direct, style-free records with checked widths/flags, grid bounds, strict source order, and non-overlapping ranges | Styles, overlap, empty or duplicate containers, invalid widths, and any repairable form |
| Row `spans` | An optional bounded first:last range within the worksheet grid | Invalid values or forms whose authoritative parser would repair |
| `pageMargins` | One complete direct record with finite bounded values | Missing/extra fields, invalid values, duplicates, or unsupported placement |
| `pageSetup` | A small value-safe `paperSize`/`orientation` subset and lexically valid strict or transitional relationship ID | Other print semantics, invalid IDs, unknown fields, or duplicates |

The accepted elements must occur in the scanner's permitted order. Unknown
elements and attributes, non-whitespace character data, CDATA or general references in metadata,
styled columns or cells, formulas or dependencies outside the selected
contract, duplicate or late metadata, and repairable or ambiguous forms mark
the worksheet ineligible and leave the authoritative parser in charge. The
scanner never skips an opaque subtree merely to reach `sheetData`. Malformed
XML, resource-limit failures, cancellation, CRC/size failures, source changes,
and other typed failures retain their existing error paths.

Admission to the newly supported metadata and row-`spans` path has an
additional budget precondition. Let `n` be the declared worksheet input bound;
the path is admitted only when `n` is no greater than each of the event-count,
event-byte, and per-event attribute-byte limits, `2n` is no greater than both
the name-byte and context-byte limits, and the event-attribute cap is at least
the legacy attribute cap. The default ceilings are approximately 512 KiB.
This prevents the metadata-enabled stream from acquiring a stream-only tail
limit that would make the authoritative baseline fallback less bounded. An
uncertain or larger part falls back at the first newly admitted metadata
barrier. Existing prefix failures and the plain numeric route retain their
current behavior.

This route does not materialize layout, view, or print semantics. In
particular, the selected-cell path does not resolve the worksheet relationship
graph or expose a `pageSetup` metadata accessor; its relationship ID check is a
bounded lexical check for the selected-cell contract. Callers that need print
settings use the authoritative print-settings accessor. A valid worksheet can
therefore still be conservatively rejected.

The change is confined to the XLSX worksheet raw owner and qualification
tests. It adds no public CRUD type, dependency edge, unsafe code, cache, or
producer-template behavior. The authoritative parser remains the fallback and
semantic reference.

## Guard and measurement scope

Both measurement roles must be rebuilt from the current source family with the
0893 guard contract
`post-endpoint-per-worksheet-tolerance-v1`; historical 0891/0892 binaries do
not qualify this candidate. The timed scope ends after the fresh source-backed
open, worksheet enumeration, full stored-cell visits, and canonical callback
digest. After the timer, allocator region, and process endpoint snapshot, the
harness forces every worksheet tolerance report, checks the aggregate report,
checks source freshness again, and validates the independent oracle.

Those post-endpoint checks are required for correctness, but they can force
authoritative worksheet/store work outside the timer and still affect whole
process RSS, allocator high-water state, and profiler or Callgrind totals.
Consequently those metrics remain descriptive, and the old aggregate-only
captures cannot be reused as candidate latency evidence.

The fixed primary producer archive is 22,739 bytes with SHA-256
`55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8`. It has
four worksheets and 4,096 cells: 3,686 numeric and 410 shared-string values.
Its canonical cell digest is
`8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b`.

The paired protocol uses four cases: this producer full-cell scan, a
marker-free XLSX full-cell control, a producer DOCX selected-paragraph
control, and a producer PPTX selected-slide control. It schedules 64 native
children and eight allocator children, sequentially on CPU 12 with a warm
filesystem cache. Native children use three warmups and 50 samples; allocator
children use three warmups and 15 samples. Native roles alternate AB/BA by
block and case; allocator roles alternate by block. Reduction uses bootstrap
seed 894, 10,000 draws, and the retained 249/9749 interval endpoints.

The capture runner validates plan-bound artifact hashes and role, build, and
review bindings before launching the sequence. A reviewer also called
`load_plan` while capture was active, causing additional artifact hashing;
its exact overlap and timing effect were not measured. The shared host had no
exclusive-use guarantee. These results support rejecting this candidate under
the recorded protocol, not a precise causal slowdown or a universal tail or
RSS claim.

The primary gate is the native producer-case candidate-over-baseline p50
paired median ratio at most `0.98` with a bootstrap 95% upper endpoint below
`1.0`. A pair above `1.05`, p50 spread above `1.20`, allocator counters, RSS,
tail metrics, and controls are review evidence rather than automatic adoption
gates. The completed reduction, independent review, and semantic verification
are reported below.

## Quality and corpus evidence

Focused tests compared the scanner with the authoritative worksheet
`Store`, including clean positives, shared-string dependency closure,
strict/transitional namespaces, prefix and x14ac aliases, active and inactive
MCE branches, malformed XML and typed error chains, limits, cancellation,
source fences, styles, overlaps, invalid values, duplicates, late ordering,
and repairable metadata. The producer route tests will check the generated
four-sheet package and controlled metadata removals. A positive scanner result
is useful only when the authoritative value and error vectors agree.

The completed [quality-v9 receipt](results/change-0894/attempts/quality-v9/quality-receipt.json)
passes all nine gates: scoped formatting; seven focused metadata tests; 1,397
full `litchi-xlsx` library tests; all-features compilation; warnings-denied
Clippy; warnings-denied rustdoc; 22 producer-shape harness tests; one
selector test; and crate-boundary checks. The earlier v8 attempt retained a
namespace-fixture failure; v9 includes its correction. Earlier failed attempts
remain retained, and this quality result makes no build, timing, or adoption
claim.

The completed [corpus-v1 run](results/change-0894/corpus-v1/run.json) covers
all 180 checked-in XLSX inputs, 394 worksheets, and 1,576 worksheet legs for
each role across cold source visits, cold source cell traversal, warm visits
after stored-extent access, and eager cell traversal. The before/after input
censuses are equal. The corrected
[comparison-v2](results/change-0894/corpus-v1/comparison-v2.json) reports zero
candidate-versus-baseline input, leg, structural, or summary differences. It
retains 12 same-operation errors on chart sheets, present identically in both
roles; no discovery or source-open error was omitted. Cell transcripts,
partial callback counts and digests, typed error chains, tolerance reports,
aggregate tolerance, and source-version stability remain part of the corpus
records. This is semantic evidence only and makes no timing claim.

The first host comparison failed while normalizing the scalar `summary.legs`
field. `compare-v2.py` and `corpus_finalize.py` reran the analysis against the
already completed probe outputs; neither probe was rerun. The original failure
and its logs remain retained alongside the corrected comparison.

## Measured decision

The [reduction](results/change-0894/reduction.json) and
[verification](results/change-0894/verification.json) receipts both pass their
analysis and custody checks for all 72 children and 3,320 measured samples.
The primary native producer full-cell p50 ratio is **1.8972120877** with a
95% bootstrap interval of **1.8743111450–1.9204038076**. The frozen gate
requires a ratio at most `0.98` and an upper endpoint below `1.0`; it fails
decisively, corresponding to an 89.72% slower candidate median.

The table reports candidate-over-baseline native ratios for the eight paired
blocks in each case. RSS is a descriptive process metric.

| Case | p50 | p95 | p99 | Peak RSS |
| --- | ---: | ---: | ---: | ---: |
| Producer XLSX full-cell scan (primary) | 1.897212 | 1.907545 | 1.903312 | 4.881975 |
| Ordinary XLSX full-cell scan | 0.996510 | 0.995612 | 1.008199 | 4.910723 |
| Selected DOCX paragraph | 1.009419 | 0.985367 | 1.012750 | 6.120170 |
| Selected PPTX slide | 0.998365 | 1.002065 | 0.949042 | 6.044943 |

Six control review-flag groups remain recorded: p99 and peak RSS for each of
the three control cases. The primary case has four review flags for p50, p95,
p99, and peak RSS; no native p50-spread flag fired. Allocator observations
changed from **25,933** to **126,171** calls and from **37,618,900** to
**47,632,498** allocated bytes, while region peak-live bytes changed from
**9,252,418** to **8,649,758**. These are descriptive observations, not an
automatic gate. The large RSS ratios across all cases cannot establish a
causal heap change because executable residency is a known confounder from
[0888](0888-executable-residency-and-mce-controls.md).

The [restoration receipt](results/change-0894/restoration.json) passes and
records the exact rejection decision: three tracked candidate files (one
implementation and two test files) were restored, and the two newly added
candidate files were removed. The trial failed before metadata admission could
be reconsidered. The next step is to profile per-event allocation in the
selected semantic stream before optimizing it or revisiting metadata
admission; no causal attribution is claimed from this timing run.

## ADR matrix

| ADR | Application to this candidate | Qualification evidence |
| --- | --- | --- |
| 0001 — priorities and API layers | Correctness, safety, and typed refusal precede work elimination; the optimization remains in a narrow raw XLSX owner and does not leak into ordinary CRUD APIs | Authoritative-parser differential, typed-error checks, and no public API change |
| 0002 / 0024 — topology and ownership | `litchi-xlsx` owns the worksheet scanner; no peer-format or archive dependency is introduced | Dependency-boundary gate and scoped source inventory |
| 0003 — snapshots, edits, and patches | The route is read-only and must retain immutable source identity and post-operation freshness checks; it changes no edit or publication semantics | Source-version fences and existing read-path tests |
| 0005 — I/O, memory, and measured performance | Selected streaming remains bounded, lazy, cancellable, and source-backed; performance claims require representative paired measurement with memory evidence | Fresh baseline/candidate builds, 72-child protocol, uncertainty interval, and RSS/allocator review |
| 0006 — validation, security, and compatibility | Unknown, repairable, ambiguous, styled, duplicate, or misordered markup falls back; limits, MCE, namespaces, relationships, and preservation behavior remain fail-closed | Differential corpus, malformed/limit tests, MCE/x14ac controls, and typed error chains |
| 0008 — migration and verification | The candidate remains a buildable, test-gated migration step and is not a support claim until the evidence gates pass | Quality receipt, source/build custody, independent review, and retained-only verifier |
| 0017 — producer template ownership | The fixed producer corpus is evidence for a format-owned template; this candidate does not alter generated assets or minifier ownership | Producer archive hash/oracle and unchanged producer-template boundary |

## Status

The quality-v9 receipt, corpus-v1 semantic replay, paired measurement,
reduction, verification, restoration, cleanup-v2, and
[verification-after-cleanup](results/change-0894/verification-after-cleanup.json)
are complete. The [cleanup receipt](results/change-0894/cleanup.json) accepts
the previously hash-verified early removal of the quality root and retains the
original cleanup failure. [Capture review](results/change-0894/capture-review.json)
and [corpus replay after cleanup](results/change-0894/corpus-v1/verification-after-cleanup.json)
pass. The [decision](results/change-0894/decision.json) rejects the candidate;
[final review](results/change-0894/final-review.json) binds the retained reports
and cleanup evidence. No 0894 worktree was created, and all owned temporary
source/build roots were removed. Older dirty worktrees were preserved. The
broader non-iWork performance goal remains open.

Evidence: [0892 current producer profile](0892-current-producer-full-cell-profile.md),
[0893 tolerance guard](0893-forced-worksheet-tolerance-guard.md), and the
[0894 qualification packet](results/change-0894/README.md).
