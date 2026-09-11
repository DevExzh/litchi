# Admission review: conditional worksheet parser/layout fusion

This review evaluates the retained candidate patch against base
`33a21e0f0`. The production and test paths have since been restored to base;
the decision below uses the candidate patch, its bound binaries, and the
captured `before`/`after` reports rather than the restored checkout. The
candidate replay was exact and its patch digest is recorded in
`candidate-patch.json`.

The speculative fusion gate for cold ordinary edits fails the cold semantic-no-op admission
condition. The independent public guard reproduces the expected exact no-op
behavior while measuring a large candidate cost in every shape and both no-op
sizes. This rejects admission of the current gate. It does not evaluate or
rank a replacement design.

## Evidence and measurement scope

The primary evidence is the standalone guard in `guard-probe/src/main.rs`,
bound by `guard-plan.json`. It uses the public XLSX API and a deterministic
two-sheet numeric workbook at 8x8 (`tiny`), 32x32 (`medium`), and 256x256
(`dense-wide`). The `before` and `after` corpora have identical byte lengths
and SHA-256 identities. Each guard receipt reports exit code zero and
`source_unchanged: true`; the guard source manifest is shared by both roles.
The raw report files are `before/guard-pilot-report.json`,
`after/guard-pilot-report.json`, and the paired `guard-allocator-r1` and
`guard-allocator-r2` reports. The timing/oracle path was reviewed in
`guard-probe/src/main.rs` (`run_scenario`, `warm_stores`, and
`prepare_edit`); the main pilot path was reviewed in
`tools/perf-baseline/src/lib.rs`, and allocator-region semantics in
`tools/perf-baseline/src/allocation_metrics.rs`.

For edit scenarios, the guard clocks only `Edit::commit`. Fresh workbook open, edit preparation,
Store warming, result destruction, serialization, and readback oracles are
outside the clock. Cold same-value edits derive their values from the
generator without querying the source Store, so they enter the candidate's
cold Store path. Warm scenarios enumerate the complete affected Stores before
the clock. Every cold and warm same-value iteration reports an empty patch and
byte-equal output; changed iterations pass their readback oracle. The cold
first-cell read also passes its numeric readback oracle.

The pilot uses 20 measured samples and two warmups per role and shape. The
allocator guard uses the unchanged shared counting allocator, 10 samples and
one warmup per role, with two repeats per role. Its `serialized_region_peak_v3`
counter records the observer-ordered live-byte high-water mark inside the
same operation region. This is allocator evidence, not an OS RSS measurement
and not attribution of every byte to one particular object.

## Cold semantic no-op result

The guard pilot's p50 elapsed times are below. The percentage is calculated as
`after / before - 1` from the raw reports.

| Shape | Cold same one cell | Cold same one percent |
| --- | ---: | ---: |
| tiny | 37,200 -> 62,530 ns (**+68.09%**) | 75,220 -> 123,790 ns (**+64.57%**) |
| medium | 483,862 -> 800,593 ns (**+65.46%**) | 976,114 -> 1,610,566 ns (**+65.00%**) |
| dense-wide | 31,451,314 -> 51,715,152 ns (**+64.43%**) | 62,569,727 -> 115,530,340 ns (**+84.64%**) |

The cold first-cell read, which does not request a layout, changes by only
`+1.62%`, `-1.90%`, and `-2.98%` for tiny, medium, and dense-wide. Warm
same-value edits, which hit the prepopulated Store branch, range from `-14.88%`
to `+3.68%`. Warm changed edits range from `-3.38%` to `-1.73%`. This control
pattern isolates the large regression to cold requested grid actions rather
than indicating a general timing shift.

The source implementation explains the result: eligibility is decided from
requested cell/row/column maps before projection knows that their values are
unchanged. A cold same-value commit therefore parses semantic data, runs the
shared snapshot observer, finishes the temporary Layout, validates and caches
the Store, and only then drops the Layout when projection finds no effective
action. The exact no-op oracle does not remove that work.

## Allocator and live-memory result

The two allocator repeats have identical per-scenario vectors. Absolute process live-byte maxima observed inside the operation
region are:

| Shape | Cold same one cell | Cold same one percent |
| --- | ---: | ---: |
| tiny | 90,731 -> 99,315 (**+9.5%**) | 118,023 -> 126,607 (**+7.3%**) |
| medium | 814,799 -> 934,312 (**+14.7%**) | 1,210,176 -> 1,329,689 (**+9.9%**) |
| dense-wide | 48,951,183 -> 56,328,006 (**+15.1%**) | 74,098,991 -> 81,475,814 (**+10.0%**) |

Subtracting region-entry live bytes per sample gives dense one-cell demand
of 27,278,293 -> 34,655,117 bytes (+27.04%) and dense one-percent demand of
39,575,534 -> 46,952,358 bytes (+18.64%). This incremental callback-ordered
peak is separate from the absolute table and from document peak or RSS.

The same reports show the extra allocation work directly:

| Shape | Scenario | Allocation calls | Allocated bytes |
| --- | --- | ---: | ---: |
| tiny | one cell | 406 -> 833 (**+105.2%**) | 54,337 -> 81,911 (**+50.7%**) |
| tiny | one percent | 807 -> 1,661 (**+105.8%**) | 108,726 -> 163,874 (**+50.7%**) |
| medium | one cell | 4,544 -> 10,101 (**+122.3%**) | 758,048 -> 1,176,888 (**+55.3%**) |
| medium | one percent | 9,121 -> 20,235 (**+121.9%**) | 1,518,594 -> 2,356,274 (**+55.2%**) |
| dense-wide | one cell | 265,295 -> 597,095 (**+125.1%**) | 47,016,362 -> 73,571,934 (**+56.5%**) |
| dense-wide | one percent | 533,203 -> 1,196,803 (**+124.5%**) | 94,143,070 -> 147,254,214 (**+56.4%**) |

Warm same-value and warm changed allocator vectors remain effectively equal
between roles, with only the expected one-byte baseline offset in region
absolute values. The cold increases therefore align with the candidate's
speculative Layout work and parser/Layout lifetime overlap. The measurement
does not establish a complete memory model for a future design.

## Main pilot and profile context

The ordinary main pilot is a changed-edit matrix and has no cold semantic
same-value scenario. Its p50 result is mixed: across the changed commit and
commit/save rows, all but one row improve by roughly `0.42%` to `6.21%`, while
the dense-wide `xlsx_one_percent_commit` row regresses by `+11.09%`. The main
pilot is useful context for changed edits, but it cannot discharge the cold
no-op condition.

The retained Callgrind profile selects the dense-wide one-percent commit/save
operation with three samples. It records 14,458,336,036 instruction
references before and 13,547,679,875 after, a `-6.30%` instruction-reference
difference for that selected operation. In the candidate profile,
`Worksheet::store_with_layout` accounts for 37.82% inclusive references and
`parse_with_layout` for 37.80%, with both `Scanner::start` and semantic
`Parser::consume_event` visible under the fused parser. This confirms the
intended shared path in the selected changed case; instruction references are
not wall-clock evidence for the no-op decision. The profile log contains a
Valgrind `brk segment overflow` diagnostic while still exiting successfully,
so the profile remains contextual evidence.

## Admission decision

The current candidate fails the cold semantic-no-op gate on both latency and
allocator evidence, across all three shapes and both one-cell and one-percent
actions. The no-op and changed readback oracles pass, so this is a performance
admission rejection rather than a correctness finding.

Retain the candidate patch and reports as review evidence, with the production
tree restored as recorded by `restoration.json`. A revised candidate must
measure a cold same-value action separately from changed edits and demonstrate
an acceptable allocation and region-peak result before the fusion gate can be
admitted. The same source-bound guard protocol should be rerun after any
change. No conclusion about whether a narrowed gate, prepass, or two-pass path
is fastest follows from this review.

No source, build, test, or benchmark commands were run by this review.
