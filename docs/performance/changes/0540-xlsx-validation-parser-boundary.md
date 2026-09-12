# 0540: XLSX worksheet validation/parser boundary attribution

Four fresh release Callgrind captures confirm that worksheet parsing and
value-only validation dominate source-backed one-percent edit planning.
Production and benchmark source remain unchanged from restored 0539. This
batch admits no optimization and makes no native latency or allocation claim.

| Shape/repeat | Planning Ir | Worksheet validation | Raw worksheet parse | Validator traversal | Parser traversal |
| --- | ---: | ---: | ---: | ---: | ---: |
| medium/r1 | 125,622,471 | 37.114% | 60.097% | 19.438% | 19.427% |
| dense-sparse/r1 | 237,073,145 | 37.849% | 59.759% | 19.859% | 19.852% |
| medium/r2 | 125,610,151 | 37.119% | 60.093% | 19.440% | 19.429% |
| dense-sparse/r2 | 237,116,049 | 37.843% | 59.765% | 19.855% | 19.849% |

All percentages use the selected planning call as denominator. Worksheet
validation and raw parsing are disjoint immediate children of the selected
snapshot loader. Each traversal column sums three disjoint immediate children
of its loop: `read_event_impl`, `process_event`, and `resolve_event`. These
traversal costs are nested within their owners and must not be added to the
validation/parsing columns. The validator-wide traversal includes the small
workbook validation call as well as selected worksheets. It is measured work,
not an estimate of removable work or a conversion to wall-clock latency.

Direct validator allocator/deallocator instruction edges account for only
1.057–1.076% of planning. That is not an allocation count or a bound on all
ownership costs. The pinned quick-xml API cannot return an input-lifetime
local-name slice through the existing event-name accessors, so direct stack
borrowing is not a valid drop-in change. A private discriminant representation
could be measured separately, but the larger traversal boundary takes priority.

The next implementation study should feed the same borrowed event to separate
validation and raw parser states only when preprocessing retains exact source
identity. It must preserve complete validation-first error precedence,
including later validator errors overriding earlier parser/preprocessing
errors, and the original x14ac error retry. Unsupported or transformed input
needs the authoritative existing path. No state may publish before semantic,
style, scalar, source-version, resource and execution checks succeed.

Before a candidate is measured, add a combined first-error differential matrix
and establish bounded provisional state on malformed input. Then freeze fresh
native, allocation and profile gates for both primary shapes and the existing
managed, one-cell, vendor-extension and noncompact guards. Include exact no-op,
invalid-input and peak-memory guards so speculative parsing cannot repeat the
0514 failure. The rejected 0516 emitted-output and 0539 attribute candidates
remain rejected; this is a different source-planning boundary.

The [evidence bundle](../results/change-0540/README.md) retains the fresh build,
symbol selection, all four profile children, all lifecycle and termination
dumps, deterministic annotations, source reviews, and read-only replay.
Six final 0539 quality checks and 1,292 final test executions are reused under
exact complete source-manifest equality; they are not newly executed tests.
No native before/after, allocator, hardware, cold/range or scaling lane was run.
OLE2/OOXML remain active, ODF is deferred, and iWork is excluded.
