# 0542: shared XLSX traversal rejected on late-error latency

The measured candidate improved valid source-backed edit planning, but three of
four late raw-error rows exceeded the frozen latency envelope. Production was
restored exactly. The standalone planning guard is retained as a measured enabler;
the candidate and its cap tests remain evidence, not live runtime changes.

| Shape | Repeat | Planning p50 reduction | Workflow p50 reduction | Workflow mean reduction |
|---|---:|---:|---:|---:|
| dense-sparse | 1 | 26.41% | 8.02% | 8.07% |
| medium | 1 | 25.30% | 7.64% | 7.40% |
| dense-sparse | 2 | 26.24% | 9.15% | 9.20% |
| medium | 2 | 23.55% | 6.94% | 6.40% |

These are matched candidate measurements, not a retained speedup. The four primary
rows passed the 10% planning and 3% total p50/mean requirements. Planning allocated
bytes fell 0.078–0.104%; incremental peaks rose only 0.0025–0.0062%, within the 1%
limit. Allocation-instrumented timings were not compared.

The late-raw medium repeat 1 and both dense-sparse repeats took 2.048×, 2.041× and
2.038× their same-shape baseline-valid planning p50, exceeding the fixed 2× limit.
Medium repeat 2 passed at 1.992×. Against each same-invalid baseline, late raw
errors were 52.1–55.3% slower and late validator errors were 164.7–168.8% slower.
The validator cost includes speculative parsing before the late fault and remains
a separate tradeoff for the next candidate. Every invalid incremental-peak row passed the
1.10× baseline-valid envelope. The current failure path discards a post-EOF
materialization error and repeats full validation/raw parsing plus the historical
x14ac error scan. The next candidate should preserve that validated error while
retaining x14ac retry precedence; this requires fresh matched evidence.

Static review also requires dropping the retained validator before fallback: its
first error can own source-sized text while authoritative validation constructs
another error. The 8 MiB source and 131,072-event caps bound provisional state;
they are not an OOM guarantee or a process-wide memory budget.

The campaign contains 2,440 main native samples, 240 allocator samples per main
phase, 720 standalone native samples and 240 standalone allocator samples. Both
baseline and candidate use the same three-file harness enabler. The standalone
fixtures are distinct 96×96 and 128×128 stored-ZIP grids, not the main CRUD corpus.
Raw receipts bind source manifests, build identities, commands, sample order and
ABBA timing. All 48 native adverse flags, 99 native repeat-drift flags, 417 guard
adverse flags and 13 guard drift flags receive individual dispositions. These
include correlated statistics and absolute allocator diagnostics, not independent
experiments. Ordinary phase allocation analysis has no >5% adverse flags.

The baseline suite passed 1,298 tests and the candidate passed 1,301, including
three cap-boundary tests. The restored final suite passed 1,298 tests (3,897 test executions across the
three runs); all eight final checks passed, including workspace check, Clippy,
rustdoc, formatting, crate boundaries and standalone guard checks.
Conditional instruction, hardware and eager lanes were not executed after the
refusal gate failed; their prepared scripts provide no measurement claim.

OLE2 and OOXML remain the optimization priority; ODF is deferred. The broad goal
remains incomplete.

Evidence: [bundle](../results/change-0542/README.md),
[decision](../results/change-0542/decision.json),
[source review](../results/change-0542/source-review.md),
[static review](../results/change-0542/review.md).
