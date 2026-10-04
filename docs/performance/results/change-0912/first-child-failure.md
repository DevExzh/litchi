# 0912 first-child native perf failure

The first native pilot is retained and rejected. Its workload completed the same 50-operation materialized route guards, but perf reported **6 lost records** totaling **304 lost events**. The decoded sample stream contains no exact `litchi_pptx::package::model::Package::opened_presentation_transaction` frame, so it cannot establish the measured snapshot interval.

The pilot recorded `16920` `cycles:u` sample records with **1008** nonempty callchains and **15912** empty stacks. Their sampled periods sum to `72417582246`. These are sampled event periods, not elapsed time, native cycle totals, guest instructions, or a route comparison.

| owner scope | samples | period |
| --- | ---: | ---: |
| measured owner | 0 | 0 |
| setup owner | 0 | 0 |
| unclassified owner | 0 | 0 |
| outside owner | 16920 | 72417582246 |

The one exact `x86_sha::compress` frame has period **4,081,079**, but its stack contains neither the snapshot owner nor `timed_iteration`; it is retained as an outside-owner diagnostic. No software-backend frame was observed.

| backend | measured samples | setup samples | unclassified samples | outside samples | outside period |
| --- | ---: | ---: | ---: | ---: | ---: |
| `sha2::sha256::x86_sha::compress` | 0 | 0 | 0 | 1 | 4081079 |
| `sha2::sha256::soft::unroll::compress` | 0 | 0 | 0 | 0 | 0 |

The stream has **75** unknown frames across **50** samples and no explicit depth-limit marker. The backend candidate itself is an incomplete scope chain; the reader does not infer whether missing frames arose from sampling or DWARF unwinding.

| timestamp | lost events |
| --- | ---: |
| 1999339.120705889 | 47 |
| 1999340.278499403 | 36 |
| 1999341.221879010 | 64 |
| 1999344.893615408 | 99 |
| 1999345.063588022 | 26 |
| 1999348.214336960 | 32 |

The exact raw `perf.data`, decoded script, mmap/build-id exports, command receipt, and semantic artifacts remain bound in [`first-child-failure.json`](first-child-failure.json). The native attribution lane stops here; any GDB follow-up can establish reachability only and cannot replace this rejected perf result or create a timing claim.
