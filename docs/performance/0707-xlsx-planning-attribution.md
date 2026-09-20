# 0707 — attribute XLSX source-backed planning before choosing a change

`performance_claim: none`

This diagnostic measures the current source-backed XLSX one-percent scalar-cell
edit/save route at revision `a664c8539894de38fea670bb34eaa66cc69058ce`. It
does not change production or harness source and makes no before/after
speedup claim. It follows the refreshed 0705 phases and the rejected 0706 XML
auditor pilot. The purpose is to attribute current planning work before
selecting another candidate.

The selector uses an instrumented in-memory positional provider, selects all
four worksheets, and runs over deterministic medium and dense-sparse workbooks.
The native interval includes open, planning, staging plus commit, and
sequential publication, including the returned publication snapshot drop. It
excludes sink setup, remaining handle destruction, reopen, and semantic or
preservation oracles. The logical read counters describe the provider and do
not measure physical storage or decompression. All current source, corpus,
output, sink, and resource identities are retained in the packet.

## Current native timing

Each child is pinned to CPU 12 and retains 100 samples after 20 warmups. The
second repeat reverses shape order. Shares use the matched phase means over
the native elapsed mean; they are not sums of percentile shares.

| Shape / repeat | Total p50 ms | p95 ms | p99 ms | Open share | Planning share | Commit share | Publication share |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| medium / 1 | 16.511 | 16.638 | 16.965 | 0.49% | 36.78% | 21.59% | 41.14% |
| dense-sparse / 1 | 32.236 | 32.449 | 32.637 | 0.27% | 35.82% | 21.20% | 42.72% |
| medium / 2 | 16.524 | 16.599 | 16.647 | 0.49% | 36.88% | 21.58% | 41.06% |
| dense-sparse / 2 | 32.168 | 32.391 | 32.612 | 0.27% | 35.75% | 21.26% | 42.72% |

Publication remains the largest measured phase, while planning is the next
largest. No native total metric has a repeat shift over 5%. The retained
same-build A/A flags are:

| Shape | Metric | Repeat 1 | Repeat 2 | Shift |
| --- | --- | ---: | ---: | ---: |
| medium | planning p99 | 6.552 ms | 6.180 ms | −5.68% |
| dense-sparse | reopen p50 | 37.068 ms | 40.220 ms | +8.50% |
| dense-sparse | reopen p95 | 37.750 ms | 40.862 ms | +8.24% |
| dense-sparse | reopen p99 | 38.047 ms | 41.044 ms | +7.88% |
| dense-sparse | reopen mean | 37.131 ms | 40.328 ms | +8.61% |

Reopen is outside the timed native edit/save interval. These are repeat-drift
diagnostics from one source build, not candidate regressions. The sample size
does not establish cross-host or universal tail behavior, and no causal
explanation is assigned to the shifts.

## Allocation attribution

The allocator lane uses a separate binary, two repeats, three samples, and two
warmups per shape. The table reports operation-region relative metrics. Six
relative metrics are retained for every sample; absolute process gauges are
kept separately and are not treated as candidate-comparable operation
metrics. Region peaks are not summed.

| Shape / region | Allocation calls | Requested bytes | Peak above region start | Net live change |
| --- | ---: | ---: | ---: | ---: |
| medium / plan | 67,845 | 12,139,796 | 3,191,768 | +2,200,064 |
| medium / staging | 389 | 1,899,381 | 28,521 | +12,227 |
| medium / commit core | 42,383 | 3,205,114 | 2,058,609 | +2,046,254 |
| medium / commit | 42,772 | 5,104,495 | 2,070,836 | +2,058,481 |
| medium / publication | 19,197 | 4,126,154 | 1,095,243 | −15,633 |
| dense-sparse / plan | 129,411 | 15,850,404 | 7,602,062 | +4,200,215 |
| dense-sparse / staging | 725 | 2,916,044 | 35,828 | +25,944 |
| dense-sparse / commit core | 80,110 | 5,968,832 | 3,931,745 | +3,902,769 |
| dense-sparse / commit | 80,835 | 8,884,876 | 3,957,689 | +3,928,713 |
| dense-sparse / publication | 36,573 | 5,747,529 | 1,604,554 | −15,633 |

Planning has the largest allocation count and requested-byte volume in this
matrix. This identifies a current owner for follow-up; it does not equate
allocation count with native latency or predict a candidate result. Allocator
instrumented elapsed time is excluded from native conclusions.

## Planning profile and candidate boundary

Four CPU-pinned Callgrind children retain three lifecycle dumps and one
measured planning dump per shape and repeat, plus the termination dump. The owner is
`litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets`; the
measured dump is selected by its positive incoming edge from
`run_xlsx_cell_values_edit_save`. Lifecycle calls are selected separately from
`run_xlsx_cell_value_lifecycle_gates`. The raw numbered dumps are retained in
the packet, and the independent analysis verifies that the immediate children
form a disjoint partition without adding nested inclusive instruction totals.

The measured profile rows are:

| Shape / repeat | Owner Ir | Worksheet parse Ir / share | Validator observe Ir / share | `validate_element` Ir / share |
| --- | ---: | ---: | ---: | ---: |
| medium / 1 | 109,457,190 | 105,471,637 / 96.36% | 20,126,878 / 18.39% | 9,374,579 / 8.56% |
| dense-sparse / 1 | 206,110,295 | 199,478,509 / 96.78% | 38,638,620 / 18.75% | 18,105,553 / 8.78% |
| medium / 2 | 109,513,690 | 105,528,465 / 96.36% | 20,126,529 / 18.38% | 9,374,676 / 8.56% |
| dense-sparse / 2 | 206,570,841 | 199,940,928 / 96.79% | 38,638,348 / 18.70% | 18,105,561 / 8.76% |

`worksheet_xml_and_parse_source` is an inclusive combined validation/parser
owner and contains the nested Validator, FactsBuilder, raw parser, and Store
diagnostics. Those rows overlap and must not be added to the owner or its
immediate-child partition. Callgrind Ir is guest-instruction attribution only,
not native latency, hardware cycles, allocation counts, RSS, or cache behavior.
The profile process uses Valgrind's allocator replacement, so allocator edge
counts from these files are not treated as native allocation measurements.
In particular, dense raw allocator-edge call fields exceed the separate
planning allocation counter. That discrepancy remains unresolved; this packet
does not use those fields to quantify validator allocation requests.

The profile identity and output/corpus/source parity checks pass for all four
measured rows. The profile result is current-source attribution, not a
before/after comparison and not a speedup claim.

Source review identifies a bounded private experiment for the next batch. The
current validator copies each Start local name into a boxed byte slice and
copies each Empty local name into a temporary vector. A possible private
representation could use static modeled names, borrow Empty names only for the
callback, and retain unknown names in owned storage. It must preserve exact
close-name matching, first-error order, strict and transitional dialect rules,
unbound-prefix refusals, copied-subtree bytes, depth limits, allocation
failure handling, and authoritative fallback. It must not introduce a global
cache, public API, parser fusion, or proof reuse. The full test and measurement
matrix is recorded in [`next-experiment.md`](results/change-0707/next-experiment.md).

No candidate has been implemented, built, compared, or admitted in this
batch. A future candidate requires a frozen plan, fresh native A/A and A/B/B/A
measurements, separate allocation and profile evidence, exact output/source/
resource parity, focused validation coverage, and review of every over-5%
regression and repeat-drift flag.

## Verification and limits

The baseline native and allocator builds use an exact 7,282-entry source
census and are bound to the packet's binary, plan, script, constraint, and
receipt hashes. The packet retains the corrected capture preflight record; no
child was overwritten or discarded. Its independent analysis verifies native
phase reconciliation, corpus/output/source/sink/resource identity, allocator
relative vectors, and the separation of absolute process gauges.

This batch adds no RSS capture, physical cold-storage test, native Office
producer, parallel-scaling test, cross-platform result, or universal workload
claim. iWork remains excluded. The broader non-iWork GOAL remains active.

All four profile stderr logs report Valgrind `brk segment overflow`. The
children exit successfully and output identities match, but this remains an
instrumentation environment limitation. No native allocator-cost claim or
causal explanation for the call-field discrepancy is drawn from those profiles.
