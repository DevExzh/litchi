# 0829 results review

This is a static review of the profile report against the retained analysis and
audit records. No reader, numerical replay, workload, or other execution was
run for this review.

The report and retained evidence agree on the primary counts:

| Measure | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| Whole process | 5,861 | 5,866 |
| Exact edit owner | 2,670 | 2,685 |
| Outside exact owner | 3,191 | 3,181 |
| Capture | 1,196 | 1,193 |
| Set-text | 607 | 612 |
| Commit/application | 867 | 879 |
| Unclassified | 0 | 1 |
| Ambiguous | 0 | 0 |

The machine-readable phase records conserve the same owner and phase counts;
the sole unclassified sample is the reported
`Transaction::set_shape_text` leaf. The retained paired estimates round to
`1.000715` for wrapped/direct and `1.023752` for frame-pointer/wrapped, with
the report's intervals matching `root-audit.json`.

The claim boundaries also match. `analysis.json` marks the phase partition as
exhaustive and mutually exclusive while omitting timing-additivity, causal
fraction, and Amdahl claims. `frame-audit.json` has the same contract. The
report explicitly describes phase rows as CPU samples rather than operations,
milliseconds, wall-clock shares, or an Amdahl model.

The independent diagnostics are consistent with the report: all seven reader
receipts have exit code 0; frame audits report no empty, malformed, lost-event,
or explicit-truncation markers; stack diagnostics report the same 0/0 empty
stacks, 0 lost events, 0 malformed frames, 0 explicit truncations, unknown
stack endings of 32/30, and maximum depths of 28/26. The retained limitation
that complete unwinding is not proven is present. The source opportunity review
also matches the report: the duplicate-catalog and raw-Scene-span ideas remain
hypotheses, the XLSX investigation is separate, and no candidate or durability
weakening is adopted.

There is one final-sealing sequencing issue. At this review point the packet
contains `cleanup.py` and `seal.py`, but not their generated `cleanup.json` or
`seal.json`. The report and README currently say cleanup and final-seal
verification are recorded. Root should perform the planned cleanup and final
seal (or make that wording pending) before treating those statements as sealed;
this does not require any metric or scope change.

No other blocker was found in the report claims or retained evidence.

Root completion note: target cleanup subsequently removed 3,688 files /
2,054,144,787 logical bytes after exact verification of both executables.
Post-cleanup final validation passed as reader attempt 7. All eight retained
reader attempts pass; the sequencing issue above is resolved before sealing.
