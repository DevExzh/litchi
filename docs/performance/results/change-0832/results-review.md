# 0832 results review

This review covers the retained parent packet, the supplemental promotion guard,
the additive reader correction, and the root adoption decision. The experiment
is bound to base `7eeaab48c0281b06f53527d4c4f4ea79050d27e7`. The parent analysis
and independent audit both report `pass`, match each other, and retain 180
reports with 20,304 samples including qualification. The supplemental analysis reports
`pass` for 48 reports and 12,040 samples with an empty `regression_flags` list.

## Parent result

For the pinned real edit, midpoint observer summaries report requested
allocation bytes falling from 2,373,966 to 276,494, a reduction of 2,097,472
bytes (88.35%). Allocation calls fall from 2,509 to 2,505. Raw region peak
falls from 1,130,221 to 89,604 bytes, while net live bytes are unchanged. The
native paired p50 ratio is 0.895998, with the six-block bootstrap interval
0.887658–0.898114. The paired RSS ratio is 1.000174, with interval
0.999346–1.000596.

The real lifecycle case retains the same 2,097,472-byte requested allocation
reduction. Its native p50 ratio is 0.990366 (interval 0.981989–0.996211) and
its paired RSS ratio is 0.999753 (interval 0.999419–1.000364). The remaining
fixed cases have no frozen native p50 lower-confidence-bound or RSS adoption
flag. `noop-medium` has a descriptive p50 estimate of 1.017549, but its lower
bootstrap bound is 0.991379 and therefore does not meet the 1.05 gate.

The parent qualification oracle is admitted and checked across all 36
reports. The pinned full-output oracle and lifecycle publication checks pass;
timed edit samples retain admitted edit outcomes without serializing every
timed result.

## Promotion guard

The guard covers two generated derivatives of the pinned workbook: a
disjoint two-record sheet and a 128-record nested wide sheet with complete
modeled properties. Fixture replay found ten ZIP members in each archive and
confirmed compressed and uncompressed payload equality for every member except
`xl/worksheets/sheet1.xml`.

Lifecycle qualification produced one output hash shared by both source legs
and both binaries for each variant:

| Variant | Shared lifecycle output SHA-256 |
| --- | --- |
| `disjoint-two-records` | `785c924300549de8a72a9bbc08cc54d63ec26ca4a637325a06e5d2c91c9782fc` |
| `overlap-128-wide-complete` | `adfc359c926c6faf3a9f8bdeb83c5170e701cf0b7b87dfcbb412778606b5109b` |

The native guard p50 and RSS comparisons remain within the frozen gates:

| Variant | p50 ratio and bootstrap interval | RSS ratio and bootstrap interval |
| --- | --- | --- |
| `disjoint-two-records` | 1.001557 [0.978550, 1.010032] | 0.999419 [0.999070, 1.000829] |
| `overlap-128-wide-complete` | 1.001599 [0.998202, 1.005419] | 1.000160 [0.999942, 1.000698] |

The observer lane reports zero change in allocation calls, requested bytes,
entry-adjusted peak, and net live bytes for both variants. Raw region peak is
three bytes lower in both A/B block comparisons. Representative raw reports
show zero failed allocation calls and bind the before/after binary hashes to
the frozen release binaries. No promotion regression flags were produced.

Promotion intervals use nearest-rank within-process quantiles, midpoint medians
across six paired blocks, and 10,000 paired bootstrap resamples with seed
`832128` and ranks 250/9749. They are per-variant summaries; they do not form
an aggregate multi-record or cross-format result.

## Correction and decision custody

The first two guard admission attempts failed before comparative capture because
the frozen reader referenced two omitted globals, `SOURCES` and `write_once`.
Attempts 06 and 08 retain their receipts and logs. The v2 correction supplies
only those two globals through `replay_reader_v2.py`, imports and runs the
unchanged frozen reader, statically verifies that these are its only missing
globals, and binds itself to the previous correction, the frozen reader,
freeze, qualification receipt, and both failure records. It changes no report
parser, workload, statistic, threshold, or frozen guard script. Corrected
qualification admission and corrected analysis replay both pass, with their
hash receipts retained.

The root decision is `adopt: true` with an empty regression flag list. It
checks the parent analysis and audit schemas, base, cardinalities, qualification
oracle state, promotion analysis schema and counts, lifecycle equality, frozen
reader hash, and correction receipts. The decision also retains the parent
entry-adjusted peak diagnostic. Its `spread_flags` and `tail_flags` remain
descriptive variance evidence and are not silently treated as adoption
regressions.

## Limits and unresolved risks

The evidence supports this candidate for the named pinned real edit and the
two controlled promotion fixtures. It does not establish a general
multi-producer distribution, nonempty-column-action latency result, cross-format
speedup, cache-miss result, or observer-lane latency result. Process RSS
includes harness setup. Guard lifecycle publications are verified and then
deleted, so the guard has no retained output archive or independent Office
validation; the parent pinned fixture supplies the independent full-byte oracle.

The source review records that dense-map reservation has no injected fault
test. Promotion remains locally staged and unpublished until replay succeeds,
and the existing parser, writer, transaction, allocation, quality, and output
gates pass. This is a documented verification limitation rather than a failed
adoption gate.

No blocking issue was found in this results review. Final cleanup, replay after
cleanup, sealing, and the exact owned-path commit remain required before the
batch is complete.
