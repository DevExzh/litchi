# 0831 terminal results review

This is a bounded, read-only factual and scope review after comparative capture,
independent replay, adoption, and owned-target cleanup. I found no results
blocker.

The terminal identities are internally consistent: `analysis.json` has SHA-256
`aaffb3db12a945e24202a38f6edc643e459b5641a7b0e66fa76cc888278d5a58`,
`audit.json` has SHA-256
`888b3fa361d9f3ed414f3a416b76b045ed80be6538337a311d1a8a53177d68a9`, and
`decision.json` adopts the candidate with the same analysis and audit hashes.
Post-cleanup `analyze.py --check`, `audit.py --check`, and `decision.py` all
have successful retained receipts.

The complete nine-case table agrees with the analysis and independent audit:

| Case | Native p50 ratio | Requested bytes after minus before | Baseline-adjusted peak after minus before |
| --- | ---: | ---: | ---: |
| real-edit | 0.980416 | -524,288 | 0 |
| real-lifecycle | 0.997905 | -524,288 | 0 |
| one-cell-tiny | 0.942858 | -524,288 | -27,086 |
| one-cell-medium | 0.993495 | -524,288 | 0 |
| one-cell-dense-wide | 0.998447 | -524,288 | 0 |
| one-percent-tiny | 0.936734 | -1,048,576 | -21,326 |
| one-percent-medium | 0.996076 | -2,097,152 | 0 |
| one-percent-dense-wide | 0.999925 | -1,048,576 | 0 |
| noop-medium | 1.000000 | 0 | 0 |

The raw real-edit observer reports show 2,898,254 requested bytes and 2,510
allocation calls before, versus 2,373,966 and 2,509 after. Absolute region
peak falls from 1,130,221 to 1,130,218 bytes, while entry live bytes fall from
63,264 to 63,261, leaving peak above entry at 1,066,957 on both legs. The
real-lifecycle reports show the same three-byte absolute-peak and entry-baseline
shift, with unchanged 1,103,104 peak above entry. This matches the report's
baseline-adjusted diagnostic and does not support a real peak-memory claim.

The retained quality logs support the stated gates on both source legs: each
has 2,087 XLSX tests and three oracle tests passing; each harness run has 555
tests passing and one explicitly ignored opt-in security-corpus test. The
qualification, native, observer and capture counts are 36/36, 108/20,160,
36/108, and 180/20,304 reports/samples respectively. The 200 command receipts
decompose into 16 quality, 4 build, 36 qualification, and 144 comparative
capture commands.

The 20 spread flags and 10 p99/p50 diagnostics listed in the report match the
analysis and decision vectors. The wording keeps observer elapsed time,
lifecycle latency, RSS, real peak memory, pooled timing, cross-format behavior,
and nonempty-column-action latency outside the adoption claim. Synthetic peak
reductions are presented as diagnostics, while the real-edit decision uses the
requested-byte gate and preserved output oracle.

The archived source diff is exactly the three-line empty-action return after
the protected-sheet check and before `Assignments::new()`. Cleanup passed after
removing 11,308 target files (8,430,128,751 bytes) and the one marked scratch
file (38 bytes), with the independent readers retaining the binary descriptors.
