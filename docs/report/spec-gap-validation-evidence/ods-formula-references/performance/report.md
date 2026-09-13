# ODS formula reference parser performance

This is the final bounded parser capture for the ODS formula reference batch. The
baseline is commit `b7a66574a541ac0efc443056c4d11f6b829fffa9`; the candidate is
the gated source whose `formula.rs` digest is
`d276e8e9190751ab949ce902deae7bfd90cc00f4f4029d24330329d066209005`. Both
runs use the same standalone harness and comparable twelve-case input set.
Extended reference forms are measured in a separate candidate-only coverage
set.

## Method

The [standalone harness](harness/) parses through
`litchi_ods::codec::formula::FormulaParser` only. Each lane uses a fresh release
process, `--locked --offline`, `taskset -c 2`, three warmup batches, and fifteen
measured batches. Ordinary lanes execute 1,000 parser calls per batch. The
256-local-reference lane and 16 KiB component lanes execute 128 calls per batch.
The harness uses the instrumented `System` allocator; `/usr/bin/time -v` records
process maximum RSS.

The table normalizes elapsed time and allocator requested bytes by calls per
batch. Allocator-call counts are likewise per parser call. `Peak live delta` is
the raw maximum observed for a complete measured batch; it is a maximum and is
not divided by the call count. RSS includes process startup and static data.
Each result includes checksums and success/refusal counts in the raw receipt.

The measurement includes formula token construction, reference parsing, error
display, and destruction. It excludes package I/O, XML parsing, formula
evaluation, decompression, publication, URI resolution, networking, and Office
interoperability. Fifteen batches on a shared host provide engineering
observations rather than statistical tail bounds.

## Comparable results

All twelve baseline and candidate lanes exited successfully and matched their
expected success/refusal outcomes. Values are measured-batch p50s normalized per
parser call. The complete batch distributions and status evidence are in the
[baseline raw CSV](baseline/raw.csv) and [candidate raw CSV](candidate/raw.csv).

| Case | Input bytes | Calls/batch | Outcome | Baseline p50 ns/call | Candidate p50 ns/call | Change | Alloc calls/call | Requested bytes/call | Peak live delta (batch max bytes) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| parse-bracket-current-cell | 9 | 1000 | success | 136.31 | 172.61 | +26.6% | 3.00 -> 3.00 | 458.00 -> 458.00 | 458 -> 458 | 2244 -> 2240 |
| parse-bracket-sheet-cell | 18 | 1000 | success | 181.84 | 269.86 | +48.4% | 4.00 -> 5.00 | 473.00 -> 697.00 | 473 -> 697 | 2240 -> 2240 |
| parse-bracket-range | 29 | 1000 | success | 298.30 | 409.33 | +37.2% | 6.00 -> 7.00 | 488.00 -> 712.00 | 488 -> 712 | 2184 -> 2296 |
| parse-quoted-doubled-sheet | 19 | 1000 | success | 184.07 | 271.78 | +47.7% | 4.00 -> 5.00 | 474.00 -> 697.00 | 474 -> 697 | 2220 -> 2288 |
| parse-local-refs-256 | 1175 | 128 | success | 18111.48 | 23371.59 | +29.0% | 265.00 -> 265.00 | 115671.00 -> 115671.00 | 58775 -> 58775 | 2196 -> 2184 |
| parse-sum-unbracketed | 15 | 1000 | success | 245.33 | 289.98 | +18.2% | 5.00 -> 5.00 | 468.00 -> 468.00 | 468 -> 468 | 2136 -> 2192 |
| parse-vlookup-unbracketed | 26 | 1000 | success | 391.16 | 474.51 | +21.3% | 8.00 -> 8.00 | 3172.00 -> 3172.00 | 1828 -> 1828 | 2136 -> 2204 |
| parse-malformed-zero-row | 9 | 1000 | refusal | 205.15 | 174.73 | -14.8% | 5.00 -> 4.00 | 114.00 -> 101.00 | 88 -> 76 | 2252 -> 2192 |
| parse-malformed-missing-separator | 8 | 1000 | refusal | 163.06 | 265.35 | +62.7% | 4.00 -> 6.00 | 128.00 -> 138.00 | 104 -> 94 | 2248 -> 2192 |
| parse-malformed-unclosed-bracket | 8 | 1000 | refusal | 154.78 | 159.76 | +3.2% | 4.00 -> 4.00 | 112.00 -> 112.00 | 88 -> 88 | 2240 -> 2240 |
| parse-malformed-missing-row | 8 | 1000 | refusal | 201.67 | 221.78 | +10.0% | 5.00 -> 5.00 | 105.00 -> 139.00 | 80 -> 114 | 2248 -> 2284 |
| parse-malformed-bad-quote | 18 | 1000 | refusal | 158.51 | 162.73 | +2.7% | 4.00 -> 4.00 | 122.00 -> 122.00 | 88 -> 88 | 2256 -> 2296 |

The candidate has material latency regressions on the accepted rich-reference
lanes, the 256-reference lane, and the unbracketed controls. The malformed
missing-separator lane also regresses materially. In this paired capture, the
accepted unbracketed controls and 256-reference lane increase 18.2–29.0%;
bracketed reference forms increase 26.6–48.4%, with sheet-qualified, range, and
quoted-sheet forms at 37.2–48.4%. The range lane's maximum RSS rises from 2184 to
2296 KiB (+5.1%). The current-cell lane and 256-reference lane retain their
baseline allocation counts and requested bytes; sheet-qualified, range, and
quoted-sheet lanes add one allocation and increase requested bytes. These
measurements do not support a speedup claim.

## Candidate-only coverage

These sixteen lanes exercise syntax deliberately absent from the committed
baseline, so their timings are not a baseline comparison. Every valid lane
succeeded on every measured batch; the over-limit IRI lane refused every call.

| Case | Input bytes | Calls/batch | Outcome | Candidate p50 ns/call | Alloc calls/call | Requested bytes/call | Peak live delta (batch max bytes) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: |
| coverage-source-cell | 28 | 1000 | success | 316.08 | 5.00 | 717.00 | 717 | 2204 |
| coverage-source-range | 32 | 1000 | success | 372.19 | 6.00 | 722.00 | 722 | 2244 |
| coverage-empty-source | 12 | 1000 | success | 212.80 | 4.00 | 685.00 | 685 | 2200 |
| coverage-unicode-escaped-source | 49 | 1000 | success | 441.40 | 5.00 | 758.00 | 758 | 2212 |
| coverage-whole-columns | 11 | 1000 | success | 256.59 | 5.00 | 685.00 | 685 | 2200 |
| coverage-whole-rows | 11 | 1000 | success | 179.97 | 3.00 | 683.00 | 683 | 2212 |
| coverage-cross-sheet-range | 25 | 1000 | success | 363.11 | 6.00 | 487.00 | 487 | 2296 |
| coverage-nested-inherited | 21 | 1000 | success | 406.76 | 8.00 | 741.00 | 741 | 2296 |
| coverage-ref-error | 11 | 1000 | success | 138.61 | 3.00 | 683.00 | 683 | 2208 |
| coverage-colon-sheet-1k | 1033 | 1000 | success | 1430.31 | 4.00 | 2506.00 | 2506 | 2208 |
| coverage-colon-sheet-4k | 4105 | 1000 | success | 5012.78 | 4.00 | 8650.00 | 8650 | 2208 |
| coverage-colon-sheet-16k | 16393 | 128 | success | 20991.89 | 4.00 | 33226.00 | 33226 | 2212 |
| coverage-source-iri-1k | 1036 | 1000 | success | 4687.57 | 5.00 | 2733.00 | 2733 | 2200 |
| coverage-source-iri-4k | 4108 | 1000 | success | 17698.26 | 5.00 | 8877.00 | 8877 | 2248 |
| coverage-source-iri-16k | 16396 | 128 | success | 69510.30 | 5.00 | 33453.00 | 33453 | 2208 |
| coverage-source-iri-over-16k | 16397 | 128 | refusal | 33565.23 | 6.00 | 16657.00 | 16437 | 2208 |

Input-size-sensitive lanes scale with decoded component length. The 16 KiB
source IRI lane costs about 69.5 microseconds per call under this run; the
16,385-byte refusal still scans and accounts for bounded input before refusal.
The coverage [raw CSV](coverage/raw.csv), per-lane stdout/status/time files, and
[coverage group manifest](coverage/group.json) retain the exact inputs and
outcomes.

## Provenance and reproduction

The [baseline source digests](baseline/source-sha256.txt),
[baseline harness digests](baseline/harness-sha256.txt), and
[baseline binary digest](baseline/binary-sha256.txt) identify the committed
parser and standalone executable. The [candidate source digests](candidate/source-sha256.json),
[candidate harness digests](candidate/harness-sha256.txt), and
[candidate binary digest](candidate/binary-sha256.txt) identify the gated source
and executable. The [candidate patch](candidate.patch) applies cleanly to a
fresh checkout of the baseline; its replay result is recorded in
[patch replay receipt](candidate/patch-replay.status). Isolated candidate checks are recorded in
[isolated candidate checks](candidate/isolated-checks.json) (47 formula unit tests, 10 reference integration
tests, and seven existing function integration tests).

Build commands, source status, run timestamps, checksums, stdout/stderr, and
`/usr/bin/time -v` reports are retained in the [baseline](baseline/),
[candidate](candidate/), and [coverage](coverage/) directories. The final [candidate harness digest receipt](candidate/harness-sha256.txt) and
[environment](environment.json) record the runner and host. The parent gate
receipt records 810 tests across 47 targets, warning-denied Clippy and rustdoc,
doctests, formatting, and final source hashes in the [gate receipt](../gates/results.json).
Temporary isolated worktrees and release targets were removed after the binary
digests, patch replay, and repository verifier passed; the cleanup record is in
[cleanup.json](cleanup.json).

Run the repository verifier from the checkout:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-references/verify.py
```

It verifies source digests against the gate receipt, patch replay status,
isolated test counts, every retained CSV field against its stdout/time/status
receipt, expected outcomes, and the final artifact manifest.
