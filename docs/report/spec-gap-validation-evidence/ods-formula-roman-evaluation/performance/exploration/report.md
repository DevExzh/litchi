# ROMAN/ARABIC formula evaluation performance

This is bounded evidence for the ROMAN/ARABIC formula-function addition. It
is not a broad performance claim. The exact baseline is
`d0c1ca700ceda177c27a820acffe0de534720390`; it has no ROMAN/ARABIC
implementation, so candidate-only rows are absolute measurements rather than
before/after comparisons. The performance disposition remains **open**: the
four-round ABAB replay leaves six comparable p50 regressions over the 5%
review trigger, and a scratch-only variant is under investigation. No root
cause or acceptance conclusion is inferred from these timings.

## Capture and custody

The harness used three warmups and 15 measured iterations, with fixed cases at
128 repeats and scale cases at 64/32/8/2 repeats. Measured children were pinned
to CPU 2 with `taskset`; the host had no hard latency isolation and unrelated
Cargo activity targeted `/home/zhuhe/litchi-goal-0557-target`. A separate
benchmark at `/home/zhuhe/litchi-goal-0557-target/retained/baseline/normal` was
observed pinned to CPU 2 at about 96% CPU during the window, with greater than
300% host noise. Treat these CPU 2 receipts as provisional diagnostic evidence,
not a clean isolated baseline; root is probing preserved binaries on CPU 6.
The environment snapshot and method boundaries are in [environment.json](environment.json)
and [harness-review.md](harness-review.md).

The baseline profile binary is 1,305,776 bytes, SHA-256
`5d092ebb4d4d79236c5950c686955a86c8b1d39a309e23ccb1ae24d8192e94be`; the
candidate is 1,316,664 bytes, SHA-256
`dedd2e26816965572cb5a9367622118195ada202b5fdd0546d50ddb02cae4e84`.
Full source custody is in [baseline/source-sha256.json](baseline/source-sha256.json)
and [candidate/source-sha256.json](candidate/source-sha256.json); key candidate
hashes are `evaluation.rs`
`4263a0acd4e75e257d98ed1806d7c6f1f91e0cb5d797d8bce22ae1741f7c1296`,
`evaluation/roman.rs`
`3441233624f68e5efc72157953d3b7137bbb86f6c9835fd809eee40674342da1`, and
the Roman test `b67a32a5a06371ed24c5df6af92cf0c4bb8488367e6e3739dd9d3361df712b7b`.

## Candidate-only absolute lanes

The table reports the `evaluate` phase. `p50 ns/batch` is the measured batch
elapsed time; `ns/op` is only the displayed `p50` divided by the declared
repeat. `requested` and `peak-live` are harness allocator/accounting values,
while RSS is process-wide `/usr/bin/time -v` maximum RSS.

| Case | Input bytes | Repeat | p50 ns/batch | Derived ns/op | Requested bytes | Peak-live bytes | RSS KiB | Outcome |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `roman-3888-format-0` | 14 | 128 | 419,701 | 3,278.9 | 65,408 | 463 | 2640 | none |
| `roman-3888-format-4` | 14 | 128 | 600,743 | 4,693.3 | 64,512 | 456 | 2680 | none |
| `roman-499-format-0` | 13 | 128 | 163,751 | 1,279.3 | 64,256 | 454 | 2632 | none |
| `roman-499-format-4` | 13 | 128 | 599,103 | 4,680.5 | 63,744 | 450 | 2652 | none |
| `roman-998-format-0` | 13 | 128 | 224,941 | 1,757.4 | 64,512 | 456 | 2680 | none |
| `roman-998-format-4` | 13 | 128 | 601,233 | 4,697.1 | 63,872 | 451 | 2652 | none |
| `arabic-uppercase` | 26 | 128 | 57,840 | 451.9 | 51,200 | 400 | 2684 | none |
| `arabic-indirect` | 22 | 128 | 610,123 | 4,766.6 | 64,512 | 456 | 2692 | none |
| `arabic-input-64` | 75 | 64 | 33,910 | 529.8 | 25,600 | 400 | 2552 | none |
| `arabic-input-4096` | 4107 | 2 | 13,230 | 6,615.0 | 800 | 400 | 2808 | none |
| `roman-concat-64` | 576 | 64 | 2,112,270 | 33,004.2 | 424,832 | 3,489 | 2684 | none |
| `roman-concat-4096` | 36864 | 2 | 4,045,609 | 2,022,804.5 | 811,612 | 201,057 | 4220 | none |
| `roman-refusal-work` | 14 | 128 | 34,860 | 272.3 | 33,792 | 264 | 2668 | resource-work |
| `roman-refusal-memory` | 14 | 128 | 585,423 | 4,573.6 | 68,608 | 488 | 2644 | resource-memory |
| `roman-refusal-stack` | 14 | 128 | 18,760 | 146.6 | 27,648 | 216 | 2644 | resource-objects |
| `roman-refusal-cancelled` | 14 | 128 | 840 | 6.6 | 0 | 0 | 2676 | cancelled |

The candidate-only corpus contains 41 cases across parse, evaluate, and
parse-evaluate (123 rows). It also covers all five ODF ROMAN format vectors
for 3888, 499, and 998, zero/truncation/logical mappings, nested ARABIC,
bounds and typed refusal/cancellation paths. The complete data is in
[candidate/roman/raw.csv](candidate/roman/raw.csv); all 123 rows completed
with status 0.

## Comparable controls and repeated flags

The 159 baseline and 159 candidate comparable rows matched exactly on the
retained semantic and allocator fields: status/failure, success/refusal
counts, checksum, allocation/deallocation counts, requested/released bytes,
live-before/live-after, tracked peak-live bytes, and output reservation. RSS
is excluded from that heap-parity statement because it is process-wide and
noise-sensitive. Raw controls are [baseline/comparable/raw.csv](baseline/comparable/raw.csv)
and [candidate/comparable/raw.csv](candidate/comparable/raw.csv).

The initial selection had 41 flags ([initial-flags.json](initial-flags.json)).
The four-round ABAB replay has 328 rows ([abab/raw.csv](abab/raw.csv)) and its
summary is [abab/summary.json](abab/summary.json). The following table lists
every replay lane with any `|delta| > 5%` in p50, p95, p99, or max RSS. A dash
means that metric stayed within 5%; no repeated max-RSS metric crossed the
threshold.

| Phase | Case | Repeat | p50 delta | p95 / p99 delta | Max RSS delta |
| --- | --- | ---: | ---: | ---: | ---: |
| `parse` | `control-utf8-left-256` | 32 | -7.39% | -7.76% / -7.76% | — |
| `parse` | `control-false` | 128 | — | -19.85% / -19.85% | — |
| `parse` | `failure-work` | 128 | — | +37.53% / +37.53% | — |
| `parse` | `failure-cancelled` | 128 | — | +21.25% / +21.25% | — |
| `parse` | `logical-and-64` | 64 | — | +6.65% / +6.65% | — |
| `parse` | `text-if-escaped` | 128 | — | -23.79% / -23.79% | — |
| `parse` | `bitwise-rshift-4096` | 2 | — | -6.55% / -6.55% | — |
| `parse` | `bitwise-error-shift` | 128 | — | -25.53% / -25.53% | — |
| `parse` | `radix-base-small` | 128 | — | -27.37% / -27.37% | — |
| `parse` | `radix-base-max` | 128 | — | -21.96% / -21.96% | — |
| `evaluate` | `control-utf8-left-256` | 32 | — | +6.08% / +6.08% | — |
| `evaluate` | `control-utf8-left-1024` | 8 | +8.40% | +34.96% / +34.96% | — |
| `evaluate` | `control-utf8-left-4096` | 2 | +11.25% | +30.01% / +30.01% | — |
| `evaluate` | `control-true` | 128 | — | +7.06% / +7.06% | — |
| `evaluate` | `control-false` | 128 | — | -11.03% / -11.03% | — |
| `evaluate` | `failure-array` | 128 | — | +19.51% / +19.51% | — |
| `evaluate` | `logical-and-64` | 64 | — | -9.25% / -9.25% | — |
| `evaluate` | `bitwise-coerce-text` | 128 | +10.05% | +12.40% / +12.40% | — |
| `evaluate` | `bitwise-lazy-selected` | 128 | — | +17.04% / +17.04% | — |
| `evaluate` | `radix-decimal-small` | 128 | — | +7.28% / +7.28% | — |
| `parse-evaluate` | `control-coerce-4096` | 2 | +5.16% | +5.35% / +5.35% | — |
| `parse-evaluate` | `control-utf8-left-1024` | 8 | +6.69% | +14.69% / +14.69% | — |
| `parse-evaluate` | `control-utf8-left-4096` | 2 | +8.78% | -11.70% / -11.70% | — |
| `parse-evaluate` | `control-escaped-4096` | 2 | — | +6.33% / +6.33% | — |
| `parse-evaluate` | `failure-stack` | 128 | — | +8.41% / +8.41% | — |

The six repeated p50 regressions are the two `evaluate` UTF-8-left lanes
(+8.40% at 1024 and +11.25% at 4096), `evaluate/bitwise-coerce-text`
(+10.05%), `parse-evaluate/control-coerce-4096` (+5.16%), and the two
`parse-evaluate` UTF-8-left lanes (+6.69% at 1024 and +8.78% at 4096). These
are observations on this shared host; the scratch variant investigation is
separate from the captured candidate and supplies no causal attribution.
The remaining table entries are tail quantile flags from only 15 samples;
p95 and p99 often coincide with the sample maximum.

## Whole-process hardware counters

Five retained `perf stat` captures used four events (`cycles`,
`instructions`, `branches`, `branch-misses`) and all exited with status 0.
Counts include process setup, preflight, warmups, the timed repetitions, and
output. They are not operation-only counters and do not establish CPU
causality.

| Capture | Cycles | Instructions | Branches | Branch misses |
| --- | ---: | ---: | ---: | ---: |
| baseline text-if-concat | 6,358,134,768 | 11,624,853,882 | 1,964,277,222 | 358,724 |
| candidate text-if-concat | 6,478,399,101 | 11,624,734,196 | 1,964,257,251 | 1,998,591 |
| candidate ROMAN 3888 format 0 | 2,652,310,785 | 7,575,817,883 | 1,347,170,395 | 3,650,648 |
| candidate ROMAN 3888 format 4 | 3,774,535,507 | 11,408,028,204 | 1,743,375,966 | 637,875 |
| candidate ARABIC input 4096 | 538,667,820 | 3,091,138,222 | 601,364,902 | 83,233 |

For the common `text-if-concat` control, candidate versus baseline changes are
+1.89% cycles, -0.001% instructions,
-0.001% branches, and +457.14% branch misses.
This is whole-process counter evidence only; it is not a latency or CPU
speedup claim. Because CPU 2 was contended during the retained window, these
receipts remain provisional until the CPU 6 probe is resolved. Receipts are in
[perf-stat/](perf-stat/).

## Reproduction and limits

Use the retained [baseline build/provenance](baseline/binary-provenance.json),
[candidate build/provenance](candidate/binary-provenance.json), and the two raw
CSV files for the comparable A/B view. The gate receipt is
[../gates/results.json](../gates/results.json). Direct use of `run.py` should
be followed by checking every raw-row status; the capture wrapper performed
that check for these receipts. Timer boundaries, allocator definitions, exact
corpus counts, and the shared-host limitation are recorded in
[harness-review.md](harness-review.md). Candidate-only absolute results should
not be generalized beyond these bounded fixtures, and the current repeated
p50 trigger remains unresolved.
