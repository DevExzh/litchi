# Const-specialized scalar-cell candidate: captures 04 and 05

These are bounded diagnostic A/B captures for the current const-specialized
scalar mode branch. They are not a full performance acceptance result. Each
capture used the preserved before ELF and the same candidate ELF, CPU 6,
three warmups, 31 measured iterations, and serial AB/BA/AB ordering. The
`p50_ns` and `p95_ns` values are raw timed batch values; the harness repeat is
inside each batch.

[`scalar-cell-pairs-04.tar.gz`](diagnostics/scalar-cell-pairs-04.tar.gz)
(SHA-256 `5e896e7f702d7aea418358ca2443b95298660428d1963167f2693f338498a516`)
and [`scalar-cell-pairs-05.tar.gz`](diagnostics/scalar-cell-pairs-05.tar.gz)
(SHA-256 `df8f8bd963e87e60f007b611cb402a946365124830416f6fabea342a64d4f088`)
each contain 162 successful child processes, 81 pairs, and zero deterministic
mismatches. The candidate ELF is SHA-256
`844120ca99cca071daa5e2377cfa6b9953f94db0c95aebe64f395a3071f88eba`; the
candidate evaluator source closure records `value.rs` as
`f191993cb9761d606664b7f9aacb575b84fde5bce0322a80498f4a6ffa911c0a`.

## Scalar lanes

| case | capture 04 p50 delta, AB / BA / AB | capture 05 p50 delta, AB / BA / AB |
|---|---:|---:|
| `reference-repeat-1` | −44.39% / −44.08% / −46.30% | −45.43% / −46.11% / −41.84% |
| `reference-repeat-16` | −54.49% / −55.18% / −54.03% | −55.86% / −54.37% / −53.74% |
| `reference-repeat-256` | −60.04% / −58.74% / −61.04% | −58.82% / −60.15% / −57.97% |
| `reference-repeat-4096` | −60.88% / −59.58% / −60.03% | −59.99% / −60.15% / −58.35% |
| `reference-distinct-1` | −46.77% / −45.85% / −44.56% | −44.51% / −44.75% / −45.75% |
| `reference-distinct-16` | −54.65% / −55.07% / −54.40% | −55.70% / −56.12% / −55.11% |
| `reference-distinct-256` | −60.33% / −59.81% / −59.15% | −58.56% / −58.02% / −59.18% |
| `reference-distinct-4096` | −57.99% / −59.48% / −58.51% | −58.67% / −59.24% / −60.16% |

The 4,096-cell scalar allocation result remains 12,304 → 16 calls and
2,163,192 → 524,792 requested bytes for both repeated and distinct references;
read, work, checksum, borrowed-text, copied-byte, and result parity stayed
exact. The practical scalar gains survive the simpler const-mode branch.

## Control review flags

The GOAL 5% review trigger is applied to positive candidate deltas. The table
lists every control row above 5% in p50 or p95, or in process maximum RSS;
RSS entries include absolute KiB before → after values.

| capture/round | control | p50 >5% | p95 >5% | RSS >5% (KiB) |
|---|---|---:|---:|---:|
| 04/AB-1 | `reference-lazy` | — | +10.21% (73,940→81,490) | — |
| 04/AB-1 | `reference-limit-work` | — | +9.51% (42,580→46,630) | — |
| 04/AB-1 | `reference-empty-arithmetic` | — | — | +8.24% (2,912→3,152) |
| 04/BA-2 | `matrix-lazy-inline-4096` | +6.47% (1,929,158→2,054,048) | +6.83% (1,935,368→2,067,559) | — |
| 04/BA-2 | `reference-range-1` | +6.53% (190,091→202,501) | — | — |
| 04/BA-2 | `reference-background-4096` | — | — | +7.44% (3,708→3,984) |
| 04/AB-3 | `matrix-lazy-inline-4096` | +5.64% (1,919,877→2,028,158) | +6.53% (1,934,288→2,060,538) | — |
| 04/AB-3 | `reference-limit-cells` | — | +10.59% (42,200→46,670) | — |
| 05/AB-1 | `reference-range-1` | +6.63% (190,560→203,201) | +7.92% (196,041→211,561) | +7.37% (2,876→3,088) |
| 05/AB-1 | `reference-empty` | — | — | +5.03% (2,940→3,088) |
| 05/AB-1 | `reference-range-256` | — | — | +7.79% (3,132→3,376) |
| 05/BA-2 | `reference-lazy` | — | +5.44% (77,800→82,030) | — |
| 05/BA-2 | `reference-cancelled` | — | +5.47% (1,280→1,350) | +6.71% (2,684→2,864) |
| 05/BA-2 | `matrix-lazy-aggregate-4096` | — | — | +5.46% (5,712→6,024) |
| 05/BA-2 | `reference-range-1` | — | — | +8.79% (2,912→3,168) |
| 05/BA-2 | `reference-range-256` | — | — | +6.78% (2,952→3,152) |
| 05/AB-3 | `reference-limit-cells` | — | — | +6.06% (2,904→3,080) |
| 05/AB-3 | `reference-range-16` | — | — | +8.76% (2,876→3,128) |

For `reference-range-1`, p50 deltas across all rounds were +1.47% / +6.53%
/ +0.53% in capture 04 and +6.63% / +0.27% / +4.78% in capture 05. The
04 `matrix-lazy-inline-4096` p50 deltas were +4.08% / +6.47% / +5.64%; no
capture-05 round exceeded 5% for that p50 metric. The flags vary between windows; they do not establish that the regressions
have been eliminated. They have no CPU
causal attribution here; RSS is process-level and timing is shared-host data.

The scalar allocation and result evidence is useful for the narrow
optimization, while acceptance remains open pending review of the control
flags and further independent evidence.

Three scalar-lane RSS observations also exceeded 5%: capture04/AB-3
`reference-repeat-1` was 2,948 → 3,140 KiB (+6.51%),
`reference-repeat-16` was 2,936 → 3,208 KiB (+9.26%), and capture05/AB-1
`reference-distinct-1` was 2,888 → 3,128 KiB (+8.31%). These process-level
flags remain visible alongside the allocation reductions; they are not hidden
by the scalar timing gains. The specialized source passes
[all five ODS gates](../gates/scalar-cell-candidate-03.json), including the 66
array/reference integration tests.
