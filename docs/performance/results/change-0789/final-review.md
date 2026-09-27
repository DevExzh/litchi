# 0789 final evidence review

This is an independent read-only review of the finalized 0789 accounting
packet. It does not modify the probe, analyzer, capture drivers, production
source, or native data. The evidence supports the diagnostic disposition in
[`0789-rss-accounting-calibration.md`](../../0789-rss-accounting-calibration.md):
the historical 0787 RSS rejection remains unchanged and no production
candidate is adopted.

## Receipt and replay scope

The matrix contains 272 children: 17 cases, two affinities, four
launcher/observer combinations, and two repeats. It has 816 full and 816 identity
checkpoint sets as specified by the plan. The original four syscall traces and
the separately labeled four `-v` trace supplements add 48 full checkpoint
observations, for 280 children and 864 full snapshots overall; the matrix's
136 identity children contribute 816 identity snapshots. The supplements are
retained separately from the primary analyzer's 276-child scope.

The following offline checks were replayed while auditing the packet:

* `accounting_summary.py --check` passed. It independently checked all 280
  receipt records, the full `smaps`/`smaps_rollup`/status RSS sums, known-page
  anonymous growth, and the four detailed traces.
* `analyze.py --check` passed after the packet finished its concurrent output
  update. Its primary scope remains 272 matrix children plus four original
  traces; it does not silently absorb the four supplements.
* The final-seal validator was not yet runnable at this review point because
  root had not written `cleanup.json` or `seal.json`. After this review is
  included, root must regenerate those custody artifacts and rerun
  `validate.py --require-final-seal`.

## Direct `wait4` and self high water

The raw matrix has 136 direct rows, and all 136 satisfy:

```text
direct wait4.ru_maxrss == maximum RSS0789 self sample
```

That equality is a scope check, not proof that the six samples captured the
process peak. In 112 of the 136 rows, the maximum is already the `startup`
self value. Those rows cover the 0, 1, 4, and 16 MiB direct cases, including
the worker variants. The direct route is launched through `taskset` and then
execs the probe; its resource high-water history survives that exec. Direct
startup self maxima are approximately 17,708–19,372 KiB, while the fresh
GNU-time probe children start around 1,364–1,480 KiB. The remaining 24 direct
rows are the 64 MiB cases, where the maximum occurs at `touched` or later.

Therefore the equality is consistent with the direct launcher retaining a
pre-exec high-water value and cannot establish that a transient was absent or
that a small mapping's physical peak was observed. The report correctly keeps
the direct parent `wait4`, probe self samples, and GNU-time child result as
different observations.

## 4 MiB, 32-worker affinity result

The four full GNU-time rows for the 4 MiB, 32-payload-worker case are the
strongest affinity contrast in the packet:

| Affinity | Repeat | Maximum sampled `smaps_rollup` RSS | GNU-time `%M` | Difference |
| --- | ---: | ---: | ---: | ---: |
| all CPUs | 0 | 6,080 KiB | 4,296 KiB | +1,784 KiB |
| all CPUs | 1 | 6,080 KiB | 4,040 KiB | +2,040 KiB |
| CPU 0 | 0 | 6,084 KiB | 6,132 KiB | −48 KiB |
| CPU 0 | 1 | 6,080 KiB | 6,132 KiB | −52 KiB |

The all-CPU versus CPU-0 change is compatible with the source-backed
per-CPU-counter hypothesis. It is not a proof of that mechanism: the observed
RSS column is the maximum of six acknowledged snapshots, not a continuous
peak, and the two repeats are descriptive observations rather than a
confidence interval. The comparison uses GNU-time `%M`; separate detailed
`-v` traces for the zero/64 MiB controls match that metric to the probe child’s
`wait4` result. The parent `wait4` on
the GNU-time wrapper is a different, pre-exec-contaminated scope and must not
be substituted into this table.

## Known 64 MiB anonymous growth

The eight full matrix rows for 64 MiB, main-thread touch, and zero workers
show an exact mapped-to-touched `smaps_rollup` `Anonymous` increase of
65,536 KiB (67,108,864 bytes), with zero KiB beyond the requested payload in
every row. The independent summary also checks the summed `smaps` RSS against
rollup and status `VmRSS` for all 864 full snapshots. This validates the
standalone probe's known-page witness and the parser's accounting for this
control. It does not identify the source of the 0787 candidate's 222 KiB
increase; the office workload has a different mapping and lifecycle.

The 36 KiB four-worker and 44 KiB 32-worker anonymous deltas are retained as
process-wide observations. They include thread and runtime mappings and should
not be described as payload-only allocator cost.

## Syscall and protocol checks

The four detailed traces match all 48 self `getrusage` calls to the probe
transcript and match each GNU-time `%M/%R/%F` triple to the child `wait4`
return. The retained triples are:

| Case | Affinity | RSS KiB / minor faults / major faults |
| --- | --- | ---: |
| 0 MiB | CPU 0 | 1,780 / 79 / 0 |
| 0 MiB | all CPUs | 1,780 / 78 / 0 |
| 64 MiB | CPU 0 | 67,016 / 16,463 / 0 |
| 64 MiB | all CPUs | 67,016 / 16,464 / 0 |

The probe checksum and mapped-byte witnesses are checked by the primary
replay. All worker task IDs are absent by the post-touch checkpoints, so the
full snapshots do not observe live worker stacks. A nonzero ACK delta remains
possible: the independent summary records 288 KiB on three full checkpoint
intervals and two identity intervals. The parent should retain that as a
sampling/handshake observation rather than erase it.

The source review's per-CPU explanation remains a testable hypothesis. Ubuntu
vendor kernel patches were not independently diffed, and the experiment does
not instrument individual per-CPU residues. No constant correction or relaxed
0787 threshold follows from this packet.

## Disposition

There is no measurement blocker in the captured evidence. The one remaining
release-gate action is custody: include this review in the final packet seal,
remove only the owned temporary executables and target, and rerun the sealed
validator. Until that succeeds, the packet is scientifically reviewable but
not yet closed as a committed artifact.
