# 0724 — XLS checkpoint cost attribution

The diagnostic contrasts reproduce the missing-query and small-file publication
costs of the rejected 0723 candidate, but do not identify a single machine-level
cause. Production remains unchanged at `0d943df447`. The prior rejection stands;
this packet has no retention gate or support promotion. The non-iWork
performance goal remains active. `performance_claim: none`.

Four variants were built with the same unchanged probes: baseline; the larger
index layout and fixed charge only; that layout plus replay-selection code with
no checkpoint constructor; and the complete archived 0723 production candidate.
The sequence is baseline → layout → selection → full, then its reverse. Six
owned-source cases cover 54016 first/late/missing, Plan1 first, Simple first and
45365 first. Each stage/case has 100 fresh measured owners (eight queries each)
and nine separate 50,000-query processes after two preparatory queries.

| Full versus baseline | Round 1 p50 delta | Round 2 p50 delta |
| --- | ---: | ---: |
| 54016 late repeated loop | −81.93% | −81.93% |
| 54016 missing repeated loop | +6.90% | +7.22% |
| Plan1 first repeated loop | +2.61% | +3.03% |
| Simple first repeated loop | +1.78% | +1.92% |
| 45365 first repeated loop | +2.24% | +5.49% |
| Simple second query / publication | +12.68% | +9.59% |

Negative latency deltas are faster. Repeated-loop percentiles describe nine
process averages, not individual-query tails. All p50/mean/p95/p99/min/max
statistics, component contrasts and mirrored drift remain in the
[analysis](results/change-0724/analysis.md) and its complete JSON. This smaller,
interleaved owned-source matrix does not replace 0723's qualification matrix.
Plan1/Simple early-loop regressions are smaller here; no observations have been
pooled or substituted to overturn the previous decision.

For missing 54016 queries, repeated-loop p50 moves from 101.815 to 102.619 to
105.813 to 108.836 ns across baseline/layout/selection/full in round one. Round
two gives 101.516, 102.426, 105.950 and 108.848 ns. The layout contrast adds
0.79–0.90%; the unpopulated selection contrast adds 3.11–3.44%; the full contrast
adds a further 2.74–2.86%. Missing queries never execute the target-checkpoint probe
and have no replay cursor walk. Therefore that final contrast cannot be called
the time spent constructing a checkpoint. It includes nonlocal compilation,
layout and runtime effects not separated by these observations.

The same limitation is pronounced in publication timing: the layout → selection contrast changes
54016 missing q2 p50 from about 612 to 518 microseconds in round one and 610 to
532 microseconds in round two, despite retaining the baseline scan path. The
full-minus-selection q2 contrast combines more than the added constructor work.
This experiment cannot apportion those changes to a specific instruction, cache
line, compiler decision or scheduler event. No hardware counters were captured.

Simple's full q2 p50 is 800 ns in both rounds, versus baseline 710 and 730 ns.
Its layout and selection variants stay at 720–730 ns, making removal of extra
publication work a concrete next hypothesis. The full late-target repeated-loop
p50 stays near 258 ns versus baseline 1,429 ns. The mechanism remains useful,
but the missing-query and construction costs still require redesign.

All mirrored repeated-loop p50 changes are within 3.05% in this capture. Native
timing has larger variation: Simple layout q2 mean falls 17.23% between rounds,
while its p50 stays 720 ns. The complete tails are retained and no outlier is
removed. Two mirrored stages are limited replication, not a confidence interval
or cross-machine guarantee.

The next bounded source seams are to avoid unused setup for an empty indexed
slot range while preserving final execution/source checks, reuse the scan's
existing borrowed path vector during publication, and record the first selected
frame offset during scanning instead of searching the retained slot vector.
The independent [source review](results/change-0724/source-review.md) documents
these proposals and the earlier-target fallback requirement. An early-target
threshold alone would leave the larger fixed layout in place; its charge cannot
be removed merely because a checkpoint is absent. No proposed seam is installed
in production by this packet.

Eight release binaries use Rust 1.95.0 on CPU 12, AMD EPYC 9R45, with warm OS
caches. The 480 processes comprise 48 native invocations with 4,800 measured
owners / 38,400 queries plus 144 warmup owners, and 432 repeated-query processes
with 21.6 million measured queries. Native missing-query budget is 1 MiB; all
other native budgets and all repeat budgets are 2 MiB. Source, fixture, plan,
primary analyzer, driver and binary identities were frozen before acquisition.
An independent audit checks exact argv, ordered run identities, filenames,
raw hashes, source contrasts, source restoration and all semantic outcomes.
Its separately computed statistics match every primary stage statistic exactly.
Six verifier controls include altered raw hashes, outcomes, repeat counts,
duplicate runs and end bindings. No allocation, RSS, physical-I/O, file-backed,
concurrent, cold-device or production-wide speedup is claimed in this packet.

The first layout build failed the existing unused-code lint because diagnostic
fields/helpers were intentionally unused. The attempt is archived. Scoped
allowances apply only to those archived layout/selection items; baseline/full
source and production warning policy remain exact. A later audit metadata check
was corrected to read `all_queries_agree` per owner instead of per report; its
initial failure is also archived. Frozen measurement tools and samples did not
change. Owned build/binary roots were removed after capture and exact identity
witnesses retained for offline replay. [Evidence index](results/change-0724/README.md).
