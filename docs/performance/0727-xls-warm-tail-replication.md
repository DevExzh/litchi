# 0727 — XLS warm-tail replication and queue correction

The diagnostic does not clear the rejected 0726 candidate. Across three fixed
cycles, its original four failure metrics produce 14 failing central checks:
ten means and four medians. The generated stored-query case has positive
median deltas in every paired process comparison, so the earlier rejection
cannot be dismissed as only a few isolated mean spikes. Production remains
unchanged at baseline `959daa11e5`; `performance_claim: none`.

| Focus cell / metric | Paired processes | p50 failures | Mean failures | p50 delta range | Mean delta range |
| --- | ---: | ---: | ---: | ---: | ---: |
| 54016 stored, owned q3 | 18 | 0 | 3 | +1.49% to +4.00% | −6.54% to +8.47% |
| 54016 stored, file q3 | 18 | 0 | 2 | 0.00% to +2.78% | −6.71% to +7.84% |
| Generated, owned q3-to-q8 mean | 18 | 4 | 5 | +2.13% to +6.31% | +0.28% to +6.35% |
| 45365 late, file q8 | 18 | 0 | 0 | −2.34% to +3.57% | −4.34% to +4.74% |
| 45365 first, file q8 control | 18 | 0 | 0 | −1.21% to +3.51% | −4.28% to +4.67% |
| 54016 missing, owned q8 control | 18 | 0 | 0 | −30.77% to −16.67% | −40.29% to −20.78% |

Negative deltas are faster. These are ranges across independent process
comparisons, not pooled-owner estimates or confidence intervals. Each process
retains all 100 measured owners. The generated focus mean fails in each of the
three cycles; its median fails in cycles zero and two. The owned 54016 q3
means fail only in cycle two, and file q3 fails once each in cycles one and two.
The earlier 45365-late q8 failure does not recur in these 18 comparisons. None
of those facts changes the original frozen retention result.

Across all seven metrics and six cells, there are 60 failed paired central
checks and 313 p95/p99/maximum regression flags above 5%. All comparisons,
A/A controls, within-phase drift, per-process distributions and raw observations
remain in the [complete analysis](results/change-0727/analysis.md) and JSON.
Both p50 and mean use the original 5% threshold; only q3, q8 and their warm mean
have the original 10 ns exception. Here these are diagnostic flags, not an
optimization admission gate. No samples were removed, replaced, trimmed or
pooled to change their interpretation.

The candidate is exactly the one-function archived 0726 change, which skips
unused path/resolver/hint setup for empty indexed slot ranges while retaining
worksheet lookup and final execution/source checks. There is no new code
candidate. It still removes one 16-byte allocation/deallocation on warm missing
queries under 0726's exact allocator evidence; this batch does not recapture or
expand that allocation claim. The positive missing-control response recurs, but
stored-query cost also changes. The experiment does not isolate instruction,
code-layout, cache, scheduler or hardware causes. Do not treat the mean failures
as noise or the missing benefit as proof of whole-matrix non-regression.

The prospective plan contains three aa1 → aa2 → a1 → b1 → b2 → a2 cycles.
Each of six fixed case/mode cells has three fresh processes per leg, with three
warmup owners and 100 measured owners taking eight queries. This yields 324
processes, 32,400 measured owners, 259,200 queries and 972 warmup owners. The
four previously failing cells are accompanied by a same-fixture-family file
stored-query control and an owned missing-result benefit control. Native runs
use the unchanged 0686 probe, Rust 1.95.0, CPU 12 on AMD EPYC 9R45, and warm
OS caches. Each case retains its original index budget and coordinate.

Both source-bound release binaries are freshly built serially. Baseline source
is restored before capture; the candidate archive and 494-file source census
prove the sole source difference and exact 0726 ancestry. Root preflight ran
before freeze; independent audit implementation continued during collection and
validated the completed packet afterward. No code or frozen analysis rule was
changed after freeze. Exact probe commands, source/fixture/tool hashes, outcomes,
headers and raw stdout/stderr are verified independently. Ten actual-analyzer
controls exercise complete evidence, corrupted hashes/outcomes/process identity,
commands, phases, end bindings and timing boundaries. Post-cleanup analysis and
independent statistics replay exactly using the two executable identity witnesses.

No production source change is retained, so no new Cargo correctness claim is
made. The candidate's exact prior correctness qualification remains in sealed
0726; this batch adds diagnostic replication and source-custody checks. Structural
claim classification and the non-iWork gate pass at terminal replay. No cold
storage, remote source, concurrent execution, RSS, Office or cross-platform
claim follows. [Evidence and replay index](results/change-0727/README.md).

The immediate decision is to stop repeating the unchanged empty-slot candidate
as a retention attempt. A future version needs a reasoned explanation or remedy
for the stored-query cost and a fresh full-matrix qualification. The corrected
[priority review](results/change-0727/priority-review.md) selects fresh DOC/PPT
length-changing-save attribution under the implemented 0663 Reuse policy as the
next qualification. Older 0617 copying fractions must be remeasured; 0663 changed
layout behavior and recorded a roughly 40% Reuse regression on picture.doc.
The existing 0652 placement decision remains authoritative. Its initial
shared-string-admission proposal was rejected during review because that work
already landed in 0667; the initial assessment is preserved separately rather
than allowed to become a stale queue instruction. The non-iWork goal remains
active and no registered claim or CRUD coverage is promoted.
