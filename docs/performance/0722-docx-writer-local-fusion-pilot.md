# 0722 — writer-local DOCX structural scan fusion

The writer-local candidate is retained. All 96 primary and 16 read-control
hard gates pass. The candidate improves edit p50 by 16.16–19.32% on
generated-medium and 6.23–6.80% on NumberedList across two ABBA cycles.
Allocation-region peak and net live bytes are unchanged in every paired
corpus/phase. The performance program remains active.

| Scenario | Paired p50 delta range | Paired mean delta range |
| --- | ---: | ---: |
| Generated edit | −19.32% to −16.16% | −18.56% to −16.13% |
| NumberedList edit | −6.80% to −6.23% | −7.16% to −6.30% |
| Generated lifecycle | −3.01% to −1.16% | −3.04% to −1.18% |
| NumberedList lifecycle | −1.01% to +0.20% | −4.25% to −0.002% |
| Generated paragraph listing | −0.67% to −0.06% | −0.51% to −0.18% |
| Pinned-media eager paragraph count | −1.72% to −0.69% | −0.58% to +0.16% |

Deltas are candidate/baseline minus one; negative latency deltas are faster.
[Every pair](results/change-0722/measurements.md) and every raw sample remain
available. No primary native tail or central repeat-drift flag exceeds 5%.
Read controls have three tail/max regression flags and ten repeat flags, all
of the latter confined to p99/max. They have no p50/mean repeat flag over 5%.
These repetitions describe observed variation, not a cross-machine guarantee
or a formal confidence interval.

The read-tail exceptions remain material limitations. Pinned-media pair 4
p99 rises from 115,751 to 137,261 ns (+18.58%); its maximum rises from 117,731
to 143,310 ns (+21.73%). Pair 3 maximum rises from 124,530 to 170,030 ns
(+36.54%). In pair 4, p50 is 101,240 → 100,490 ns, mean is approximately
102,065 → 102,224 ns, and p95 is 111,590 → 113,321 ns. The three largest
candidate observations are 137,261, 137,311 and 143,310 ns. They are retained
without exclusions. No read-tail improvement or universal read non-regression
is claimed, and the data do not identify a hardware or scheduling cause.

Retention follows the frozen central-timing and memory gates, with these tail
flags reviewed separately. Independent review recomputed all 16 read-control
sample vectors and confirmed the statistics, parity and unchanged read-path
custody. The p99 regression is not uniform across the four pairs; it remains
a recorded limitation rather than a discarded observation. Future work for
p99-sensitive callers needs a separately planned larger process cohort, not
resampling this packet to change its result.

The implementation replaces the mutable writer's adjacent alternative-anchor
scan and active-block scan with one private dual-state reader in
`alt/codec/document_scan.rs`. The original public codec stays at its original
module path, with only child-module declarations added. The complete shared
`namespace.rs` is byte-identical. `document_part.rs` only loses the unused
writer helper and its imports; existing read consumers remain unchanged.
[Source guard](results/change-0722/source-guard.json) verifies these exact
transformations against the archived baseline.

The private helper preserves independent namespace predicates, capture
suppression, depth/node/anchor caps, checked offsets, fallible reservation
labels, and alt-first refusal precedence. A range failure is deferred until
the alt parser and its first MCE decision succeed. Both ordered MCE calls and
the final body-preservation reader remain. Reader and range-parser storage are
dropped before MCE selection. The retained range-state copy creates a
maintenance obligation; the independent frozen differential scanner tests
must remain when either grammar changes. Identical host allocator-exhaustion
scheduling is not established by these explicit resource-policy checks.

| Corpus / phase | Allocation requests | Requested bytes | Peak above region start | Net live bytes |
| --- | ---: | ---: | ---: | ---: |
| Generated edit | 8,277 → 8,264 | 1,768,329 → 1,766,562 | 390,629 → 390,629 | 388,061 → 388,061 |
| Generated lifecycle | 15,616 → 15,603 | 3,199,408 → 3,197,641 | 1,049,866 → 1,049,866 | 461,654 → 461,654 |
| NumberedList edit | 2,556 → 2,538 | 436,173 → 431,414 | 21,310 → 21,310 | 18,976 → 18,976 |
| NumberedList lifecycle | 4,674 → 4,656 | 2,661,485 → 2,656,726 | 710,663 → 710,663 | 122,437 → 122,437 |

Both allocator pairs have these medians. [Aligned memory vectors and formulas](results/change-0722/memory-diagnostics.json)
retain all three samples per invocation. These counters cover individual
operation regions; owner destruction is outside those regions. They are not
RSS or leak measurements and are never summed across phases. The prior 0721
NumberedList edit peak increase is absent in this candidate's matched results;
this combined experiment does not isolate the cost of each source change.

Temporary debug tracing preserves exact public output and repeated stderr.
All 21 boundaries preserve ordered MCE inputs/outputs and structured raw and
selected scanner metadata. The public oracle covers 19 XML cases and two
complete edited-package publications, with byte-identical baseline/candidate
reports. At the representative body boundaries:

| Corpus | Main XML bytes | Baseline structural reads | Candidate structural reads | Ordered MCE input counts |
| --- | ---: | ---: | ---: | --- |
| Generated-medium | 21,517 | 2,820 | 1,410 | 0, 200 |
| NumberedList | 4,563 | 356 | 178 | 0, 5 |

These counts include successful EOF reads and exclude the separate final body
capture. Baseline range reads are counted before admission, including a read
that later fails a depth/node check. Their borrowed-event end positions are
unknown and remain null. Observer callbacks have distinct refusal boundaries
and are never counted as reader calls. [The complete trace matrix](results/change-0722/trace-details.md)
and [independent review](results/change-0722/design-review.md) retain the
refusal cases. No physical-I/O or instruction-count claim follows from these
parser event counts.

The baseline revision is `a1b259201d57b5a98c6ac053ac5fe65c45a3fa7e`.
Builds use Rust 1.95.0 with locked dependencies and release optimization;
measurements use CPU 12 on AMD EPYC 9R45, Linux 7.0.0-1012-aws, glibc 2.43,
and the observed ext4 filesystem. Physical storage topology is not attributed.
[Host metadata](results/change-0722/host.json), source censuses, binary hashes,
commands, plans and retained reports bind the evidence.

The frozen plan uses two ABBA cycles, one fresh outer invocation per
corpus/phase/stage, 100 warmups and 200 native samples. The first four stages
are repeated in the separate allocator lane with no warmups and three samples.
There are 64 outer invocations: 32 primary native, 16 native read controls,
and 16 allocator. They retain 9,600 native samples and 48 allocator samples.
Each of the eight pinned-media controls uses 100 warmup, 200 priming and 200
measured inner processes, totaling 4,000 isolated filesystem children. No
resampling or removal of observations occurred.

Edit timing excludes package open; lifecycle includes open, edit and atomic
save. Output validation, owner drop and cleanup remain outside both clocks.
Native binaries were frozen before debug instrumentation and were never
rebuilt during capture. The native and allocator measurements are separate;
there is no throughput, cold-cache, RSS, hardware-counter, scaling or general
cross-format claim.

Candidate qualification passed formatting, all-feature/all-target DOCX tests
(1,471 unit/integration tests), warning-denied Clippy, doctests (75 passed,
31 existing ignored), and warning-denied rustdoc. The performance harness
passed 535 tests with one existing opt-in test ignored. All six repository
evidence gates and 11 corruption checks pass. Terminal audit and exact replay
pass after removing the four owned scratch/build roots; eight binary identity
witnesses remain. The artifact manifest seals the complete packet.

The first terminal audit exposed an assertion expecting dictionary metadata
where the frozen analyzer emits allocation formula strings. Only the audit
assertion was corrected. [The original validator, check result and exact diff](results/change-0722/terminal-audit-correction/correction.json)
are retained; the 11 corruption checks and post-cleanup terminal audit then
passed. Native captures, source states, frozen analyzers and analysis outputs
were unchanged.

The accepted constraints are unchanged. The change stays within the DOCX
owner, adds no public API or runtime dependency, and preserves publication and
source-byte contracts. It adds no unsafe code, cache, parallel execution or
ambient library I/O. No scenario-coverage promotion is claimed.

[Evidence and replay index](results/change-0722/README.md).
