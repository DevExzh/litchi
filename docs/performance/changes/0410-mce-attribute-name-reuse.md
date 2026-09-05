# Change 0410: reuse expanded attribute names in the MCE stream

Date: 2026-09-04

`performance_claim: scoped`

`claim_authorized: true`

The retained optimization reduces warm selected-cell query p50 by 3.94% and
4.14% in the two matched AB pairs. Mean reductions are 4.03% and 4.19%; p95
reductions are 4.21% and 4.71%; p99 reductions are 4.43% and 3.64%. All four
statistics pass the scoped ABBA direction/drift policy. Operation-local
allocation calls fall 5.745% and allocated bytes fall 3.567%, with unchanged
live bytes after the operation. The strict registry entry is
`claim-0410-mce-attribute-names` and covers only this selector/corpus/build.
The generic comparator continues to exclude the filesystem selector from its
ordinary latency gate; this claim uses the explicit clean 500-sample ABBA
package and per-change review.

## Implementation and hypothesis

The 0409 selected-cell profile attributed 16.22% of selected-ancestor leaf
weight to `clone_bounded_name_part`; 75.50% of that helper's direct-caller
weight came through expanded-name cloning. This motivated a narrow ownership
change in the shared MCE stream.

Commit `e4d477466718a8fad38cd55b9babe0b826e7f3a7` checks duplicate attributes
using a pre-reserved `HashSet<&Name>` and moves admitted attribute names into
semantic events after their existing checks. Namespace expansion still creates
the initial owned name. Raw observer names and element/frame names retain the
copies required by their public owned-value contract. Lexical fields keep their
existing borrowed representation.

Duplicate equality/hash behavior, input-order error detection, expanded-name
limits, namespace/MCE filtering, raw/semantic callback order, and end-event
ownership remain in place. Added tests retain event names across callback/input
buffer reuse and exercise the exact and one-under expanded-name boundary. No
public signature, runtime, dependency, parallelism or unsafe code is added.

The change follows ADR 0005's bounded allocation and measurement contract,
ADR 0006 preservation/validation rules and ADR 0008 verification gates. It stays
within the shared grammar owner under ADRs 0001/0002/0010/0011/0024 and does not
change snapshot, edit or patch behavior under ADR 0003.

## Measurement scope

Control is `972dc25be0dbd6690c74429839a48288d637e2d5`; candidate is
`e4d477466718a8fad38cd55b9babe0b826e7f3a7`. Each role has a separately copied
normal and allocator executable, with SHA-256/build ID evidence. Captures use a
clean detached worktree at the corresponding revision. Later OPC fixes and a
test-module lint rename are outside these measured revisions.

Rust 1.98.1 release builds use debug level 1, forced frame pointers/unwind
tables, four build jobs and no incremental compilation. The separate harness
workspace does not inherit root release LTO settings. Every measured command
is pinned to CPU 2 on AMD EPYC 9R45, Linux 7.0.0-1011-aws, perf 7.0.12, KVM,
32 affinity CPUs and about 128 GiB RAM. No task compilation, tests or profiling
overlapped the ABBA measurements. The shared guest was not an exclusive host.

Order is A1/B1/B2/A2. Each leg runs a normal selected query (500 samples,
20 warmups), source/eager one-cell edit/save guardrails (500/20 each), then
operation-local selected-query allocations (30/3). Allocation-instrumented
elapsed values are excluded from normal latency comparisons.

The medium deterministic corpus has four 48-by-48 sheets (9,216 cells),
4 MiB media and 17 ZIP members. Its archive SHA-256 is
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.
The selected query opens/prepares outside the timer, then selects `bEnCh01`
as `Bench01` at position 1 and reads M29, exact numeric lexical value `1028012`.
Every sample runs in a fresh child with warm filesystem cache; full semantic
verification and deterministic oracle construction are outside the timer.
External perf/time totals include surrounding work and inherited children.

The edit/save guardrails time open, selector planning, commit and sequential
publication. They change Sheet1!A1, then reopen and verify outside the timer.
They are in-process selectors, despite the shared filesystem configuration
flags. Outputs retain the same 4,226,480 bytes, 201 writes, 32,768-byte maximum
write and SHA-256
`9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1`.
Source-backed samples retain 257 reads/4,233,005 bytes and unchanged-member
digest `7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d`.
Workbook/selected/unselected compressed-range overlap is 226/6,816/20,330
bytes. These counters include raw publication copies and do not imply parsing.

## Results and regression review

| Normal query statistic (ns) | A1 | B1 | B2 | A2 |
| --- | ---: | ---: | ---: | ---: |
| p50 | 3,537,662 | 3,398,112 | 3,398,202 | 3,544,863 |
| Mean | 3,544,045.732 | 3,401,184.956 | 3,403,096.096 | 3,552,030.136 |
| p95 | 3,598,072 | 3,446,642 | 3,446,312 | 3,616,823 |
| p99 | 3,653,393 | 3,491,592 | 3,549,843 | 3,683,863 |

The retained reports contain all samples, sample-order mappings and 95%
Student-t intervals for each mean. Same-role p50 drift is 0.204% control and
0.00265% candidate; mean drift is 0.225% and 0.0562%. No aggregate across
unrelated scenarios is used.

All 30 allocator samples in both control legs have 81,918 allocation calls
(including 12 reallocations), 81,903 deallocations, 10,690,444 allocated bytes
and 10,689,061 deallocated bytes. Both candidate legs have 77,212 allocation
calls (including 12 reallocations), 77,197 deallocations, 10,309,094 allocated
bytes and 10,307,711 deallocated bytes. Failed allocations are zero throughout.
The reductions are 4,706 calls and 381,350 bytes per query. Live bytes remain
20,253 before / 21,636 after; lifetime peak-after is 8,494,671 control versus
8,494,637 candidate. Those absolute lifetime snapshots are not operation peak
measurements or a peak-memory improvement claim.

External normal-command maximum RSS is 115,992/116,472/116,236/115,988 KiB in
ABBA order, paired changes +0.414%/+0.214%. Guard-command RSS changes are
+0.189%/-0.121%; allocator-command changes are +0.266%/+0.155%. These process
high-water observations include surrounding work, rather than isolated API
allocation peaks.

The initial eager edit/save guard is adverse: p50 changes +0.533%/+5.659%,
mean +0.578%/+5.688%, and p99 +1.537%/+5.975%. Candidate p50/mean drift also
exceeds 5%. This triggered review and one additional full guard-only ABBA
block, retaining the original reports. The diagnostic repeat uses identical
500/20 counts, case order, corpus, executable pairs and CPU affinity. Its eager
p50 values are 8,439,552/8,346,226/7,892,474/7,932,535 ns; candidate reductions
are 1.106%/0.505%, while control and candidate same-role p50 drift are -6.008%
and -5.437%. Both implementations therefore exhibit substantial process-to-
process variation. Neither block authorizes an eager latency claim, and the
initial adverse result is not replaced by the favorable repeat.

Static timer/corpus review finds no execution of the changed MCE stream in
this generated eager fixture: ordinary parsing uses the legacy processor,
`dyDescent` is absent, and no shared-string member is present. The selected
query always invokes the changed streaming path. This supports retaining the
measured narrow optimization; it does not establish the cause of eager timing
variation or a general absence of regressions. Initial source-backed guard
p50 reductions are 0.069%/1.430%; repeat reductions are 0.375%/0.191%.
These small guard movements are descriptive and no edit/save claim is made.

## Residual CPU profile

A separate candidate capture retains cycles:u at 999 Hz, frame-pointer
callchains, 300 samples and 10 warmups, the raw perf data, no-inline text export
and a whole-process flamegraph. The exact selected-cell ancestor contains
2,929 stack blocks and 10,134,187,241 weighted event periods.
`clone_bounded_name_part` accounts for 10.91% of selected leaf weight;
`parse_element` is 5.89% self / 23.42% inclusive. Lexical byte cloning is
0.068% self / 0.239% inclusive. This is residual hotspot evidence, not a paired
CPU/cycle reduction: the 0409 and 0410 sample denominators differ, frequency
adaptation affects fresh-child sampling, and inclusive rows overlap.
No new cache-counter claim is inferred. The native L2/generic-L1/LLC
availability findings from 0409 remain the host inventory.

## Independent correctness repair

The broader common/XLSX tests exposed three row-visibility publication failures
that reproduce on the control. OPC topology publication compared replacement
payloads during a signed/no-op probe and read them again during changed-member
validation. For uncached large payloads this repeated source I/O. Commit `6d1588602fe4f682c57992e897583113e03efff6`
retains one verified read without retaining additional `PartData` across
planning. Signed changed replacements refuse before unrelated physical-index
planning; exact signed no-ops remain byte-exact.

Managed changed publication also needs memory for bounded topology planning in
addition to retained/in-flight payloads. The row-visibility fixture and API
documentation make that allowance explicit while retaining payload one-under
checks and a planning-budget refusal/no-output/release regression. These fixes
are correctness work, with no timing claim inferred from the MCE experiment.

## Verification and reproducibility

The final all-feature OPC/common/XLSX run passes 1,918 unit, integration and
doc tests, including all 17 row-visibility integration cases. The earlier
common/XLSX run passed 1,487 and failed three; those same failures were
reproduced against the control before the independent OPC repair. Scoped
formatting passes. A nested MCE test module was renamed to remove a preexisting
Clippy module-inception lint; that test-only rename is outside the measured
candidate revision.

Warning-denied common/OPC Clippy, all-three-crate rustdoc and the crate-boundary
check pass. Two preexisting OPC test-helper lints were corrected after the full
test run and typechecked by the final all-target Clippy check. The boundary
graph remains 64 packages, 240 internal declarations and 14 iWork debt items. The broader XLSX all-target Clippy run remains blocked
by preexisting diagnostics: eight `chunks_exact_to_as_chunks`, one useless
`as_ref`, two field-reassign/default diagnostics and 17 `result_large_err`
diagnostics (28 total in the library-test build). No green workspace-wide lint
or native Office gate is claimed. There is no dedicated MCE fuzz target in the
current tree; the existing MCE malformed/limit/recovery suites are exercised.

The [artifact bundle](../results/change-0410/) contains clean build identities,
all report/catalog pairs and command journals, the explicit regression repeat,
allocation and corpus checks, CPU data/attribution, full gate logs, and a
standalone inventory verifier. The initial allocation verifier incorrectly
required identity sample order; it now validates the sort permutation and raw
alignment, with duplicate-order and raw-elapsed mismatch rejection checks.
The first adapted profile parser mistyped the corpus hash while changing its
version label; the identity check rejected it. Both failed scripts/logs remain
available. Corrected postprocessing used the original captures throughout.

## Limits

This is one warm medium XLSX corpus and a single worker. It establishes no
physical-cold, remote-source, real-producer, other-format, scaling, copied-byte,
decoded-byte or operation-local cache-counter improvement. The larger non-iWork
CRUD matrix and the program goal remain open.
