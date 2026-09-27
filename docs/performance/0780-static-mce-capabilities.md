# 0780 — static baseline MCE capabilities

Status: retained as a narrow staged-commit and allocation improvement, with an explicit lifecycle tradeoff. Public PPTX staged commit improves 5.012–9.167% across three corpora, while the large full lifecycle regresses 2.398%. The equally weighted lifecycle summary is effectively flat. Adoption rests on the useful public staged operation and reduced allocation traffic without added live memory; it does not establish an end-to-end speedup.

The OOXML markup-compatibility baseline recognizes seventeen fixed namespace URIs. Its constructor previously allocated seventeen strings and grew a hash set on every call. The current constructor diagnostic confirms 21 allocations and 2,398 cumulative requested bytes per profile. These counters include allocation requests, not physical copies or resident memory. The historical 0760 PPTX commit profile motivated this investigation; its percentage is not a current 0780 attribution.

The candidate represents the fixed profile with a private Baseline variant. Exact membership scans seventeen static references. The first custom registration materializes the existing baseline set plus the caller namespace; custom profiles continue to use the owned HashSet. Extensions remain independently owned. No public method signature, global cache, dependency or unsafe code is introduced. Derived Debug displays the new private representation.

A target/compiler layout test checks that the enum and capability owner occupy the same sizes as the previous HashSet owners. The DOCX settings envelope still covers its nineteen materialized namespace strings and existing owner sizes. This is a checked current-target property, not a portable promise about Rust enum layout.

The frozen plan measures public PPTX capture, staged one-edit commit and capture-through-publication across 3×4, 12×8 and 100×100 slide/shape corpora. Package ingress and corpus construction are outside the clock. The lifecycle includes capture, edit, commit, apply and serialization. Native timing uses six alternating before/after blocks, thirty samples after three warmups; allocation uses two separate blocks with three samples and no warmups. The constructor mode is a separate 1,000-iteration diagnostic and includes profile destruction inside its loop.

All publications are reopened after clocks and allocation regions, compared against exact generated text, and bound by source/output SHA-256. Whole-process RSS and perf stat include setup and verification, and must not be treated as operation-local metrics. Six process blocks provide descriptive uncertainty; no physical cold-cache, remote-source, concurrency, scaling, native Office or complete CRUD coverage claim follows.


## Paired public results

Times are the median of six process p50 values, in milliseconds; each process quantile uses nearest rank. Changes and 95% bootstrap intervals use the six matched after/before ratios (10,000 resamples, seed 780078), so they need not equal the ratio of the displayed medians. RSS is the median of six whole-process maxima in KiB. Negative changes are improvements. These intervals describe this capture, not population guarantees.

| Corpus and phase | Before → after p50 (ms) | Paired change | 95% interval | RSS before → after (KiB) |
|---|---:|---:|---:|---:|
| tiny capture | 0.257717 → 0.246676 | -4.262% | -4.821…-3.879% | 5,058 → 4,850 |
| tiny commit | 0.231371 → 0.219921 | -5.012% | -5.312…-4.555% | 4,892 → 4,846 |
| tiny lifecycle | 1.454846 → 1.437817 | -1.188% | -1.505…-1.064% | 4,966 → 4,814 |
| medium capture | 0.515807 → 0.496383 | -3.855% | -5.436…-2.667% | 5,148 → 5,178 |
| medium commit | 0.328112 → 0.306941 | -6.592% | -7.327…-5.873% | 5,370 → 5,236 |
| medium lifecycle | 2.095544 → 2.071189 | -1.335% | -1.462…-1.071% | 5,086 → 5,238 |
| large capture | 21.920530 → 21.886357 | -0.225% | -0.636…+1.722% | 16,540 → 16,472 |
| large commit | 1.464937 → 1.331586 | -9.167% | -9.380…-8.940% | 16,534 → 16,418 |
| large lifecycle | 31.993240 → 32.663036 | +2.398% | +1.677…+3.106% | 16,512 → 16,576 |

The equally weighted geometric mean of the three paired p50 ratios is −2.798% for capture, −6.939% for staged commit and −0.056% for lifecycle. This is a descriptive summary of these three corpora, not a workload mix. The large lifecycle regresses in all six blocks (+1.650% to +3.476%), despite the faster staged commit. Its +2.398% median regression must not be hidden behind the 5% flag threshold or the constructor diagnostic. No whole-lifecycle speedup is established.

All 20 spread flags and four paired metric flags are retained in [all-flags.csv](results/change-0780/all-flags.csv). The four metrics contain five above-threshold blocks: medium lifecycle RSS +6.571%/+7.033%, tiny commit RSS +5.276%, tiny lifecycle p99 +21.199%, and tiny lifecycle RSS +6.529%. Their median paired ratios remain below +5%. The spread flags cover tails and RSS; none covers public-operation p50. See [all process values](results/change-0780/native-processes.csv), [paired values](results/change-0780/native-pairs.csv) and [distribution summary](results/change-0780/native-summary.csv). No flagged process is discarded or resampled.

## Allocation and independent diagnostics

Separate allocator builds retain two process blocks with three samples each. The following operation-region counters agree across both blocks. All thirteen replayed metrics are available in [allocation-summary.csv](results/change-0780/allocation-summary.csv); no allocation spread or regression flag exceeds 5%.

| Corpus and phase | Allocation calls before → after | Requested bytes before → after |
|---|---:|---:|
| tiny capture | 1,548 → 1,338 | 141,615 → 117,635 |
| tiny commit | 2,073 → 1,863 | 172,144 → 148,164 |
| tiny lifecycle | 6,117 → 5,670 | 1,075,792 → 1,023,938 |
| medium capture | 3,268 → 2,869 | 274,193 → 228,631 |
| medium commit | 3,460 → 3,061 | 278,724 → 233,162 |
| medium lifecycle | 10,393 → 9,568 | 1,426,880 → 1,331,862 |
| large capture | 74,353 → 72,106 | 5,024,525 → 4,767,939 |
| large commit | 17,745 → 15,498 | 1,368,582 → 1,111,996 |
| large lifecycle | 108,852 → 104,331 | 9,060,009 → 8,542,943 |

Net live bytes retained at region exit and peak live bytes above region entry are unchanged in all nine public cases. Absolute live gauges differ by one byte because the retained probe binary name differs; that is not a production memory saving. There is no retained-cache growth or operation-peak reduction claim.

The 1,000-profile constructor diagnostic measures 1.361931 → 0.055670 ms median process p50 (paired −95.926%), and 21,000 → 0 allocation calls / 2,398,000 → 0 requested bytes. Heaptrack independently resolves complete allocation ancestry: direct constructor events fall from 21,000 / 2,398,000 bytes to zero. The enclosing diagnostic function retains three post-clock reporting allocations totaling 413 bytes. Exact trace event counts and byte sums match the histogram; event counts also match the print summary. Attribution comes from the complete trace, including inline ancestry, not a sum of truncated printed stacks.

Twelve separate perf-stat processes measure whole-process user counters, including corpus setup and readback. Median paired instructions/cycles change −92.604%/−92.807% for the constructor diagnostic, and −0.611%/+1.064% for the large staged-commit process. These are not operation-local instruction counts or a causal phase fraction. They are not pooled with native timing. [Observer analysis](results/change-0780/observer-analysis.json) binds all counters, trace/decode receipts and exact commands.

## Correctness, architecture and replay

Six final gates pass: formatting, all-feature/all-target checks, all-feature tests, warning-denied library Clippy, warning-denied rustdoc and crate boundaries. The test log independently totals 5,683 passed, zero failed, 35 ignored across 237 suite summaries. Five new tests cover the independent 17-URI oracle, owner layout, legacy/stream equivalence, custom registration/default/clone/extensions/MustUnderstand, and malformed/unbound-prefix/input-limit parity. Custom-registration performance was not measured; its first nonbaseline registration materializes the owned set.

The initial quality attempt failed on a test-only array-iterator pattern. The failed draft and log remain archived; [applied-model.patch](results/change-0780/candidate/applied-model.patch) is the corrected, formatted source used by all successful gates and candidate builds. No production logic changed in that test correction. Probe build warnings concern unused diagnostic helpers; production warning-denied gates pass.

Shared MCE ownership, immutable snapshot behavior and preservation contracts remain under ADRs 0001–0006, 0008, 0010/0011 and 0024; no payload memo under 0032 is introduced. See [design](results/change-0780/design.md), [source review](results/change-0780/source-review.md), [candidate review](results/change-0780/candidate-review.md) and the 35 unchanged [architecture input hashes](results/change-0780/architecture-inputs.json). No CRUD registry row is promoted. ODF remains deferred by the prior owner decision, and iWork is excluded.

The packet retains 120 native, 40 allocator, ten baseline qualification and fourteen profiler processes. Exact build/source/probe/lock identities, generated fixture/output digests, commands, environment, raw process samples, all flags and failed quality history remain available. Offline replay uses `python3 -B docs/performance/results/change-0780/validate.py`, followed by `python3 -B docs/performance/results/change-0780/tables.py --check`. The final seal covers every packet file except itself. Observer replay checks retained binaries or exact cleanup witnesses and works after path relocation.

## Integration and cleanup

The production change is retained under the explicit [disposition](results/change-0780/disposition.json). All four executable identities were verified before the owned target was removed (6,287,734,597 file bytes). Offline replay passes with cleanup witnesses after removal. Commit `59bdb64f16` was fast-forwarded into `feat/office-format-completeness`. All 648 sealed payload files plus the seal were checked against staged Git blobs before commit. The owned worktree `/home/zhuhe/code/litchi-worktrees/0780-static-mce-capabilities` and branch `perf/0780-static-mce-capabilities` were removed after verifying and removing their copied lockfile and three exact reference symlinks. All pre-existing worktree records and the three unrelated main-worktree file hashes remain unchanged. Main replay and all five generated-table checks pass both before and after removal of the original worktree and executables. The non-iWork goal remains active; this batch is a narrow measured improvement with an accepted lifecycle tradeoff.
