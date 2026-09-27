# 0786: finite execution-budget read-session scaling

This batch adds a reusable standalone benchmark for the three public low-level
read sessions under finite shared-budget dimensions. It addresses an evidence
gap after the correctness-focused read-budget composition in 0676. Production
source remains byte-identical to `76dde712f53f007d51be6b29a64b575c3a4c3eb9`.

## Scope and design

The frozen matrix has 120 cases: OPC eager opening, CFB bulk stream reads, and
source-backed ordered Part reads; 32 members of 4 KiB, 256 KiB, or 31 large plus
one small member; task floors of zero and 64 KiB; and widths 1, 2, 4, 8, 32.
Fresh sessions cover all three sizes. Primed sessions cover the large corpus.
OPC ZIP members are Deflate-compressed. All CFB streams use regular FAT,
including the 4096-byte streams at the MiniFAT cutoff.

The source is immutable memory already available to the process. Fresh is a
session/package state, not physical page-cache cold. OPC times the public eager
open; CFB and Parts time payload reads after metadata setup. Primed OPC/CFB
reuse scheduling state; primed Parts retain payloads as a cache-hit control.
Absolute durations across these three scopes are not interchangeable.

Each sample owns a new finite Budget root. Worker and I/O limits equal the
requested width, CPU tasks are capped at one million, memory at 128 MiB, and
in-flight work at 32 tasks / 16 MiB. Inputs, outputs, objects, depth and work
are also finite. The aggregate parallel threshold remains 64 KiB. Production
code applies the separate per-task floor to every selected OPC/Part member, or
to every member of a CFB batch. One small member therefore serializes the mixed
request in this one-batch corpus at the 64 KiB floor. Structural OPC content
types and root relationships are not payload tasks.

Corpus generation, metadata preparation where applicable, priming, byte/order
verification and teardown are outside the wall interval. Process CPU-clock
reads bracket that interval and therefore include slightly more work. RSS is
whole-child peak memory, including corpus generation, both retained container
representations, warmups, verification and all samples. It is not an operation
allocation measurement. Source counters are enabled only in a separate
observer binary; observer timings are excluded from native results.

## Architecture and claims

| Constraint | Treatment |
| --- | --- |
| ADR 0001/0002/0024 layering | Standalone tool calls existing low-level owners; no facade dependency change. |
| ADR 0003 immutable snapshots | Reads only; no edits, publication or patch behavior changed. |
| ADR 0005 evidence and explicit execution | Finite caller-created contexts; serial child orchestration; raw timings and source diagnostics retained separately. |
| ADR 0006 preservation and safety | Every returned member is compared byte-for-byte and SHA-256/order checked outside timing; production source unchanged. |
| ADR 0010/0011 ownership | ZIP generator stays in the benchmark; OPC/CFB operations use existing public APIs. |
| ADR 0031 budgets | Worker/I/O/CPU-task limits explicit; post-drop permit checks; cumulative CPU-task charge is not treated as a releasable reservation. |

This is low-level CRUD category 15 evidence. It does not promote native-format
CRUD coverage or close physical-cold, delayed-range, cross-session contention,
allocation, hardware-counter or whole-program performance requirements.

The retained [plan](results/change-0786/plan.json),
[source review](results/change-0786/source-review.md), and
[packet instructions](results/change-0786/README.md) describe the exact scopes
and replay contracts. Measured results follow; integration and cleanup are recorded at the end.

## Measured scaling

All 120 qualification children, 720 native children and 240 separate observer
children passed: 22,200 measured samples, including 21,600 native samples.
Six native blocks use forward/reverse case orders fixed before capture, with
30 samples and three warmups per process. Each reported latency is the median
of six nearest-rank process p50s; speedup is the median of the six within-block
width-one/width-N p50 ratios. Bootstrap intervals use 10,000 resamples, seed
786078, and nearest-rank 2.5%/97.5% endpoints. No historical timings are pooled.

Large corpus, zero per-task floor (8 MiB logical payload):

| Route / state | Width-one p50 (ms) | Speedup at 2 | At 4 | At 8 | At 32 |
| --- | ---: | ---: | ---: | ---: | ---: |
| opc / fresh | 0.842658 | 0.5441× | 0.8024× | 1.0265× | 0.6627× |
| opc / primed | 0.839209 | 0.5636× | 0.8650× | 1.2377× | 1.5677× |
| cfb / fresh | 1.233670 | 0.9898× | 1.4938× | 1.9194× | 1.4972× |
| cfb / primed | 1.162970 | 0.9731× | 1.5205× | 2.3949× | 3.1642× |
| parts / fresh | 0.776834 | 0.5405× | 0.8617× | 1.1977× | 0.9773× |
| parts / primed | 0.003045 | 0.0132× | 0.0131× | 0.0101× | 0.0052× |

Primed CFB at width 32 reaches **3.1642×** speedup (95% interval
[3.0674, 3.2992]), or about 9.89% requested-width efficiency. Fresh CFB
performs better at width eight than width 32. Fresh OPC at width 32 is slower
than serial: speedup **0.6627×** [0.6538, 0.6725]. Fresh Parts at width 32
are near serial, **0.9773×** [0.9462, 1.0112]. These are distinct scopes,
not an absolute performance ranking of container formats.

Primed Parts expose the clearest follow-up: width-one p50 is **3.045 µs**,
while width-32 p50 is **578.022 µs**. Paired speedup is only **0.005244×**
[0.005084, 0.005402]. Both per-task floors show the same large degradation.
All primed observer Part samples have zero source calls; their CPU-task charge
still rises from 32 to 64. Source inspection shows scheduling eligibility uses
declared member sizes before reads encounter the payload cache. This warrants
a separately profiled and paired cache-aware scheduling experiment. It does
not yet prove which scheduling component causes the cost or authorize
changing budget/refusal semantics without correctness tests.

At the 64 KiB floor, width-32 small/mixed speedups range from 0.99398× to
1.00237× across the three routes; every corresponding interval includes one.
These controls are serial by eligibility policy despite the requested width.
More requested workers therefore do not imply more admitted parallel work.

All individual cases, tails, RSS and CPU observations are retained in the
[machine-readable analysis](results/change-0786/analysis.json),
[scaling table](results/change-0786/scaling.md), and
[CSV](results/change-0786/scaling.csv). The
[independent raw audit](results/change-0786/raw-audit.json) reconstructs all
96 deterministic payloads and reproduces all 120 paired speedup curves.

## CPU, memory, tails and model limits

Wider execution also costs memory. At width 32 versus width one in the large,
zero-floor cases, median paired whole-child peak-RSS changes are +61.495% /
+62.210% for fresh/primed OPC, +20.572% / +20.202% for CFB, and +31.105% /
+31.027% for Parts. These exceed the program's 5% review threshold. The packet
retains every per-block RSS value; it does not attribute these increases to a
specific allocation or claim an operation-local peak.

Across 120 scaling rows, max/min block spread exceeds 5% in 58 p50 series,
86 p95 series, 111 p99 series, and 22 RSS series. Some spread flag occurs in
112 rows, and 77 rows have at least one latency/RSS increase above 5% against
their same-block width-one reference. These are width comparisons and noise
flags for a baseline study, not before/after production regressions. Thirty
samples make the nearest-rank process p99 the maximum observed sample; tails
therefore remain descriptive and noisy. No blanket tail improvement is claimed.

Process CPU/wall ratios are also not worker counts. For example, primed large
CFB at width 32 records a median CPU/wall ratio of 18.93 while speedup is only
3.16×. The wider CPU interval includes its own clock-boundary overhead, which
matters most for the few-microsecond serial cache-hit control.

The retained Amdahl model fits normalized time as `s + (1-s)/W`. Both an
unconstrained fit and a fit constrained to `0 <= s <= 1` retain their residuals;
per-width apparent fractions remain unclamped. For large zero-floor CFB,
unconstrained fitted `s` is about 0.610 fresh and 0.438 primed, with normalized
RMSE 0.127 and 0.180. The curves are not exact Amdahl behavior. Cached Parts
produce an unconstrained fitted value around 145.5 and a per-width apparent
fraction approaching 197, which cannot be physical serial fractions. They
expose the model's inadequacy when scheduling overhead grows with requested
width. Neither fit establishes a causal CPU decomposition.

## Validation and reproducibility

Six standalone quality gates pass: formatting, all-target/all-feature checking,
three focused tests, warning-denied all-target Clippy, warning-denied rustdoc,
and the repository boundary checker. The tests cover deterministic corpus
identity and argument limits, public read routes with ordered bytes and permit
release, and zero-I/O refusal before a CFB payload read. This is not a claim
that all production-crate test suites were freshly rerun.

The first diagnostic compiler pass found two malformed format strings and an
unused import; the first quality pass found three dead helpers under Clippy.
All were corrected before the successful frozen release build. Their logs
remain retained. No primary capture was retried, discarded or excluded.

The separate native/observer binaries use Rust 1.95.0, release optimization
level 3, thin LTO, one codegen unit, debug level 1, and no incremental build.
The KVM guest exposes 32 AMD EPYC 9R45 CPUs; affinity is 0–31, with no finite
CPU quota in the current cgroup ancestry. Both binaries use the same frozen
standalone lockfile; all 109 external package versions/checksums come from the
workspace lockfile. Exact commands, environment, source and executable hashes
are retained. All 9,196 production files and 35 architecture/goal inputs are
checked against their origin Git blobs by replay.

## Evidence closure

The packet seals 3,304 payload files plus the seal itself. Both executable
hashes were verified before removing 808,040,280 owned target file bytes.
Full replay and the independent raw audit pass with the target absent.
No production source changed. Evidence commit `f23126471b` was fast-forwarded
into `feat/office-format-completeness`. All 3,309 committed packet/tool blob
hashes match retained custody.

The owned worktree, branch, copied root lockfile and three reference symlinks
were removed after their identities were checked. Every pre-existing worktree
and the three unrelated local-file hashes remain unchanged. From the main
checkout, with both the owned target and worktree absent, final sealed replay
and the independent raw audit pass again: 1,080 reports, 22,200 samples,
120 scaling rows and six quality gates. The program-level goal remains open.
