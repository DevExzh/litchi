# 0499: reuse workers within an ordered Part batch

0498 measured substantial many-small local-read regressions from explicit
parallel batches. Its scheduler creates one scoped thread for each admitted
Part, even when a batch spans many waves. The hypothesis is that operation-local
worker reuse can remove this repeated work while preserving deterministic wave
admission. No global pool, runtime dependency, or automatic facade parallelism
is proposed. The before revision is b27718fe8 and its exact benchmark executable
is retained before any production edits.

Keep the existing PartBatch API, output ownership, cache path, prepared metadata,
source/cancellation fences, request cap, serial fallbacks, typed errors and actual
monotonic accounting. Workers must not start later-wave reads until the entire
current wave succeeds. All admitted workers must terminate and join on every
result, including setup failure and an unwinding provider panic. New command or
result storage must be bounded and reserved. Preserve worker/task/declared-byte
limits and conservative stack accounting. Avoid a general-purpose executor.

Capture matched before and after executables with an unchanged harness: two
corpora, owned/warm-file/synthetic short-delay sources, ordinary serial and
batch widths 1/2/4/8, three warmups and thirty measurements for two repeats.
Retain per-repeat/aggregate latency, throughput, whole-child RSS, counters,
byte oracles, cleanup and budget-release evidence, and every adverse flag above
five percent. A separate strace capture counts clone/clone3 calls; traced latency
is not accepted timing evidence. Hardware counters remain whole-child scope.
Neither warm files nor synthetic delay prove controlled cold-cache or real
remote service behavior. Superlinear results and four-Part caps cannot be
misrepresented as simple Amdahl fits.

Use Rust 1.98.1, release debug=0, incremental=0, build jobs=2, dedicated target
/home/zhuhe/.cache/litchi-goal-0499/target. Measurements use CPUs0-7; build/check
work uses CPUs16-31 on the shared host. Protect unrelated ODG/iWork work. Run
focused concurrency/error/ownership tests, OPC default/allfeatures tests, lint,
formatting, documentation and downstream checks. Commit the completed batch,
retain replay executables/evidence, and remove only owned scratch. The full
non-iWork docs/GOAL.md objective remains open beyond this optimization.
