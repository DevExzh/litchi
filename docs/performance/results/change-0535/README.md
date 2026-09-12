# 0535: instruction-address attribution of the CFB exact chain collector

The previous batch made progress: it rejected the slower physical role/FAT
paired-prefix loop and retained independent regression coverage. This batch
keeps the runtime unchanged and investigates the largest remaining named
exclusive constructor owner, `SectorChainScratch::collect_exact`.

The frozen protocol selects XLS owned-source open/one-cell and direct CFB
few-large opens. Two native repeats use 20 warmups and 1,000 samples per case
(4,000 native samples). Four constructor profile children use no warmups and
five samples each: 20 timed dumps plus two separately classified CFB setup
dumps. Callgrind records instruction addresses, with line positions disabled,
absolute positions and jump metadata enabled. Generated assembly comes from
the same new normal binary. There is no candidate or before/after gain claim.

The validated analysis maps exclusive collector instruction costs to the bound
binary's instruction boundaries, after validating relocation consistently.
It keeps constructor inclusive costs, collector exclusive costs and direct
callee costs separate. Positive incoming edges distinguish timed constructors
from fixture setup. Call and jump metadata may contain collection-off activity;
those counts are not operation-scoped allocations or branch counts. Profile
elapsed times are excluded from native timing evidence.

The four guarded allocation vectors were unchanged in the prior experiment,
and no allocation or source/harness change occurs here. The normal executable
reports allocation metrics unavailable. No allocator-instrumented timings,
mixed-workload aggregate speedup, hardware cache inference or causal timing
explanation is introduced by this diagnostic batch.

Current production and harness source exactly match the final 0534 manifest
`9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75`.
Reuse of that source's 14 quality gates and 4,382 test executions requires
manifest, receipt, prior verification and seal custody checks. Python analyzer
checks are separate and do not increase those Rust test execution counts.

The coordinator owns the serial build and captures. Storage is confined to
`/home/zhuhe/litchi-goal-0535-target` and its `/tmp/litchi-goal-0535` alias;
owned TMPDIR and `--vgdb=no` apply from the first relevant child. Both paths
have been removed after cumulative verification. Retained evidence replays
without the executable. The standalone harness has its own workspace and no explicit
release LTO profile.

OLE2 and OOXML remain active; ODF is deferred until their optimization goal
completes, and iWork is excluded. The paired-prefix loop, visited-bit fusion
and freshness-session proposals remain rejected. This diagnostic scope does
not complete the broader CRUD, producer, cold/range or scaling program.

The completed capture inventory has 19 successful serial receipts. The native
analysis retains all 13 same-build variation flags, including CFB few-large
p50 drift of +10.0959%; no timing cause or gain is inferred. The instruction
report maps all 22 dumps and proves that collector instruction self costs
reconcile with function self costs. The valid loop already contains checked
bitset updates and direct vector append; it has no hot call to the separate
`CheckedBitSet::insert` symbol. See [mechanism review](mechanism-review.md)
and [conditional next experiment](next-candidate.md).

The verifier confirms exact-source reuse of the prior 14 quality gates and
4,382 Rust test executions. Five in-memory verifier tamper probes also pass.
No new Rust test execution or runtime optimization is claimed.
