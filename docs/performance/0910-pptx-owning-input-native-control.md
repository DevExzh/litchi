# 0910 — comparable native PPTX owning-input control

The existing owning-input route has a pooled median of **32.180007 ms**, versus
**259.703332 ms** for materialized ReadAt, on the qualified generated fixture.
Across twelve process pairs, the median materialized/owned ratio is
**8.072147**, with a paired bootstrap 95% interval of **7.960538–8.240714**.
This is a same-build comparison of existing routes, not a production change
or a historical before/after speedup. It corrects the scope of the earlier
eager/source comparison: owning input already avoids most of the measured
publication cost.

## Native results

The [raw reduction](results/change-0910/analysis.json) validates all 24 children
and 480 measured samples. Every warmup and measured operation passes the same
semantic and byte-identity gates, totaling 1,200 validated operations.

| Metric | Materialized ReadAt | Owning `from_vec` |
| --- | ---: | ---: |
| Total p50 | 259.703332 ms | 32.180007 ms |
| Total p95 | 267.737926 ms | 33.901569 ms |
| Total p99 | 285.221701 ms | 36.383925 ms |
| Input clone p50 | 0.810824 ms | 2.903215 ms |
| Ingress p50 | 6.194567 ms | 2.971650 ms |
| Opened snapshot p50 | 12.349500 ms | 15.865198 ms |
| Stage edits and commit p50 | 2.868150 ms | 2.942265 ms |
| Apply, serialize and sink p50 | 236.840172 ms | 5.575988 ms |
| Whole-process maximum RSS, median | 95,990 KiB | 96,166 KiB |

Phase medians are not additive. Total also includes route-local destruction
and timing bookkeeping. Owning input spends more time in the clone and
snapshot phases; those increases exceed the program's 5% review threshold
and remain visible despite the lower total. Opening and snapshot capture
must be interpreted together because lazy ingress can defer payload work.
The copy phase is also affected by allocator and cache state, which this
experiment does not isolate.

The median paired publication ratio is 42.497504 (95% bootstrap interval
41.593130–43.239137). This is an elapsed-time phase comparison, not a count of
codec operations. Owning snapshot capture is now the largest measured phase.
That is the next attribution target; another broad recompression change is
not justified for this owning route by this evidence.

BA/AB pooled total-p50 ratios are 0.999057 for materialized and 0.982161 for
owning. Whole-process RSS medians differ by about 0.18%; the observed ranges
are 94,984–96,956 KiB and 95,080–96,988 KiB respectively. These process envelopes
do not establish equal operation allocations or peak live bytes. Raw phase
vectors, per-process results, tails, order splits and bootstrap data remain
available; the interval is conditional on this shared-host capture.

## Workload and measurement boundary

Both routes begin with the same borrowed archive and clone its bytes inside
the primary timer. The owning route moves the clone into `Package::from_vec`.
The materialized route constructs an uninstrumented `OwnedSource` and
`Arc<dyn ReadAt>`, opens `SourceBackedPackage`, consumes it with
`into_opc_package`, and constructs `Package::from_opc_package`.

Each then captures the same opened transaction, replaces 64 text boxes across
eight of 200 slides, commits, applies, calls `to_bytes`, and writes an identical
sequential sink in chunks of at most 64 KiB. Sink reservation is outside the
timer. Route-local destruction is included; sink destruction and semantic
verification are excluded. Patch replay is outside this production workload;
its bounded correctness qualification remains in 0909.

This differs from 0907, which excluded the input clone and included inverse
and forward patch replay. Its build also differs. Neither the 0907 native
numbers nor the 0908 guest instruction counts are a numeric baseline here.

## Reproduction and custody

The source parent is `468fe2931d6532c32f4c6608590816b8eaa88e31`, with the
retained 0907 lockfiles and a three-file harness overlay. The isolated build
uses Rust 1.95.0, release optimization, LTO disabled, 16 codegen units, debug
information disabled, and two Cargo jobs. Existing defaults and selector
registrations remain unchanged.

The [capture plan](results/change-0910/capture-plan.json) specifies 12 paired
blocks, alternating materialized/owned and owned/materialized order. Each
fresh process runs 30 warmups and 20 measured operations on CPU 12. Captures
run serially on a shared KVM host with an AMD EPYC 9R45 and 32 visible
logical CPUs; CPU pinning does not establish exclusive host access. The
generator, semantic checks and output hashing are outside each operation
timer, but can affect allocator/cache state and process RSS.

Raw reports retain every measured operation and phase duration. Inference uses
the 12
paired process medians. A deterministic bootstrap resamples complete pairs
10,000 times with seed 910; pooled p95/p99 are descriptive. Whole-process
maximum RSS includes generation, all warmups, samples, checks and artifact
writes, and must not be interpreted as operation peak live memory.

## Preservation and interpretation

The fixed source contains 229 logical OPC Parts and 445 ZIP members, including
eight deterministic 2 MiB media Parts. Every measured output is checked for
all requested and untouched text, slide topology, relationships, content types,
untouched logical payloads and media. Representative outputs from both routes
pass the [independent 0909 raw ZIP census](results/change-0910/census-runs.json):
exactly eight selected slide members change and 437 remain raw-component
identical, excluding central-directory relocation offsets. Both physical and
central member order and the archive comment are unchanged. All 24 outputs
are 17,017,145 bytes with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.
The source is 17,017,139 bytes with SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The generated archives use data descriptors and no ZIP64; broader descriptor
and ZIP64 forms are not qualified here.

Owning ingress retains archive preservation provenance. The materialized
conversion does not. This supplies a source-supported mechanism hypothesis;
output equality alone does not establish which members were copied or
regenerated, or how much decompression, compression or allocation occurred.

No cold-cache, filesystem, remote/range, concurrency, allocation, hardware
counter, or real-producer result is established by this generated in-memory
control. ADRs 0003, 0005, 0006, 0011, 0024 and 0030 remain unchanged: existing
transactions, preservation, ownership boundaries and ingress APIs are used.
Broader non-iWork performance work remains open; ODF is deferred.

## Validation and cleanup

The release build, scoped formatting, and warning-denied Clippy all pass.
The [source review](results/change-0910/source-review.md) binds the final
three-file harness overlay; [analysis](results/change-0910/analysis.json)
checks exact schedule, source/build/binary custody, output identity, sink
bounds and phase arithmetic. The independent raw ZIP checks pass for both
routes. The source review's initial sink-boundary correction was completed
before the frozen build and capture.

Offline replay reconstructs all 9,355 source files, validates the retained
binary and all capture artifacts, reruns both censuses, and reproduces the
canonical statistics byte-for-byte. Replay passes before and after cleanup.
Cleanup removed 1,538,498,852 logical bytes from the owned source/build roots;
no worktree was created. The native executable and generated archives remain
as explicit Gitignored local evidence with tracked identities. A fresh clone
must rebuild the binary and regenerate those archives; the packet does not
embed them in Git.

To run the standalone control from the reconstructed source, use the build
flags and command in [the build receipt](results/change-0910/build-v1/command.json)
and then invoke the example as `owned 30 20 NEW_OUTPUT_DIR` or
`materialized 30 20 NEW_OUTPUT_DIR`, pinned to CPU 12. The source manifest,
Git archive paths, lockfile overlays and binary overlay patch identify the
complete build input. The [capture plan](results/change-0910/capture-plan.json)
records the exact 24-child order and the [runner](results/change-0910/run.py)
records the commands and environments. Each artifact directory must be new.

```sh
python3 -B docs/performance/results/change-0910/replay.py \
  --output /tmp/new-0910-replay.json
```

Target-branch integration remains pending the unrelated shared iWork
cherry-pick. The next measurement should attribute the owning route's cold
opened-snapshot capture before selecting a production optimization.
