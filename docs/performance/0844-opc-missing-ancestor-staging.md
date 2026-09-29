# 0844 — avoid repeated OPC Part-name ancestor allocations

Adopt the bounded `PartNameIndex::insert` change against `1b6072fc32`. In this generic OPC experiment, 4,096-Part publication improves 3.13% with shared ancestors and 8.15% with shared deep paths. Owner allocation events fall 24.95% and 59.93%, respectively. Unique deep paths cost approximately 0.7–2.3% more p50 time: this is a measured tradeoff, below the frozen 5% review trigger. These synthetic substrate results do not establish semantic Office or mixed-catalog performance. The broader `docs/GOAL.md` program remains open; iWork is excluded and ODF remains deferred.

## Source and contract

Only `crates/litchi-opc/src/members.rs` changes in production. Accepted Part names contribute every proper ancestor, making the ancestor map prefix-closed. The implementation searches borrowed prefixes until the first missing ancestor, then stages only that suffix with one exact Vec reservation. Shared ancestors no longer require temporary Strings or Vec storage. Root paths stage nothing. There is no cache, new hash strategy, worker change, unsafe code, or validation bypass.

Duplicate, equivalent and derived-name checks retain their order. All missing owned strings are prepared before the original names/map reservations and mutation sequence. `or_insert(at)` retains the first descendant representative. Three tests cover folded shared ancestors, zero-capacity deep and partially shared paths, and unchanged maps followed by valid insertion after all four conflict forms. Allocation-error ordering is reviewed statically; allocator fault injection was not performed. See [source review](results/change-0844/source-review.md) and [actual patch](results/change-0844/applied.patch).

## Matched protocol

Both legs use the same standalone probe, root-aligned lock/dependencies, Rust 1.95/LLVM 22, fat LTO, opt-level 3, one codegen unit, debug information, default allocator and CPU 12. Source inventories, binary hashes, symbols and command receipts are retained. Shapes are root `/pNNNNN.bin`, shared `/custom/pNNNNN.bin`, deep-shared `/a/b/c/d/e/f/g/h/pNNNNN.bin`, and deep-unique `/uNNNNN/b/c/d/e/f/g/h/pNNNNN.bin`; the deep paths have nine components and eight ancestors. Counts are 4, 64, 1,024 and 4,096. Each payload is 256 deterministic bytes; the last Part's first byte is XORed with one, with N+2 ZIP members.

Publication times public topology preparation, preservation and emission, including consumed-package drop, while excluding source opening and replacement-plan construction. Lifecycle additionally includes opening and plan construction; fixture cloning and target URI construction remain outside. Hashing, readback and output drop are outside both clocks. Fixture/oracle preparation and validation make this a warm protocol. The independent reader reconstructs payloads and CRCs, checks untouched local spans and changed-target local/central stable metadata, and rejects 34 mutations. All 64 qualification artifacts match exactly across legs and modes.

The main matrix retains 64 qualifications, 384 native reports (six paired blocks, 30 samples after three warmups), and 64 Heaptrack reports (two repeats per shape/count/leg). That is **512 reports / 11,648 samples**, plus two separate one-sample observer preflights. Before any native capture, an order-only schedule repair gave every case three before-first and three after-first pairs; the original build-bound plan remains retained. No formal workload was replaced. Native and instrumented timings are separate.

Ratios below are medians of six paired block ratios, not ratios of the displayed marginal medians. Intervals use 10,000 paired bootstrap draws, seed 844084, endpoints 249/9749. The native p50/RSS review gate requires ratio >1.05 and lower interval >1; every p95/p99 point >1.05 is separately reviewed. All [224 native metrics and 32 allocation comparisons](results/change-0844/tables.md) are retained and independently recomputed.

## Main results

| Shape | Parts | Publish p50 ratio | Lifecycle p50 ratio |
|---|---:|---:|---:|
| shared | 4 | 1.000024 | 0.988665 |
| shared | 64 | 0.980285 | 0.981062 |
| shared | 1024 | 0.971212 | 0.988225 |
| shared | 4096 | 0.968699 | 0.987930 |
| root | 4 | 0.997319 | 0.992147 |
| root | 64 | 0.991569 | 1.002319 |
| root | 1024 | 0.984923 | 1.003015 |
| root | 4096 | 0.982840 | 1.007186 |
| deep-shared | 4 | 0.974781 | 0.972148 |
| deep-shared | 64 | 0.937042 | 0.957152 |
| deep-shared | 1024 | 0.917177 | 0.953441 |
| deep-shared | 4096 | 0.918484 | 0.949176 |
| deep-unique | 4 | 1.007273 | 1.008431 |
| deep-unique | 64 | 1.009688 | 1.014204 |
| deep-unique | 1024 | 1.015912 | 1.023295 |
| deep-unique | 4096 | 1.015340 | 1.021662 |

At 4,096 Parts, shared publication marginal p50 is 3.653081 → 3.534995 ms, paired ratio 0.968699 (95% interval 0.962436–0.972737). Deep-shared publication is 4.967072 → 4.554850 ms, ratio 0.918484 (0.910249–0.928302); lifecycle improves 5.08%. Deep-shared 1,024-Part publication improves 8.28%. Deep-unique 1,024-Part lifecycle increases 2.33% (ratio 1.023295, interval 1.004484–1.027914), and 4,096-Part lifecycle increases 2.17% (1.021662, 1.020427–1.026131). These costs remain visible in the adoption decision.

The largest process RSS point increase is 3.03% (deep-unique, 4 Parts, publish); no p50/RSS gate, allocation gate or block-spread gate fires. RSS covers the full process, including fixture/oracle preparation and validation, and is not operation peak live memory.

At 4,096 Parts, both allocation repeats agree:

| Shape | Owner allocation events before → after | Requested bytes before → after |
|---|---:|---:|
| shared | 32,826 → 24,636 | 11,419,801 → 11,288,761 |
| root | 24,634 → 24,634 | 11,112,063 → 11,112,063 |
| deep-shared | 61,498 → 24,643 | 12,901,064 → 11,770,844 |
| deep-unique | 61,500 → 61,500 | 16,669,037 → 16,669,037 |

Full Heaptrack stacks attribute the exact event and requested-byte reductions to `PartNameIndex::insert`; root and deep-unique controls are unchanged. Requested bytes fall 1.15% and 8.76% for shared and deep-shared shapes. This demonstrates avoided allocation traffic, not peak-live memory, copied bytes, hardware cycles or separately timed phase costs. See [mechanism replay](results/change-0844/mechanism.json).

## Separate tail confirmation

The original 30-sample measurements produce six p99 point estimates above 1.05, all with intervals including 1. With this sample count p99 is the process maximum. Those results remain unchanged. A separately frozen follow-up covers exactly those six cases, six balanced blocks, 1,000 samples after three warmups per child: **72 reports / 72,000 samples**. The same binaries, CPU, clock and artifact contracts apply. The populations are never pooled. Across the main matrix and confirmation there are 584 reports / 83,648 samples, excluding the two observer preflights.

| Shape | Parts | Mode | Original p99 ratio [95% interval] | Confirmation p99 ratio [95% interval] |
|---|---:|---|---|---|
| shared | 4 | lifecycle | 1.063042 [0.742396, 1.446497] | 1.009030 [0.994376, 1.021066] |
| shared | 64 | publish | 1.067077 [0.939803, 1.125040] | 0.976626 [0.971064, 0.997787] |
| root | 4 | publish | 1.068688 [0.948547, 1.173280] | 0.993342 [0.955590, 1.009949] |
| root | 4 | lifecycle | 1.091424 [0.966781, 1.217671] | 0.991642 [0.955446, 1.009464] |
| deep-unique | 4 | publish | 1.184306 [0.737333, 1.683942] | 1.015329 [0.971853, 1.057638] |
| deep-unique | 4 | lifecycle | 1.068557 [0.960863, 1.162011] | 1.011310 [1.007063, 1.029067] |

Confirmation has zero regression, tail-point or spread flags. It resolves the >5% review trigger without proving universal tail equivalence. Deep-unique/4 lifecycle retains small costs: p50 +0.756%, p95 +2.317% (interval 1.006–1.035), p99 +1.131% (1.007–1.029). All remain below the review gate. The [42 confirmation metrics](results/change-0844/tail-tables.md) have an independent raw-sample/interval cross-check. The final closure binds the original six cases, balanced order, plan, capture manifest's exact 72 receipts, analysis and rendered tables.

## Quality, review and disposition

Fresh baseline OPC tests pass: 995 passed, zero failed, one ignored. Candidate affected-owner tests pass: 6,934 passed, zero failed, 38 ignored. Probe tests pass six per leg; independent Heaptrack reader tests pass nine, and ZIP mutation tests reject all 34 malformed inputs. Formatting, workspace all-feature library/test checks, affected-owner all-target checks, workspace library and affected-owner all-target Clippy with warnings denied, affected-owner rustdoc with warnings denied, and crate-boundary checks pass. The baseline gate reuse is source-bound and freshly replayed; iWork examples/full-workspace rustdoc retain their recorded exclusions. See [quality receipts](results/change-0844/quality-after.json).

Independent source/protocol/measurement review supports adoption with the disclosed unique-path cost and synthetic scope. Original source/plan/draft evidence is preserved, including the never-applied invalid handwritten patch and rejected reservation proposals. No formal failing run was dropped. All admitted files remain immutable; the tail confirmation is additive.

The [decision](results/change-0844/decision.json) adopts this batch only. Replay the complete evidence with `python3 -B docs/performance/results/change-0844/close_v2.py verify`; this extends the frozen original closure with confirmation custody and arithmetic. Owned build/scratch roots are removed after verification; cleanup witnesses and the exact committed-file seal remain in the packet. Normative documents and unrelated user changes are preserved. Mixed real-catalog confirmation and remaining non-iWork performance work remain open.
