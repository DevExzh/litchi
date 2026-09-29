# 0845 — retain indexed entry IDs through OPC publication

Adopt direct indexed-entry iteration against `d6619c101f`. In this warm generic OPC experiment, 4,096-Part publication p50 improves 7.24–11.17% across four path shapes; lifecycle p50 improves 2.31–4.18%. Owner allocation counts and requested bytes are unchanged. No frozen latency, RSS, allocation or block-spread review gate fires. These synthetic substrate results do not establish semantic Office or mixed-catalog performance. The broader `docs/GOAL.md` program remains open; iWork is excluded and ODF remains deferred.

## Source and contract

The ZIP owner adds a doc-hidden `IndexedArchive::file_entries()` iterator yielding opaque IDs and normalized names in physical local-header order. Two OPC publication scans retain those IDs for metadata access and physical-member lookup construction, eliminating repeated string lookup. The size/reservation pass remains unchanged. No cache, new hash strategy, worker, unsafe code, compression or durability change is introduced.

The immutable index guarantees ID/name correspondence, so the removed missing-ID branch was unreachable for a valid index. Limits, checked totals, progress/source checks, relationship classification, folded ambiguity detection, reservation ordering and first representative behavior remain intact. IDs stay inside the existing ZIP/OPC owner seam. The focused test reverses central-directory order, excludes directory entries, checks normalized `dir/../second.bin`, and verifies ID/name/metadata correspondence and exact iterator length/exhaustion. See [source review](results/change-0845/source-review.md) and [applied patch](results/change-0845/applied.patch).

## Matched protocol

The [host record](results/change-0845/host.json), source inventories, binary hashes, symbols and command receipts bind both legs. The recorded host is AMD EPYC 9R45 on x86_64 Linux 7.0.0-1012-aws. Both use the identical 0844 probe and dependencies, Rust 1.95/LLVM 22, fat LTO, opt-level 3, one codegen unit, debug information, system allocator and CPU 12. Retaining the probe's 0844 schema is deliberate; the drivers and plan identify batch 0845.

Shapes are root `/pNNNNN.bin`, shared `/custom/pNNNNN.bin`, deep-shared `/a/b/c/d/e/f/g/h/pNNNNN.bin`, and deep-unique `/uNNNNN/b/c/d/e/f/g/h/pNNNNN.bin`. Counts are 4, 64, 1,024 and 4,096. Each payload is 256 deterministic bytes; the final Part's first byte is XORed with one, with N+2 ZIP members.

Publication includes public topology preparation, preservation, emission and consumed-package drop. Lifecycle additionally includes opening and replacement-plan construction. Fixture cloning and target URI construction precede both clocks; hashing, readback and output drop follow them. Fixture/oracle preparation and validation make this a warm protocol. Independent ZIP reconstruction validates payloads, CRCs, logical digests, untouched raw spans and changed-target stable local/central metadata. All 64 qualification artifacts agree exactly across legs and modes.

The main matrix has **512 reports / 211,328 samples**: 64 qualifications, 384 native reports and 64 one-sample Heaptrack reports. Each native case has six paired blocks, three before-first and three after-first, with three warmups and 1,000 samples at counts 4/64 or 100 at counts 1,024/4,096. Two one-sample Heaptrack preflights and a separately planned 16-report / 16-sample instruction population remain separate. No native and instrumented timing populations are pooled.

Ratios are medians of six paired block ratios. The 95% intervals use 10,000 paired bootstrap draws, seed 845085. The p50/RSS review gate requires ratio >1.05 and interval lower bound >1; every p95/p99 point >1.05 is separately reviewed. Allocation ratios >1.05 and block p50 max/min >1.20 also trigger review. All [224 native metrics and 32 allocation comparisons](results/change-0845/tables.md) are retained; the native metrics were independently recomputed from raw samples.

## Native and allocation results

| Shape | Parts | Publish p50 ratio | Lifecycle p50 ratio |
|---|---:|---:|---:|
| shared | 4 | 0.980700 | 0.987107 |
| shared | 64 | 0.933294 | 0.976422 |
| shared | 1024 | 0.906610 | 0.966004 |
| shared | 4096 | 0.888293 | 0.958164 |
| root | 4 | 0.978520 | 0.988337 |
| root | 64 | 0.948912 | 0.980274 |
| root | 1024 | 0.929497 | 0.970714 |
| root | 4096 | 0.911912 | 0.964444 |
| deep-shared | 4 | 0.980154 | 0.988230 |
| deep-shared | 64 | 0.935892 | 0.976577 |
| deep-shared | 1024 | 0.912894 | 0.966847 |
| deep-shared | 4096 | 0.904003 | 0.968907 |
| deep-unique | 4 | 0.974647 | 0.981034 |
| deep-unique | 64 | 0.949515 | 0.978765 |
| deep-unique | 1024 | 0.933319 | 0.976818 |
| deep-unique | 4096 | 0.927642 | 0.976926 |

Shared 4,096-Part publication marginal p50 is 3.549550 → 3.156874 ms, paired ratio 0.888293 (95% interval 0.884338–0.891285). Root, deep-shared and deep-unique publication ratios at that count are 0.911912 (0.901928–0.921693), 0.904003 (0.899019–0.905905), and 0.927642 (0.923515–0.934196). All 32 p50 point estimates improve. The highest p99 ratio is 1.007768 for deep-shared/4 publication; the largest RSS point increase is 1.89% for root/1,024 publication. There are zero latency/RSS, tail-point, allocation or block-spread flags, so no separate tail confirmation was required.

Both allocation repeats agree exactly before/after in all 16 shape/count cases, for both owner allocation events and requested bytes. This matches removal of lookup work without changing allocation structure. Process RSS includes fixture/oracle preparation and validation; it is not operation peak live memory. No copied-byte, cache-miss, hardware-cycle, IPC, parallel-scaling or cold-cache improvement is claimed.

## Instruction evidence and reader repair

The instruction comparison uses the same binaries and scoped publication owner, for shared/deep-unique at 64/4,096 Parts, two repeats per leg. Instrumented elapsed times do not enter native estimates. Full self-cost conservation, one positive owner call and zero-cost terminal dumps are mandatory.

The fourth observer command exited successfully, but the inherited reader ignored valid Callgrind `fe` source-return records and then falsely rejected a function source-file change. The [Callgrind format specification](https://valgrind.org/docs/manual/cl-format.html) defines these source-file switches. The additive `reader_v2.py` recognizes `fe` and the `cfl` callee-file alias while preserving the function-identity and cost-conservation checks. Its regression suite replays 34 inherited mutations, all 64 qualification ZIPs, compressed/plain source-return cases and all four retained profiles. The legacy reader's failure remains reproducible. The original reader, runner, admission, failed parent receipt and all raw captures remain unchanged.

The original three completed run receipts are preserved byte-for-byte. A pre-resume inventory binds all four existing captures, the missing fourth receipt is completed from its successful command and corrected validation, and only the remaining twelve scheduled observers are executed. The additive manifest and analyzer use their actual runner/reader identities; no successful original capture manifest is fabricated.

All **16 reports / 16 samples** pass the repaired full reader. Owner guest-instruction counts are:

| Shape | Parts | Repeat | Before Ir | After Ir | After/before |
|---|---:|---:|---:|---:|---:|
| shared | 64 | 0 | 2,128,247 | 2,018,865 | 0.948605 |
| shared | 64 | 1 | 2,128,706 | 2,019,602 | 0.948746 |
| shared | 4096 | 0 | 73,212,905 | 66,530,622 | 0.908728 |
| shared | 4096 | 1 | 73,560,347 | 66,731,271 | 0.907164 |
| deep-unique | 64 | 0 | 3,135,700 | 2,971,572 | 0.947658 |
| deep-unique | 64 | 1 | 3,134,574 | 2,970,150 | 0.947545 |
| deep-unique | 4096 | 0 | 136,589,838 | 127,216,060 | 0.931373 |
| deep-unique | 4096 | 1 | 137,363,455 | 126,442,415 | 0.920495 |

Shared 4,096-Part owner Ir falls 9.13–9.28%; deep-unique falls 6.86–7.95% across the two repeats. The selected `lookup_member_name` self cost falls exactly 75% in each case, and the selected `entry_id` function is absent from the after profiles. This supports the intended removal of repeated lookup work. Inclusive regions overlap and are not summed; these counts do not estimate native phase time or hardware cycles. See [full instruction analysis](results/change-0845/profile-analysis-v2.json) and [resume custody](results/change-0845/profile-resume.json).

## Quality and ADR boundaries

Fresh baseline ZIP/OPC tests pass: 1,653 passed, zero failed, three ignored. Candidate affected-owner tests pass: 6,935 passed, zero failed, 38 ignored. Each probe leg passes six tests; Heaptrack reader tests pass nine, and the inherited independent reader rejects 34 malformed inputs. Formatting, workspace all-feature library/test checks, affected-owner all-target checks, workspace library and affected-owner all-target Clippy with warnings denied, affected-owner rustdoc with warnings denied, and crate-boundary checks pass. Baseline gate reuse is source-bound and freshly replayed. The recorded iWork examples/full-workspace rustdoc exclusions remain explicit. See [quality receipts](results/change-0845/quality-after.json).

| Accepted ADR boundary | Evidence in this change |
|---|---|
| 0001/0002/0010/0011/0024: API layers and archive/OPC ownership | Opaque indexed IDs remain in the existing low-level seam; no facade representation or dependency change. |
| 0003/0006: atomic publication, lossless preservation, validation | Same checks and emission path; exact before/after source and output qualification; unchanged conflict and resource semantics. |
| 0005/0032: measured work removal and bounded memory | Matched native/observer evidence; unchanged allocations; no retained metadata memo, cache or worker. |
| 0030: fallible lazy OPC access | Existing fallible metadata access remains; only its admitted lookup key changes. |

The [decision](results/change-0845/decision.json) applies to this batch only. Replay with `python3 -B docs/performance/results/change-0845/close_v2.py verify`; the additive closure binds the reader repair and the original failed parent command without accepting any failed workload. Owned build/scratch roots are removed only after verification, with cleanup witnesses and an exact committed-file seal retained. Normative documents and unrelated user changes are preserved.

The next priority is current semantic DOCX/XLSX/PPTX mixed-catalog publication and durability diagnosis, including the large changed-Part control. No further synthetic OPC micro-optimization is justified before confirming its end-to-end relevance. The remaining non-iWork program is not complete.
