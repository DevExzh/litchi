# Non-iWork `docs/GOAL.md` audit

## Current evidence through 0432

[0432](changes/0432-xlsx-streaming-operation-memory.md) closes the missing
operation-allocation observation for the existing XLSX scalar streaming writer.
Across 64/8,192/131,072 rows, all 180 allocator samples retain the same
420,110-byte incremental peak and zero exit live-byte change. This supports
the tested creation path; it does not close richer authoring, logical append,
all-feature memory bounds, cold/remote I/O or explicit scaling. Medium normal
latency repeat drift remains flagged. The [next-work audit](results/change-0432/next-work.md)
identifies ODS bounded scalar-row creation; the [native audit](results/change-0432/native-gap.md)
records the PPTX identity/fixture gap. The non-iWork goal remains active.

## Prior evidence through 0431

[0431](changes/0431-verified-compressed-source-transfer.md) implements and
measures verified compressed source-part transfer in the source-backed PPTX/OPC
workflow. Matched synthetic media API medians improve 86.5–89.0% for bytes,
warm files and short ranges, and 19.4–20.1% for simulated delayed range. The
initial delayed-range regression and its bounded-read refinement remain
reviewable. Refined validation includes 1,755 tests, strict checks and ASAN
smoke. This closes one measured publication hotspot; representative CRUD,
broader native/size coverage, true cold/remote I/O, semantic streaming/append,
allocator/physical-copy attribution, concurrent scaling and remaining strict
gates stay open. See [batch scope](results/change-0431/goal-scope.md). The
broader non-iWork goal remains active.

## Prior evidence through 0430

[0430](changes/0430-pptx-publication-cpu-attribution.md) closes the missing
publication CPU attribution in the synthetic media-rich provider experiment.
Frame-pointer capture on the unchanged binary resolves the Deflate callers;
its measured iteration share is 83.22% / 83.38% for bytes/warm files. The
[source and transfer audit](results/change-0430/transfer-design.md) identifies
compressed-media reuse as the next optimization, subject to preserved
validation, source authority, budgets and archive ownership. This batch makes
no production change or performance-improvement claim. Normal matched
captures, transfer implementation and adversarial validation remain required;
the broader non-iWork goal remains active.

## Prior evidence through 0429

[0429](changes/0429-pptx-provider-native-baselines.md) completes the 32-process/960-sample provider and native selected-
image baseline. It adds a bounded ZIP short-read correctness fix, not a measured
speedup. All final managed gauges are zero, and non-RSS phase observations match
exactly across samples and repeats. All 27 repeat flags are RSS points already
different at entry; comparative RSS conclusions remain withheld. Files are warm,
ranges are explicit simulations, and native selected-image ownership does not
establish native cross-copy. The 100-sample CPU-profile validator correction is
explicitly amended with original validators and unchanged captures retained.
Broader producer/size and CRUD coverage, true cold I/O, allocator/physical-copy
attribution, semantic streaming and explicit scaling remain open. The global
goal is still active; see [batch scope](results/change-0429/goal-scope.md).

## Prior evidence through 0428

[0428](changes/0428-managed-pptx-cache-lifetimes.md) completes the fixed synthetic
managed-cache/budget matrix: 16 processes, 480 samples, exact admission and
refusal, pinning/eviction, oversized bypass, and three publications at an exact
cumulative output ceiling. Every final releasable caller budget is zero. It
adds no production optimization and does not close the global goal. All 27
repeat flags are RSS points with differing entry values; comparative RSS
claims require a new setup/warmup attribution protocol. Native-producer and
cold/range-source lifecycles, the full CRUD taxonomy, streaming, and explicit
scaling remain open. See the [next-work record](results/change-0428/next-work.md).

## Prior evidence through 0427

[0427](changes/0427-pptx-allocator-drop-checkpoints.md) adds explicit allocator
phase/drop observations for the current plain/media-rich PPTX APIs. Eight
fresh processes retain 240 samples, all returning to their entry callback
live-byte value after the final sink drop, with identical repeat phase changes.
This closes the narrow missing caller-drop callback observation. It does not
close the full retention/cache/managed-budget/near-limit gate or establish RSS
release, object-owned bytes or a leak finding. The
[next-work record](results/change-0427/next-work.md) identifies additive
fallible PPTX cache-diagnostic forwarding using existing managed constructors.
Composite harness coverage is 288 passes and one ignored after a baseline-
reproduced stale selector assertion is corrected; the full failed command is
retained. Existing harness strict debt remains the same 29 findings.

An earlier retained allocation optimization is
[0424](changes/0424-staged-pptx-payload-reuse.md), based on production revision
`d18bf7db4` and evidence revision `340cc91ae`. Its matched source-backed media
lifecycle records 14.266% fewer requested allocation bytes and a 7.317% lower
mean V3 operation-region peak. Logical reads and endpoint retention remain
essentially unchanged; normal timing is diagnostic. The standalone bundle
replays 16 reports, four traces and 104 mutation probes after original
worktree/binary cleanup. It establishes no release latency, physical-copy,
managed-budget or post-drop claim.

The intervening [0419](changes/0419-pptx-bounded-archive-growth.md),
[0420](changes/0420-opc-owned-payload-reuse.md),
[0421](changes/0421-allocator-peak-counter.md),
[0422](changes/0422-operation-region-allocator-peak.md), and
[0423](changes/0423-matched-source-backed-pptx-lifecycles.md) records distinguish
allocation requests, retained endpoints, corrected lifetime peaks, V3 region
peaks and aligned lifecycle boundaries. Historical high-water values affected
by the 0421 correction remain unsuitable for current comparisons. There are
nine strict registered claim replays and 15 index categories with 32
representative selectors in the retained 0424 verification; coverage remains
representative rather than complete.

The 0425 audit finds 64 workspace members: 45 non-iWork leaf/shared packages,
the separately configured `litchi` facade, 17 iWork owners, and the Python
binding that unconditionally enables iWork. Shared ZIP and XML owners remain
in scope. The current strict layout finding in ODF's `ArchiveReaderKind`
requires a measured ownership decision; its borrowed reader is retained inline.
Mechanical constant-chunk and test setup findings are closed in the 45 leaf/shared
packages, with strict Clippy passing for 44 and the ODF layout gate open.
[0425](changes/0425-non-iwork-verification-maintenance.md) retains composite
coverage of 15,778 passing tests and 98 ignored tests. The separate facade had
six baseline-reproduced test failures, now closed by
[0426](changes/0426-xls-formula-ancillary-preservation.md): five fixture/assertion
corrections and bounded standard BIFF8 Formula ancillary preservation. The
final facade suite passes 461 tests with 11 ignored; XLS passes 1,341 tests
with one ignored doctest and its scoped strict gate passes. This is correctness
evidence, with static metadata layout cost documented separately and no
performance claim. The facade strict rerun retains the same 18 findings from
the prior audit; no allowance is added.
The standalone harness has 29 lint findings outside the modified code, and
native-resave requires a lockfile refresh before its locked gate can run.
These strict-gate debts remain open; no blanket strict pass is claimed.

Resolve the remaining scoped strict-gate debt and return to the performance
evidence program with facade correctness restored.
Explicit caller-drop snapshots, near-limit memory/cache evidence, broad native
producer matrices, cold/range sources, bounded semantic streaming and append,
scaling/CPU counters, full CRUD coverage and final strict gates remain open.
The full non-iWork goal is not achieved. Sections below retain earlier
revision-specific audits and their historical counts; this section supersedes
only the current facts explicitly listed above.

Change [0418](changes/0418-pptx-cross-copy-candidate-reuse.md) retains a scoped
owned PPTX candidate-reuse optimization. Media lifecycle p50 improves
38.27–38.69%, with an explicitly reviewed approximately 8% whole-process RSS
increase. Two opt-in lifecycle selectors bring the registry to 427; the default
36-case/198-row matrix and representative index statuses remain unchanged.
The batch retains 6,400 normal observations, 480 separate allocator observations,
paired profiles and source/output checks. Retained-memory, broader corpus,
source-backed, native/cold/remote and scaling work remains open.

Change [0417](changes/0417-representative-crud-baseline.md) broadens current
descriptive evidence to all 30 representative selectors across 14 executable
categories at clean revision `b10d6c25a13242ca260a8c897946f4d80ae06c61`.
It retains 30,000 normal observations, 1,800 separate allocator observations,
30 preflight checks, explicit timer boundaries, repeat uncertainty and portable
verification. Five selectors have >5% repeat drift on at least one quantile;
28 lack operation allocation attribution. This is progress on baseline coverage,
with no speedup or completion claim.

The index's default/full-run status contract is unchanged, and several cases
time only an already-open query or commit. Aligned end-to-end workflows,
operation allocation, broad corpus/producer matrices, cold/remote, native,
failure and scaling evidence remain open. The media-rich PPTX path provides
a concrete investigation target with an existing source-backed counterpart
capability; matched timing boundaries and preservation checks are required.

Change [0415](changes/0415-zip64-streaming-deflate.md) adds explicit streamed
ZIP64 Deflate framing, automatic Office transport selection from raised limits,
and strict support for descriptor-backed zero local-size placeholders. Its
transport tests and independent large Python OPC corpus narrow the ZIP64 gap;
semantic streaming, broader append/failure matrices and the performance program
remain open.

Change [0416](changes/0416-local-zip64-read-preservation.md) adds local-only
forced-ZIP64 read/preserve compatibility: 13 deterministic Python fixtures, 10
focused ZIP integration cases, and 7 OPC no-op/edit/failure cases, with strict
path fuzz coverage. The fixtures cover signed and unsigned descriptors,
descriptor-signature CRC collision, seekable no-descriptor output, empty and
many-small members, and central/local or ZIP32-tail combinations. This records
capability coverage only; the source identity, performance guards, and final
gate status belong to the 0416 change record. The overall non-iWork goal
remains open.

The prepared ODF catalog's existing `has_zip64_metadata` observation is based
on central-directory/tail metadata. It does not classify a seekable local-header
ZIP64 sentinel when the central and tail metadata remain ordinary. ODF catalog
parity for that input form, together with semantic, cold/remote, native,
concurrency/scaling, and broader failure-matrix evidence, remains open.

Change [0414](changes/0414-zip64-output-promotion.md) implements preservation
output promotion at ZIP32 local-offset and member-count boundaries, including
OPC provenance for reopened packages above 65,535 members. Generated ZIP64
Deflate through that revision's streaming local-header path remained a typed
refusal; 0415 replaces that refusal with explicit framing. These capability
changes do not close the performance program.

Change [0413](changes/0413-cfb-chain-scratch-reservation.md) retains a scoped
production CFB optimization with nine XLS ABBA comparisons, exact allocation
and logical-I/O checks, a reviewed ~3% CFB guard cost, paired CPU/PMU evidence
and eight strict registered claims. This is verified progress, not completion
of the non-iWork program. Broad scenario, native/cold/remote and scaling gaps
below remain open.

**Audit date:** 2026-09-05
**Audit basis:** the 0418 candidate `f8f9e6667` against control `79dfee502`,
the 0417 representative baseline at `b10d6c25a` and retained evidence,
with the 0416 local-framing candidate `d18cd04a2` and retained evidence,
with the 0415 streaming source batch and its retained evidence,
with the 0414 source batch and its retained verification evidence,
with the 0413 committed control `6b632726b` and the 0412 captured candidate at
`63c95bc22d5883c8ecab0872030757e5584254f7`, with the verified 0411 baseline at
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d` and the earlier
`9c6742c5212dd0e7ff2367da585abe357aae8975` ZIP64 control retained for the
historical comparison. iWork is outside this audit by the user's instruction.

**Disposition: OPEN.** The repository has substantial correctness and scoped
performance work, but the definition of done in `docs/GOAL.md` is not met. The
coverage index explicitly describes itself as representative, and the current
records do not provide a complete, independently reproducible baseline and
optimization result for the non-iWork CRUD matrix.

## 0412 current implementation

Change [0412](changes/0412-xls-observer-isolation.md) at committed revision
`b8f61970d` corrects the XLS observer's category range catalog by coalescing
only exactly adjacent spans, preserving overlap multiplicity, and disabling
the unused generic repeated-read union. It adds three explicit opt-in plain
`OwnedSource` lifecycle selectors: `xls_owned_source_open`,
`xls_owned_source_open_list_worksheets`, and `xls_owned_source_open_one_cell`.
Their timed reports expose operation/allocation metrics with `source: None`;
separate instrumented observations retain the logical locality evidence.

Eleven focused XLS observer/owned/allocator tests plus the registry test pass,
along with scoped formatting, crate boundaries, coverage-index validation, the
seven-claim strict checker, and report classification. The clean candidate
`63c95bc22` now has 18 schema/corpus-verified timing, allocation, profile and PMU
reports. Plain one-cell p50 is 0.166–0.169 ms; separate allocator captures record
126 calls / 223,774 bytes. Instrumented locality remains unchanged across the
observer correction. All 91 comparator tests pass, including mismatched observer
rejection. This is a measurement enabler and baseline, with no production
speedup claim. The registry is now 425 selectors; the default remains 36 cases / 198 rows and the
index remains 15 categories / 30 representative selectors. The non-iWork goal
remains open.

## 0411 current update

Change [0411](changes/0411-xls-read-allocation-baseline.md) now retains a
verified, descriptive baseline for six explicit opt-in XLS/CFB lifecycle
selectors: eager and source-backed open, open plus worksheet listing, and open
plus one-cell selection. The fixed generated corpus is
`xls-comments-opaque-heavy` (`litchi-xls-comments-opaque-heavy-v1`): two sheets
(`Comments`, `Untouched`), selected `Untouched!E21 = 42.0`, 257 logical entries,
10 archive members, 16,995,840 archive bytes, an 80,946-byte `Workbook` stream,
and eight 2 MiB opaque streams. The eager and source-backed paths use the same
two-sheet semantic oracle; this does not assert parity with the separate real
producer corpus discussed in older records.

The [0411 evidence bundle](results/change-0411/) contains four fresh CPU-2,
one-worker normal processes with 20 warmups and 500 samples per selector, plus
two allocator processes with 3 warmups and 30 samples per selector. Samples
share each process; the protocol is warm in-memory corpus evidence, not cold,
remote, concurrent, or per-sample process-isolated evidence. Source-backed
reports retain classified `ReadAt` counters and prove zero worksheet/opaque
payload reads for open/list and selected-only worksheet reads for one-cell.
Allocator reports retain operation-scoped allocation vectors; their peak values
are process-lifetime snapshots, and neither allocation nor timing is an A/B
speedup claim.

The capture and corpus verifier passed on clean revision
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d`. All-feature, all-target
`litchi-xlsx` Clippy passed; the full XLSX test run passed 1,238 tests with no
failures, and the combined harness/allocator test selection passed 9 tests with
four test threads. This evidence expands the current descriptive baseline but
does not promote the selectors into the default matrix or claim completion of
the non-iWork goal. The index remains 15 categories and 30 representative
selectors; the default remains 36 cases / 198 rows and the selector registry
remains 422.

## 0410 current update

Change [0410](changes/0410-mce-attribute-name-reuse.md) is a bounded MCE
expanded-attribute ownership experiment on the XLSX selected-cell path. It
uses candidate `e4d477466718a8fad38cd55b9babe0b826e7f3a7` and control
`972dc25be0dbd6690c74429839a48288d637e2d5`, the fixed four-sheet 9,216-cell,
17-member corpus with 4 MiB of media, and Rust/Cargo/Rustdoc 1.98.1 release
builds on CPU 2. The primary ABBA p50 candidate-minus-control changes are
`-3.9447%` and `-4.1373%`; operation-local allocator calls move from 81,918
to 77,212 and allocated bytes from 10,690,444 to 10,309,094. The seventh
strict-registry claim entry has been added, and the strict checker passes all
seven claims.

The initial eager edit/save guard is adverse at +0.53% / +5.66% p50 elapsed
time. A diagnostic repeat is lower by 1.106% / 0.505%, while same-role p50
drift is roughly 5–6% for both roles. The edit claim is therefore withheld and
the initial adverse result retained. Review found that the eager fixture does
not execute the changed MCE stream; this limits causal interpretation and does
not prove an eager no-regression result. The residual selected-path profile
attributes 10.91% of leaf weight to `clone_bounded_name_part` and reports
`parse_element` at 5.89% self / 23.42% inclusive. These are sampled profile
figures, not paired CPU evidence.

The final all-feature run passes 1,918 tests across the three crates. The
OPC/common warning-denied Clippy passes after two preexisting test-lint fixes;
XLSX all-target Clippy remains blocked by 9 library diagnostics and 28
library-test diagnostics, including 19 additional test diagnostics. This
update records scoped progress only; the non-iWork goal remains open.
The [0410 evidence bundle](results/change-0410/) retains the build identities,
ABBA reports, allocator captures, profile attribution, and gate logs. Rustdoc,
crate-boundary, and scoped-format checks pass.

## Authorities and reading boundary

This audit uses the current goal, the accepted ADRs, and the current
non-iWork performance records:

- [`docs/GOAL.md`](../GOAL.md), especially the mission, measurement contract,
  CRUD matrix, deliverables, and definition of done.
- ADR 0001 (public layers and typed refusals), ADR 0002 (downward crate
  ownership), ADR 0003 (immutable snapshots and atomic edits), ADR 0005 (the
  `ReadAt`/budget/output/measurement contract), ADR 0006 (preserve-by-default
  validation and security), ADR 0008 (migration gates), ADR 0010/0011 (facade
  and OPC physical ownership), and ADR 0024 (current topology).
- [`REPORT.md`](REPORT.md), [`CRUD_COVERAGE.md`](CRUD_COVERAGE.md),
  [`crud-coverage-index-v1.json`](crud-coverage-index-v1.json), and the
  claim-registry policy.

The decisive constraints for the open ZIP work are: preserve is the default;
validation must not mutate; an owned changed OPC source may publish only
through a proven preservation plan; unsupported framing must return a typed
refusal before output; and physical ZIP ownership stays in `soapberry-zip`.
Those constraints come from ADR 0005's exact-source amendment, ADR 0006, and
ADRs 0010/0011.

## Current evidence

| Goal requirement | Evidence at audit time | Assessment |
| --- | --- | --- |
| Reproducible CRUD baseline | The coverage index maps 15 categories and 30 representative selectors. Change 0411 now retains an independently verified six-selector opt-in XLS/CFB descriptive baseline, while measured rows still require a validated `target/perf/container-baseline.json`; the index itself is only an identity/contract artifact. | Partial. The full required scenario matrix and metric set are not evidenced. |
| Scoped claims | The earlier strict checker validated 6 claims; 0410 adds a seventh strict-registry entry, and the current strict checker passes all 7 claims. The report classifier snapshot has 167 rows: 0 `strict_claim`, 145 historical, 14 descriptive, and 8 withheld. | Claims are deliberately scoped; the current check and snapshot counts do not establish program completion. |
| Correctness and boundaries | 0409 checks report 862 Python tests run with 20 skips and no failures, 258 release harness tests passing with one ignored, and a passing crate-boundary check. The final all-feature three-crate run passes 1,918 tests; 0411 adds 1,238 passing all-feature XLSX tests and 9 passing harness/allocator tests under four threads. Current rustdoc, crate-boundary, and scoped-format checks pass. The existing graph has 64 packages, 240 internal declarations, and 14 iWork debts. | Useful gates, but they do not establish latency, memory, I/O, or full CRUD coverage. |
| OPC/ZIP/ODF correctness | The earlier integrated run recorded 362 ZIP, 316 OPC and 287 ODF library tests (1,227 total passed, two ignored). The final all-feature three-crate run now passes 1,918 tests, including the ZIP64 and row-visibility coverage. | Strong package/format correctness evidence for exercised fixtures; not a complete workspace performance or native-producer gate. |
| Full gate health | The prior ODF MIME/unsafe-text false positive and late ZIP64 topology assertion were repaired. OPC/common warning-denied Clippy passes after two preexisting test-lint fixes, and 0411's all-feature/all-target XLSX Clippy now passes. | These scoped gates are green; they do not establish the full performance or CRUD goal. |
| Hardware/resource profiling | `perf stat` counters, 32 affinity CPUs, 128 GiB RAM, Heaptrack, strace, fincore, and cargo-flamegraph are available. | 0406 retains a pinned release OPC materialization baseline, perf counters and self/inclusive reports, Heaptrack, RSS, syscall traces, and a CPU flamegraph. 0408 adds reusable expectations, operation-local allocations and decoded-byte counters, and usable frame-pointer caller attribution. 0409 extends this to XLSX selected-cell and matched edit/save, corrects member-range attribution, and validates native L2 events. 0410 measures the selected MCE ownership candidate and reports residual `clone_bounded_name_part` leaf weight of 10.91% and `parse_element` at 5.89% self / 23.42% inclusive; these are not paired CPU deltas. Generic L1 zeroes are unusable and exact LLC events are unavailable in the guest. The wider matrix remains open. |

The passing library suites establish behavior for the exercised fixtures. They
do not establish the `docs/GOAL.md` requirements for p50/p95/p99, throughput,
allocations, peak RSS, copied/decompressed/recompressed bytes, physical I/O,
lock wait, cold/warm behavior, or Amdahl scaling. The [0404 correctness evidence](results/change-0404/validation.json) retains
commands, source hashes, and compressed logs, including the explicit ODF lint
failure. ZIP/OPC warning-denied Clippy and warning-denied documentation builds
for all three crates pass.

## ZIP64 passthrough audit

The low-level foundation is now materially further along than the last
committed control. `soapberry-zip` retains ZIP64 field origins and tail bytes;
its preservation plan preflights the complete output, patches only movable
offset fields, retains central/local metadata and ZIP64 extensible data, and
has malformed, multi-disk, truncation, descriptor, limit, offset, sink, and
no-output tests.

The earlier integration changed public preservation construction to
`AllowZip64` and admitted already-ZIP64 sources. Change 0414 also promotes
generated/copied central offsets and synthesizes ZIP64 tails for ZIP32 sources.
OPC size/count capability guards and the provenance count ceiling are removed;
arithmetic, limits and structural refusals remain. Earlier synthetic tests
continue to cover:

- `zip64_source_targeted_save_preserves_untouched_records_and_tail` changes
  one Part in a ZIP64 package, retains a different Part's raw local and central
  records, retains ZIP64 tail extensible bytes and the comment, reopens the
  output, and compares `to_bytes` with `write_to_stream`.
- `projected_zip64_descriptor_targeted_save_preserves_untouched_member`
  changes one Part while a different member uses ZIP64 data-descriptor fields,
  then checks raw preservation, reopen, and stream output.

Focused source-backed coverage now also includes
`zip64_one_part_overlay_preserves_raw_unknown_members_and_cold_work` and
`topology_zip64_add_remove_preserves_untouched_records_and_tail`. Both pass in the final integrated test run. The
package-writer suite also has a ZIP64 partial-sink failure case.

These tests are meaningful end-to-end package-writer correctness evidence, but
they are synthetic in-memory fixtures. They leave the following gates open:

1. Extend the source-backed selected-Part and topology matrix to signed/encrypted inputs,
   both descriptor forms, non-seek sinks, cancellation, configured metadata
   limits, and atomic filesystem finalization. Each refusal must leave the
   sink/destination untouched where the API promises that property.
2. Exercise ZIP64 add/remove/topology plans with more than the current narrow
   synthetic graph, including unknown physical members and dependency closures.
   The existing focused add/remove test is valuable correctness evidence but
   does not certify every topology operation.
3. Complete generated-size and streaming creation coverage. Change 0414
   implements offset/count promotion and known-size Store/precompressed ZIP64
   headers. Public sparse tests cover generated offsets immediately below, at
   and above `u32::MAX`; OPC tests cover count promotion and repeated owned
   publication beyond 65,535 members. Copied-offset promotion has focused
   layout/metadata tests. Change 0415 adds explicit one-pass streamed Deflate
   framing and raised-limit Office transport admission, including generated
   metadata charging and descriptor-backed local placeholders. Preservation
   regeneration still buffers payloads; transport-level bounded-memory evidence
   does not certify semantic row/paragraph/slide creation or append.
4. Add at least one real-producer or independently generated large ZIP64 OPC
   corpus and validate all semantic Part bytes, raw member identity, archive
   layout, and reopen behavior. Change 0415 supplies an independently generated
   Python OPC source with a real 4 GiB logical member and a small XML overlay
   preservation check. The 0416 batch adds deterministic Python forced-ZIP64
   local-framing fixtures and focused ZIP/OPC no-op, edit, and refusal cases;
   these remain synthetic compatibility evidence, and the final gate status is
   deferred to its change record. Native Office producer coverage, ODF
   prepared-catalog classification of local-only sentinels, and the broader
   dependency/topology matrix remain open.
5. Expand the scoped gates in
   [`changes/0404-zip64-preservation-integration.md`](changes/0404-zip64-preservation-integration.md)
   to the full non-iWork feature and native-producer matrix. The scoped final
   commands, source bindings, and pass/fail logs are now retained.

The control behavior is also important: at the control revision, a changed
owned ZIP64 source was refused rather than normalized. That refusal remains the
safe fallback for unsupported framing, suffixes, opaque topology, and
unrepresentable generated output. A successful targeted test does not authorize
a normalizing fallback for a different unsupported source.

## Prioritized remaining work

| Priority | Requirement from `docs/GOAL.md` | Next reviewable evidence |
| --- | --- | --- |
| P0 | Close the ZIP64 integration and preservation contract | 0414 adds offset/count promotion and repeated OPC publication beyond the count sentinel. 0415 adds streaming Deflate framing and independent large Python OPC coverage. The 0416 batch adds local-only forced-ZIP64 reader/source-backed compatibility fixtures and focused refusal cases. Semantic bounded-memory creation/append, native producers, ODF local-sentinel catalog classification, and the broader failure-atomicity matrix remain open. |
| P0 | Establish the Phase-1 baseline before selecting further optimizations | 0411 supplies a clean, descriptive six-selector XLS/CFB warm baseline with normal and allocator observations. Continue the full non-iWork capture with p50/p95/p99, throughput, `perf stat` counters, allocation/peak RSS, source calls/bytes/ranges, decompressed/recompressed/copied bytes, output bytes, lock-wait fields, and cold/warm plus explicit bounded-worker cases. |
| P0 | Turn cache correctness into accepted observation | 0405 now retains validated direct-lock acquisition observations, cache counters, and explicit timing scope in 24 normal and 24 observed smoke rows. The release harness passes 256 tests with one ignored. Extend this descriptive evidence to representative workloads with operation-local attribution. Keep cache observations separate from latency claims until an accepted ABBA protocol exists. |
| P0 | Capture current hardware/resource evidence | 0406 binds machine, corpus, revision, binary, raw samples, counters, allocation traces, RSS, syscall evidence, and profiling limitations. 0408 improves caller unwinding and verification efficiency and adds operation-local allocation/ZIP evidence. 0409 records XLSX query/edit/save and usable native L2 events; exact LLC is documented unavailable on this guest. 0410 measures the expanded-name ownership candidate, with residual selected-path attribution but no paired CPU delta. 0411 adds the six-selector XLS/CFB lifecycle and allocation baseline, still without a speedup claim. 0412 isolates diagnostic ReadAt observer cost with plain-source selectors; 0413 retains a scoped CFB reservation optimization with paired CPU/PMU evidence and a reviewed CFB guard cost. Remaining CFB FAT/stream/physical validation costs need attribution before further production changes. The broader semantic CRUD, cold-source, and resource matrix remains open. |
| P1 | Finish source-backed OPC CRUD adoption | Extend selective open/read/edit/save across format facades and topology changes; measure the selected-Part/compressor buffer, physical I/O, allocation/RSS, and semantic phase boundaries. The current report calls this migration incomplete. |
| P1 | Cover the high-impact CRUD categories | Add or explicitly classify conversion, stream append, structural edits, deletion/sanitization, cross-document dependency copy, merge/split, patch and inverse timing, repair/normalize, dynamic calculation, security, malformed, and real-producer scenarios. Correctness-only selectors cannot close the timing requirement. |
| P1 | Finish CFB consumer evidence | Move exact-range and overlay substrate work into DOC/XLS/PPT semantic owners; add physical-cold and high-latency range-source cases, FAT-tail behavior, and operation-local memory/resource attribution. |
| P1 | Validate output and source variants | Cover filesystem cold cache, caller-supplied range/remote sources, sequential non-seek sinks, cancellation, source-version changes, signed/encrypted/macro-enabled files, and atomic replacement across the non-iWork formats. |
| P2 | Demonstrate bounded parallelism | Use explicit execution contexts, then publish scaling curves and serial-fraction/Amdahl analysis. Lock-wait data is currently missing from the accepted report. |
| P2 | Consider layout and SIMD changes | Only after the release profiles identify a hot loop and show a material benefit with scalar fallback and differential correctness tests. |

No row in this audit should be read as a completion claim. The current
evidence supports targeted capability progress and an integrated ZIP64 change
with retained validation; it does not close the full non-iWork performance goal.
