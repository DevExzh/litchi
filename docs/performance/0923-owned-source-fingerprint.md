# 0923 — owned-source artifact fingerprint reuse

## Opportunity and scope

The 0915 DOCX profile attributes 111,584,396 guest instructions to source
fingerprinting in the query snapshot and 111,584,412 in the publication's
current snapshot. Each contains 109,404,406 SHA-compression instructions.
These establish repeated work; guest instruction counts do not establish
native SHA time or predict a speedup. The 0914 generic-source baseline's
query/publication phases together account for 89.98% of timed nanoseconds.

The proposed change concerns the library's move-owned source ingress.
`SourceBackedPackage::from_vec` currently wraps immutable owned bytes and then
erases that ownership distinction into `Arc<dyn ReadAt>`. A private proof tied
to that exact owner could permit reuse of its artifact digest. A generic
provider's version token cannot provide this proof: bytes may change while
its token remains equal. Generic and path-backed sources must retain complete
freshness checks and hashing. The emitted artifact's digest is a distinct
identity needed for inverse authorization and must still be computed.

The new owned-input benchmark must use the same generated DOCX sections,
media, editor, six timed phases, output, preservation, patch, refusal and sink
controls as 0914. It is a separate input route, not a replacement or a claim
about the generic `ReadAt` scenario. Its baseline includes the already adopted
0916 target-reuse change. Both versions will use identical benchmark code;
all candidate work stays isolated until source review and measurement pass.

This batch does not reapply rejected 0915, 0918, 0920, 0921 or 0922 patches.
The prior turn is classified as progress: it committed the complete 0922
experiment and independently reproducible rejection. The broader performance
goal remains unproven; iWork is excluded.

## Fresh pre-implementation baseline

Production source is `4e2532a0ce958f5928496b0ad9b356d8636916e1`, with an
identical, separately bound harness extension for all compared versions. Before
candidate implementation, each route completed 30 warmups and 200 measured
samples pinned to CPU 12. All per-route semantic, preservation, exact inverse, cache,
source-observation and sink checks passed. The shared negative-case suite
runs through generic `ReadAt`; its combined flag does not prove owned-route
negative-case coverage. Focused owned-route tests are required separately.

| Ingress route | Baseline p50 (ms) | p95 (ms) | p99 (ms) | Observed input |
| --- | ---: | ---: | ---: | --- |
| Generic `ReadAt` | 4.5047135 | 4.539419 | 4.569819 | 202 positional reads, 8,409,488 bytes |
| Owned vector | 4.732154 | 4.764059 | 4.836450 | 2,101,427 moved bytes; no positional observer |
| Owned reader | 4.8427695 | 4.884550 | 4.936350 | 258 sequential reads, 2,101,427 bytes |

These are separate baseline characterizations, not a cross-route speed
comparison or candidate result. Paired baseline/candidate runs will compare
each route with itself. The generated source is 2,101,427 bytes with SHA-256
`183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8`;
the expected output is 2,101,446 bytes with SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`.

The prospective native schedule is eight alternating AB/BA pairs per route:
48 processes, 24 per version, each with 30 warmups and 200 measured samples.
Both owned routes must achieve a median paired baseline/candidate p50 ratio
of at least 1.05 with a lower 95% bootstrap bound above 1.0. The generic
control must retain a lower bound of at least `1 / 1.05`. Tail latency and
whole-process RSS regressions remain explicit review flags. A failed or
invalid capture fails the batch; it cannot be selectively replaced.

The independent baseline audit reproduces the source archive and harness
patch and verifies the source tree, binary and report identities. Original
build/run start timestamps and build logs were not retained; receipt creation
times must not be treated as those events. The initial source receipt also
names a listing hash without retaining its file. Any reconstructed listing
is identified as post-capture evidence. These custody limitations remain
visible and are not repaired by inventing historical records.

## Reviewed implementation and correctness

The candidate changes only `crates/litchi-opc/src/source_backed.rs`. Explicit
owned constructors retain their exact immutable owner alongside a shared
one-cell digest memo. Generic, filesystem, validation and managed ingress
remain uncached, including `OwnedSource` erased behind generic `ReadAt`.
The first fingerprint keeps the existing full scan and error order; only a
successful digest initializes the cell. Hits check freshness before and after
lookup, and debug builds rederive the digest from the retained bytes. New
ingress creates fresh state. Publication hashing and authenticated inverse
copying are unchanged.

The cell uses actual fallibly reserved storage with an exact capacity check.
On this target its logical heap cost is 36 bytes for the cell and 48 bytes for
the memo plus Arc control block, excluding allocator metadata. Each retaining
package/artifact adds an 8-byte optional Arc field. The strong source handle
shares existing bytes; it adds no payload copy. Reservation occurs after the
existing successful open checks, so it cannot mask an earlier malformed-source
or input-limit error.

Independent quality evidence is in
[shared-owner-quality-v3.json](results/change-0923/shared-owner-quality-v3.json).
The broad tested revision passes 622 OPC library tests, 1,231 DOCX library
tests, 79 DOCX source-backed tests and 29 OPC restore/read-ahead tests. The
final revision differs only in four formatting hunks and passes three focused
fingerprint tests, ten callback tests, and OPC/DOCX warning-denied Clippy and
rustdoc. The focused public-API tests also pass the baseline and exercise
managed repeated-call accounting/cancellation, stable-version generic mutation,
short reads and invalid read lengths. Full formatting checks reproduce
pre-existing baseline differences; final added lines have no formatter drift.

The initial isolated test failures were missing fixture errors, retained in
the evidence. Restoring the exact fixtures enabled the successful full rerun.
Benchmark and quality lockfiles are separately bound.

## Build and measurement custody

The machine reports AMD EPYC 9R45, Linux 7.0.0-1013-aws,
132,553,801,728 bytes of memory and 4 KiB pages. All 48 children are pinned
to CPU 12. Their environment field `rustc_version` reports the runtime
`rustc --version` probe (1.98.1), not the compiler that produced the executable.
Actual build receipts and fingerprints bind Rust 1.95.0. Those two observations
are kept distinct; binary identity comes from the executable SHA-256.

The final production patch is 15,772 bytes, SHA-256
`77c2a9952335d5f16c24fbfc8d87e120301808258f70e990324fa21b8ea8e75d`.
An independent comparison of all 9,355 files per source tree finds exactly
that one production difference; applying the patch in memory reconstructs
the candidate exactly. The identical harness patch is 17,046 bytes, SHA-256
`5404cf09f7f840aa3a587d1bbdce6659d0d3a8cf35d2c44bbd42b747ddf217e0`.

The native binary SHA-256 values are
`ed2d5609cc7a54459cd33627a80416d5ce2ef765b32916b5235b91209e19db28`
(baseline, 69,962,624 bytes) and
`8d98309501f63f4911828f52c42653f5950f3f5019bd040cf98f787173528861`
(candidate, 70,379,960 bytes). All 238 shared build-fingerprint records agree
on compiler, profile, features, target, flags, configuration and compile kind.
A candidate receipt mistakenly copied root-workspace profile settings into
its `profile_overrides` field. A separate corrective audit records that those
settings were not applied to the standalone benchmark workspace. The original
receipt is preserved; profile parity rests on the actual build fingerprints.

The native manifest is frozen at SHA-256
`b3089f7dcd5cc1517d22615a066e16088a50b4edbd4e703f7594bc5e66a568ae`.
Heaptrack will provide a separately scheduled, whole-process allocation
diagnostic on these same native binaries. The existing allocator entry point
has no operation region for this case, so an attempted allocator build was
stopped and retained as setup evidence. No operation-scoped allocation or
profiler-time latency claim is made.

## Native result

All 48 children completed successfully, with 9,600 measured samples in total
and no replacement runs. Source closure was revalidated afterward. All three
paired timing gates pass. Ratios below are baseline p50 divided by candidate
p50, reduced over eight process pairs; confidence intervals use the fixed
10,000-draw bootstrap and seed 923.

| Route | Paired median ratio | 95% interval | Gate |
| --- | ---: | --- | --- |
| Generic `ReadAt` | 1.005503 | [1.001920, 1.006873] | Pass |
| Owned vector | 1.271812 | [1.269746, 1.276512] | Pass |
| Owned reader | 1.265597 | [1.260773, 1.322012] | Pass |

The following quantiles pool all 1,600 raw samples per route/version.
They describe the captured population; the paired intervals above govern
timing eligibility. Root uses the midpoint median for even sample counts.
The supplemental reducer also records nearest-rank p50, which differs by
less than 50 ns here; p95 and p99 agree exactly.

| Route | Baseline p50 (ms) | Candidate p50 (ms) | p50 change | Baseline → candidate p95 (ms) | Baseline → candidate p99 (ms) |
| --- | ---: | ---: | ---: | --- | --- |
| Generic `ReadAt` | 4.5074135 | 4.486534 | −0.46% | 4.546789 → 4.523239 | 4.577289 → 4.553019 |
| Owned vector | 4.717960 | 3.705361 | −21.46% | 4.751660 → 3.732636 | 4.775951 → 3.762306 |
| Owned reader | 4.8601805 | 3.832221 | −21.15% | 5.085882 → 4.057707 | 5.109252 → 4.080597 |

Whole-process median maximum RSS changes by −0.53%, +0.52%, and −1.41%,
respectively. No total-latency or RSS 5% review flag fires. This does not
establish a general generic-provider improvement: its role is a regression
control. The owned-route benefit is scoped to this generated DOCX section-edit
and publication workflow.

During reduction, root found that the frozen runner's `pooled_latency_review`
pooled child summary quantiles rather than raw samples; its phase summary
also summarized child medians. Those fields are retained as historical output
but must not be interpreted as pooled operation quantiles. The paired p50
bootstrap is correct and unchanged.
[root-raw-sample-summary.json](results/change-0923/root-raw-sample-summary.json)
independently binds the raw reports and supplies the actual pooled values
above. No timing data were recaptured or discarded to correct reporting.

The phase data are consistent with removing the second source hash during
publication: owned-vector publication p50 falls from 2.687537 to 1.691903 ms
(−37.05%), and owned-reader publication from 2.684862 to 1.681707 ms
(−37.36%). Initial-query p50 changes by less than 1% in either owned route.
Owned-vector open p50 rises by 0.835 µs (+1.90%); reader open rises by
2.330 µs (+0.61%). The generic control's preparation p50 rises by 5.769 µs
(+16.88%), while its total p50 falls slightly. This small phase regression is
retained rather than hidden by the total; no cause is established. Phase
medians are descriptive and need not add to the median total.

## Memory diagnostic and protocol deviation

The prospective memory lane was 24 Heaptrack processes, separate from native
timing. The first process completed, but the analysis command expected
`heaptrack.zst` while Heaptrack had written `heaptrack.zst.zst`. The workflow
then incorrectly restarted a full 24-process cohort despite an instruction
to preserve the first capture and wait before continuing. Thus **25 processes
were captured**, and the original 24-process qualification is invalid. The
original profile, failed analysis logs, first protocol/runner, separate rerun
and corrective analysis are all retained. The earlier “no recapture” claim
is explicitly superseded by
[allocation-corrective-supplement-v2.json](results/change-0923/allocation-corrective-supplement-v2.json).
This failure did not affect the native capture or its frozen artifacts.

An independent reviewer reparsed all 25 profiles and verified their identities:
[independent memory disposition](results/change-0923/independent-memory-disposition.json).
The original baseline profile has exactly the same allocation totals as the
rerun's generic baselines. The values below are descriptive whole-process
diagnostics, including setup, correctness checks, serialization and teardown.
They are not a clean 24-process qualification or operation-scoped memory result.
No captured profile is discarded.

| Route | Allocation events, baseline → candidate | Requested bytes, baseline → candidate | Peak live heap bytes, baseline → candidate |
| --- | --- | --- | --- |
| Generic `ReadAt` | 312,717 → 312,717 | 151,804,036 → 151,804,057 | 24,234,801 → 24,234,805 |
| Owned vector | 312,705 → 312,706 | 151,803,020 → 151,737,589 | 24,234,029 → 24,234,117 |
| Owned reader | 312,964 → 312,965 | 423,388,539 → 423,323,108 | 24,234,083 → 24,234,171 |

Every profile ends with 544 live bytes. Owned routes show one additional
allocation and 88 additional peak-live bytes, while requested bytes fall by
65,431. This is consistent with the small retained memo and removal of one
64 KiB fingerprint scratch allocation, with process setup differences also
included. It is not evidence that all allocation changes occur inside the
timed operation. The memory measurements show no material broad regression,
but their protocol deviation remains a limitation.

## Disposition and limits

Adopt the exact reviewed production patch and focused tests. The unchanged
native gates pass with a practically useful owned-route latency improvement,
correctness/refusal checks pass, and the fixed memory cost is explicit. The
allocation profiles are supplemental descriptive risk evidence; adoption does
not declare their original qualification valid.

The result covers one generated 2.1 MB DOCX section-edit/publication fixture,
two owned-input routes and a generic regression control on this machine. It
does not establish other format, corpus, cold-cache, concurrency, hardware
counter or program-wide gains. A memo miss still scans the complete source,
and every newly opened owned package pays the fixed memo allocation even if
no fingerprint is requested. The broader `docs/GOAL.md` objective remains
active and unproven; this batch is progress, not completion or a blocker.

## Reproduction and retained artifacts

Native capture/replay is implemented by
[comparison-capture-v2.py](results/change-0923/comparison-capture-v2.py); the
retained packet is
[capture-final](results/change-0923/capture-final/capture-final-receipt.json).
Actual pooled statistics are independently reproducible with
[comparison-pooled-v3.py](results/change-0923/comparison-pooled-v3.py).
The root replay independently copies the committed packet twice, verifies
local hashes and obtains canonical-equal results for all 48 children:
[root-native-replay.json](results/change-0923/root-native-replay.json).

Each native child uses `taskset -c 12`, the appropriate binary and
`LITCHI_DOCX_INGRESS_ROUTE=generic_read_at|owned_vec|owned_reader`, with
`--case docx_source_backed_existing_section_layout_edit_save --samples 200
--warmup 30 --json REPORT`. Each source tree builds with
`CARGO_TARGET_DIR=TARGET cargo build --locked --offline -j2 --release
--manifest-path SOURCE/tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`
using Rust 1.95.0; exact commands, lockfiles, environment and binary identities
are bound in the manifest and corrective profile audit.

Large artifacts stay outside the repository under
`/tmp/litchi-performance-0923/`. The pinned source archive
`base-4e2532a0ce958f5928496b0ad9b356d8636916e1.tar` is 170,065,920 bytes,
SHA-256 `a6880fb4f638fbdcff23afb728382ea0c807f3af201648f579844309492b10a9`.
Regenerate it with `git archive --format=tar --prefix=litchi/
4e2532a0ce958f5928496b0ad9b356d8636916e1 .cargo Cargo.toml clippy.toml
crates tools/perf-baseline`; add the two bound compile fixtures and lockfile,
then apply the retained harness patch and, for the candidate, production patch.
Complete source listings are 1,437,118 bytes each, with baseline SHA-256
`988d3c1d40cbeefc8f937eb2a3f635c1fd003ea36589f1329029f5fef6537509`
and candidate SHA-256
`5eb6840368d0e601e2db4756ae2dbc4c7fdbc4d0b8bfd810934652a95090a83a`.
Their reconstruction and all other external artifact size/hash/command
bindings are recorded in the custody receipts. Binaries are never committed.

Cleanup removed 11,062,934,160 logical bytes of build/test intermediates and
duplicate source/archive copies, retaining final binaries, fingerprint records,
measured source closures, locks, all raw captures and all failure/correction
records. The redundant `base-head.tar` archive is retained byte-identically
under the pinned archive name above. See
[cleanup receipt](results/change-0923/adoption-cleanup-post-v1.json).

The adopted harness additionally includes six formatter-only hunks applied after
measurement. The measured source and binaries remain unchanged; the separate
[formatting receipt](results/change-0923/harness-format-validation-v1.json) binds
both patches and source hashes and verifies equivalence after whitespace and
optional trailing-comma normalization. The final harness passes rustfmt. No
rebuild or recapture was used for this formatting-only adoption step.
