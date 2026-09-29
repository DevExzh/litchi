# 0834 — aligned-source filesystem harness repair

This batch repairs the harness evidence that blocked the filesystem baseline
in [0833](0833-filesystem-cold-qualification.md). Production readers, writers,
and operation timers are unchanged. The evidence packet is
[change-0834](results/change-0834/README.md).

## Cause established before changing the classifier

A separately frozen diagnostic build passed. Its warm PPTX selected-slide
lifecycle passed, while its cold-only request failed during `verified-prime`.
The timed cold child had not run. The retained failure contains chronological
read offsets, requested and returned lengths, the constructor/query boundary,
compressed payload ranges, and the existing aggregate counters.

Independent reconstruction found 883 reads / 169,707 returned bytes. During
construction, one exact 65,536-byte read at offset 16,953,344 overlapped 16,230
bytes of unselected-slide compressed payload. All construction payload overlap
came from that read. The query made two reads / 615 bytes, covering the selected
slide's full 522-byte payload with no unselected-slide or media overlap. This
establishes the classification failure's cause, not a cold latency result.

## Bounded harness correction

PPTX now proves the private archive's exact EOCD zero-padding transformation,
including the pinned ordinary source, aligned source identity, page size,
unchanged EOCD position and original bytes, and zero suffix. Exactly one full
tail request is permitted during construction. Additional constructor payload
overlap is rejected. Query coverage is calculated from query observations
alone; it receives no metadata allowance. Raw counters are preserved, and the
report retains enough ranges and observations for independent reconstruction.
Ordinary unaligned sources keep their existing classification.

OPC save proofs retain each route's raw output hash and length. The source
route preserves the private alignment comment; the eager route emits the
ordinary comment-free archive. A per-route proof checks fixed EOCD bytes,
comment policy, member ordering, and raw unchanged-member bytes. The normal
semantic verifier and exact output digest checks remain active. Combined cold
routes compare a digest with only the proven comment difference removed;
published outputs are not normalized. Warm and advisory-cold routes still
require raw byte parity.

## Validation and retained correction

The first repaired source passed formatting, compilation, and the full library
suite: 569 passed, one ignored, none failed. Clippy then rejected a manual
divisibility predicate. The retained `repaired-v2` amendment changes only
`aligned.len() % page_size != 0` to
`!aligned.len().is_multiple_of(page_size)`; a preceding check already rejects
zero page size. The original source freeze and failed Clippy receipt remain.

The v2 source passed formatting, compilation, all nine affected helper tests,
Clippy, rustdoc, and the boundary check. Fresh qualification passed both cache
states for OPC open and PPTX selected-slide lifecycle. The aligned PPTX replay
retains all 16,230 incidental unselected-slide bytes and the cold child passes
its eligibility gate. The three OPC save commands failed because the new
helper rejected data-descriptor framing present in the generated corpus. That
terminal failure is retained; extending the raw-member proof to cover the
descriptor bytes is required before admission. Ten independent diagnostic-reader
mutation tests passed.

The final `repaired-v3` helper uses the ZIP crate's existing descriptor
validation and extends each raw local span to the next physical local header,
or the central directory for the last member. It therefore compares descriptor
bytes too. Only the helper changes from v2; the main harness and README remain
byte-identical. Fresh formatting, compilation, twelve focused helper tests,
Clippy, rustdoc, and boundary checks all pass, as does the release build.

All seven final qualification commands pass: each of the six selectors in
warm and verified-cold mode, plus the combined OPC save pair in both modes.
The independent reader admits all seven reports / sixteen samples. Across
the diagnostic and both qualification attempts, twelve retained reports
contain twenty-five samples. Four workload failures and the Clippy failure
remain retained. The independent custody audit passes, and 31 reader mutation
tests reject altered diagnostic, PPTX, or OPC proof claims.

The complete source suite was run before the two helper amendments; each
amendment received fresh helper tests and the other quality gates. This does
not represent a second full-suite run on v3. Full corpus warm/cold qualification
provides the final integration check. Source review found no correctness
blocker. Cleanup removed both marked temporary roots: 6,658 files /
3,203,973,694 logical bytes, retaining descriptors for all three release
executables. Report replay, all mutation checks, and the independent audit
pass after cleanup. The final seal binds the exact owned source, report, and
evidence paths and is checked against committed Git blobs.

## Scope

This is a harness repair, not a measured production optimization. OPC eager
open drops its package inside the timer, whereas source open retains it through
post-timer diagnostics. PPTX logical read counters are from an untimed replay.
Verified cold evidence concerns observed page-cache residency and process
`read_bytes`, not physical-device reads. The prospective six-block baseline
has 72 reports / 2,160 samples. It is deferred to a separate batch using the
committed repaired harness. No formal performance sample was collected here.
iWork and unrelated user files are outside the batch.
