# Change 0496: retain verified-cold cache observations before and after the operation

Date: 2026-09-10

Status: harness capability; no measurement or performance claim

## Problem and scope

The existing opt-in `cold-verified` protocol already isolated each sample in a
fresh child, opened a private page-aligned source read-write, issued
`fsync`/`posix_fadvise(DONTNEED)`, and required a strict external `fincore`
record with zero resident, dirty, and writeback bytes before the timed source
operation. It retained a positive process `/proc/self/io` `read_bytes` delta.
That proved the admission state and process-I/O observation, but the report did
not retain an observation of the same file after the operation.

The harness now performs one post-operation `fincore` query in the same fresh
child, immediately after the timed interval and process-I/O snapshot, before
untimed correctness work. The query is outside the timer.
`Sample.fincore_post` records the post status, file size,
resident/dirty/writeback counters, and the same privacy-preserving tool
provenance and fallback fields as the pre-operation probe. The existing flat
`fincore_size_bytes`, `resident_bytes`, `dirty_bytes`, and `writeback_bytes`
fields remain the pre-operation observation for schema compatibility.

## Admission and fallback policy

The caller must select `--filesystem-cache cold-verified` explicitly. The
default `warm,cold-requested` selection is unchanged. A post probe is accepted
only when it returns one strict record for the expected source size and its
tool identity/method match the pre-operation probe. Post residency is recorded
as an observation; it is not required to be zero because the operation may have
populated the page cache. The pre-operation zero-residency gate and positive
`read_bytes` gate remain unchanged.

If the post probe is unavailable, malformed, has a size/path error, or changes
tool provenance, the sample receives `ineligible_post_fincore` and does not
produce a timed result. The nested post record retains the underlying fincore
status or `provenance_changed` marker. There is no silent fallback to
`cold-requested` or `warm`, and no global `drop_caches` operation.

The protocol remains limited to per-file page-cache residency,
dirty/writeback counters, and process `read_bytes`. It does not claim physical
media I/O, device-cache state, storage latency, or production behavior. The
same harness-only boundary applies to the DOCX verified-cold lifecycle path.

## Host feasibility check

On the review host, a private 8,192-byte file on the allowlisted ext4
filesystem was `fsync`ed and passed a direct `posix_fadvise(DONTNEED)` call.
The installed `/usr/bin/fincore` from util-linux 2.41.3 returned one strict
JSON record with size 8,192 and zero resident, dirty, and writeback bytes. The
probe then read the private file and observed a second size-matched record with
8,192 resident bytes and zero dirty/writeback bytes. It used 4,096-byte pages,
created and removed only its private temporary file, and did not write
`drop_caches` or alter global cache state. This is a capability check for the
per-file protocol, not a workload capture or a cold-cache performance result.

## Verification

The change adds unit coverage for a successful post observation with populated
resident pages, source-size mismatch, unavailable fincore, and provenance
fallback. The canonical Python report validator accepts the additive
`fincore_post` object and the explicit `ineligible_post_fincore` status while
retaining the existing fail-closed schema checks. No workload capture is part
of this change.
