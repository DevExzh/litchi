# Change 0473: DOCX fresh streaming operation memory

`performance_claim: none; descriptive fresh-creation memory baseline`

`claim_authorized: false`

The new opt-in `docx_streaming_create` selector measures the public
StreamingDocumentWriter. Earlier DOCX small-creation cases build an owned
Package and Vec, so they could not establish the streaming writer's memory
behavior. This is a harness-only addition; production authoring, the allocator
observer and the 37-case default matrix are unchanged. Selectable cases
increase from 439 to 440. Implementation revision is
`a1bd623b3dd0ab23259d2d6f692392ee275475e0`.

The frozen matrix uses 64, 8,192 and 131,072 single-run paragraphs, two reversed
repeats, normal and allocator executables, fresh processes on CPU 2, three
warmups and thirty samples per process. All twelve formal captures pass,
retaining 360 samples; two separate one-sample pilots are excluded.

| Paragraphs | DOCX bytes | Normal p50 ms R1 / R2 | Operation peak above entry |
| ---: | ---: | ---: | ---: |
| 64 | 1,233 | 0.119590 / 0.118971 | 414,732 bytes |
| 8,192 | 24,571 | 11.010543 / 10.973853 | 414,732 bytes |
| 131,072 | 376,848 | 176.711590 / 176.334184 | 414,732 bytes |

All 180 allocator samples have an identical 414,732-byte incremental requested-
heap peak, zero live-byte change at exit and zero failed allocations. This
supports a stable observed operation peak across the tested 2,048-fold range
of paragraph counts. The 64-byte XML-escaping scratch reservation and zero
retained sink output are separate properties; neither is substituted for
measured heap. ZIP/Deflate, context and text-generation allocations are inside
the measured operation, while allocator-internal realloc overlap, physical
copies and process RSS are outside the observer's accounting.

Requested allocation work grows with document size: 107 / 8,235 / 131,115
calls and 1,245,741 / 1,895,981 / 11,726,381 requested bytes for the three
shapes, with four reallocations per operation. Constant observed live peak
does not imply constant allocation work. Source text contains deterministic
UTF-8 and XML-significant characters; each paragraph has exactly one run.

All normal repeat changes in mean/p50/p95/p99 are within five percent; the
largest absolute change is tiny p99 at -3.263%. These thirty-sample timings
remain descriptive. There is no normal-versus-allocator timing comparison,
production speedup claim or new registered latency claim. Full elapsed
samples, alignment, dispersion and Student-t mean intervals are retained and
independently validated. Full-process RSS is 82,608–82,740 KiB for normal
captures and 82,684–82,740 KiB for allocator captures. It includes setup and
the untimed materializing oracle, and does not establish writer RSS causality.

Before timing, an artifact is fully reopened to verify the exact ordered
three-member package, paragraph counts, every paragraph and its one-run text,
full-text and paragraph digests, and finalized document XML identity. The
artifact and XML are dropped before samples. Each timed output must match
its archive hash, bytes and checked writer counters. The pre-finish XML getter
excludes the known closing suffix, which is accounted against the complete
preflight part. Timer/allocation boundaries include context construction,
generated text, writer creation, all streaming calls, ZIP finalization and
writer/context destruction; sink construction/digest extraction and observer
reads are outside.

Validation passes 363 release harness/allocator-target tests with one ignored,
plus 18 writer unit and four integration tests. Eleven Python evidence tests
include semantic, statistic, allocator-balance, scope, identity, chronology and
seal mutations. Formatting, warning-denied Clippy/rustdoc, crate boundaries and
the strict ten-claim registry pass. The first focused compile's unused-local
warning and the initial full-suite stale-count failure remain retained with
their passing final checks. Independent initial source audit completed;
follow-up review agents hit a service usage limit, and root completed final
source/evidence review locally.

Both measurement executables come from the same clean source checkout, Rust
1.98.1, frame pointers/unwind tables and release debug level 1. The inventory
binds 7,033 tracked Rust/TOML/lock files, including 39 historical probe files,
and two compile fixtures. A temporary-filesystem quota error occurred after
successful compilation while copying the first binary. Sparse checkout freed
unrelated tracked files while retaining all authenticated sources. The two
ignored fixture paths then needed exact-hash restoration; the formal capture
precheck refused before any sample until they were restored. Both incidents,
source custody, successful binary copies and chronology remain documented.
No formal capture was discarded or rerun, and the frozen protocol is unchanged.

The [bundle](../results/change-0473/README.md) contains reproducible commands,
raw reports, strict schema/arithmetic checks, source/build/capture bindings,
seals and portable replay. The copied strict report helpers are adapted from
0432's existing checks; DOCX semantic identities are independently regenerated
in Python. Output archives are not exported, so portable replay verifies the
retained producer oracle/digest evidence rather than independently reopening
those discarded archives.

This proves only the tested fresh plain-text creation measurement. Logical
append, package Part addition and arbitrary edits/repackaging remain separate.
PPTX fresh streaming needs its own measurements and explicit directory-metadata
accounting; see [next work](../results/change-0473/next-work.md). Native Office,
source variants, feature breadth, worker scaling and the full non-iWork goal
remain open.
