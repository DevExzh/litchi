# 0913 — owned-source deferred ZIP reads

The candidate is **rejected**. Bypassing the temporary ZIP member-span buffer
for retained owned OPC input did not meet the frozen native gate: the paired
owned total ratio was **1.005714×**, with a 95% bootstrap interval of
**0.973308–1.017486**. The required ratio was at least 1.05 with a lower bound
above 1. Both production files were restored and the three candidate-only
tests removed from production; their exact bytes remain in the evidence packet.

The evaluated candidate changed two production files. `soapberry-zip` adds a specialized
owned-source read method that selects the existing passthrough reader. Generic
`ReaderAt` sessions retain their bounded member-span optimization. The deferred
OPC payload loader calls the specialized method. Both routes retain the same
local-header parser, ZIP64 handling, output reservation, overrun sentinel,
Store/Deflate decoder, CRC/size/descriptor verification and payload accounting.
No unsafe code, shared decoder, new cache or ordinary format API is introduced.

The candidate bypassed the speculative member-span allocation, zero-fill and copy
from an already retained immutable byte vector. Output decoding and its owned
allocation remain. This is a source-backed mechanism, not a measured allocation
count or a claim that all document copies disappear.

## Native result and rejection

All 48 children passed, covering 960 measured samples and 2,400 guarded
operations. The independent reducer exactly matches all 40 compared metric
summary and bootstrap groups. Four raw ZIP censuses each verify 445 members,
437 unchanged members, exactly the eight expected changed slides, physical
ordering and archive comments. Source and output digests are identical across
all children. These correctness results do not turn the failed timing gate
into an adoption.

| Route | Binary | Total p50 ms | p95 ms | p99 ms | Process max-RSS median KiB |
| --- | --- | ---: | ---: | ---: | ---: |
| materialized | baseline | 257.804112 | 265.534590 | 300.143136 | 95,528 |
| materialized | candidate | 262.021552 | 269.942558 | 286.545322 | 95,386 |
| owned | baseline | 32.179962 | 33.915470 | 34.467653 | 95,614 |
| owned | candidate | 32.259566 | 34.194186 | 50.413489 | 96,590 |

The materialized paired total ratio was 0.984090× (95% interval
0.973052–0.990175). Owned snapshot p50 was 15.838492 ms before and
15.836998 ms for the candidate. No route showed a >5% RSS regression;
no allocation-count or peak-live-byte benefit was measured.

All nine >5% flags remain visible in `analysis.json`:

- Materialized clone p95/p99: +5.51%/+10.41%; ingress p95/p99:
  +7.67%/+11.49%; snapshot p99: +9.62%.
- Owned ingress p99: +30.93%; snapshot p99: +16.83%; publication p99:
  +53.52%; total p99: +46.26%.

These tails are observations on a shared host; this packet does not assign
all outliers causally to the source change. They are not removed or rerun.
Order strata also vary: owned baseline AB/BA total p50 is 32.910934/31.894078 ms,
and candidate is 32.428125/31.959316 ms. Both primary admission conditions fail
without relying on an interpretation of these secondary flags.

A post-capture static admission census narrows the original hypothesis. The
baseline member-span window is capped at 64 KiB. All eight media members have
2,097,797 compressed bytes and already use passthrough in the baseline. Across
all 445 physical members, 437 are span-eligible; one read of each eligible
member would admit 211,315 span bytes. This is an archive-layout bound, not a
measured call count, allocation count or dynamic copied-byte total. The
16,777,216 decoded media bytes still require the existing decode/output path.
The candidate therefore did not remove an intermediate span copy of the large
media. Future payload optimizations must first qualify actual size-class and
force-path exposure; the 0911 guest memcpy counts alone do not establish it.

## Evidence boundaries

The candidate is built from the exact 9,355-file 0910 source reconstruction,
plus two production overlays, two test overlays and an ownership marker. The
retained 0910 executable is the baseline. The toolchain and release flags match:
Rust 1.95.0, opt-level 3, LTO disabled, 16 codegen units, debug information off,
incremental compilation off, two build jobs, locked offline dependencies.

The frozen plan runs 12 matched baseline/candidate process pairs for each of
the materialized and owned routes, alternating AB/BA order. Each process runs
30 warmups and 20 measured operations on CPU 12. The generated input has 200
slides, eight text boxes per slide, eight 2 MiB media members, 229 OPC parts and
445 ZIP members. Each operation edits 64 text boxes over eight slides. Source
clone, ingress, transaction snapshot, edit/commit, apply/serialization, the
preallocated sequential sink and route-local destruction stay inside the timer.
Digest checks, semantic guards and artifact writes stay outside it.

Ratios compare baseline against candidate within each route. Pooled p50/p95/p99
are descriptive; paired process medians are the bootstrap unit, with 10,000
draws and seed 913. Adoption requires an owned total median ratio of at least
1.05 and a 95% lower confidence bound above 1. Any route's latency or
whole-process maximum RSS regression above 5% requires review. RSS includes
fixture generation, warmups, guards and output artifacts; it is not operation
peak live bytes. Capture children receive the same explicit PATH, locale and
UTC environment.

## Correctness and architecture

Focused tests cover Store/Deflate bytes and accounting, local ZIP64 metadata,
data descriptors, CRC and size failures, truncated framing, source bounds,
limits, concurrent first success/refusal, decoded allocation sharing and exact
source publication. Archived Office fixtures and the supplemental adversarial
test are admitted only for quality checks, then removed to restore the exact
frozen benchmark inventory before capture.

Quality attempt v8 passed 8,096 tests across 330 suites, with 50 ignored and
none failed. Scoped formatting, warning-denied Clippy and rustdoc passed for
ZIP, OPC, DOCX, PPTX, XLSX and XLSB. The boundary check passed separately.

| Constraint | Candidate behavior |
| --- | --- |
| ADR 0001/0002/0010/0011/0024: ownership and layers | Physical ZIP work stays in `soapberry-zip`; only OPC selects it. |
| ADR 0003/0006: preservation and validation | Source authority, verification, edits and publication are unchanged. |
| ADR 0005/0031: bounded work and execution | One temporary copy is bypassed; no new parallelism or ambient I/O. |
| ADR 0030: deferred access and limits | Existing first-force cells, stable failures and declared/actual budget checks remain. |

Source review identified existing ADR 0030 wording debt: its clone paragraph
describes independently cloned cells, whereas the current implementation shares
an `Arc<DeferredPayload>` and its outcome across clones. This candidate preserves
that implementation behavior; the tests establish no new clone contract, and
this batch does not claim to resolve the prior wording mismatch.

The full-file formatter has one existing `office.rs` census-formatting delta.
The scoped gate requires exactly the same baseline delta and no new formatting
violations. It does not claim a clean whole-workspace formatting gate.

## Scope

Offline replay passed both before and after cleanup, with byte-identical
results. It reconstructs all 9,355 baseline and 9,358 candidate source files,
checks the 48 capture packets and eight quality attempts, and reproduces the
statistics, four ZIP censuses, span-eligibility census and immutable iWork
commit checks. The owned performance source/build roots were removed,
reclaiming 4,790,351,009 logical bytes. Together with iWork cleanup, this batch
removed 5,939,537,187 logical bytes and its only created worktree.

The [evidence packet](results/change-0913/) retains the failed attempts,
candidate patch and five exact candidate source files. The candidate executable
and 96 generated capture PPTX files remain local and gitignored; retained native
quality fixtures are tracked. Full historical binary/artifact replay requires
those local files. A fresh clone can inspect the tracked source, recipes,
digests and measurements, but cannot run that custody replay without restoring
the retained binary and generated artifacts.

The packet retains each quality attempt. Early attempts exposed a stale root
lock, an unavailable offline dependency, and missing compile-time or runtime
fixtures from the archived source inventory. The quality lock preserves all
benchmark registry package versions; supplementary native inputs are retained
with checksums. The seventh attempt reached the supplemental adversarial test
and exposed an incorrect test assumption about the existing outer I/O error
classification. Its failed receipt is retained separately from the correction.

An environment continuation removed the temporary source directory before the
eighth quality attempt. `source-recovery.json` records exact reconstruction of
the 9,358-file candidate inventory and readmission of 1,445 quality fixtures.
The retained benchmark binary did not change. The kernel changed from
`7.0.0-1012-aws` to `7.0.0-1013-aws`; CPU details, Rust versions and capture-tool
hashes remained identical. Fresh measurements of both binaries use the latter
environment, bound from `environment-after-recovery.json` in every child.

The executed freeze explicitly includes both production files and the two
initial test files. Do not rerun the immutable `run.py freeze` in place. The
measurement entry point is `capture.py`; the unused `run.py capture` branch
has a directory-creation collision and is not part of the executed protocol.
`analyze.py --output NEW.json` reduces the retained measurements, and
`replay.py --output NEW.json` reconstructs source and checks the packet without
building or running either benchmark executable.

This is one warm, generated media-rich PPTX workflow on a shared Linux host.
It does not qualify real producer diversity, cold/range workloads, parallel
scaling, allocation counts or native hardware-counter cost shares. OLE2/OOXML
coverage remains incomplete; ODF is deferred and iWork is excluded. Integration
of performance results on the shared target branch is tracked separately from
this batch's private-index commit.

The user additionally requested resolution of the pending iWork cherry-pick.
That separate integration admits the neutral ingress feature with its exact
required clone-inventory predecessor; it contributes no performance result to
this experiment. Its isolated source, quality attempts and cleanup are retained
under the packet's `iwa-*` records. The initial shared-index preservation
condition is superseded only for this explicitly requested integration.
Commit `1b12cef9a97bbe3597039f7e3caadc2c45c905ce` completes that cherry-pick
after 302 passing iWork tests, doctests, Clippy, rustdoc and boundary checks.
Its temporary worktree and build directory were removed, reclaiming
1,149,186,178 logical bytes. The root lock file was unchanged; the isolated
quality lock refresh retained every existing registry package version.
