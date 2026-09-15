# 0578: ZIP passthrough is already bounded — a refuted hypothesis

Status: refuted, no production change. `performance_claim: none` — this record
carries deterministic peak-retained-byte and allocation counts only. **No timing
is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## The hypothesis

The plan for this change was that
[`ZipArchiveWriter::write_precompressed_file`](../../crates/soapberry-zip/src/writer.rs)
and its `_with_accounting` / `_classified` siblings take the compressed payload
as `compressed: &[u8]`, a fully materialized buffer, and that
`IndexedArchive::read_entry_precompressed_*` allocates a `Vec` of
`compressed_capacity` to produce it. Copying one unchanged 50 MB embedded video
from source to output would therefore materialize 50 MB of compressed bytes, and
peak memory would scale with the largest unchanged member rather than with a
bounded window — leaving the DEFINITION OF DONE clause

> unchanged large media and package members can flow from source to output
> without unnecessary decompression or **logical-byte copies**

unmet on the copy side. The planned remedy was a streaming passthrough entry
point with a named bounded scratch, in the shape of change
[0570](0570-cfb-fat-run-batching.md)'s discipline.

**The hypothesis is wrong, and nothing was implemented.** Both halves of it are
individually true as statements about those two function signatures, and both
are irrelevant, because neither function is on the path an unchanged member
takes from source to output. That path was already bounded before this
investigation, by a mechanism the plan did not account for.

The headline measurement: a source-backed OOXML save whose unchanged media
member grows from 64 KiB to 64 MiB holds its peak retained heap **flat at 532,626
bytes and its allocation count flat at 1,167**, and at the ZIP layer a single
**4.06 GiB** unchanged member reaches a sequential sink behind **1,916 retained
bytes and fifteen allocations**.

This record is the evidence for the refutation. Changes
[0567](0567-ooxml-single-index-per-open.md) and
[0569](0569-ooxml-detect-then-open-priced.md) are the precedent for retaining one.

## What was removed

**Nothing.** No production code, test, or fixture was changed. The only files
added are this record and its evidence directory.

## The three things the plan did not account for

### 1. `write_precompressed_file*` has no production callers

A repository-wide census of every occurrence, with each call site classified by
whether it sits inside an in-crate `#[cfg(test)]` module or under a `tests/`
integration-test directory:

| symbol | definitions | internal delegations | in-crate test call sites | `tests/` call sites | **production** |
| --- | ---: | ---: | ---: | ---: | ---: |
| `write_precompressed_file` | 1 | 0 | 4 | 3 | **0** |
| `write_precompressed_file_with_accounting` | 1 | 1 | 4 | 0 | **0** |
| `write_precompressed_file_classified` | 1 | 2 | 0 | 0 | **0** |
| `write_generated_deflate_file_with_accounting` | 1 | 0 | 1 | 0 | 1 |

The four in-crate test call sites of the public entry point are `preserve.rs:5225`
and `office.rs:14159`, `:14181`, `:14282`; the three `tests/` ones are
`directory_spool.rs:80`, `:1134`, `:1252`. The `_with_accounting` test sites are
`writer.rs:3248`, `:3258`, `:3434`, `:3810`. The internal delegations are the
public entry point calling `_with_accounting` (`writer.rs:1491`) and
`_with_accounting` and the generated-Deflate helper both calling `_classified`
(`writer.rs:1517`, `:1536`).

The single production call in the table is the last row's: `office.rs:6765`,
inside the streaming **create-from-scratch** writer. It compresses a member the
library has just generated, so it is not a passthrough. **No production code
calls `write_precompressed_file` in any form.**

A second fact closes this off independently: `VerifiedPrecompressedEntry`'s
accessor is

```rust
pub(crate) fn compressed_payload(&self) -> &[u8]   // office.rs:772
```

so an out-of-crate caller **cannot** feed a verified token into
`write_precompressed_file` even if it wanted to. The probe written for this
record hit exactly that compile error. The `&[u8]` signature the plan targeted is
reachable only by a caller that has already materialized the bytes itself, and no
such caller exists in the repository.

### 2. The real copy path retains a range, not bytes

An unchanged member is planned as `PreservationAction::Copy(id)` and prepared as

```rust
PreparedLocal::Copy(source_entry.local_span.clone())   // preserve.rs:900, a Range<u64>
```

`PreservedEntry` holds `local_span: Range<u64>` and no payload bytes at all. The
emission is `preserve.rs:2374`'s `copy_range`, which walks that range through a
caller-supplied buffer:

```rust
source.read_exact_at(&mut buffer[..len], offset)?;
write_all_counted(sink, &buffer[..len], accounting, AccountingWriteKind::RawUnchangedSource)?;
```

and the buffer is one **stack** array allocated once per `write_to` call, not per
member:

```rust
const COPY_CHUNK_SIZE: usize = 64 * 1024;             // preserve.rs:31
let mut copy_buffer = [0u8; COPY_CHUNK_SIZE];          // preserve.rs:778
```

This is precisely the bounded named-constant shape change 0570 asks for. It
predates this investigation.

### 3. Both production save paths already use it

| | Path A — `PackageWriter` | Path B — source-backed |
| --- | --- | --- |
| entry | `litchi-opc/src/pkgwriter.rs:884` | `litchi-opc/src/source_backed.rs:9566` |
| index built at | `pkgwriter.rs:209` | `source_backed.rs:9587` |
| unchanged member | `PreservationAction::Copy`, `pkgwriter.rs:641/661/673` | `PreservationAction::Copy`, `source_backed.rs:9771` |
| edited member | `RegeneratedEntry::new_shared` + Deflate, `pkgwriter.rs:777` | same, `source_backed.rs:9757` |
| written by | `index.write_to(&plan, Chunked { .. })`, `pkgwriter.rs:704` | `index.write_to(..)`, `source_backed.rs:9832/9867` |
| source | `ZipArchive::from_slice` — already resident | `IndexedArchive<SourceReader>` — positional |

Neither path constructs a precompressed entry for an unchanged member. Path A
has no precompressed path at all. `litchi-docx`, `litchi-xlsx`, `litchi-pptx`
and `litchi-ooxml-common` contain no callers of the precompressed API; they reach
`PackageWriter` or `write_topology_to_stream` and nothing else.

## Measured effect

Peak retained bytes and allocation counts for a save that passes unchanged
members from a positional `FileReader` source to a **sequential, non-seek** sink
that accepts and discards every byte. Measured through a process-global counting
allocator with the same wrapper shape as
`tools/perf-baseline/src/bin/support/counting_allocator.rs`, over a detached git
worktree of `32d25e08806d93f792ffd4954d83acc9db9c5301` with an isolated
`CARGO_TARGET_DIR`. Environment and the probe source are in
[`results/change-0578/`](results/change-0578).

Three scenarios are measured against the same synthetic package — 21 small
deflated XML members plus one large `word/media/big.bin` member:

- **A**, copy-through: `Copy` for every member, one small member `Regenerate`d.
  This is the production save shape.
- **B**, precompressed token: capture the large member as a
  `VerifiedPrecompressedEntry` and republish it through
  `RegeneratedEntry::new_precompressed_shared`. This is the cross-document part
  copy shape.
- **C**, the `&[u8]` API the hypothesis named, fed by a caller that materializes
  the member's compressed range itself, because nothing else can call it.

### Axis 1: peak against the size of the largest member

Stored member, 21 other members held constant. Every figure is region peak
retained bytes.

| largest member | output bytes | **A copy-through** | B token | C `&[u8]` API |
| ---: | ---: | ---: | ---: | ---: |
| 64 KiB | 69,406 | **427,498** | 146,285 | 66,034 |
| 256 KiB | 266,014 | **427,498** | 539,501 | 262,642 |
| 1 MiB | 1,052,446 | **427,498** | 2,112,365 | 1,049,074 |
| 4 MiB | 4,198,174 | **427,498** | 8,403,821 | 4,194,802 |
| 16 MiB | 16,781,086 | **427,498** | 33,569,645 | 16,777,714 |
| 64 MiB | 67,112,734 | **427,498** | 134,232,941 | 67,109,362 |

**A is flat to the byte across a 1,024-fold growth in the member**, and so is its
allocation count (61 allocations, 431,719 allocated bytes, at every size). The
Deflate series is identical: 427,498 at every size, 62 allocations. B is
`2 × member`; C is `1 × member`.

The flat 427,498 is not the passthrough. It is the preservation index plus one
Deflate encoder state for the single *edited* member — both independent of the
member being copied. Removing the edited member isolates it:

| largest member | output bytes | A pure copy-all peak | allocations |
| ---: | ---: | ---: | ---: |
| 64 KiB | 69,411 | **12,971** | 53 |
| 256 KiB | 266,019 | **12,971** | 53 |
| 1 MiB | 1,052,451 | **12,971** | 53 |
| 4 MiB | 4,198,179 | **12,971** | 53 |
| 16 MiB | 16,781,091 | **12,971** | 53 |
| 64 MiB | 67,112,739 | **12,971** | 53 |

At 64 MiB that is **5,173 output bytes per retained byte**. The Deflate series is
again identical at 12,971 / 54.

### Axis 2: what the residual actually scales with

Member size fixed at 4,096 bytes, member count varied, pure copy-all:

| members | archive bytes | peak | allocations | peak per member |
| ---: | ---: | ---: | ---: | ---: |
| 20 | 84,302 | 12,124 | 30 | 606 |
| 100 | 421,422 | 58,316 | 112 | 583 |
| 500 | 2,107,022 | 282,364 | 514 | 565 |
| 2,000 | 8,428,022 | 1,129,456 | 2,016 | 565 |
| 8,000 | 33,712,022 | 4,517,824 | 8,018 | 564 |

Peak is linear in the **central-directory record count**, at roughly 565 bytes
per member, and independent of payload size. That is the retained physical index
OPTIMIZATION WORKSTREAM B asks a successful open to keep, not a payload copy.

### Axis 3: real corpus fixtures

Pure copy-all of the whole package. The corpus's largest single member is the
882,682-byte stored TIFF in `ArtisticEffectSample.pptx`.

| fixture | archive bytes | largest member | peak | allocations |
| --- | ---: | ---: | ---: | ---: |
| `ArtisticEffectSample.pptx` | 972,788 | **882,682** | **33,583** | 69 |
| `saut_page.docx` | 2,959,626 | 217,495 | **17,022** | 70 |
| `no_drawing_patriarch.xlsx` | 672,414 | 336,090 | **7,536** | 22 |
| `ConditionalFormattingSamples.xlsx` | 654,688 | 57,356 | 84,124 | 145 |
| `EmbeddedVideo.pptx` | 201,418 | 101,799 | 24,169 | 50 |

The largest media member in the entire corpus flows through at a peak **26 times
smaller than the member itself**. `ConditionalFormattingSamples.xlsx` has the
highest peak of the five despite the smallest largest-member, because it has 132
members — axis 2 again.

### Axis 4: past the ZIP64 promotion boundary

A single stored member of 4,362,076,160 bytes (4.06 GiB), generated from a
deterministic incompressible stream, in a real on-disk archive of 4,362,076,822
bytes, read through `FileReader` under explicit finite limits (not
`ArchiveLimits::UNBOUNDED`):

| member bytes | archive bytes | output bytes | **peak** | **allocations** |
| ---: | ---: | ---: | ---: | ---: |
| 4,362,076,160 | 4,362,076,822 | 4,362,076,822 | **1,916** | **15** |

**4.06 GiB of unchanged media reaches a sequential sink behind 1,916 retained
bytes and fifteen allocations** — a ratio of 2,276,658 output bytes to one
retained byte. The fixture's SHA-256 is
`3fc0338da7d31fc6fb9b50659c2764701edaf7504e1d70a0d6c3ffbbc8932c10`; it was
deleted after measurement and is reproducible from the retained probe source.

### Axis 5: end to end, through both production save paths

The four axes above measure the ZIP layer. This one drives the two production
OOXML save paths from `litchi-opc`, over fixtures built from
`test-data/ooxml/docx/drawing.docx` by adding one large binary media part and one
small one, then editing the **small** part and publishing to a sequential,
non-seek sink. Only the large media part's size varies.

- **A** — `OpcPackage::open` → `get_part_mut().set_blob()` →
  `PackageWriter::write_to_stream`.
- **B** — `SourceBackedPackage::from_path` → `write_part_overlay_to_stream`.

| media member | archive bytes | output bytes | **B peak** | B allocs | **A peak** | A allocs | A/B peak |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 KiB | 164,712 | 164,712 | **532,626** | 1,167 | 1,161,217 | 2,269 | 2.2x |
| 1 MiB | 1,148,053 | 1,148,053 | **532,626** | 1,167 | 3,127,598 | 2,389 | 5.9x |
| 4 MiB | 4,294,743 | 4,294,743 | **532,626** | 1,167 | 9,420,016 | 2,773 | 17.7x |
| 16 MiB | 16,881,495 | 16,881,495 | **532,626** | 1,167 | 34,589,680 | 4,309 | 64.9x |
| 64 MiB | 67,228,503 | 67,228,503 | **532,626** | 1,167 | 135,268,336 | 10,455 | **253.9x** |

**The source-backed save is flat to the byte — peak and allocation count both —
across a 1,024-fold growth in the unchanged media member.** 67 MB of output
leaves a 532,626-byte high-water mark, 126 output bytes per retained byte. The
DEFINITION OF DONE clause is satisfied end to end, not merely at the ZIP layer.

**The two paths produce byte-identical output.** An FNV-1a digest of the complete
output stream, folded in the sink, matches between A and B at every size:

| media member | output digest (FNV-1a 64) |
| ---: | --- |
| 64 KiB | `4778955ffbd36738` |
| 1 MiB | `c7b9d54f9027cede` |
| 4 MiB | `cc84c8b54dfb8b1f` |
| 16 MiB | `de13ce9c3bb4c8a6` |
| 64 MiB | `395c924099cdf4b8` |

That is the byte-identity result this investigation set out to produce, arrived at
from the other direction: rather than proving a new streaming path agrees with a
materializing one, it shows the two **existing** production paths already agree
byte for byte while differing 254-fold in peak retained memory. The difference is
not the ZIP passthrough, which both share; it is that Path A's `OpcPackage` owns
the source bytes and materializes every part blob, while Path B keeps a
positional source and materializes only the part it edits. Path A's allocation
count also grows with the media member (2,269 to 10,455) where Path B's does not
move at all.

## Validation preserved

Nothing was changed, so every check keeps its position and identity by
construction. The points worth recording are why the existing design is already
correct on the axes the plan intended to defend:

- **ZIP framing identity is not at risk, because framing is not rebuilt.** A
  copied member's local header, name, extra fields, payload, and data descriptor
  are reproduced by copying the source's `local_span` byte range verbatim. The
  central record is likewise copied (`PreparedCentral::Copy`), with only the
  local-header offset patched when members move. There is no second code path
  that could disagree with a first, which is the whole risk a differential test
  would have been written to catch.
- **ZIP64 promotion, descriptors, and offsets are already exercised at the
  boundary.** `crates/soapberry-zip/tests/preservation_zip64_promotion.rs`'s
  `generated_store_offset_promotes_at_each_zip32_boundary` runs the copy path
  with a generated member landing at `u32::MAX - 1`, `u32::MAX`, and
  `u32::MAX + 1` over a multi-gigabyte sparse source.
- **The bounded-memory property is already pinned by an assertion, not merely
  true.** That same test asserts `sink.max_write <= COPY_CHUNK_SIZE.max(4096)`,
  that source requests stay within the locator probe bound, and — explicitly —
  `sink.retained_bytes() < 1024 * 1024`, under the comment *"region sink must not
  retain the multi-gigabyte source payload"*. A future change that rewrote
  `copy_range` into a single full-member read would fail it. This is the guard
  the plan proposed to add; it exists.
- **Sequential, non-seek output is what was measured.** Every figure above was
  produced against a sink that implements only `Write`, accepts each buffer
  whole, and discards it. Both production paths additionally wrap the sink in a
  `Chunked` adapter that caps each `write` at 64 KiB.
- **Short reads and partial-sink failure at a chunk boundary already have a
  test**: `copy_all_handles_short_reads_and_partial_sink_at_chunk_boundary`
  (`preserve.rs:3441`).
- **No `unsafe` was added to the repository.** The counting allocator lives only
  in the scratchpad probe, which is not part of the build; `deny(unsafe_code)` in
  `soapberry-zip` is untouched. Limits stayed finite and enforced throughout,
  including in the 4 GiB run.

## Correctness evidence

**No test fails against the pre-change code, because there is no change.** That
is the finding, not an omission: the property a new differential test would have
proven is already asserted by
`generated_store_offset_promotes_at_each_zip32_boundary`, and the property a new
streaming passthrough would have provided is already provided by `copy_range`.

Adding a test here was considered and rejected. The obvious candidate — an
allocator-instrumented assertion that peak does not scale with member size —
cannot be written as an in-crate test without installing a process-global
allocator, and the mechanism-level version of it (bounded writes, bounded source
requests, bounded sink retention at ZIP64 scale) is exactly what the existing
test already asserts. A second test asserting the same invariant with weaker
instrumentation would be filler.

Gates run, all against a detached worktree of `32d25e088` with an isolated
`CARGO_TARGET_DIR` so no concurrently edited tree could affect them:

- `cargo test -p soapberry-zip` — **570 tests across 11 binaries, zero
  failures**, 2 ignored. That figure matches the one change
  [0573](0573-zip-single-local-header-read.md) reported for the same crate.
- `cargo test -p soapberry-zip --test preservation_zip64_promotion` — 3 passed,
  0 failed, including `generated_store_offset_promotes_at_each_zip32_boundary`,
  the test this record leans on. Its bounded-retention and bounded-write
  assertions are live and passing, not merely present.
- `tools/check_perf_claims.py --mode structural` — reports the same pre-existing
  `claim-0251-xlsx-xml-borrowed` strict-evidence message at pristine `32d25e088`
  as it does with this record applied. This record adds no registry entry,
  because `performance_claim: none` records such as 0570 and 0573 carry none.

The two gates that are red at `32d25e088` — three facade `document::doc` tests
and `tools/check_example_targets.py` on three duplicate iWork example targets —
are unrelated and untouched. No workspace source file was modified, so no
formatting or lint gate could change state.

## Limitations

No timing, cold-cache, physical-device, or cross-platform result is claimed. The
figures are callback-ordered heap accounting from a single-threaded probe on one
machine over warm fixtures; they exclude allocator-internal fragmentation, RSS,
and the page cache holding the source file.

**Two genuine materializations were found, and neither is the clause's subject.
Both are left in place.**

**The precompressed token retains twice the member.**
`IndexedArchive::capture_precompressed` (`office.rs:4069`) `try_reserve_exact`s
the full `compressed_capacity`, and `read_entry_precompressed_and_decoded_with_progress`
additionally collects the whole decoded payload; `VerifiedPrecompressedEntry`
then holds an `Arc<Vec<u8>>` of the compressed bytes. Scenario B measures the
result at `2 × member`, up to 134,232,941 bytes for a 64 MiB member. Its sole
production caller is `litchi-opc/src/source_backed.rs:7871`, which republishes a
part copied **from a different package** as a topology addition. That is
cross-document copy, CRUD scenario 9, not "unchanged members flow from source to
output"; the member is by definition not unchanged in the destination. The cost
is also not hidden: `source_backed.rs` reserves it explicitly and twice through
`reserve_topology_memory`, plus a fixed term, before capture begins, and the
comment there states the `C + 2*name + 4096` bound. Making it streaming would
mean dismantling the verification model in which the decoded bytes are compared
against the capture before the compressed bytes are allowed to reach the
preservation writer — the token is a security artifact, documented as "the only
way this compressed payload can reach the preservation writer" — and that is a
larger design question than a buffer size. It is recorded here as a candidate,
not attempted. `crates/litchi-opc/src/` is also being worked in concurrently by
another agent.

**Path A holds its whole source in memory.** `pkgwriter.rs:197` opens the
preservation source with `ZipArchive::from_slice(source)`, so a
`PackageWriter::write` save has the entire source archive resident as `&[u8]`
before the copy loop starts. Axis 5 prices it: **135,268,336 peak bytes against
Path B's 532,626 for byte-identical output**, and Path A's peak tracks roughly
twice the archive size because `OpcPackage` additionally materializes every part
blob. The copy loop underneath is the same bounded one in both paths and still
copies no member twice; the residency is `OpcPackage`'s eager ownership model,
not the ZIP writer's. Fixing it means making the eager package lazy or routing
more callers to the source-backed door, both of which live in
`crates/litchi-opc/src/` — currently being worked in by another agent — and
neither of which is a ZIP-layer change. It is recorded as the largest measured
opportunity this investigation found, and deliberately not attempted here.

Finally, axes 1 to 4 measure the ZIP layer directly and axis 5 measures the two
`litchi-opc` save doors. No format facade above `litchi-opc` —
`litchi-docx`, `litchi-xlsx`, `litchi-pptx` — was instrumented; the production
call chain from those crates down into the measured doors is traced above by file
and line, but not executed. Axis 5's fixtures are built by adding binary parts to
one real DOCX, so they exercise a media-heavy shape rather than a
many-small-parts one; axis 2 is the evidence for the latter.
