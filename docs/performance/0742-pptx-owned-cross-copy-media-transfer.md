# 0742 — the owned PPTX cross-copy frames copied images from their verified source-compressed bytes: media-rich lifecycle p50 410 → 183 ms

Status: retained, implemented in `litchi-opc` and `litchi-pptx`, with two
additions to `soapberry-zip`.
`performance_claim: none` — no claim-registry entry; the paired medians and
counts below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `009d515bef`; production commits `317920af5c`, `52db88c24c` and
`b2132486af` (after the first review), `172501ac89` and `d2b2aa3d75` (after
the second) on `perf/0742-pptx-owned-cross-copy-media-transfer`. Evidence
packet: [`results/change-0742/`](results/change-0742/README.md).

## Result

On the generated media-rich pair (8 × 2 MiB incompressible PNGs per deck,
16.8 MB archives), `pptx_cross_copy_media_rich_lifecycle` falls from a median
process p50 of **410.081 ms to 183.024 ms**; the median of eight paired
after/before ratios is **0.4439** (bootstrap 95% [0.4406, 0.4585]). Both legs
were built with the identical command, from the base and from `d2b2aa3d75`.
Planning falls from 292.6 ms to 65.7 ms; commit and publication are
unchanged. The plain control, whose closure has no image, publishes the same
bytes as the base and stays within 1% (lifecycle ratio 1.0070 [0.9985,
1.0189], non-lifecycle 1.0097 [0.9902, 1.0228]). The owned and source-backed
routes publish byte-identical image members.

Which copied images transfer is decided from the bytes the source publishes.
Application depends only on the recorded revisions and the two packages'
bytes, so a redo after an undo, or a copy into any destination with the
recorded revisions, publishes the same bytes as the first copy. The second
review's fixes that make this so did not slow the measured path: the matrix
of `b2132486af` measured 0.4609 on the same case.

## What changed

An owned cross-presentation slide copy builds a candidate package, serializes
it, and reopens it. Every copied part became a new archive member, and the
targeted writer deflated each one again from its decoded bytes. Change 0740
located that Deflate at 78.8–79.5% of strictly attributed planning cycles on
the media-rich corpus. The copied images' compressed bytes already exist in
the source archive, and the source-backed route already transfers them
(`AuthorizedPrecompressedPart`). The owned route now does the same.

`crates/soapberry-zip`:

- `IndexedArchive::precompressed_layout_provable` runs, without reading a
  payload byte, the checks a verified precompressed capture runs first:
  - Store or Deflate;
  - a valid strict stream target (resolved ZIP64 fields, one disk, not
    encrypted);
  - consistent Store sizes;
  - the strict local layout: local header against central record, a bounded
    span, the data descriptor.

  A disproof is `Ok(false)`, and the proof is memoized exactly as a capture's
  is.
- `Error::is_content_fault` (new) names the errors that are properties of the
  archive's own bytes: malformed, inconsistent or unsupported records and
  payloads. Only these disprove a layout. Allocation, limits, I/O,
  cancellation and any other kind are returned as errors.

`crates/litchi-opc`, `package/compressed_transfer.rs` (new) keeps two
questions apart:

- **Eligibility.** `OpcPackage::compressed_transfer_size` never decodes a
  payload. It returns the member's declared compressed size when all of these
  hold:
  - the package retains the owned archive it was opened from, with a source
    member for the part;
  - the part still holds the payload allocation it was opened with (pointer
    identity with the preservation provenance);
  - its content type is unchanged, it has no relationships, and it is not XML
    by name or content type;
  - the package has no signature infrastructure;
  - the member's layout is provable from its headers;
  - its compressed size is at most its decoded size, plus 5 bytes per started
    4 KiB, plus 64 bytes (see *The size guard*).
- **Verification.** `OpcPackage::authorize_compressed_transfer` decodes the
  part through its own route if it is still deferred, then re-checks the
  member's declared sizes against the `ReadLimits` the archive was admitted
  under. It captures the member's exact compressed span through
  soapberry-zip's `IndexedArchive::read_entry_precompressed_with_progress`,
  which decodes the capture, compares every decoded byte with the payload and
  records the actual CRC.
  - When the member's own bytes disprove the capture, it returns `Ok(None)`.
    One example is a Deflate stream followed by bytes it does not consume,
    which the ordinary reader tolerates.
  - Every other failure is a typed `OpcError`.

Other `litchi-opc` changes:

- A deferred part uses the index its own decode builds. An eagerly
  materialized owned package builds one transfer index per open, shared by
  clones. The index cell stores only a built index or a content fault; an
  allocation or limit failure is not stored, so a later call or a clone tries
  again.
- `OpcPackage::source_read_limits` (new) reports the limits the owned archive
  was admitted under.
- `payload.rs`: `PartPayload::Transferred` holds the decoded allocation and the
  verified capture as one value. `BlobPart::with_compressed_transfer`
  (`part.rs`) builds a part over it. `set_blob`, `set_blob_shared` and
  `set_content_type` replace or demote the payload, so no payload or
  content-type change keeps the capture.
- `pkgwriter.rs`: `part_entry` frames a transferred payload with
  `RegeneratedEntry::new_precompressed_shared`. That means fresh known-size
  headers, no data descriptor, zero timestamps, no source extras and the
  source method. It does so for appended and regenerated members, and only
  when the part's visible allocation is the one the capture was verified
  against. The full (non-preserving) writer still re-encodes the decoded
  bytes.

`crates/litchi-pptx` (`opened/cross_copy_plan.rs`):

- `classify_media` applies the format rule: a relationship-free, non-XML
  `image/*` part. It then asks the source's *owned view* whether each such
  part is eligible (`owned_view`). The owned view is the source itself when it
  is an unmodified owned source, and otherwise the reopen of its
  serialization.
- `preflight_parts` then checks the candidate estimate, including the eligible
  members' declared compressed sizes, against `max_patch_bytes`. Only after
  that does `MediaTransfers::capture` take any capture.
- A member whose capture is disproved by its own bytes is recompressed.
- The copied-media encoding (`CopiedMedia::{Recompressed, SourceCompressed}`)
  follows from the members that transfer. It is recorded in the plan and in
  the durable patch. Every later proof (the fresh re-plan in
  `apply_cross_slide_copy_plan`, and both routes of
  `apply_cross_slide_copy_patch`) rebuilds under the recorded encoding.
- A transferring copy's candidate is built from the destination's owned view.
  Into a destination that is not an unmodified owned source, it publishes the
  candidate reopened from its archive and carries the destination's save
  preferences onto it. A recompressing copy keeps the clone-and-apply
  publication it had. See *Deciding from bytes*.
- The retained candidate archive of change 0656 records the plan indexes of
  its transferred members. It is reused only when a fresh classification
  yields the same list, and a release build that reuses it skips the captures.
  Otherwise the application serializes its own candidate, which then fails the
  plan comparison exactly as it would without retention.
- The durable format moves to `LPCP0004`: one byte after the presentation
  relationship ID records the encoding. `LPCP0002` and `LPCP0003` are refused
  by name before any header field is read, and an unknown encoding byte is
  `Error::Invalid`.

No `unsafe`, no new dependency, no new thread, clock or I/O; litchi-pptx still
has no archive dependency.

## Authority

The owner decisions of 2026-09-16
([0652](0652-owner-decisions-for-the-third-wave.md)) apply as follows:

- **Trade-off 1** authorizes the format bump and the published-byte changes
  below.
- **Trade-off 2** decided three points:
  - The capture is verified by decode-and-compare, although pointer identity
    already proves the payload.
  - A member whose own bytes disprove its capture, or that fails the size
    guard, keeps the re-deflating route. The ordinary reader accepts such
    members, and the base copied them. Every resource failure stays a typed
    error, never a quiet change of route.
  - A transferring copy into a destination with caller-defined parts publishes
    the reopened candidate rather than keeping source captures in the caller's
    package.
- **Trade-off 3** is the scope. The common benign path (untouched images
  copied between unmodified owned packages) gets the transfer at no extra
  cost. Modified packages pay one bounded serialization and reopen.

Other authority:

- Decision 4's precedent (the `LPRM`/`LPCP` bump with a typed refusal by name,
  implemented in [0655](0655-pptx-memoized-revision-proof.md)) is followed for
  `LPCP0003 → LPCP0004`. `LPRM` is unaffected.
- **The dispatch** fixed the transfer rule, the typed-refusal rule, the
  no-unbounded-foreign-retention rule and the proofs that stay mandatory.
- **The first review** led to the header-only layout proof. A member the
  ordinary reader accepts, but whose local header disagrees with its central
  record, keeps today's route.
- **The second review** (verdict: merge after fixes) required three things:
  - application depends only on recorded revisions and bytes;
  - a capture failure caused by the member's bytes selects recompression at
    planning;
  - a header-only size guard, with the captures charged against
    `max_patch_bytes`.

  A measured unit replaces its prescribed guard unit (see *The size guard*).
- **ADR 0030** (lazy decode): eligibility decodes nothing. Authorization
  decodes through the package's own route, so its refusal is the one
  `get_part` reports.
- **ADR 0005:**
  - no capture outlives the candidate build;
  - the retained candidate's new field is a short index list, inside the
    budget change 0656 declared;
  - an owned view lives only for one planning or application call.
- **ADR 0006:** untouched destination members are still copied verbatim. A
  transferred member is framed only from bytes the ZIP reader decoded,
  compared and checksummed.
- **ADR 0003:** plans and patches stay source-checked, reversible and
  deterministic, now also across undo/redo and byte-identical packages.
- **ADRs 0010/0011:** the ZIP token stays inside litchi-opc.

### Deciding from bytes

The branch's first version decided transfer from package state that no
recorded revision captures:

- a copy transferred only when the destination was an unmodified owned source;
- a source image transferred only while it still held the allocation it was
  opened with.

So two packages with identical semantic and physical revisions could accept
or refuse the same patch. The library's own undo publishes a restored clone,
which is a modified package, so the redo of a transferring copy failed. The
second review reproduced both.

Every decision is now a function of bytes and recorded revisions:

- **Source.** Eligibility and captures are asked of the source's owned view.
  - In an unmodified owned source, every part still holds its opened
    allocation, because every mutating entry point revokes that status, even
    for a no-op. Pointer identity there is therefore a property of the
    archive.
  - Any other source is serialized and reopened, and every part of the reopen
    is untouched.
  - So a source whose image was replaced, even with equal bytes, lends the
    member it now publishes.
- **Destination.** A transferring copy's candidate is built from the
  destination's owned view, so the candidate is a function of the
  destination's bytes.
- **Application** first checks the recorded semantic and physical revisions
  of both packages, then rebuilds under the recorded encoding. Any source and
  destination with those revisions therefore rebuild the same candidate and
  publish the same bytes.

Tests cover undo then redo (by patch and by plan), a byte-identical modified
destination, a modified source, and a re-provenanced source.

Cost and bounds of the normalization:

- It serializes a package through the same bounded writer whose hash is that
  package's physical revision, so it is charged against `max_patch_bytes`. No
  package whose revision could be taken is refused for size.
- The reopen re-admits the bytes under the package's own admitted read limits.
- It runs only in two cases:
  - a source that is not an unmodified owned source and has at least one
    copied image of an eligible format;
  - a transferring copy into a destination that is not an unmodified owned
    source.

  The unmodified owned packages of the measured cases pay nothing.
- While it runs, the call holds one bounded serialization per normalized
  package, besides the candidate's.

**What changed for such destinations.** For a transferring copy into a
destination that is not an unmodified owned source:

- The base published a clone of the live destination with the patch's decoded
  resources applied. That keeps caller-defined `Part` implementations and save
  preferences.
- That clone cannot reproduce a candidate that framed source-compressed bytes,
  because the targeted writer deflates those resources again.
- Making it carry the captures would keep copies of source bytes inside a
  caller's package for its whole lifetime, against ADR 0005's retained-state
  rule (the 0656 amendment).
- So the copy now publishes the reopened candidate:
  - save preferences are carried, since they do not change what a package
    serializes to;
  - caller-defined `Part` implementations become built-in parts with the same
    bytes.
- A recompressing copy keeps the clone-and-apply publication.

The inverse still publishes the restored clone, as the base did. The reviewer
suggested reopening it, but that would:

- move the serialization and reopen, which a transferring redo pays, onto
  every undo;
- drop caller-defined parts a destination may hold.

The apply-side rule already makes the redo byte-identical.

### The size guard

A member is eligible only when its declared compressed size is at most its
decoded size, plus 5 bytes per started 4 KiB, plus 64 bytes. The check reads
the central record only. A member that fails it is recompressed,
deterministically.

The review prescribed one 5-byte stored-block header per 65,535 bytes, which
allows 229 bytes on a 2 MiB member. But zlib at its default memory level frames
incompressible data in 16 KiB stored blocks. So does zlib-rs in soapberry-zip's
writer, which produced the harness corpus. The packet's
`size-guard/zlib_framing.py` measures zlib 1.3.1 on 2 MiB of seeded random
bytes:

| encoder on 2 MiB of incompressible bytes | framing bytes | verdict |
|---|---:|---|
| zlib, memLevel 9, levels 1/6/9 | 320 | transferred |
| zlib, memLevel 8 (the default), levels 1/6/9 | 640 | transferred |
| zlib, memLevel 4 | 10,240–10,245 | recompressed |
| zlib, memLevel 2 | 41,100–41,110 | recompressed |
| zlib, memLevel 1 | 82,523–82,555 | recompressed |
| allowed by the prescribed 65,535-byte unit | 229 | — |
| allowed by the 4 KiB unit | 2,624 | — |

- **Why the prescribed unit fails.** Under it, every zlib-default member above
  about 273 KiB would have been recompressed. That includes every image of the
  harness corpus and most photographs.
- **What the 4 KiB unit admits.** It admits the default encoders with four
  times their framing. A litchi-opc test checks the default encoder's framing
  on a 1 MiB member and that the member transfers.
- **What it still bounds.** A transferred member stays within about 0.12% plus
  64 bytes of its decoded size. In the reviewer's source, 200,000 empty stored
  blocks (1 MB) sit in front of the flat image; that image is now published
  recompressed, not padded. A PPTX test checks this.
- **Low-memory encoders.** Their members are recompressed, which also makes
  them smaller.

A member its producer Stored is transferred Stored, as the source-backed route
does and as the producer chose. For a compressible image, that output is larger
than recompression would give: 113,791 against 81,255 bytes (+40%) in the
reviewer's example. The review noted that this runs against decision 10's "make
files smaller where possible". The coordinator kept the transfer for Stored
members for these reasons:

- it matches the source-backed route and the producer's choice;
- the guard cannot tell compressible from incompressible Stored bytes without
  the decode-and-deflate this change removes;
- images are almost always in compressed formats, where Store costs nothing;
- the published member is the source's own representation in fresh framing.

## Breaking changes (relative to the base)

| item | before | after |
| --- | --- | --- |
| `CrossSlideCopyPatch::to_bytes` | `LPCP0003` | `LPCP0004`, one extra header byte (the copied-media encoding) |
| `CrossSlideCopyPatch::from_bytes*` on `LPCP0003` | parsed | `Error::DurablePatchRevisionFormat { found: CrossSlideCopyV3, expected: CrossSlideCopyV4 }` before any header field is read. `LPCP0002` now reports `expected: CrossSlideCopyV4`. An unknown encoding byte is `Error::Invalid` |
| published bytes of an owned copy with eligible images | copied images deflated again, with a data descriptor | the source member's compressed bytes and method, in fresh sized framing; target physical revisions change accordingly. Copies without eligible images are byte-identical to the base |
| a copied image its producer Stored | deflated | Stored, as in the source; larger for compressible bytes (+40% in the reviewer's example) |
| a transferring copy (plan or forward patch) into a destination that is not an unmodified owned source | a clone of the destination with the patch applied; caller-defined parts kept | the reopened candidate: the same published bytes, save preferences carried, caller-defined `Part` implementations replaced by built-in parts with the same bytes |
| a source that is not an unmodified owned source and has a copied image of an eligible format, or such a destination of a transferring copy | not re-read | serialized (charged against `max_patch_bytes`, as its physical revision already is) and reopened under its own read limits; a read-limit refusal of that reopen is returned |
| `CrossSlideCopyPlan` / `CrossSlideCopyPatch` `Debug`, `PartialEq` | — | carry the encoding; the plan's candidate slot reports `transferred_members` |

The branch's first commits also refused two things: a transferring copy into
a modified destination, and a source whose image had been re-provenanced. The
second review's fixes removed both refusals, and neither exists relative to
the base.

Additive:

- `soapberry_zip::Error::is_content_fault`;
- `soapberry_zip::office::IndexedArchive::precompressed_layout_provable`;
- `litchi_opc::CompressedPartTransfer`;
- `litchi_opc::OpcPackage::compressed_transfer_size`;
- `litchi_opc::OpcPackage::authorize_compressed_transfer`;
- `litchi_opc::OpcPackage::source_read_limits`;
- `litchi_opc::BlobPart::with_compressed_transfer`;
- `litchi_pptx::DurablePatchFormat::CrossSlideCopyV4` (the enum is
  `#[non_exhaustive]`);
- `CrossSlideCopyPlan::transfers_source_compressed_media`;
- `CrossSlideCopyPatch::transfers_source_compressed_media`.

## Evidence that motivated it

- [0739](0739-pptx-cross-copy-current-baseline.md): media-rich owned lifecycle
  p50 403.95 ms, with planning 71.06% and application 22%, and 277.8 MB
  allocated per lifecycle.
- [0740](0740-pptx-cross-copy-native-profile.md): generated-entry Deflate
  under `build_candidate` is 78.79–79.54% of strict planning cycles.
- The coordinator's sweep on the base binary: owned media-rich lifecycle
  428.5 ms, against 19.5 ms for the source-backed route on the same corpus.

## Measurement

Method:

- **Harness:** `tools/perf-baseline`, unchanged by this change.
- **Builds:** both legs used the identical command, `cargo build --release
  --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin
  litchi-perf-baseline`, plus `--features allocator-metrics --bin
  litchi-perf-baseline-alloc` for the allocator lane. The before leg was built
  from the read-only base checkout, the after leg from `d2b2aa3d75`.
- **Runs:** every process was pinned with `taskset -c 4`. Each case ran four
  rounds of before, after, after, before (16 processes, eight per arm).
  Media-rich cases took 20 samples after 3 warmups; the other cases took 40
  after 3. The allocator lane ran the two lifecycle cases, 3 samples after 1
  warmup.
- **Statistics:** the median of process p50s per arm, with its min–max, and
  the median of the eight paired ratios ((s0, s1) and (s3, s2) in each round),
  with a percentile bootstrap (10,000 draws, seed 742).
- **No run was discarded or selectively repeated.** Three earlier complete
  matrices were superseded, by code changes and by the build-matching rule;
  their summaries are kept (see below).

| binary | SHA-256 |
| --- | --- |
| before, native | `b375de3695473a3f0bf870ae25d9c2e7a8f9b8124aa43370698f6557a75f3385` |
| before, allocator | `75c5ee1707d0f45ee2366af3e77f26b9748cb34474e14550cea44e566eeddab2` |
| after, native (`d2b2aa3d75`) | `9d714e54b1f34a13d1778ab34a0f7ae41711ffab41760f7864e796db12521d2e` |
| after, allocator (`d2b2aa3d75`) | `d8ff3f80738cf97aa27ba0fc27d6e7a51d7bc2a2f2b207df2ab90d4a34391434` |

The before binaries were rebuilt for this matrix, with the same command and
from the same tree as the `b2132486af` matrix's before binaries. Their hashes
differ from that build's (`0f20b4d0…`/`3e304e2d…`); the cause was not
investigated. Host: AMD EPYC 9R45, Linux 7.0.0-1012-aws, shared with other
agents' builds during the run. Rust 1.95.0 from `rust-toolchain.toml`,
release, `--locked --offline`.

### Native latency

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 410.081 [405.783–416.086] | 183.024 [176.960–194.569] | 0.4439 [0.4406, 0.4585] |
| `pptx_cross_copy_media_rich` | 386.716 [383.267–397.587] | 158.999 [153.792–180.286] | 0.4125 [0.3985, 0.4413] |
| `pptx_cross_copy_plain_lifecycle` (control) | 8.360 [8.225–8.485] | 8.401 [8.372–8.576] | 1.0070 [0.9985, 1.0189] |
| `pptx_cross_copy_plain` | 7.091 [6.962–7.245] | 7.147 [7.054–7.226] | 1.0097 [0.9902, 1.0228] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` (control) | 17.022 [12.103–17.492] | 16.678 [12.028–17.084] | 0.9830 [0.7077, 1.0106] |

Median process p95 and mean:

| case | p95 before → after (ms) | mean before → after (ms) |
|---|---:|---:|
| media-rich lifecycle | 411.407 → 184.384 | 409.923 → 183.405 |
| media-rich | 392.249 → 160.486 | 386.846 → 159.288 |

Phases, median of process medians (ms):

| case | plan | commit | publication |
|---|---:|---:|---:|
| media-rich lifecycle | 292.579 → 65.673 | 89.158 → 89.123 | 6.508 → 6.298 |
| media-rich | 291.186 → 64.028 | 88.901 → 89.157 | 6.392 → 6.416 |
| plain lifecycle | 3.304 → 3.323 | 3.805 → 3.812 | 0.001 → 0.001 |
| plain | 3.294 → 3.313 | 3.786 → 3.815 | 0.001 → 0.001 |

The superseded matrices, each also 80 native and 32 allocator processes:

| matrix | media-rich lifecycle | media-rich | plain lifecycle | plain | source-backed control |
|---|---:|---:|---:|---:|---:|
| `b2132486af`, identical-command builds | 405.441 → 185.703 ms (ratio 0.4609) | 0.4292 | 1.0097 | 1.0215 | 0.9973 |
| `317920af5c`, prebuilt base binary | 408.126 → 194.768 ms (0.4742) | 0.4276 | 1.0006 | 0.9999 | — |
| `52db88c24c`, prebuilt base binary | 411.384 → 186.599 ms (0.4522) | 0.4231 | 0.9924 | 0.9968 | — |

The `b2132486af` matrix is the commit the second review examined. The fixes
moved classification and the captures ahead of the candidate build; its plan
phase measured 68.7 ms against 65.7 ms here.

### Regression flags

No measured case regresses by more than 5%.

- **Plain copies.** The plain cases are within 1% (1.0070 and 1.0097), and both
  intervals include 1.0. The `b2132486af` matrix measured the non-lifecycle
  plain case at +2.15%, and the two before it at −0.0% and −0.3%. A plain copy
  has no image, so the only new work on its path is O(1): one empty
  classification, one encoding byte, one empty-list comparison, and one small
  cell per owned open. No direction is claimed.
- **Source-backed control.** Its paired ratios include 0.696, 0.714, 0.708 and
  1.386. Its path is unchanged, with identical output. Each outlier pairs a
  process whose large buffers happened to fault fresh with one whose buffers
  did not (next paragraph). Its median ratio is 0.9830.

The media-rich spread in both arms follows first-touch page faults, which the
harness records for every lifecycle sample (`faults.py` → `faults.json`).
One freshly mapped 33.6 MB buffer costs about 8,204 faults and 4–7 ms here:

| arm | fresh 33.6 MB mappings per lifecycle | samples | lifecycle ms | plan ms | commit ms |
|---|---:|---:|---:|---:|---:|
| before | 1 | 100 | 408.647 | 290.372 | 89.140 |
| before | 2 | 60 | 413.194 | 295.566 | 88.805 |
| after | 0 | 20 | 176.960 | 63.973 | 89.112 |
| after | 1 | 70 | 181.145 | 63.673 | 89.052 |
| after | 2 | 60 | 187.800 | 69.946 | 94.407 |
| after | 3 | 10 | 194.945 | 70.260 | 95.968 |

- At equal fault counts the commit phase is unchanged: 89.14 against 89.05 ms
  with one mapping.
- The lifecycle ratio at one mapping is 0.443.
- How many large buffers fault fresh is per-process allocator state. It is
  visible in the unchanged source-backed control too: processes that fault
  nothing run 12.0–12.1 ms, the rest about 17 ms.
- The mechanism that selects a process's mode is not established.

### Allocations (allocator lane, lifecycle region, median of process medians)

| field | media-rich before → after | plain before → after |
|---|---:|---:|
| allocation calls | 56,356 → 56,383 (+27) | 46,613 → 46,618 (+5) |
| deallocation calls | 46,068 → 46,086 | 38,273 → 38,275 |
| reallocation calls | 6,271 → 6,272 | 5,298 → 5,298 |
| allocated bytes | 277,809,048 → 272,763,540 (−1.82%) | 16,613,707 → 16,615,187 (+1,480) |
| region peak live bytes | 305,068,449 → 305,071,038 | 1,313,024 → 1,313,469 |

- **Media-rich.** The eight captures (16.8 MB) replace the eight generated
  Deflate buffers and compressor states. Peak live bytes are unchanged within
  2.6 KB.
- **Plain.** The five extra allocation calls are the transfer-index cells of
  the lifecycle's four owned opens and one classification vector. The 1,480
  extra bytes are not attributed further; each `OpcPackage` value is also 192
  bytes larger, for `source_limits` and the cell pointer.
- **History.** `b2132486af` made the cell owned-source-only and boxed its
  index, after the second superseded matrix measured ten allocations and
  4.6 KB per plain lifecycle.

### Output bytes

- **Media-rich.** Output goes from 33,599,873 to 33,599,745 bytes: the eight
  16-byte data descriptors are gone. SHA-256 goes from `6a3536fc…` to
  `68be0ce5…`, with one digest per arm across all 24 media-rich processes of
  that arm. These are the same digests as in the `b2132486af` matrix, so the
  fixes do not change what the corpus publishes.
- **Plain.** Outputs are byte-identical to the base (`3e9ae280…`, 31,545
  bytes).
- **Source-backed control.** Its output is unchanged (`809c6172…`, 33,599,715
  bytes).

**Owned vs source-backed (informational).** The two routes are not
byte-identical on the harness corpus (33,599,745 vs 33,599,715 bytes). On
`b2132486af`, the packet's `route-compare` probe published one copy of a
harness-shaped pair through both routes:

- All 16 image members are byte-identical: the same names, compressed bytes
  and framing.
- The archives differ only because the owned route names the copied slide
  `slide3-copy1.xml`, where the source-backed route keeps `slide3.xml`.
- That naming difference also changes `[Content_Types].xml`, the presentation
  relationships and the member order.

## Where the remaining time goes, and the next opportunity

A frame-pointer build of `b2132486af` (SHA-256 `8b54a242…`, packet
`attribution/`) was profiled on the media-rich lifecycle: cycles,
`--call-graph fp`, 12 samples. Shares are of cycles within a phase, not wall
time. The second review's fixes move the captures ahead of the build but do
not add a pass over the payloads.

**Commit (about 89 ms)** is 80.0% SHA-256, in five passes of about 16% each:

- the live source and destination semantic re-fingerprints
  (`package_fingerprint`);
- their physical re-fingerprints (`to_stream` into a hash sink);
- the fallible copy and digest of the retained archive;
- the reopened candidate's semantic capture in `build_candidate`;
- the same candidate's capture again, in `validate_application_candidate`.

The eager reopen (inflate, CRC, donation compare) is 6.9%.

**Plan (about 69 ms in that build):**

| component | share |
|---|---:|
| serializing the candidate with its archive digest | 35.9% |
| the snapshots' physical revisions | 22.0% |
| the candidate's semantic capture | 21.8% |
| the eager reopen | 9.5% |
| this change's capture verification (decode, compare, CRC) | 4.2% |

The header-only layout proofs do not reach the listed paths.

**Next opportunity (not implemented here):** stop re-hashing bytes whose
identity is already proven.

- `validate_application_candidate` recaptures the candidate that
  `build_candidate` just captured.
- The candidate capture could consult the snapshots' part-digest memos: the
  reopened candidate's copied and untouched parts share donated allocations.
- The live re-fingerprints at the start of application could use the facade's
  memo (change 0655) and a per-package memo of the immutable retained archive's
  physical revision.
- Owned ingress that accepts a shared archive would remove the retained
  archive's copy. Its digest was deliberately kept by 0656.

Each needs its own ADR 0005 memo argument and freshness proof. Page-fault
sites of the same build are summarized in `attribution/faults-by-site.txt`:
the snapshot decodes, the candidate serialization buffer and the publication
sink.

## What is not claimed

- No claim-registry entry.
- No Office or other native-application validation of the published archives.
  Litchi reopens them and decodes every copied member.
- Only the two generated corpora and the unit fixtures were measured. No
  real-producer media: the size guard is calibrated on zlib-family encoders
  and generated data, and members from encoders with more framing keep the
  recompressing route.
- The normalization path is not timed: no harness case copies from or into a
  modified package.
- No cold-cache, remote-source, concurrency, RSS or instruction-count result.
- No statement that the lifecycle is faster by the same factor in every host
  state. The fault-mode spread above depends on the host and the allocator.
- The source-backed control's 0.9830 is not an improvement.
- The attribution shares are one diagnostic capture of a frame-pointer build of
  an earlier commit, not ordinary-release timings.

The broader non-iWork goal remains active.

## Verification

**First review, of `317920af5c`.** A separate reviewer found no correctness
bug on any publication route, and seven lesser findings. `52db88c24c` and
`b2132486af` act on them:

- the layout-disproof classification (its main risk);
- the pointer check in the writer;
- a direct test of the capture's own verification;
- the per-open transfer index;
- documentation corrections and two stronger tests.

The reviewer also noted that an `LPCP0003` patch could in principle be read
losslessly as the recompressed encoding. The dispatch's decision to refuse the
old format by name stands. Its note that `apply_cross_slide_copy_patch`'s
forward route maps a failed re-plan to the generic `UnsafeEdit` predates this
change, and that behavior is unchanged.

**Second review, of `ac92e4b3d9`.** The verdict was merge after fixes. Its
findings and the resolution in `172501ac89` and `d2b2aa3d75`:

- **Should-fix: redo after undo was refused.** Decisions now depend only on
  recorded revisions and bytes (*Deciding from bytes*).
- **Should-fix: a member the ordinary reader accepts could fail the whole
  copy.** A capture disproved by the member's bytes now recompresses that
  member, and the decision is made at planning.
- **Nit, made required: nothing bounded a transferred member's compressed
  size, and captures were allocated before `max_patch_bytes` was checked.**
  The size guard now bounds it, and preflight charges the captures first.
- **Nit: an inaccurate comment.** It is gone with the code it described.
- **Nit: two infallible clones.** The transferred-member list is now plan
  indexes, reserved fallibly, and it is moved into the retained candidate, not
  cloned.
- **Nit: a late refusal in `apply_plan`.** The refusal no longer exists.
- **Nit: an index-build failure was reported as an error and cached.**
  Byte-caused failures are now an ineligible verdict, and allocation or limit
  failures are no longer cached.
- **Missing tests.** Added, as listed below.

**New tests.**

`soapberry-zip`, 2 tests:

- the layout predicate classifies Store, sized and descriptor Deflate members
  as provable and a renamed local header as disproven, and a provable member
  captures;
- content faults are classified apart from allocation, limit, I/O,
  cancellation and buffer errors.

`litchi-opc`, 21 tests in `package/compressed_transfer/tests.rs`:

- **Exact span and fresh framing:**
  - Deflate with and without a descriptor;
  - a Store member with a source descriptor, timestamp and extras;
  - a ZIP64-framed member, republished with fresh ZIP32 framing.
- **Eager and deferred captures are equal.**
- **Eligibility never decodes.** Replaced-with-equal-bytes, retyped,
  relationship-bearing, XML, borrowed and authored parts are refused by type;
  signed packages are refused by policy.
- **Typed refusals:** a corrupt stream, a CRC mismatch and a data-descriptor
  mismatch are each refused with a stable typed error.
- **Header-only disproofs:**
  - a local/central name mismatch is classified ineligible without decoding,
    while the lenient reader still decodes it;
  - raw names that normalize to one member name are refused at open.
- **The capture's decode-and-compare** refuses bytes that are not the
  member's.
- **Read limits:** they are re-checked, and the admitted limits are reported
  while the archive is retained.
- **Mutation** discards the capture. A forwarding custom part publishes the
  bytes it shows.
- **Other writer paths:** regenerated members, and the full-writer fallback.
- **The second review's cases:**
  - bytes after the final block give a deterministic `Ok(None)`;
  - empty stored-block padding is ineligible (15 blocks admitted, 17 and
    20,000 not);
  - a default encoder's 1 MiB incompressible member stays eligible, although
    the prescribed unit would reject it.

`litchi-pptx`, 19 tests in `opened/cross_copy_plan/media_transfer_tests.rs`:

- **Span and framing** through the PPTX route, with renamed media.
- **Byte identity:** retained, released, durable and independently planned
  outputs are byte-identical, and the inverse is exact.
- **Source state:** a modified source lends the members it publishes. Stale
  and foreign sources are refused identically with and without retention; a
  re-provenanced source is accepted.
- **Format rule:** an SVG picture is still refused; plain copies record
  `Recompressed`; the transferable-media rule includes its XML guard.
- **Recompressed members:** a disagreeing local header, bytes after the final
  Deflate block (the reviewer's counterexample), and 200,000 empty stored
  blocks (the reviewer's padding case).
- **Undo and redo:** the inverse applies to a modified destination, and a redo
  after undo (by patch and by plan) is byte-identical to the first copy (the
  reviewer's other counterexample). A byte-identical modified destination
  accepts the same copy.
- **Durable format:**
  - a tampered encoding byte: `2` is `Error::Invalid`, and `1 → 0` is
    `UnsafeEdit` in both directions with the destination untouched;
  - genuine `LPCP0003` patches from the base tree are refused by name in both
    directions, with their input hashes re-derived and the changed target
    physical revision shown.
- **Budget:** captures are charged against `max_patch_bytes`. One byte below
  the charge is refused by name, although the copy fits without the captures.
- **Chaining:** a published candidate lends its media to a further copy,
  which frames the original source's bytes, and the durable patch round-trips.
- **Destination state:** a transferring copy keeps save options but not
  caller part types, with output equal to the same copy into the destination's
  bytes.
- **Stored images:** a Stored compressible image is transferred Stored.

The existing superseded-magic test now covers `LPCP0002` and `LPCP0003`
against `LPCP0004`. The legacy fixtures are
`test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-{forward,inverse}.patch`
(5,685 bytes each), written by the packet's `legacy-fixture/` generator
against the base tree. soapberry-zip's existing capture tests cover a corrupt
compressed payload, a truncated Deflate stream, an expected-byte mismatch and
ZIP64 members.

**Gates on `d2b2aa3d75`**, run with a fresh target directory (commands, exit
codes and counts in [`gates.txt`](results/change-0742/gates.txt)):

- `cargo fmt --all --check`;
- `cargo check --all-targets` of `soapberry-zip`, `litchi-sign`, `litchi-opc`,
  `litchi-pptx` and their in-scope dependents;
- Clippy `-D warnings` on the three libraries, and on `soapberry-zip` and
  `litchi-opc` all targets. `litchi-pptx --all-targets` fails only on three
  `err_expect` lints at `opened/tests.rs:464/538/557`, present unchanged on the
  base.
- Tests of the three crates (2,361 passed), of their seven in-scope dependents
  (5,205) and of the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (382);
- rustdoc `-D warnings` for the three crates;
- the crate-boundary, non-iWork and structural claim gates.

`soapberry-zip`'s ODF and iWork dependents were not rebuilt; its change is two
additive methods.

## Cleanup

Binary identities are recorded above and in
[`cleanup.json`](results/change-0742/cleanup.json), taken before removal.

- **Removed:**
  - the target directories `targets/0742` (fresh for this round),
    `targets/0742-before` and `targets/0742-after`;
  - the scratch directory's contents;
  - the raw reports of the `b2132486af` matrix.
- **The first round's cleanup** is in `cleanup-b2132486af.json`.
- **Kept:** the worktree and branch, and the coordinator's shared base build.
