# 0742 — the owned PPTX cross-copy frames copied images from their verified source-compressed bytes: media-rich lifecycle p50 405 → 186 ms

Status: retained, implemented in `litchi-opc` and `litchi-pptx`, with one new
predicate in `soapberry-zip`.
`performance_claim: none` — no claim-registry entry; the paired medians and
counts below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `009d515bef`; production commits `317920af5c`, `52db88c24c` (review
follow-ups) and `b2132486af` (transfer-index allocation) on
`perf/0742-pptx-owned-cross-copy-media-transfer`. Evidence packet:
[`results/change-0742/`](results/change-0742/README.md).

## Result

On the generated media-rich pair (8 × 2 MiB incompressible PNGs per deck,
16.8 MB archives), `pptx_cross_copy_media_rich_lifecycle` falls from a median
process p50 of **405.441 ms to 185.703 ms**; the median of eight paired
after/before ratios is **0.4609** (bootstrap 95% [0.4460, 0.4739]), with both
legs built by the identical command. Planning falls from 288.6 ms to 68.7 ms;
commit and publication are unchanged. The plain control, whose closure has no
image, publishes the same bytes as the base and stays within 1% (ratio 1.0097
[0.9957, 1.0172]); the non-lifecycle plain case measured +2.15% in this
matrix and −0.0% and −0.3% in the two earlier ones (see *Regression flags*).
The owned and source-backed routes now publish byte-identical image members;
their archives still differ in how the copied slide is named.

## What changed

An owned cross-presentation slide copy builds a candidate package, serializes
it, and reopens it; every copied part became a new archive member, and the
targeted writer deflated each one again from its decoded bytes. Change 0740
located that Deflate at 78.8–79.5% of strictly attributed planning cycles on
the media-rich corpus. The copied images' compressed bytes already exist in
the source archive, and the source-backed route already transfers them
(`AuthorizedPrecompressedPart`); the owned route now does the same.

`crates/soapberry-zip`: `IndexedArchive::precompressed_layout_provable` runs
the checks a verified precompressed capture runs before it reads a byte (Store
or Deflate, a valid strict stream target — resolved ZIP64 fields, one disk,
not encrypted — consistent Store sizes, and the strict local layout: local
header against central record, bounded span, data descriptor) and reports a
disproof as `Ok(false)`. Only allocation, transport and cancellation failures
are errors. The proof is memoized exactly as a capture's is.

`crates/litchi-opc`:

- `package/compressed_transfer.rs` (new): `OpcPackage::compressed_transfer_eligible`
  never decodes a payload: the package retains the owned archive it was
  opened from, with a source member for the part; the part still holds the
  payload allocation it was opened with (pointer identity with the
  preservation provenance — a payload replaced even with equal bytes is not
  eligible); its content type is unchanged; it has no relationships; it is not
  XML by name or content type; the package has no signature infrastructure;
  and the member's layout is provable from its headers alone.
  `OpcPackage::authorize_compressed_transfer` decodes the part through its own
  route if it is still deferred, then issues a `CompressedPartTransfer` from
  soapberry-zip's `IndexedArchive::read_entry_precompressed_with_progress`,
  which captures the member's exact compressed span (Store or Deflate),
  decodes the capture, compares every decoded byte with the payload and
  records the actual CRC. The member's declared sizes are re-checked against
  the `ReadLimits` the archive was admitted under (now stored on the package
  as `source_limits`). Every failure is a typed `OpcError`. A deferred part
  uses the index its own decode builds; an eagerly materialized owned package
  builds one transfer index per open, shared by clones.
- `payload.rs`: `PartPayload::Transferred` holds the decoded allocation and the
  verified capture as one value. `BlobPart::with_compressed_transfer`
  (`part.rs`) builds a part over it; `set_blob`, `set_blob_shared` and
  `set_content_type` replace or demote the payload, so no payload or
  content-type change keeps the capture.
- `pkgwriter.rs`: `part_entry` frames a transferred payload with
  `RegeneratedEntry::new_precompressed_shared` — fresh known-size headers, no
  data descriptor, zero timestamps, no source extras, the source method — for
  appended and regenerated members, and only when the part's visible
  allocation is the one the capture was verified against. The full
  (non-preserving) writer still re-encodes the decoded bytes.

`crates/litchi-pptx` (`opened/cross_copy_plan.rs`):

- `build_candidate` classifies every copied part before building anything: a
  relationship-free, non-XML `image/*` part that the source package reports
  eligible is transferred; everything else — including a member the ordinary
  reader decodes but whose local header disagrees with its central record —
  keeps the re-deflating route, deterministically.
- The copied-media encoding (`CopiedMedia::{Recompressed, SourceCompressed}`)
  is decided once, at first planning, and only transfers when the destination
  snapshot is an unmodified owned source. It is recorded in the plan and the
  durable patch, and every later proof — the fresh re-plan in
  `apply_cross_slide_copy_plan`, both routes of `apply_cross_slide_copy_patch`
  — rebuilds under the recorded encoding rather than deciding again.
- `validate_application_candidate` refuses (`Error::UnsafeEdit`) to publish a
  transferring copy into a destination that is not an unmodified owned
  source, before the clone-and-apply rebuild that route would run (see *Why
  modified destinations are refused*).
- The retained candidate archive of change 0656 records its transferred
  target list and is reused only when a fresh classification yields the same
  list; otherwise the application serializes its own candidate, which then
  fails the plan comparison exactly as it would without retention.
- The durable format moves to `LPCP0004`: one byte after the presentation
  relationship ID records the encoding. `LPCP0002` and `LPCP0003` are refused
  by name before any header field is read.

No `unsafe`, no new dependency, no new thread, clock or I/O; litchi-pptx still
has no archive dependency.

## Authority

- Owner decisions of 2026-09-16 ([0652](0652-owner-decisions-for-the-third-wave.md)):
  trade-off 1 authorizes the format bump and the new refusals below;
  trade-off 2 decided three points here — the capture is verified by
  decode-and-compare although pointer identity already proves the payload,
  every post-eligibility failure is a typed error rather than a quiet return
  to Deflate, and a modified destination is refused rather than made to retain
  source captures; trade-off 3 is the scope — the common benign path (an
  untouched image copied into an unmodified owned destination) gets the
  transfer, and the minority keeps today's route or a typed refusal.
- Decision 4's precedent (`LPRM`/`LPCP` bump with a typed refusal by name,
  implemented in [0655](0655-pptx-memoized-revision-proof.md)) is followed for
  `LPCP0003 → LPCP0004`; `LPRM` is unaffected.
- The dispatch fixed the transfer rule, the typed-refusal rule, the
  no-unbounded-foreign-retention rule and the proofs that stay mandatory. This
  record adds two conditions to the rule: the destination snapshot must be an
  unmodified owned source at first planning (see below), and the member's
  layout must be provable from its headers, so a member the ordinary reader
  accepts but whose local header disagrees with its central record keeps
  today's route rather than failing the copy (a read-only review of the first
  commit found that the capture alone would have refused such copies).
- ADR 0030 (lazy decode): eligibility decodes nothing; authorization decodes
  through the package's own route, so its refusal is the one `get_part`
  reports. ADR 0005: no capture outlives the candidate build; the retained
  candidate's new field is a short list of part names inside the budget
  change 0656 declared. ADR 0006: untouched destination members are still
  copied verbatim, and a transferred member is framed only from bytes the ZIP
  reader decoded, compared and checksummed. ADR 0003: plans and patches stay
  source-checked, reversible and deterministic. ADRs 0010/0011: the ZIP token
  stays inside litchi-opc.

### Why modified destinations are refused

When the destination is not an unmodified owned source, application publishes
a clone of the live destination with the patch's decoded resources applied
(so caller-defined parts and save options survive), and the targeted writer
deflates those resources again. That package cannot reproduce a candidate that
framed source-compressed bytes. Making it carry the captures would keep copies
of source bytes inside a caller's package for its whole lifetime, and dropping
them later would change what that package serializes to — both contrary to
ADR 0005's retained-state rule (change 0656 amendment). So a copy planned
against a modified destination records `Recompressed` and behaves exactly as
before, and a transferring plan or forward patch applied to a destination that
has since become modified is refused with `Error::UnsafeEdit`; planning the
copy against the current destination is the remedy. The inverse direction
publishes the restored clone and has no such requirement on the destination;
both directions still need the source to lend the same members. Recording the
encoding is what keeps the inverse route and repeated application
deterministic: they rebuild under the encoding the forward copy used,
whatever the state of the package they run against.

## Breaking changes

| item | before | after |
| --- | --- | --- |
| `CrossSlideCopyPatch::to_bytes` | `LPCP0003` | `LPCP0004`, one extra header byte (copied-media encoding) |
| `CrossSlideCopyPatch::from_bytes*` on `LPCP0003` | parsed | `Error::DurablePatchRevisionFormat { found: CrossSlideCopyV3, expected: CrossSlideCopyV4 }` before any header field is read; `LPCP0002` now reports `expected: CrossSlideCopyV4` |
| published bytes of an owned copy with eligible images | copied images deflated, data descriptor | the source member's compressed bytes and method, sized headers; target physical revisions change accordingly. Copies without eligible images are byte-identical to the base |
| forward application of a transferring plan/patch to a modified destination | published | `Error::UnsafeEdit` ("…requires an unmodified owned destination; plan the copy against the current destination") |
| a source whose eligible image was re-provenanced (replaced with equal bytes, or reopened after that) between planning and application | published | the fresh candidate differs and is refused with `Error::UnsafeEdit`, identically with and without a retained candidate |
| `CrossSlideCopyPlan` / `CrossSlideCopyPatch` `Debug`, `PartialEq` | — | carry the encoding; the plan's candidate slot reports `transferred_members` |

Additive: `soapberry_zip::office::IndexedArchive::precompressed_layout_provable`;
`litchi_opc::{CompressedPartTransfer, OpcPackage::compressed_transfer_eligible,
OpcPackage::authorize_compressed_transfer, BlobPart::with_compressed_transfer}`;
`litchi_pptx::DurablePatchFormat::CrossSlideCopyV4` (the enum is
`#[non_exhaustive]`); `CrossSlideCopyPlan::transfers_source_compressed_media`
and `CrossSlideCopyPatch::transfers_source_compressed_media`.

## Evidence that motivated it

[0739](0739-pptx-cross-copy-current-baseline.md): media-rich owned lifecycle
p50 403.95 ms, planning 71.06%, application 22%, 277.8 MB allocated per
lifecycle. [0740](0740-pptx-cross-copy-native-profile.md): generated-entry
Deflate under `build_candidate` is 78.79–79.54% of strict planning cycles.
The coordinator's sweep on the base binary: owned media-rich lifecycle
428.5 ms against 19.5 ms for the source-backed route on the same corpus.

## Measurement

Harness `tools/perf-baseline`, unchanged by this change. Both legs were built
with the identical command (`cargo build --release --locked --offline
--manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`,
plus `--features allocator-metrics --bin litchi-perf-baseline-alloc` for the
allocator lane), the before leg from the read-only base checkout and the after
leg from `b2132486af`. Every process pinned with `taskset -c 4`; four rounds of
before, after, after, before per case (16 processes, eight per arm);
media-rich cases 20 samples after 3 warmups, the other cases 40 after 3.
Statistics: median of process p50s per arm with its min–max; the median of the
eight paired ratios ((s0, s1) and (s3, s2) in each round) with a percentile
bootstrap (10,000 draws, seed 742). No run was discarded or selectively
repeated; two earlier complete matrices (below) were superseded by code
changes and by the build-matching rule, and their summaries are retained.

| binary | SHA-256 |
| --- | --- |
| before, native | `0f20b4d07456ebb6493c8f70a11876cf2b88c93f6068dee5f33272f6c6004bf3` |
| before, allocator | `3e304e2d0b1d364fe25b0857d65dd35f20a9ddc70e986c476200d6c5910c6b09` |
| after, native (`b2132486af`) | `6237d2950fb324821e0e51e5d496b71b882f8b6fbaa1f7ac619ee859e6cdacf1` |
| after, allocator (`b2132486af`) | `191dfeeb042ced08aec52959ec8adf092bbaf5fbf76089693905b90d10f30bf1` |

Host AMD EPYC 9R45, Linux 7.0.0-1012-aws, shared with other agents; Rust 1.95.0
from `rust-toolchain.toml`, release, `--locked --offline`.

### Native latency

| case | before median p50 ms [min–max] | after median p50 ms [min–max] | median paired ratio [95% bootstrap] |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 405.441 [400.360–413.634] | 185.703 [180.666–197.265] | 0.4609 [0.4460, 0.4739] |
| `pptx_cross_copy_media_rich` | 384.408 [378.707–393.460] | 165.343 [152.640–169.300] | 0.4292 [0.4196, 0.4339] |
| `pptx_cross_copy_plain_lifecycle` (control) | 8.292 [8.216–8.487] | 8.359 [8.305–8.542] | 1.0097 [0.9957, 1.0172] |
| `pptx_cross_copy_plain` | 7.062 [6.954–7.167] | 7.193 [7.106–7.410] | 1.0215 [1.0144, 1.0413] |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` (control) | 16.806 [12.030–17.036] | 16.746 [12.038–17.069] | 0.9973 [0.9838, 1.3858] |

Median process p95: 406.229 → 187.947 ms (media-rich lifecycle), 385.008 →
165.777 ms (media-rich); means 405.366 → 185.756 and 384.285 → 165.367.

Phases, median of process medians (ms):

| case | plan | commit | publication |
|---|---:|---:|---:|
| media-rich lifecycle | 288.573 → 68.674 | 88.585 → 88.850 | 6.451 → 6.336 |
| media-rich | 289.123 → 70.135 | 88.673 → 88.736 | 6.410 → 6.482 |
| plain lifecycle | 3.269 → 3.300 | 3.767 → 3.801 | 0.001 → 0.001 |
| plain | 3.283 → 3.342 | 3.786 → 3.854 | 0.001 → 0.001 |

The superseded matrices, each also 80 native and 32 allocator processes,
measured `317920af5c` and `52db88c24c` against the coordinator's prebuilt base
binary: media-rich lifecycle 408.126 → 194.768 ms (ratio 0.4742) and 411.384 →
186.599 ms (0.4522); media-rich 0.4276 and 0.4231; plain lifecycle 1.0006 and
0.9924; plain 0.9999 and 0.9968.

### Regression flags

No measured case regresses by more than 5%. Two smaller movements are listed
rather than averaged away:

- `pptx_cross_copy_plain` +2.15% (7.062 → 7.193 ms; interval [1.0144, 1.0413];
  one pair at +5.2%). A plain copy has no image, so the only new work on its
  path is O(1) — one classification vector, one encoding byte, one empty-list
  comparison, one small cell per owned open — and the two superseded matrices
  measured the same case at −0.0% and −0.3%. Build-to-build layout shifts of
  2.7–3.4% on untouched paths were reported by the coordinator for this host;
  the direction here is not stable across builds, and no mechanism is claimed.
- The source-backed control's paired ratios include 1.419 and 1.386 (and 0.716):
  its path is unchanged (identical output), and each outlier pairs a process
  whose large buffers happened to fault fresh with one whose buffers did not
  (next paragraph). Its median ratio is 0.9973.

The media-rich spread in both arms follows first-touch page faults, which the
harness records for every lifecycle sample (`faults.py` → `faults.json`). One
freshly mapped 33.6 MB buffer costs about 8,204 faults and 6–7 ms here:

| arm | fresh 33.6 MB mappings per lifecycle | samples | lifecycle ms | plan ms | commit ms |
|---|---:|---:|---:|---:|---:|
| before | 0 | 20 | 400.360 | 288.105 | 88.610 |
| before | 1 | 100 | 405.383 | 288.316 | 88.564 |
| before | 2 | 40 | 412.450 | 295.764 | 88.573 |
| after | 1 | 70 | 180.809 | 63.758 | 88.691 |
| after | 2 | 49 | 187.549 | 70.465 | 88.756 |
| after | 3 | 32 | 195.537 | 70.704 | 95.802 |
| after | 4 | 9 | 199.089 | 79.958 | 95.575 |

At equal fault counts the commit phase is unchanged (88.56 vs 88.69 ms with one
mapping), and the lifecycle ratio at one mapping is 0.446. How many large
buffers fault fresh is per-process allocator state, visible in the unchanged
source-backed control too (processes that fault nothing run 12.0–12.1 ms, the
rest about 16.8 ms). The first superseded matrix happened to put more after
processes in the higher modes, which showed up there as a +7.8% commit-phase
median; it does not recur here. The mechanism that selects a process's mode is
not established.

### Allocations (allocator lane, lifecycle region, median of process medians)

| field | media-rich before → after | plain before → after |
|---|---:|---:|
| allocation calls | 56,356 → 56,407 (+51) | 46,613 → 46,619 (+6) |
| deallocation calls | 46,068 → 46,100 | 38,273 → 38,276 |
| reallocation calls | 6,271 → 6,274 | 5,298 → 5,298 |
| allocated bytes | 277,809,048 → 272,764,566 (−1.82%) | 16,613,707 → 16,615,405 (+1,698) |
| region peak live bytes | 305,068,513 → 305,071,849 | 1,313,024 → 1,313,576 |

The eight captures (16.8 MB) replace the eight generated Deflate buffers and
compressor states; peak live bytes are unchanged within 3.4 KB. The plain
path's six extra allocations are two classification vectors and the transfer
index cells of the lifecycle's four owned opens, about 0.4 KB each (the cell
was made owned-source-only and its index boxed in `b2132486af` after the
second superseded matrix measured ten allocations and 4.6 KB per plain
lifecycle).

### Output bytes

Media-rich output 33,599,873 → 33,599,745 bytes (the eight 16-byte data
descriptors), SHA-256 `6a3536fc…` → `68be0ce5…`, one digest per arm across all
24 media-rich processes of that arm. Plain outputs are byte-identical to the
base (`3e9ae280…`, 31,545 bytes); the source-backed control's output is
unchanged (`809c6172…`, 33,599,715 bytes).

**Owned vs source-backed (informational).** Not byte-identical on the harness
corpus (33,599,745 vs 33,599,715 bytes). The packet's `route-compare` probe
publishes one copy of a harness-shaped pair through both routes on the final
commit: all 16 image members are byte-identical (same names, compressed bytes
and framing); the archives differ only because the owned route names the
copied slide `slide3-copy1.xml` where the source-backed route keeps
`slide3.xml`, which also changes `[Content_Types].xml`, the presentation
relationships and the member order.

## Where the remaining time goes, and the next opportunity

A frame-pointer build of `b2132486af` (SHA-256 `8b54a242…`, packet
`attribution/`) was profiled on the media-rich lifecycle (cycles, `--call-graph
fp`, 12 samples); shares are of cycles within a phase, not wall time.

- **Commit (about 89 ms)** is 80.0% SHA-256, in five passes of about 16% each:
  the live source and destination semantic re-fingerprints
  (`package_fingerprint`), their physical re-fingerprints (`to_stream` into a
  hash sink), the fallible copy and digest of the retained archive, the
  reopened candidate's semantic capture in `build_candidate`, and the same
  candidate's capture again in `validate_application_candidate`. The eager
  reopen (inflate, CRC, donation compare) is 6.9%.
- **Plan (about 69 ms)**: serializing the candidate with its archive digest
  35.9%, the snapshots' physical revisions 22.0%, the candidate's semantic
  capture 21.8%, the eager reopen 9.5%, and this change's capture
  verification (decode, compare, CRC) 4.2%; the header-only layout proofs do
  not reach the listed paths.

Next opportunity, not implemented here: stop re-hashing bytes whose identity
is already proven. `validate_application_candidate` recaptures the candidate
`build_candidate` just captured; the candidate capture could consult the
snapshots' part-digest memos (the reopened candidate's copied and untouched
parts share donated allocations); the live re-fingerprints at the start of
application could use the facade's memo (change 0655) and a per-package memo
of the immutable retained archive's physical revision; and owned ingress that
accepts a shared archive would remove the retained archive's copy (its digest
was deliberately kept by 0656). Each needs its own ADR 0005 memo argument and
freshness proof. Page-fault sites of the same build are summarized in
`attribution/faults-by-site.txt`: the snapshot decodes, the candidate
serialization buffer and the publication sink.

## What is not claimed

No claim-registry entry; no Office or other native-application validation of
the published archives (Litchi reopens them and decodes every copied member);
only the two generated corpora and the unit fixtures; no real-producer media;
no cold-cache, remote-source, concurrency, RSS or instruction-count result; no
statement that the lifecycle is faster by the same factor in every host state
(the fault-mode spread above is host- and allocator-dependent); the
source-backed control's 0.9973 is not an improvement and the plain +2.15% is
not attributed to code; the attribution shares are one diagnostic capture of a
frame-pointer build, not ordinary-release timings. The broader non-iWork goal
remains active.

## Verification

A read-only review of `317920af5c` by a separate reviewer found no correctness
bug on any publication route and seven lesser findings; `52db88c24c` and
`b2132486af` act on them: the layout-disproof classification (its main risk),
the pointer check in the writer, a direct test of the capture's own
verification, the per-open transfer index, the documentation corrections and
two stronger tests. It also noted that an `LPCP0003` patch could in principle
be read losslessly as the recompressed encoding; the dispatch's decision to
refuse the old format by name stands, and the documentation now says it is a
format decision rather than an impossibility. Its note that
`apply_cross_slide_copy_patch`'s forward route maps a failed re-plan to the
generic `UnsafeEdit` predates this change and is unchanged.

New tests: one in `soapberry-zip` (the layout predicate classifies Store,
sized and descriptor Deflate members as provable and a renamed local header as
disproven, and a provable member captures); 15 in `litchi-opc`
(`package/compressed_transfer/tests.rs`: exact span and fresh framing for
Deflate with and without a descriptor and for a Store member with source
descriptor, timestamp and extras; eager and deferred captures equal;
eligibility never decodes; replaced-with-equal-bytes, retyped,
relationship-bearing, XML, borrowed and authored parts refused by type;
signed packages; corrupt stream, CRC mismatch and data-descriptor mismatch each
refused with a stable typed error; a local/central name mismatch classified
ineligible without decoding while the lenient reader still decodes it; the
capture's decode-and-compare refusing bytes that are not the member's; read
limits re-checked; mutation discards the capture; a forwarding custom part
publishes the bytes it shows; regenerated members; full-writer fallback) and
11 in `litchi-pptx` (`opened/cross_copy_plan/media_transfer_tests.rs`: span
and framing through the PPTX route with renamed media; retained, released,
durable and independently planned outputs byte-identical and the inverse
exact; a replaced payload recompressed beside a transferred one; SVG still
refused; plain copies record `Recompressed`; the transferable-media rule
including its XML guard; modified destination refused with its message until
re-planned; the inverse applies to a modified destination; stale, foreign and
re-provenanced sources refused identically with and without retention; a
disagreeing local header recompressed beside a transferred image; genuine
`LPCP0003` patches from the base tree refused by name in both directions, with
their input hashes re-derived and the changed target physical revision shown).
The existing superseded-magic test now covers `LPCP0002` and `LPCP0003`
against `LPCP0004`. The legacy fixtures are
`test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-{forward,inverse}.patch`
(5,685 bytes each), written by the packet's `legacy-fixture/` generator
against the base tree. soapberry-zip's existing capture tests cover a corrupt
compressed payload, a truncated Deflate stream, expected-byte mismatch and
ZIP64 members.

Gates on `b2132486af` (commands, exit codes and counts in
[`gates.txt`](results/change-0742/gates.txt)): `cargo fmt --all --check`;
`cargo check --all-targets` of `soapberry-zip`, `litchi-sign`, `litchi-opc`,
`litchi-pptx` and their in-scope dependents; Clippy `-D warnings` on the three
libraries and on `soapberry-zip` and `litchi-opc` all targets
(`litchi-pptx --all-targets` fails only on three `err_expect` lints at
`opened/tests.rs:464/538/557`, present unchanged on the base); tests of the
three crates, of their seven in-scope dependents and of the facade with
`doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`; rustdoc `-D warnings` for the three
crates; the crate-boundary, non-iWork and structural claim gates.
`soapberry-zip`'s ODF and iWork dependents were not rebuilt; its change is one
additive method.

## Cleanup

Binary identities are recorded above and in
[`cleanup.json`](results/change-0742/cleanup.json), taken before removal. The
target directories `targets/0742`, `targets/0742-before`, `targets/0742-fp`
and `targets/0742-probe`, the scratch directory with its perf captures, and
the raw reports of the two superseded matrices were removed; the worktree and
branch are kept, as is the coordinator's shared base build.
