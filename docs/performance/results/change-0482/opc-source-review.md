# OPC source-backed splice review

Status: read-only correctness, budget, and security review of the current
`source_backed/splice.rs` and its XML reader seam. The worktree source was
last inspected at 2026-09-09 00:11 UTC. This review does not authorize a
merge or a performance claim.

The reviewer development commands and their exact outputs are retained in
[`reviewer-development-attempts.json`](reviewer-development-attempts.json).
Those commands include earlier Cargo checks and focused tests run before the
root-only test rule was established. They are development observations, not
root acceptance gates; no Cargo, test, profile, or measurement command was run
during this focused pass. The root-owned receipt
`validation/opc-dev-tests-04.json` reports 20 focused OPC cases at 23:37;
the decoder-bound and owner-fixture edits continued afterward, so that
receipt is not the final gate. The XML/harness development receipts likewise
are development observations. Root still owns the final build, lint, and test
decisions.

## What is correct in the current shape

`prepare_source_part_splice` classifies the target with
`xml_minifier::audit::package::is_xml_part(name, content_type)`
(`crates/litchi-opc/src/source_backed/splice.rs:578-581`). XML targets use the
generic authored audit. Binary Parts use the same complete decoded source and
candidate hash/length proofs and drain the candidate stream without applying
XML grammar. This keeps the public operation generic while retaining an
explicit XML-only audit boundary.

The source proof and candidate proof are separate fresh verified-reader passes
before a plan is returned (`splice.rs:1067-1184`). The candidate pass presents
prefix, fragment, and suffix as one stream. Publication opens another fresh
reader for each replay callback and applies `verify_authored_reader` to the
complete candidate for XML targets (`splice.rs:407-429,1211-1285`). A format
callback returning a proof is therefore not accepted as the generic XML
validation authority.

The exact no-op path independently rehashes the decoded source and requires
the candidate hash to equal the source hash and the fragment hash to be the
empty digest (`splice.rs:305-325,1107-1123`). It then copies the retained
physical source artifact. XML auditing, changed-signature refusal, and changed
encryption refusal are intentionally skipped for this exact-copy route, so
malformed and signed sources retain the same parity as ordinary source-artifact
publication. The no-op archive ceiling is checked as
`ReadResource::ArchiveTotalBytes` before the sink is touched.

The changed route retains the source compression method, rejects signature
infrastructure and encrypted entries, rejects trailing bytes after the located
archive, and keeps ZIP types private to OPC. The public plan exposes proof
scalars and a publication method, not `PreservationIndex`, entry IDs, or raw
ZIP records. Source, context, output-budget, chunking, and accepted-byte
wrappers surround the replay sink; final source freshness is checked after the
last flush (`splice.rs:337-558`).

## Findings requiring resolution

### 1. Replay and decoder bounds are owner-scoped with explicit sink-adapter accounting

The lower layer now exposes a replay bound instead of making OPC guess the
compressor size. `replay_memory_upper_bound()` charges the three simultaneous
64 KiB arrays (replay input window, compressor output window, and preservation
copy buffer), the pointer-shaped
`size_of::<ReplayPayloadWriter<'static, &'static mut ReplaySink<io::Sink>>>>()`
writer layout, `size_of::<ReplaySink<io::Sink>>()` for fixed adapter state, and
a 512 KiB envelope for the locked zlib-rs Deflate allocation
(`crates/soapberry-zip/src/preserve/replay.rs:37-65`).
This covers the previously missing third fixed array and is above the 380,032
byte allocation recorded in the independent change-0475 owner receipt. The
bound is conservative for the locked `flate2` 1.1.10 plus `zlib-rs` 0.6.7
configuration, provided that the dependency lock and allocator formula remain
part of the owner contract.

The replay sink seam is now covered explicitly. The emit writer uses the
pointer-shaped marker and the scalar adds the fixed `ReplaySink<io::Sink>`
layout; the owner test uses the same two terms. The contract's exclusion of
arbitrary caller sink storage and private buffering remains intact
(`opc-splice-contract.md:231`), while the replay-owned pointer and progress /
error bookkeeping are now charged. This closes the prior adapter-bound gap.

The verified-reader seam now exposes a compression-specific bound and reserves
it for every callback, including source/candidate preparation and no-op proof
(`office.rs:1124-1137,3411-3459`; `source_backed.rs:3658-3677`). Store readers
charge the 16 KiB callback buffer plus the concrete positional source counter
and reader state. Deflate readers add flate2's 32 KiB input buffer, the
decoder wrapper layout, `size_of::<Decompress>()`, and a 64 KiB zlib-rs inflate
envelope. The XML workspace reservation separately covers the
`SpliceAuditReader` adapter and the XML auditor's declared bound; changed
binary targets reserve the same adapter term.

This removes the earlier concrete undercharge, including the previously
omitted decoder/source wrapper state. Acceptance still depends on the owner
keeping these backend envelopes synchronized with the locked dependencies and
on root's build evidence for the production target. The owner comments are
source-derived and conservative, but they are not a portable claim for a
different compressor backend or feature set.

### 2. Preservation now exposes a conservative owner envelope; breadth evidence remains needed

`preservation_memory_requirement` now consumes the ZIP owner's scalar bound
and adds OPC's separately allocated 64 KiB index scratch
(`splice.rs:1037-1045`). The owner formula uses `size_of` for the private
`PreservedEntry`, action, prepared-entry, offset-patch, and promoted-central
types; it separately charges prepared/layout vectors, twelve ZIP64 source
ranges per entry, promotion extra bytes, a prepared-tail term, and one 64 KiB
copy buffer (`crates/soapberry-zip/src/preserve.rs:335-395`). The ODF
`INSERTION_OBJECT_BYTES = 512` peer constant is no longer used as evidence.

The formula also charges `metadata_bytes * 4`. This is materially stronger
than the former `metadata * 2`: `metadata_bytes` includes every source central
record, ZIP64 tail, and archive comment, while replay target preparation retains
generated local framing and central bytes and ZIP64 promotion can retain a
second generated central representation. For a long target name, the source
central record plus generated local and central records is roughly three
name-sized copies, with a possible fourth promotion copy. The four metadata
terms therefore cover the raw-byte peak, while the explicit private-type and
range terms cover the containers around it.

On source inspection this removes the earlier concrete undercharge. The bound
is still a source-level contract rather than a measured whole-process peak:
the owner should retain long-name and ZIP64-promotion tests, keep the
`size_of` formula synchronized with private layout changes, and keep the
locked target assumptions visible. The formula's small fixed tail allowance
(`+128`) should remain accompanied by those owner tests; it must not silently
be treated as a universal allocator or error-string allowance. The first
long-name owner-test attempt failed because its 60,000-byte fixture exceeded
the test archive's default 4,096-byte member-name admission limit; the live
fixture now supplies an explicit raised limit, and the final ZIP owner suite
passes that fixture.

### 3. Inverse authorization and output quota are fenced in the current path

The earlier candidate-freshness and current-output-quota defects are fixed in
the current path. Inverse publication enables monitoring before fingerprinting,
checks the current candidate while hashing and copying, wraps the copy in the
current operation's `OutputBudgetedSink` when managed, and checks both source
snapshots and contexts after the final flush (`splice.rs:769-870`). The exact
candidate artifact SHA remains the authorization; reopening an identical
physical artifact does not require matching the original source version or
lineage. The inverse path reserves one 64 KiB workspace and reuses it for
fingerprinting and copying. This is a bounded operation-local claim; it does
not include arbitrary caller sink buffering.

The current source uses an explicit internal copy path rather than nesting the
retained artifact's original output-budget wrapper. That preserves the
selected current-context policy without double-charging a shared budget, while
retained-context cancellation and freshness remain checked independently.

### 4. Fragment work and physical read cancellation are bounded; decoded cadence is unspecified

Fragment hashing charges `Resource::Work` in 64 KiB chunks and checks source
freshness before and after each charge (`splice.rs:1194-1211`). During source,
candidate, and replay passes, fragment consumption repeats the Work charge and
checks freshness around it (`splice.rs:1604-1666`). Since
`ExecutionContext::consume` checks cancellation before accepting a charge, the
focused large-fragment cancellation path is sound and failed preparation does
not retain a plan or reservation.

The source branch of `SpliceAuditReader` delegates reads to the verified
reader. Its `SourceReader::read_at` calls
`read_source_at_with_context`, which checks the execution context before and
after each bounded physical source read and checks the monitored source
version at the same boundaries. The verified-reader wrapper also checks source
and context before and after the complete callback. This satisfies a
physical-read and operation-finalization interpretation of the contract, and
it prevents a cancelled or changed source from being reported as a successful
publication.

There is no separate check after a fixed number of *decoded* bytes that were
already buffered by the Deflate reader. A highly compressible physical read
can therefore yield substantially more decoded bytes before the next physical
read check. The ADR says the source is checked “periodically during scanning
and copying” (`opc-splice-contract.md:198-204`), but does not define whether
that interval is physical input, decoded output, or a maximum latency; the
reader contract says only that declared decoded work is consumed and reader
state is reserved (`opc-splice-contract.md:281-287`). This is a contract
clarification and latency note, not a demonstrated production blocker. If the
accepted contract requires a decoded-byte cadence, add that explicit bound
and a check at the splice adapter's bounded fill/consume boundary, retaining
the current typed pending-failure path.

## Error, accounting, and format checks

The changed path merges both decoded-reader reports and the low-level replay
report after the source/context decision (`splice.rs:517-547`). The source
publication helper preserves accepted-byte progress and gives source changes,
cancellation, and flush failures precedence over secondary accounting errors.
The no-op path uses the existing exact-artifact accounting route and checks the
physical archive ceiling before any copy. These paths are structurally sound;
the root run should retain the exact typed errors for partial writes and
accounting overflow.

The final root-owned receipts now report successful all-feature OPC tests,
OPC clippy/rustdoc, workspace check, the format/harness check, and the ZIP
owner test/clippy/rustdoc suites. The earlier focused receipt at 23:37 and
the intermediate full OPC and ZIP failures remain useful development history:
the former used the old 16 KiB reader expectation, and the latter used the
old 4,096-byte name limit. Both were superseded by the final receipts after
the bound and fixture updates. The current XML test/no-default/clippy receipts
also pass; the recorded XML rustdoc failure from 23:49 predates the
`audit.rs` intra-doc-link fix at 23:50:55 and is superseded by the final
rustdoc receipt. These receipts are still not a complete whole-process memory
measurement or a performance claim.

Physical replay tests should continue to cover raw untouched local spans,
central records, descriptors, ZIP64 tails, comments, and opaque members. The
current source preserves the existing low-level replay authority and keeps raw
ZIP types private, but the focused OPC fixture is not itself evidence for every
validated ZIP64 topology.

## Acceptance state

| Requirement | Current review result | State |
| --- | --- | --- |
| XML classification for generic Parts | `is_xml_part(name, content_type)` gates only the generic XML audit; binary paths hash and drain | Code review pass |
| Independent source and candidate proofs | Fresh complete passes run before a plan; replay callbacks reopen fresh readers | Code review pass; dev receipt green |
| Generic authored audit before output | Complete candidate audit is repeated in each replay callback for XML targets | Code review pass |
| Exact malformed/signed no-op | Authenticated no-op proof plus exact source-artifact copy | Code review pass; dev receipt green |
| Source/version and partial-output precedence | Source/context/output wrappers and final flush checks are present | Code review pass; final OPC receipt pass |
| Store/Deflate and opaque-member replay | Existing replay API is wired and the final ZIP owner suite passes | Final receipt pass |
| No raw ZIP types in the public format API | Physical index and entry IDs remain private to OPC | Code review pass |
| Fragment/parser/adapter reservations | Fragment capacity, XML/binary adapter workspace, and owner-scoped Store/Deflate decoder state are reserved | Code review pass; final OPC receipt pass |
| Replay/compressor reservation | Owner bound includes three fixed arrays, pointer-shaped writer layout, fixed replay-sink adapter state, and a 512 KiB zlib-rs envelope | Code review pass; final ZIP owner receipts pass |
| Preservation reservation | Owner bound includes private type sizes, explicit ZIP64 ranges, scratch, and `metadata*4` raw-byte terms | Code review pass; final raised-name-limit fixture passes |
| Exact inverse authorization and freshness | Candidate SHA authorization, pre-fingerprint monitoring, write/flush freshness, and post-output checks are present | Code review pass; final OPC receipt pass |
| Inverse active output quota | Current managed candidate context owns the one output-budget wrapper; retained context remains independently checked | Code review pass; final OPC receipt pass |
| Mid-callback cancellation latency | Fragment Work chunks and every bounded physical source read check cancellation; decoded output already buffered by Deflate has no separately specified cadence | Code review pass; contract cadence remains unspecified |
| Root acceptance and performance evidence | Final code, lint, rustdoc, workspace, boundary, fuzz, and harness receipts pass; no formal performance run is claimed | Code gates pass; performance evidence intentionally absent |

No complete-memory or performance claim should be attached to this batch until
any formal performance evidence is recorded separately. The current code
review found no remaining source-integrity, publication-safety, or named
memory-reservation blocker.
