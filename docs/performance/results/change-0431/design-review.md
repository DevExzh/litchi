# 0431 compressed source-part transfer: design review

This is a source-only review. I read the accepted ownership, preservation,
validation, snapshot, and I/O ADRs; the 0430 transfer design; the current ZIP,
OPC, and PPTX source-backed paths; and the unapplied ZIP handoff in
`/tmp/litchi-goal-0431-drafts/zip`. No Rust code, builds, tests, scripts, or
measurements were run for this review. No OPC/PPTX implementation draft was
available when the initial review was written, so its integration disposition
was conditional; the applied-path review below supersedes that note and remains
source-only.

## Disposition

The physical contract is sound as the smallest next integration step:
`soapberry-zip` issues an opaque, fully verified Store-or-Deflate token;
`litchi-opc` authorizes that token against the source artifact and topology;
PPTX retains its existing decoded semantic checks and asks OPC to attach the
token only during publication. The token must be consumed through the existing
prepared-entry/preflight path, with a fresh destination wrapper.

The ZIP draft has one concrete correctness gap before it can be accepted:
`/tmp/litchi-goal-0431-drafts/zip/preserve.rs:2042-2048` classifies every
Deflate regenerated payload as `GeneratedDeflate`. A precompressed Deflate
token therefore loses the `AccountingWriteKind::Precompressed` classification;
`precompressed_payload_bytes_emitted()` remains zero and the output is charged
as newly generated Deflate. The Store case can remain `Stored`, but the
precompressed Deflate case must be `Precompressed`, with a partial-sink
accounting regression. The handoff itself calls for this at
`api-proposal.md:72-78`.

## Required physical proof

The proposed `VerifiedPrecompressedEntry` is the right capability boundary.
Its fields must stay private, and no second public constructor may accept raw
compressed bytes plus caller-supplied CRC and sizes. The issuer at
`office.rs:3609-3770` should continue to use the strict layout proof and copy
only `data_start_offset..data_end_offset`. It must then:

- resolve and validate the data descriptor, including signed and unsigned
  forms, descriptor width, ZIP64 placeholders, and the authoritative
  descriptor CRC;
- require Store compressed size to equal decoded size;
- require Deflate to reach end of stream and consume exactly the bounded
  compressed range, rejecting short, trailing, or malformed data;
- compare decoded bytes with the already validated logical bytes in bounded
  chunks, check the declared decoded size, and compute the actual CRC; and
- retain the actual CRC in the token. A zero source CRC may follow the current
  compatibility policy, but it must never be emitted as an unchecked zero in
  the destination wrapper.

The draft's `expected.valid(actual)` is the correct CRC-zero policy. Reusing
the ordinary strict callback finalizer would be wrong: current
`VerifiedEntryBufReader::finish` calls `valid_strict`, which rejects a
nonzero actual CRC when the expected CRC is zero. Keep this policy explicit and
test nonempty and empty Store/Deflate zero-CRC entries separately.

## Callback, source, and error behavior

The revised handoff reports bounded compressed and decoded progress and aborts
immediately on callback failure. That is appropriate for cancellation: it
avoids draining an arbitrarily large expansion. The callback error must never
be converted into a valid token or into a logical recompression fallback.
Source-backed callers should check authority before allocation, during every
compressed and decoded progress callback, after descriptor/decoded validation,
and immediately before topology publication. A source mutation, transport
failure, CRC/size failure, descriptor failure, or decode failure must remain a
typed failure. Only an explicitly classified, preflighted optimization or
budget refusal may select the existing logical path.

The token must carry or be wrapped by private source lineage/version authority;
URI, content type, and equal decoded bytes are not sufficient. Issue it after
`current.matches(plan)` and the existing `verify_candidate` gates. Do not put a
ZIP `EntryId`, local span, or descriptor range into PPTX. `litchi-opc` must
remain the physical boundary and must retain its signature, active-content,
source-freshness, and cancellation refusals.

## Writer, limits, and preservation semantics

`RegeneratedEntry::new_precompressed_shared` must accept only the opaque token.
It should use the existing known-size precompressed writer, emit canonical
destination local and central records, and preserve the token's method and
actual CRC. It must not copy source local headers, descriptors, flags,
timestamps, extras, or offsets. Source ZIP64 framing is not automatically
destination ZIP64 framing: the destination wrapper must be selected from the
actual target size, offset, entry count, and output layout. Test ZIP32 and
ZIP64 size/offset/count cases.

Before the sink sees a byte, the integration must have admitted the source
entry, reserved compressed and logical memory, checked aggregate limits and
`u64` to `usize` conversions, validated topology/collisions/content types and
relationships, and prepared the complete output layout. Charge the retained
decoded allocation, compressed token capacity, and any one-entry staging copy;
charging only `Vec::len()` is insufficient if allocator capacity is larger.
The output bound must include local and central records, names/extras,
descriptor policy, and ZIP64 tail/promotions. Do not catch all token errors and
fall back: corruption and source errors must happen before output and must not
be hidden by recompression.

Payload transfer preserves the source compressed bytes and method under a new
canonical target wrapper. It is not cross-archive raw-member preservation.
Untouched destination members continue using `PreservationAction::Copy` and
retain their existing raw records. Before landing, the product contract must
confirm that this payload-only behavior is acceptable for `Preserve`; if
source timestamps, extras, flags, descriptors, or exact member framing are
required for the newly added part, this capability must refuse rather than
silently approximate it.

## Focused acceptance cases

ZIP tests should cover Store and Deflate, signed and unsigned descriptors,
ZIP64 descriptors and resolved sizes, zero and nonzero CRCs, method/local/
central mismatches, truncated/overlapping ranges, trailing Deflate bytes,
unsupported methods, callback cancellation in both progress phases, source
mutation, and allocation/limit refusal. Append a verified token through the
preservation writer, reopen it, compare the transferred compressed range byte
for byte, and check method, CRC, decoded size, canonical metadata, ZIP32/ZIP64
framing, accounting, and no output on preflight failure.

OPC tests should prove that only an OPC-authorized token enters a topology,
that source lineage/version and destination checks remain enforced, and that
signature/active-content refusals, content types, relationships, ordering,
unknown untouched members, cancellation, source mutation, and partial sink
behavior are unchanged. Verify compressed and logical budgets independently.

PPTX tests should retain the existing decoded image/chart validation and
dependency-closure checks, including shared images, chart XML, relationship
rewrites, stale/foreign sources, size mismatch, read failure, and sink failure.
Add Store and Deflate fixtures and assert logical equivalence plus exact
transferred compressed payloads; do not replace semantic validation with a
compressed-byte comparison.

## Recommended next step

Implement one private OPC-owned topology payload variant for source-backed
media first. Keep the current decoded `PreparedImage`/`PreparedChart` data and
all `verify_candidate` reads, issue the ZIP token only after those gates, and
consume it in the existing preflighted preservation writer. Keep generated XML
and relationships on the Deflate path. Land the accounting fix and focused ZIP
tests before wiring PPTX; then add OPC authorization and one media lifecycle
case before considering chart transfer or any staging-copy optimization.

## Applied-path review

The current OPC/PPTX wiring follows the boundary above. `SourceBackedPackage::authorize_precompressed`
checks source freshness and execution policy before and after the bounded capture, checks the central
declared sizes against the already retained decoded allocation, applies the per-entry and compressed-byte
limits before reading the payload, and retains lineage, revision, content type, decoded bytes, and the
managed reservation in the opaque token. `write_topology_to_stream` fences every token source during
publication and validates XML additions from the decoded allocation. The PPTX path still runs
`verify_candidate` and the existing relationship/dependency checks, then resolves each media or chart URI
against the source package that issued the plan. Its fallback is limited to an OPC managed
`Resource::Memory` admission refusal; source mutation, cancellation, work limits, read limits, archive
corruption, transport, CRC/size, signature, and relationship refusals remain hard errors. That is the
right fail-closed split, and the existing output writer remains responsible for the destination wrapper.

The publication error precedence must stay as implemented in the current path: if the preservation writer
returns an error after accepting a prefix, its `IncompleteOutput { written, .. }` value must win over the
final transferred-source fence. A bare post-publication `SourceChanged` or cancellation error would erase
the accepted-byte count. The current match in `write_topology_to_stream` preserves an existing publication
error and only synthesizes `IncompleteOutput` when the writer returned success but the final source fence
observed a transfer-source failure. Keep that behavior in the frozen version.

The final precompressed reservation models the captured compressed `Vec` (`C`) plus the generated-member
payload/fixed capacity (`C + fixed`); the separately charged target-name reservation (`2*name`) accounts
for the generated buffer's two name copies. The resulting `2*C + fixed + 2*name` reservation is coherent
for those explicitly modeled payload buffers. It is not a
whole-operation heap bound: temporary `String`/URI normalization, ZIP metadata objects, hash-table
entries, allocator rounding, and preservation-index allocations remain covered only by their separate
topology/index reservations or by the general execution accounting. Documentation should keep that scope
explicit and make no direct full-heap or zero-copy claim. The Memory-only PPTX fallback is sound only for
the explicitly scoped admission refusal.

The final Clippy layout correction boxes only the large `Precompressed` enum variant; ordinary decoded
additions remain inline. This changes the representation and per-addition allocation shape, but preserves
token ownership, source guards, reservations, and publication ordering. It does not establish a zero-copy
or whole-operation allocation claim.

## Deflate termination review

`verify_captured_deflate_payload` now requires the decompressor to return
`Status::StreamEnd`; it does not stop when the expected decoded bytes have been produced. It then requires
both the decoder's consumed-input count and the tracked input offset to equal the captured compressed
range, followed by the decoded length and CRC checks. With the pinned `flate2` `zlib-rs` backend, a
truncated stream that emits the complete expected prefix but lacks its final block reaches a non-terminal
status; the zero-input/zero-output guard rejects it. A valid stream therefore has a real end-of-stream
proof, while trailing compressed bytes fail the exact-consumption check.

The frozen loop feeds 64 KiB input slices with `FlushDecompress::None` and uses `Finish` only for the
final slice, retrying with an empty final input slice when more output is pending. This avoids the
alternate miniz backend's first-Finish whole-output assumption while retaining a real terminal-status
proof. It also invokes the progress callback after every decoder iteration, including zero-output
iterations, so cancellation and source fences are observed during finalization. The existing large-output
fixture (128 scratch-buffer lengths) covers bounded output continuation; retain a truncation regression
that otherwise produces the expected bytes and a trailing-byte rejection so the terminal-status and
exact-consumption rules remain explicit.

The source-oracle audit is consistent with this contract: it compares decoded and physical untouched
destination members, explicitly excludes the three regenerated package-level XML members, and separately
states that decoded PPTX equality does not prove compressed-token reuse. The harness/tool tree was unchanged
in this source review; no builds, tests, scripts, or measurements were run.
