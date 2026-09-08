# Change-0476 reusable Deflate source review

## Review scope

This is the independent source/design review for the reusable owned-ZIP
Deflate candidate. The reviewed baseline is the 0475 transport audit at
`cb6e4b7180a791bfdb1f704f2680b846a7bfc1fd`; the production diff was not yet
present when this initial review was written. The review is intentionally
read-only with respect to production code and tests.

## Findings to carry into the implementation review

### 1. The baseline finish path emits a sync flush

`ZipDataWriter::finish` calls `self.flush()` before returning the descriptor.
For an owned Deflate entry this dispatches to
`DeflateEncoder::flush`. In flate2 1.1.10, `zio::Writer::flush` runs the
codec with `FlushCompress::Sync`, drains its 32 KiB output buffer, and then
flushes the wrapped sink. `OwnedCompressor::finish` subsequently calls
`DeflateEncoder::finish`, which emits the final `FlushCompress::Finish`
bytes. Thus a raw-compressor implementation that only runs `Finish` will
usually produce a different member byte stream even when decompression,
CRC, and sizes agree. Complete archive bytes or hashes must be compared to
the baseline; if exact determinism is required, the replacement must model
the baseline's sync-flush sequence and its output ordering.

The sync flush is a per-member boundary in the current stream. It must not be
replaced by a `Full` flush or by concatenating members into one codec stream.
If the implementation deliberately removes the baseline sync marker, the
change needs an explicit output-contract decision and updated evidence rather
than an unqualified determinism claim.

### 2. Pending compressed output must be separated from codec progress

`zio::Writer::write_with_status` first drains all pending output, then invokes
`Compress::compress_vec`, and returns only the input bytes consumed. A direct
loop needs equivalent state: a fixed output buffer, the valid pending-output
range, and the codec's consumed-input delta. It cannot call the codec again
while a prior output prefix is still undelivered. The codec's
`total_out` counts bytes placed in the caller's buffer, not bytes accepted by
the archive sink, so output accounting must advance only after
`OwnedCompressedEntry::write` accepts bytes.

For a short successful sink write, retain the remainder of that exact output
range and continue draining it. For a zero write, return `WriteZero` and mark
the entry failed. For an error after a partial write, retain the accepted
compressed-byte progress and poison the operation; do not claim that all input
whose codec call already consumed was accepted by the public writer after an
error. This is the same failure shape as `zio::Writer::dump`.

### 3. Uncompressed acceptance and CRC must follow the codec's consumed prefix

`ZipDataWriter` updates its uncompressed count and CRC from the byte count
returned by the compressor. A direct compressor must return the input prefix
whose processing was completed according to the existing `Write` contract,
and must not advance the outer count for input still held in a pending codec
call after a sink failure. A loop that consumes all input before draining
output can over-report accepted input on an output-limit, sink, or
`WriteZero` failure. This is especially important because
`OwnedCompressedEntry` enforces compressed limits before writing generated
bytes, while the outer Office writer translates that error into a typed
compressed-size limit.

### 4. Finish must prove stream completion and discard dirty state

Use `FlushCompress::Finish` until `Status::StreamEnd`, draining every output
prefix before each next codec call. A `BufError` or an iteration with neither
input consumed nor output produced must become a typed compression/progress
failure, rather than an unbounded loop. The reusable state may be reset only
after all final Deflate bytes have been accepted, the descriptor has been
written, and the archive's `FileHeader` has been appended. Any codec error,
limit failure, sink error, zero progress, descriptor failure, or incomplete
entry must drop or otherwise discard the state. In particular, an entry's
`Drop` must not return a partially finished `Compress` to the next member.

`flate2::DeflateEncoder`'s `zio::Writer` drop path attempts to finish and
ignores errors. A reuse pool must not inherit that behavior: cleanup on the
failure path must be a discard path, and only a proved-success path may call
`Compress::reset` and publish the state for the next member.

### 5. The codec configuration and call boundaries are part of output identity

The replacement must retain raw Deflate (`zlib_header = false`),
`Compression::default()`, and the same effective input boundaries as the
baseline unless complete archive-hash comparisons prove that the backend is
invariant for the selected calls. A fixed staging input window may improve
allocation behavior but can change the bitstream when it changes how the
codec receives input. The implementation should compare empty, one-byte,
multi-write, large, and highly compressible/incompressible payloads, then
compare complete archive bytes, not only decompressed data.

### 6. Archive framing remains outside the reusable codec

The local header and preselected ZIP32/ZIP64 descriptor width are emitted
before payload bytes. All Deflate output must precede the descriptor. The
reuse state must not own or mutate `ZipArchiveWriter::files`,
`file_names`, `StreamingArchiveWriter::names`, OPC name indexes, or their
metadata charges. The compressed scratch capacity is not output-budget
credit; generated bytes continue through `OwnedCompressedEntry` and
`BoundedOutput` so compressed, total, output, and metadata limits retain their
current semantics.

## Required focused tests for the production diff

The test batch should include:

1. Two sequential Deflate entries after reuse, with exact archive-byte/hash
   equality to the pre-change writer and successful independent decompression.
2. Empty, one-byte, multi-call, large, compressible, and incompressible
   payloads, including a case whose generated output exceeds the scratch
   buffer.
3. A short-write sink that accepts small positive prefixes, a zero-progress
   sink, and a sink error after partial output. Assert typed progress and
   poisoning, then prove no dirty compressor is reused by a subsequent entry
   in a fresh operation.
4. Compressed-limit failures during ordinary writes and during final output,
   plus ZIP32 and explicit ZIP64 descriptor paths.
5. Drop of an unfinished entry and failure while writing the descriptor or
   final archive. These paths must not publish the name/header or recycle a
   live codec.
6. Repeated sequential members at a fixed payload and varying payload sizes,
   verifying that reset restores counters and compression configuration.

The implementation review should record whether output equality is exact or
whether the new stream is only semantically equivalent. Performance evidence
must keep temporary codec allocation effects separate from retained archive
metadata and central-directory storage.

## Implementation-diff review (writer.rs, current worktree)

The new `OwnedDeflateState` correctly uses raw `Compress`, retains a 32 KiB
scratch array, mirrors the `Sync` then `None` drain in `flush_to`, checks
consumed/produced deltas, and only resets the codec after descriptor and
central-record publication. The ownership shape also drops the boxed state
on an incomplete or failed entry instead of returning it to the parent.

The pending-output semantic gap identified in the initial review is resolved
in the current diff. `OwnedDeflateState` now carries `pending_start` and
`pending_end`, writes codec output into the remaining fixed 32 KiB capacity,
drains output that was pending at the start of a call, and leaves newly
produced bytes pending until the next drain. `flush_to` performs the sync
call before dumping the residual range, then follows the baseline's no-flush
drain loop; `finish_to` drains before and after each finish call. This matches
flate2's `zio::Writer` ordering and preserves `ZipDataWriter`'s consumed-input
and CRC accounting on sink and compressed-limit failures.

The added direct `Write::write` differential tests compare accepted input,
sink-visible output after each call, complete archive bytes, partial sink
failure timing, and compressed-limit timing against a fresh
`DeflateEncoder`. The fixed scratch range also handles short positive writes,
zero writes, and errors without returning an incomplete state to the parent:
only `OwnedCompressor::finish` after descriptor publication resets and stores
the codec, while every failing finish path consumes and drops the active
state.

No blocking issue remains in this area. A future maintenance test could add an
`Interrupted` sink retry explicitly, but the current drain path preserves the
pending range on all sink errors and retries it before invoking the codec, so
it does not duplicate codec-consumed input or recycle dirty state.

## Primary source references

- `crates/soapberry-zip/src/writer.rs`: `ZipDataWriter`,
  `OwnedCompressedEntry`, `OwnedCompressor`, descriptor finalization, and
  compressed-limit accounting.
- `crates/soapberry-zip/src/office.rs`: streaming entry admission,
  poisoning, typed limit conversion, and output progress.
- `crates/soapberry-zip/src/preserve/replay.rs`: existing direct
  `flate2::Compress` loop with consumed/produced deltas and finish progress
  checks.
- flate2 1.1.10 `src/zio.rs` and `src/deflate/write.rs`: buffering, sync
  flush, finish, and drop behavior.
- zlib-rs 0.6.7 `src/stable.rs`: `Compress::reset`, total counters, and
  backend allocation/deallocation.
