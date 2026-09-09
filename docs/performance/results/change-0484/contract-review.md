# Change 0484 contract review — live refresh

Status: **review cleared; bundle unsealed**. This is a bounded, read-only security and ownership
review of the current DOCX replayable-stream and OPC replay sources. I changed
this review document only; no production code or Git state was changed.

The three earlier production findings are resolved in the current source.

* A durable apply authenticates the current raw archive before resolver use and
  rechecks the package context after the resolver. [`ResolvedInput::seal`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:1440)
  binds the fresh handle to that package context. [`BoundReplayHandle::open`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:1265)
  checks the context before and after opening, uses
  `open_for_package` when replacing a retained provider context, and returns a
  [`BoundReplayReader`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:1125)
  that checks before and after bounded reads, charges package `Work`, and
  checks before and after terminal `finish`. The memory handle supplies the
  current context to its reader and suppresses its old `Work` charge on a
  reopened package. The resolver-cancellation test observes typed cancellation
  with zero replay opens and zero output.
* [`MemoryReplayStore::prepare_for_operation`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:810)
  rejects a store larger than the selected replay ceiling, reserves its
  retained bytes and owner object before the producer, and checks exact
  capacity. [`MemoryReplayStore::finish`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:870)
  moves the single `Vec` into `Arc<MemoryReplayStorage>`; it does not create a
  second payload copy. [`MemoryReplayHandle::open_reader`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:753)
  reserves one `Objects` term for every live reader. The default
  `AuthoredReplayStore` hook remains explicitly caller-owned storage and does
  not pretend that an arbitrary external allocation is package-accounted.
* [`StreamingParagraphEncoder::new_with_cursor_reservations`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream/encoder.rs:371)
  and [`EncodingReader::new`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream/encoder.rs:720)
  reject a buffer whose actual capacity differs from the reserved replay
  window. The OPC adapter has the same equality fence and always owns a fresh
  per-pass window; an optional retained-storage lease only covers caller-owned
  replay bytes.

The public default policy is usable: `ParagraphStreamLimits::default()` has a
finite 64 KiB replay window and finite 512 MiB explicit-retention ceiling, and
the normal tiny append passes that policy. The 512 MiB value is a selected
ceiling, not an allocation made by constructing a store; callers choose the
`MemoryReplayStore::new` ceiling explicitly. Durable references copy one
bounded token into immutable shared storage and validate the selected patch
limit at every binding boundary.

The raw durable identity is also complete. Forward apply binds the recorded
source archive length and SHA-256 through [`ensure_current_archive`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:1852),
then reruns semantic and physical preparation. It does not serialize or trust
the process-local source version as durable identity. The expected-artifact OPC
publication previews the complete candidate archive before touching the caller
sink, so a stale candidate or replay proof fails with zero caller output.
Exact inverse authentication checks the current length before hashing and
opens the explicit original provider only after current identity succeeds.

## Callback error and cancellation ordering

[`ParagraphStreamPatch::resolve_replay`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream/patch.rs:474)
maps a resolver `Err` directly to the typed `PatchError::ReplayResolver`
variant. The later package-context check runs when the resolver returns a
handle, so a resolver that cancels and returns a valid handle is refused before
any reader or output, while a terminal resolver error keeps its provider
provenance. On that error branch,
[`apply_tail_append_stream_patch`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs:1794)
reauthenticates the current archive before returning, so a simultaneous source
change wins over the resolver error while a stable source preserves the typed
provider error. [`apply_exact_inverse`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream/patch.rs:861)
uses the same ordering when the original provider mutates the current source
and returns an error. During physical publication, the OPC final source fence
still gives a simultaneous source change priority over a callback error and
wraps accepted output with the exact `IncompleteOutput` count.

The focused regressions
`durable_resolver_error_yields_source_mismatch_when_resolver_mutates_current`
and the mutating original-provider inverse case both assert the typed source
error and an untouched sink. No remaining security, ownership, or
cancellation-precedence blocker was found; the old stale-context,
missing-object-lease, over-capacity, and resolver-fence findings no longer
describe the live implementation.

## Current evidence

The unchanged-source dev58 receipt reports 40 existing DOCX tail tests and 21
stream tests passing. The stream set includes zero-object refusal, selected
replay-capacity refusal before the producer, managed memory refusal and lease
release, per-reader object retention, durable resolver cancellation with no
opens/output, proof/reference mismatch, durable reopen, and expected-artifact
publication. The unchanged-source dev61 receipt reports all-targets,
all-features Clippy with warnings denied passing. The later unchanged-source
dev68 and dev69 receipts report the full all-features and no-default-features
DOCX suites passing (1,408 and 1,389 tests respectively, with 31 ignored in
each); both include the resolver and inverse source-precedence regressions.

The bundle remains unsealed pending the parent’s final integrated validation
and source-custody decision. No heavy checks were run for this review.
