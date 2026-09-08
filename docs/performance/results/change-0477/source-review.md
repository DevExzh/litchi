# 0477 source review: explicit central-directory spool

Status: source review closed against the current coder diff. The optional
spool path is functionally ready for the root test and evidence gates. No
production correctness blocker remains in the reviewed ZIP or OPC paths.

This is a bounded source review, not a test or benchmark result. The root
agent owns the build, focused tests, capture freeze, and evidence receipts.

## Capability and ownership

`soapberry-zip::DirectorySpoolLimits` gives the caller separate serialized
central-directory and replay-window limits. `DirectorySpool` accepts only an
owned `Read + Write + Seek + Send + Sync + 'static` provider, seeks to the
provider's end to establish a generation offset, and never opens a path or
creates ambient temporary storage. The optional `DirectorySpool` is now behind
`Option<Box<_>>` in `ZipArchiveWriter`, so the ordinary writer avoids carrying
the spool's large inline state; the opt-in path pays its allocation explicitly.
The default writer still carries the two lifecycle state flags required for
fail-closed poisoning, so this is a spool-footprint reduction rather than a
claim that the complete pre-change struct layout is unchanged.

The spool retains only its admitted extent, entry count, ZIP64 observation,
one fixed replay buffer, and one serialized record at append time. Spooling
does not claim package-wide constant memory: the ordinary ZIP name index in
`StreamingArchiveWriter` and the OPC `PartNameSet` remain retained owners.

The provider contract is deliberately exclusive and coherent. The caller
must not mutate or reposition the provider, including through aliases of its
backing storage, until the archive is finished. The implementation does not
hash the records to detect an alias mutation. This is an explicit capability
contract rather than an unchecked provider assumption; checksum verification
would be a separate hardening step if a caller cannot provide that contract.

## Record, ZIP64, and limit correctness

Each finalized `FileHeader` is serialized as one complete central record in
emission order. The record length uses the finalized **central** extra-field
size, so local-only fields do not consume the central spool budget. Sized
stored and precompressed routes finalize central metadata before capacity
preflight. Directory and streaming routes account for timestamps, explicit
ZIP64 size fields, and offset-only ZIP64 fields before publication, then run
the exact finalized-record check. This covers the prior directory-at-
`u32::MAX` offset omission and keeps regular bytes identical between the
spooled and in-memory routes.

`finish` replays exactly the admitted generation extent, validates the
`SeekFrom::Start` return position, and derives ZIP64 tail state from the
record count, directory offset/size, and per-record ZIP64 state. It rejects
zero or over-reporting reads/writes and retries `Interrupted` transfers.
The focused test file includes exact and one-under ceilings, ZIP64 count and
offset parity, short and interrupted transfers, wrong seek returns,
truncation, zero transfers, and local-only extra fields.

## Lifecycle and failure publication

`ZipArchiveWriter` now has archive-level poisoning. Sink, central-record,
allocation, count, and finalization failures prevent later routes and
`finish` from publishing a second archive. A borrowed entry sets the pending
state only after its local header succeeds, keeps that state on drop or failed
descriptor/central publication, and clears it only after successful central
publication. All sized, directory, borrowed, and owned entry starts check the
pending and poisoned states.

Spool I/O failures preserve a typed operation and underlying `source()` while
their default ZIP and OPC displays omit provider paths and error text. Output
failures continue to report accepted output progress through the streaming
failure types. `Interrupted` remains retryable; zero-byte and over-reporting
transfers are treated as failures. The explicit `Send + Sync` provider bound
keeps the existing writer auto-trait behavior, and the compile-time/test
assertions cover the spooled writer shape.

## OPC boundary

`PhysPkgWriter::with_writer_and_metadata_spool` forwards only the explicit
ZIP capability and maps central-spool I/O and limit variants to typed
`OpcError` values. Existing exact, ASCII-equivalent, and ancestor/descendant
OPC name validation remains in `PartNameSet`; the adapter makes no claim that
central spooling bounds that index. Prior output progress is wrapped as
`IncompleteOutput` when a spool failure occurs after publication begins.

The remaining verification obligations are operational: root must complete
the focused suite and freeze the capture source before measurement. The source
review found no additional concrete correctness defect after the outer spool
boxing and final lifecycle/ZIP64 fixes.
