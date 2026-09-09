# 0483 OPC foundation review

Status: read-only review of the current OPC foundation changes. The review
covered the preallocated fragment handoff, positional-read retry path, typed
execution-error conversion, and the independent preservation assertions in
[`source_part_splice.rs`](../../../../crates/litchi-opc/tests/source_part_splice.rs).
The obligations were taken from
[`adr-matrix.md`](adr-matrix.md), especially the requirements that the source
remain immutable, that preparation be authenticated before output, that ZIP
preservation remain an OPC/ZIP concern, and that no archive implementation
types cross the public semantic boundary.

## Verdict

I found no current production ownership or retry correctness blocker in the
reviewed implementation. The fragment lease is transferred exactly once, the
managed interrupted-read path does not double-charge input bytes, and bare
`litchi_core::ExecutionError` values remain typed when they cross an I/O or ZIP
adapter.

The earlier preservation-oracle gap is resolved in the current test source.
The focused integration test now has an independent central-record parser and
compares both local and central bytes for the untouched opaque member. No
current production or acceptance blocker remains in this reviewed scope.

## Fragment reservation and ownership

The new handoff follows the ADR closely:

* `SourcePartSpliceFragment` retains a private package reference, a fixed-length
  `Vec<u8>`, and one optional `Arc<Reservation>`
  ([`splice.rs:215-249`](../../../../crates/litchi-opc/src/source_backed/splice.rs#L215)).
  The public mutable view cannot resize the vector.
* `allocate_source_part_splice_fragment` checks the source and execution
  context before charging `Resource::Memory`, uses `try_reserve_exact`, refuses
  a capacity larger than the charge, and zero-fills in 64 KiB publication
  chunks while checking freshness/context and consuming `Resource::Work`
  ([`splice.rs:602-664`](../../../../crates/litchi-opc/src/source_backed/splice.rs#L602)).
  Every error path drops the local reservation through RAII.
* `prepare_source_part_splice_with_fragment` checks pointer identity against
  the exact allocating `SourceBackedPackage`, moves the `Vec` into the plan,
  and moves the existing reservation instead of reserving the capacity again
  ([`splice.rs:674-693`](../../../../crates/litchi-opc/src/source_backed/splice.rs#L674)).
  A foreign package cannot substitute an identical source or budget. A failed
  proof, limit, audit, or freshness check drops the reservation with the
  partially constructed plan.

The retained `Arc<Vec<u8>>` route still charges its reported capacity in the
ordinary path, while the owner-created route uses the already charged fixed
capacity. The focused validation artifacts record the foreign-package,
reservation-release, cancellation, and successful inverse cases as passing.
No double-charge or package-escape path was found. An allocator-specific test
that forces `try_reserve_exact` to report excess capacity would be useful as a
future regression guard, but is not a current correctness finding because the
runtime check is present before the fragment is returned.

## Interrupted reads, work, cancellation, and freshness

`read_source_at_with_context` reserves one positional input window, keeps that
reservation across retries that accepted no bytes, and commits only the final
accepted byte count ([`source_backed.rs:10953-11039`](../../../../crates/litchi-opc/src/source_backed.rs#L10953)).
An `Interrupted` result consumes one unit of managed `Resource::Work` before
the next attempt. `ExecutionContext::consume` checks cancellation itself, so a
source that cancels on the interrupted attempt stops before retry. The source
freshness check also runs before each retry and after a successful read once
publication monitoring is enabled. The verified-reader and publication paths
enable that monitor before their physical reads.

This gives the intended behavior:

* an interrupted attempt does not consume `InputBytes` twice;
* repeated interruptions consume finite managed work and can fail with the
  typed execution limit;
* cancellation wins before another physical read; and
* a source revision observed around a retry is carried through the I/O marker
  and restored to `OpcError::SourceChanged` by the OPC transport mapper.

The checked-in interrupted tests cover one retry, input accounting, and
cancellation. They do not include a source-version mutation between the
interrupted error and the retry. The code has the required freshness fence,
but that mutation case should be added as a regression test. For an unmanaged
source, an adapter that returns `Interrupted` forever remains able to retry
forever; the reviewed ADRs require periodic cancellation/freshness checks for
managed work but do not specify a finite retry count for an unmanaged
`ReadAt` provider, so I do not classify that compatibility behavior as a
0483 blocker.

## Bare execution-error conversion

`error.rs` recognizes both a bare public
`litchi_core::ExecutionError` in `std::io::Error` and OPC's private
`ExecutionIoError` marker before falling back to ordinary I/O. Cancellation
maps to `OpcError::Cancelled`; every other execution variant remains inside
`OpcError::Execution` ([`error.rs:365-404`](../../../../crates/litchi-opc/src/error.rs#L365)).
ZIP I/O conversion goes through the same mapper, and `SourceReader` wraps typed
execution failures with the private marker at its sequential-reader boundary
([`source_backed.rs:2659-2677`](../../../../crates/litchi-opc/src/source_backed.rs#L2659)).
This preserves the execution type instead of reducing a format callback's
failure to a display string. The existing focused error tests cover both the
private marker and a bare public execution value.

The later `From<OpcError> for litchi_core::Error` conversion intentionally
classifies a generic `OpcError::Execution(_)` as the umbrella `Other` variant;
that is separate from the OPC boundary mapping reviewed here and does not
change the typed `litchi_opc::Result` behavior.

## Public ownership boundary

The new public signatures expose `SourceBackedPackage`, `PackURI`, scalar
proofs, finite splice limits, the fixed fragment view, and caller-owned
`Write` sinks. `PreservationEntryId`, archive indexes, replay writers, and
compressor objects remain local to OPC publication. No raw ZIP type or archive
handle is present in the fragment, plan, or publication API. The test's direct
use of `StreamingArchiveWriter` is confined to fixture construction.

## Preservation oracle

`local_record` in
[`source_part_splice.rs:245-279`](../../../../crates/litchi-opc/tests/source_part_splice.rs#L245)
manually locates a ZIP member through the EOCD and central-directory fields,
then returns the complete local header, compressed payload, and optional data
descriptor. The descriptor-mode test is useful and independent: it verifies
that a signature-like byte sequence inside compressed data is not mistaken for
the end of a descriptor.

The current `central_record` helper manually walks the central directory,
returns the complete selected record, and zeroes only the relative local-header
offset at bytes 42..46. It rejects ZIP64 offsets and the generated fixture
requires a comment-free ZIP32 EOCD, making the fixture scope explicit
([`source_part_splice.rs:281-303`](../../../../crates/litchi-opc/tests/source_part_splice.rs#L281)).
The negative oracle test mutates external attributes, which are central-only,
and proves that the local record remains equal while the central record differs;
it then changes only the relocation field and proves the normalized records
remain equal ([`source_part_splice.rs:305-329`](../../../../crates/litchi-opc/tests/source_part_splice.rs#L305)).
The Store and Deflate splice test now compares both complete local and complete
central records for the untouched opaque member
([`source_part_splice.rs:582-631`](../../../../crates/litchi-opc/tests/source_part_splice.rs#L582)).
This closes the prior concern without depending on `PreservationIndex` for the
oracle. The existing ZIP64 and broader topology suites cover their own wider
archive shapes.

Optional strengthening would compare every untouched central record and the
EOCD comment/tail in this focused fixture, plus put an untouched member after
the changed target to exercise a real offset relocation. Those would improve
coverage, but the current independent full-record comparison and metadata
negative test are sufficient for this scoped ZIP32 fixture and are not a
0483 blocker.

## Validation evidence

No Cargo, test, lint, or profile command was run for this review, per the
coordinator's serialized validation instruction. I inspected the recorded
validation artifacts only. They show the focused fragment, interrupted-read,
oracle, and error-mapping attempts passing after the initial fragment-test
retry, and the recorded final OPC run reports all 344 unit/integration tests
plus 28 `source_part_splice` cases passing. The newly inspected central-record
changes close the former oracle gap; this review itself still did not execute
those gates.
