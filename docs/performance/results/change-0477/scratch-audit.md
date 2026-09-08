# 0477 scratch and storage capability audit

This is a read-only source audit for the next streaming implementation slice.
It records the capability boundary, retained owners, and failure contract for
replaying ZIP central records into a sequential sink. It does not claim that
the streaming memory requirement is implemented or measured.

The accepted ADR inventory and unchanged hashes are recorded in
[`adr-refresh.json`](adr-refresh.json). The governing goal is
[`docs/GOAL.md`](../../../GOAL.md), especially its requirements for explicit
execution budgets, no ambient filesystem or runtime dependencies, bounded
streaming windows, and typed failures.

## Constraints from the accepted records

ADR0005 makes scratch storage an explicit capability. Supported providers are
memory, encrypted temporary storage, and caller-defined stores; plaintext
temporary files must not appear implicitly, and absent capability must produce
a typed resource error (`docs/adr/0005-io-memory-and-performance.md:8-28`).
The same record permits a finalized document to target a sequential,
non-seekable sink when layout is planned or explicit scratch is supplied, and
requires caller-owned non-atomic sinks to report incomplete output and accepted
bytes (`docs/adr/0005-io-memory-and-performance.md:32-41`).

The record's later package-scale amendment requires fallible staging and says
that retained memory and work are charged to existing transaction budgets; it
does not turn a local index or spool into a package-wide peak-memory claim
(`docs/adr/0005-io-memory-and-performance.md:2189-2231`). Any new central
spool therefore needs an explicit byte ceiling and separate evidence for the
fixed replay window, retained indexes, and provider storage.

The crate boundary rules keep `litchi-core` neutral. It owns sources, budgets,
execution, and diagnostics, but no ZIP or XML dependency
(`docs/adr/0002-crate-topology.md:8-10`, `:1042-1048`). The facade must not
expose archive types (`docs/adr/0001-priorities-and-api-layers.md:14-35`),
`soapberry-zip` owns ZIP grammar through its format-owner boundary
(`docs/adr/0010-facade-archive-ownership.md:14-75`), and `litchi-opc` owns
physical OPC packaging and translation to the selected ZIP implementation
(`docs/adr/0011-ooxml-physical-package-ownership.md:15-52`). These constraints
favor a ZIP-owned capability forwarded by OPC rather than a new core trait in
the first slice.

## Existing capability and retained owners

`litchi-core::ReadAt` is an immutable positional input source and requires
`Send + Sync`; it cannot append or replay central records
([`source.rs:21`](../../../../crates/litchi-core/src/source.rs#L21)). Its
`Resource` enum contains memory, input/output bytes, objects, depth, and work,
but no scratch or retained-index resource
([`budget.rs:18-34`](../../../../crates/litchi-core/src/budget.rs#L18-L34)).
`ExecutionContext` can reserve and consume those existing resources
([`execution.rs:180-260`](../../../../crates/litchi-core/src/execution.rs#L180-L260)),
but introducing a ZIP-specific store there would add a dependency direction
and public abstraction before another crate needs it.

The ZIP implementation currently retains the central directory in two vectors:
`ZipArchiveWriter<W>` owns `Vec<FileHeader>` and concatenated name bytes
([`writer.rs:184-190`](../../../../crates/soapberry-zip/src/writer.rs#L184-L190)).
`finish` scans those headers for ZIP64 state and emits records in insertion
order only at finalization ([`writer.rs:1452-1538`](../../../../crates/soapberry-zip/src/writer.rs#L1452-L1538)).
The scratch path can serialize each finalized record immediately, then retain
only entry count, central byte count, output offset, and ZIP64 state.

The higher streaming wrapper has a separate growing ZIP-name index:
`StreamingArchiveWriter` and `StreamingArchiveEntry` carry
`HashSet<String>` values ([`office.rs:5326-5361`](../../../../crates/soapberry-zip/src/office.rs#L5326-L5361)).
Duplicate validation and insertion occur at
[`office.rs:5880-5955`](../../../../crates/soapberry-zip/src/office.rs#L5880-L5955)
and its fallible reservation currently reports an `InvalidInput` message at
[`office.rs:6025-6032`](../../../../crates/soapberry-zip/src/office.rs#L6025-L6032).
A central spool does not bound this index.

OPC has a second semantic index. `PartNameSet` stores folded full names and
ancestor/descendant values in two hash maps
([`phys_pkg.rs:946-1038`](../../../../crates/litchi-opc/src/phys_pkg.rs#L946-L1038)).
It rejects exact and ASCII-equivalent duplicates as well as ancestor and
descendant conflicts. ZIP central records cannot replace those checks. ODF
also retains manifest and member indexes in its own writer
([`writer.rs:593-611`](../../../../crates/litchi-odf-common/src/core/writer.rs#L593-L611)).
Those are separate follow-up owners.

There are useful bounded scratch precedents, but none is the required
replayable store. Preservation accepts caller-provided `&mut [u8]`
([`office.rs:2962-2978`](../../../../crates/soapberry-zip/src/office.rs#L2962-L2978));
the private `CompressedScratch` bounds one deflated payload
([`office.rs:5278-5304`](../../../../crates/soapberry-zip/src/office.rs#L5278-L5304)).
Neither can append finalized records and seek back over them. The filesystem
helper in OPC is a durable publication mechanism and must not become an
implicit scratch provider.

## Recommended first interface

Keep the ordinary constructors and defaults unchanged. Add an advanced
low-level path in `soapberry-zip` that accepts a caller-selected replay store;
thread it through an advanced `PhysPkgWriter` entry point if OPC needs it.
The ZIP implementation must not open a path, call `tempfile`, select a global
runtime, or silently fall back to plaintext filesystem storage.

The minimal provider contract can be an archive-owned object-safe trait built
from synchronous `Read + Write + Seek` operations. The writer should record the
provider's starting end offset, append only its own bytes, enforce a checked
maximum before each record, and replay that range through one fixed-size read
window. Requiring the store to be initially empty would be unnecessarily
restrictive; an owned generation offset avoids stale bytes and avoids adding a
`truncate` requirement to every caller-defined provider.

The existing low-level writer can carry a private optional boxed provider, but
its auto-trait contract must be preserved exactly. The provider object in that
unchanged `ZipArchiveWriter<W>` path must be held as
`Box<dyn CentralDirectoryStore + Send + Sync>` (or inside a private wrapper
with those same auto traits). `ZipArchiveWriter<W>` currently derives
`Debug` and stores the output writer, central vectors, and reusable Deflate
state ([`writer.rs:184-190`](../../../../crates/soapberry-zip/src/writer.rs#L184-L190)).
The existing type is conditionally `Send + Sync` when its `W` and remaining
fields permit it. A bare `Box<dyn CentralDirectoryStore>` would be neither
`Send` nor `Sync`, so merely adding that field would regress callers whose
`W` is thread-safe. `Send + Sync` bounds on the boxed provider are therefore a
compatibility requirement for the unchanged type, even though the write
operation itself is synchronous.

The provider trait itself need not require `Debug`. Because the existing
writer derives `Debug`, put the box behind a private redacted wrapper with a
manual `Debug` implementation, or implement `ZipArchiveWriter`'s `Debug`
manually. The debug form should expose only provider presence or bounded
configuration, never provider paths, bytes, or lower-layer diagnostic text.
Do not add a public `Debug` bound merely to make the derive compile.

The boxed path necessarily favors an owned provider (normally `'static`) and
providers that are safe to move and share at the type level. That covers an
owned `Cursor<Vec<u8>>`, an owned encrypted temporary-store adapter, and other
caller-defined stores with the required auto traits. A caller with a
single-threaded or borrowed `Read + Write + Seek` store must use a separate new
generic wrapper, for example an explicitly named
`SpoolingZipArchiveWriter<W, S>`. That wrapper may leave `S` unconstrained in
its declaration; its own `Send + Sync` behavior then naturally depends on `S`.
Do not weaken the existing `ZipArchiveWriter<W>` to support that case, and do
not add a second generic parameter to every existing streaming entry type just
for this provider.

Record serialization should write directly to the spool at entry finalization
so the old `files` and `file_names` vectors are absent in the spooled mode.
The central record includes fields known only after payload completion: CRC,
compressed and uncompressed sizes, local-header offset, flags, name, extras,
and any ZIP64 values. At finalization, seek to the generation start and copy
records through a fixed replay window before writing ZIP64 and EOCD. Keep
central-record ordering and the existing local/data-descriptor behavior.

The existing `StreamingArchiveLimits.max_metadata_bytes` is an aggregate
variable-metadata ceiling ([`office.rs:4795-4880`](../../../../crates/soapberry-zip/src/office.rs#L4795-L4880));
it is not a fixed-memory promise and does not include every fixed central
header byte. Add a distinct checked ceiling for serialized central spool bytes
(for example, `max_central_directory_bytes` or a
`CentralDirectorySpoolBytes` resource). Keep the fixed replay-window
reservation separate from that cumulative spool total. An in-memory provider
may still retain the whole spool; only an explicitly selected external or
encrypted provider can make the retained central bytes leave the process.

## Typed failures and mapping

The ZIP error model already has typed allocation and limit variants plus
underlying I/O sources ([`errors.rs:35-124`](../../../../crates/soapberry-zip/src/errors.rs#L35-L124),
[`errors.rs:209-236`](../../../../crates/soapberry-zip/src/errors.rs#L209-L236)).
Extend it with a storage operation that identifies append, seek, or replay and
with a distinct spool-byte limit resource. A practical shape is a
content-free `ScratchIo { operation, source }` plus the existing structured
limit form, rather than mapping provider failures to generic `IO`. The current
generic I/O display forwards the underlying text; a scratch-specific display
must omit paths, provider identity, and content while retaining the original
error through `source()` or an explicit accessor for callers that request it.

The failure stages should remain distinguishable:

| Stage | Required result |
| --- | --- |
| Strict mode has no provider | Typed scratch-capability/resource error before output begins. The ordinary fallback mode may retain its existing vectors. |
| Record would exceed the configured spool ceiling | Typed limit with resource, observed bytes, and maximum; do not publish that central record. |
| Provider append, seek, short write, short read, or replay fails | Typed scratch operation error with a redacted default display and preserved source. |
| Fixed replay-window allocation fails | Existing typed allocation error with a stable scratch/replay resource name. |
| Output sink fails after bytes were accepted | Existing streaming progress and `IncompleteOutput` semantics, preserving accepted byte count and the scratch error as its source where applicable. |

`soapberry_zip::Error` is translated by OPC today, while
`PhysPkgWriter<W>` owns the physical archive and delegates finalization
([`phys_pkg.rs:1064-1081`](../../../../crates/litchi-opc/src/phys_pkg.rs#L1064-L1081),
[`phys_pkg.rs:1294-1303`](../../../../crates/litchi-opc/src/phys_pkg.rs#L1294-L1303)).
The mapping should add an OPC storage/central-spool resource rather than
pretending it is a read limit. `PackageWriter::write_to_stream` already counts
accepted sink bytes and returns typed incomplete output
([`pkgwriter.rs:890-947`](../../../../crates/litchi-opc/src/pkgwriter.rs#L890-L947));
that behavior must survive central replay failures.

## Boundary and implementation sequence

The first implementation should stay in `soapberry-zip`, with OPC forwarding
only the explicit capability and mapping errors. No scratch type should appear
in ordinary `litchi` CRUD signatures or in the neutral core. Preserve the
current `StreamingArchiveWriter<W>` and `StreamingArchiveEntry<W>` defaults
([`office.rs:5326-5361`](../../../../crates/soapberry-zip/src/office.rs#L5326-L5361));
the boxed provider inside the low-level owner avoids propagating a storage type
through every downstream generic. If a non-`Send + Sync` or borrowed provider
is required, add a separate generic advanced wrapper and keep it below the
ordinary facade.

Implement and verify in this order:

1. Add the private redacted provider box and central-record spool accounting,
   retaining conditional `Send + Sync` for the existing writer and no provider
   `Debug` bound.
2. Serialize records at entry completion, replay them with a fixed window, and
   remove the central vectors only in the explicit spooled mode.
3. Add exact small archives, ZIP64, ceiling, short-read/write, seek failure,
   provider failure, output partial-failure, and deterministic-order tests.
4. Thread the capability through `PhysPkgWriter` without changing ordinary
   package signatures or OPC name validation.
5. Address ZIP duplicate names and OPC `PartNameSet` separately. They require
   exact, collision-safe indexes or a closed generated-name proof; central
   spooling alone does not bound them.
6. Measure fixed replay memory, provider bytes, ZIP/OPC retained indexes, and
   output progress separately. Do not report a constant-memory or
   package-wide claim until all retained owners are covered.

No build or benchmark was run for this audit.
