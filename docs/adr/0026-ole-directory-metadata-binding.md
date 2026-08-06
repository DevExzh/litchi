# ADR 0026: Shared OLE directory metadata binding

- Status: Accepted
- Date: 2026-08-06
- Scope: `litchi-ole-common::object`

## Decision

The common OLE object owner projects the validated CFB directory fields needed
by DOC, PPT, and XLS into `object::directory::{Metadata, EntryKind, Links,
Sid}`. `Storage::directory()` always exposes typed storage metadata and
`Stream::directory()` exposes the captured stream metadata without copying its
payload. CLSIDs reuse the common Property Set `Guid` value, while `NOSTREAM`
links become `Option<Sid>`.

The projection covers directory identity, object kind, class identifier,
red-black-tree sibling/child links, starting sector, parsed stream size, and
MiniFAT placement. It does not activate an OLE class, interpret an object
payload, or duplicate CFB sector/tree validation already owned by
`litchi-cfb`.

## Validation and publication

The directory codec rejects unknown object kinds, invalid regular SIDs,
self-referential links, stream child/CLSID fields, storage stream-only fields,
and invalid root links. Existing `Limits` remain the operation ceilings for
directory traversal, stream count, nesting, and retained bytes; the metadata
itself is a small `Copy` value and adds no payload allocation.

Every staged object edit renders and reopens the candidate CFB before
publishing it. The editor now assigns that freshly captured package to its
internal state before rebuilding the public object catalog, so sector, size,
and MiniFAT metadata cannot remain stale across a subsequent snapshot edit.
Failed edits still operate on a cloned candidate and leave the source
snapshot, bytes, and metadata unchanged.

## Authority and non-goals

The field mapping follows [MS-CFB §2.6 directory entries](../../3rdparty/specs/[MS-CFB]/2%20Structures/2.6%20Compound%20File%20Directory%20Sectors.md),
including the `NOSTREAM` terminator, object kinds, CLSID rules, stream size,
and MiniFAT distinction. The inert object boundary remains compatible with
[MS-OLEDS §2.3 OLE2.0 structures](../../3rdparty/specs/[MS-OLEDS]/2%20Structures/2.3%20OLE2.0%20Format%20Structures.md),
while property-set names and values remain owned by the existing
[MS-OLEPS binding](../../3rdparty/specs/[MS-OLEPS]/2%20Structures/2.24%20Standard%20Bindings.md)
and semantic codec. Shared toolbar and envelope vocabulary remains under
[MS-OSHARED](../../3rdparty/specs/[MS-OSHARED]/2%20Structures/2.3%20Common%20Objects.md).

State bits and timestamps are intentionally not fabricated: the current
`litchi-cfb::DirectoryEntry` API does not expose those raw fields. Adding
activation, host classification, or format-specific `CompObj`/`ObjInfo`
semantics belongs in a later owner-layer migration.
