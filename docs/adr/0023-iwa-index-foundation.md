# ADR 0023: Archive-free IWA object-index foundation

- Status: Accepted
- Date: 2026-08-06

## Context

The existing IWA object index combines package traversal, archive decoding,
protobuf interpretation, native object identifiers, byte locations, and
reference extraction in `litchi-iwa`. That coupling makes the index expensive
to compile and makes its archive concepts difficult to reuse from the future
Pages, Numbers, and Keynote crates. It also leaves raw physical identity close
to ordinary lookup APIs.

## Decision

Introduce `litchi-iwa-index` as a neutral leaf. Its only internal workspace
dependency is `litchi-iwa-graph`; it does not depend on `litchi-iwa`, a package
crate, ZIP, Snappy, protobuf bindings, or `IWorkPackage`.

The public model is deliberately small:

- `FragmentId` is a non-zero adapter-local ordinal, never a package entry name
  or native archive identifier.
- `ByteSpan` stores checked half-open `u64` endpoints and rejects overflow or
  reversed ranges before publication.
- `ObjectRecord` combines a graph `ObjectId`, a `FragmentId`, and a
  `ByteSpan`; it owns no payload or archive handle.
- `Reference` contains two validated graph identities and has an explicit
  nullable-boundary constructor that reports typed source/target errors.
- `IndexBuilder` reports duplicate fragments, duplicate objects, duplicate
  references, missing endpoints, and allocation failures through
  `IndexError`.
  `build` requires both endpoints to be indexed; the explicit
  `build_allow_missing_targets` adapter path permits an absent target while
  continuing to require an indexed source.
- `ObjectIndex` is immutable. It stores sorted boxed slices, exposes borrowed
  lookups and deterministic iteration, and provides incoming, outgoing,
  reachability, and cycle queries through an immutable graph snapshot. It does
  not implement `Index` or `IndexMut`.

Native unknown fields and payload bytes remain with the format adapter. The
neutral index stores only location and graph metadata, so constructing it does
not discard unsupported content.

The dangling-target build path preserves a graph edge even when the target has
no `ObjectRecord`. Graph queries can therefore report the reference and its
reachable identity, while object lookup remains absent for the target. This
keeps incomplete archive sets observable without making the neutral leaf aware
of archives or package state.

## Adapter boundary

The IWA integration is implemented as a private adapter in
`litchi-iwa::object_index`. It validates each archive object, assigns private
`FragmentId` ordinals, translates checked locations into `ObjectRecord`, and
translates native references into `Reference` values. Package traversal,
protobuf decoding, fallback reference extraction, resource budgets, unknown
field preservation, and transactional publication remain below the public
semantic API. No native identifier or archive type is added to this crate to
make the adapter convenient.

The adapter retains only the archive name and validated source position needed
to resolve an already-parsed object. Public entry metadata exposes typed
`FragmentId` and `ByteSpan` values; native archive names, source positions, and
raw numeric compatibility queries remain private to the adapter.

## Consequences

The index leaf is cheap to build and incrementally compile, deterministic to
inspect, and safe to share across format owners. Final lookups use binary
search over compact immutable storage; temporary builder hash sets are dropped
after `build`. The adapter still owns the cost of reading and decoding IWA
archives, and the follow-up integration must preserve that ownership rather
than reintroducing archive dependencies into the leaf.
