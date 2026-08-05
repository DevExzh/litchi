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
- `ObjectIndex` is immutable. It stores sorted boxed slices, exposes borrowed
  lookups and deterministic iteration, and provides incoming, outgoing,
  reachability, and cycle queries through an immutable graph snapshot. It does
  not implement `Index` or `IndexMut`.

Native unknown fields and payload bytes remain with the format adapter. The
neutral index stores only location and graph metadata, so constructing it does
not discard unsupported content.

## Adapter boundary

The IWA integration is intentionally deferred as a private follow-up. A future
`litchi-iwa` adapter will validate each archive object, assign private
`FragmentId` ordinals, translate checked locations into `ObjectRecord`, and
translate native references into `Reference` values. Package traversal,
protobuf decoding, fallback reference extraction, resource budgets, unknown
field preservation, and transactional publication remain below the public
semantic API. No native identifier or archive type is added to this crate to
make that adapter convenient.

## Consequences

The index leaf is cheap to build and incrementally compile, deterministic to
inspect, and safe to share across format owners. Final lookups use binary
search over compact immutable storage; temporary builder hash sets are dropped
after `build`. The adapter still owns the cost of reading and decoding IWA
archives, and the follow-up integration must preserve that ownership rather
than reintroducing archive dependencies into the leaf.
