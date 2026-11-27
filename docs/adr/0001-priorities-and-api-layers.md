# ADR 0001: Priorities and API layers

- Status: Accepted
- Date: 2026-07-31

## Context

Litchi must support the full range of Office CRUD workflows without losing
unrecognized content or forcing developers to understand package IDs, XML
relationships, BIFF records, locking wrappers, or runtime-specific I/O. The
existing monolithic format crates mix those concerns and expose APIs that can
rebuild an opened document from an empty writer model.

Backward compatibility is not a constraint for this refactor.

## Decision

The design is evaluated on five dimensions: ergonomics, safety, measured
performance, modularity, and production readiness. Correctness and safety win
all conflicts. The common path stays concise; specialized zero-copy or low-level
paths are explicit.

There are three strict public layers:

1. `litchi` is the ergonomic facade. Its root contains only format detection,
   the concrete `File` enum, and universal basics. Normal APIs live under short
   modules such as `litchi::docx`, `litchi::pptx`, and `litchi::xlsx`.
2. Concrete format and semantic-vocabulary crates expose complete, typed Office
   models without container details in ordinary signatures.
3. Narrow `raw` modules and low-level container/common crates expose validated,
   lossless XML, record, part, and relationship operations.

Raw types never leak into ordinary CRUD signatures. Public low-level APIs are
supported APIs, not unrestricted mutable byte/XML escape hatches: reading is
lossless, construction is validated, and mutation is tracked.

## Required properties

- No public operation deliberately panics. Collection misses are `Option`; bad
  syntax, ambiguity, I/O, invalid state, and unsafe edits are typed `Result`s.
  Unavoidable process failures such as allocation aborts are not modeled as
  recoverable document errors.
- No public `Index`/`IndexMut` facade and no `*_unchecked` API.
- No public `Arc<RwLock<T>>`, source generic, runtime handle, or package ID in a
  normal document/worksheet/slide signature.
- Small immutable handles may be `Clone`; potentially expensive duplication is
  named `duplicate`, `detach`, or `to_owned` and is budgeted.
- High-level crates forbid unsafe code. Audited SIMD, mmap, and zero-copy unsafe
  code is isolated in narrowly scoped low-level crates with tests and safety
  arguments.
- Unsupported content remains readable and preservable. Unsupported editing is
  a typed capability error, never a silent approximation.

For the IWA chart seam, this means the public reference-line model is a
validated semantic value rather than an archive record: finite custom values,
bounded names, packed visibility, and checked lossless future kinds belong to
the common layer. Protobuf extensions, object ownership, graph ordering, and
package transactions remain below the facade boundary.

The same boundary applies to drawable geometry. Fixed-size document points,
sizes, and the horizontal/vertical flip selector live in
`litchi_iwa_common::shape::geometry`; the full IWA `DrawableGeometry` adapter
retains optional-field presence, native reflection flags, rotation conventions,
and protobuf/wire mutation because those details are not neutral semantic
vocabulary.

## Consequences

The current API may be replaced outright. Migration scaffolding is internal to
the feature branch and is not a compatibility promise or a released shim.
