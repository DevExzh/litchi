# Change 0259: shared lazy OPC structural members

## Status

Landed in `2a2baf5af`. The change is accepted for private ownership-shape
correctness and removal of one deterministic payload clone. It has no latency,
allocation-count, peak-memory, RSS, or end-to-end performance claim.

## Scope

`litchi-opc` reads `[Content_Types].xml` and relationship manifests through the
private `read_structural_member` seam. A deflated member in
`soapberry_zip::office::LazyArchiveReader` was already decompressed into the
reader's cached `Arc<Vec<u8>>`, but the generic `ArchiveAccess::read` path then
cloned that complete `Vec<u8>` before XML parsing.

The private archive adapter now has an optional shared-read hook. Only
`LazyArchiveReader` implements it, delegating to the existing validated
`read_shared` path. `StructuralMember` can retain that `Arc<Vec<u8>>` for the
duration of structural parsing. The selection order is deliberately:

1. borrow a validated stored member directly from the source;
2. retain the lazy reader's shared materialization for a deflated member;
3. use the existing owned `Vec<u8>` fallback.

`IndexedArchive` keeps the default `None` shared hook and therefore remains on
the owned fallback. Positional/file/remote sources do not acquire a new `Arc`
allocation or an unsound borrowed lifetime. The trait, enum, and handoff remain
crate-private; no facade, format-owner, dependency, or public API changes.

## Preserved contracts

The shared lazy path is the same `LazyArchiveReader::read_shared` operation
previously reached indirectly by `read`. Decompression, CRC verification,
declared and materialized size checks, ZIP limits, cache reservations,
single-flight behavior, and error propagation are unchanged. Stored-member
borrowing still bypasses the cache. Session/cancellation code and all OPC
catalog, relationship, content-type, unknown-member, source, edit, patch, and
publication behavior are untouched.

Focused tests bind the ownership variants rather than infer them from timing:

- a deflated lazy member is `StructuralMember::Shared`, and `Arc::ptr_eq`
  proves it is the cache's exact allocation;
- a stored lazy member is `StructuralMember::Borrowed`, points into the exact
  validated source slice, and leaves the decompression cache empty;
- an `IndexedArchive` member is `StructuralMember::Owned`.

## Verification

- `cargo test -p litchi-opc --lib`: 227 passed.
- `cargo test -p litchi-opc --tests`: all passed.
- `cargo test -p soapberry-zip --lib`: 257 passed.
- `cargo clippy -p litchi-opc --all-targets -- -D warnings`: passed.
- `RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc --no-deps`: passed.
- Rustfmt and `git diff --check`: passed.
- Independent review found no public-surface, limit, CRC, error-order,
  cancellation, cache, panic, unsafe-code, lint, rustdoc, or ADR blocker.

## Claim boundary and follow-up

This change establishes the exact private ownership handoff and the absence of
the former second `Vec` allocation/copy for eligible lazy deflated structural
members. It does not quantify whole-package allocation calls, copied bytes,
peak memory, RSS, latency, throughput, decompressed bytes, physical I/O, cold
or remote behavior, or any DOCX/PPTX/XLSX semantic operation.

The existing synthetic `opc_open` corpus has few relationship manifests and is
not sufficient to measure this seam. Any performance claim requires a fixed
relationship-heavy corpus plus clean CPU-pinned A1/B1/B2/A2 and resource
evidence; ordinary `opc_open`, `opc_open_owned`, and source-backed opens remain
guardrails.
