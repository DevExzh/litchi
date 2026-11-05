# 0360en — XLSB admitted physical-reader handoff

- Status: source and static checks complete on `uo2/w4b-xlsb-raw-policy`;
  focused Cargo validation and native performance qualification remain open.
- Base: `947545d9d3` (0360em replay copy admission).

## One admitted physical index

Structural and generated-image replay first admits the recipient image with a
borrowed `PhysPkgReader`. The reader is now consumed by
`OpcPackage::from_shared_vec_with_admitted_reader`, so the deferred package
does not rebuild the ZIP central-directory index. The image remains one
`Arc<Vec<u8>>`; no retained source slice was added to ordinary physical
readers. Source identity is proved through the backing
`LazyArchiveReader::as_bytes()` view and a pointer/length check before any
package publication. The reader's admitted `SourceLimits` become the package
policy.

The handoff is used by both the structural `changed_candidate` path and the
durable replay path. Recipient byte, archive, member, and declared-size
admission still precede ownership, while workbook projection and worksheet
validation remain required before publication.

## Regression coverage

The OPC tests cover a reader borrowed from a different allocation returning a
typed `SourceArtifactMismatch`, and a matching reader retaining both its exact
`SourceLimits` and source `Arc`. Rustfmt and `git diff --check` pass. Cargo
tests and performance measurements are intentionally pending the coordinator's
release; no native timing, profiler, binary, or large evidence file is
committed.
