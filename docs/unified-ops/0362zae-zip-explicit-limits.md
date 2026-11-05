# 0362zae — explicit limits for the generic ZIP substrate

- Date: 2026-10-24.
- Base: `e35cede32a`.
- Scope: the public raw ZIP constructors and every tracked consumer of
  `soapberry-zip`, including formats/shared code, tests, examples, fuzz
  targets, and performance tools.

## Canonical API shape

The raw ZIP methods now take the caller-selected limit value on their one
canonical name. The old default twins and the `_with_limits`/`with_limits`
spellings are removed.

| Surface | Canonical method | Required limits |
|---|---|---|
| borrowed ZIP reader | `office::ArchiveReader::new(data, limits)` | `ArchiveLimits` |
| positional indexed archive | `office::IndexedArchive::{from_reader, from_zip_archive}(.., limits)` | `ArchiveLimits` |
| indexed preservation | `office::IndexedArchive::preservation_index(scratch, limits)` | `ArchiveLimits` |
| located archive | `office::LocatedArchive::index(limits)` | `ArchiveLimits` |
| lazy reader | `office::LazyArchiveReader::new(data, limits, cache_limits)` | `ArchiveLimits` plus `LazyArchiveCacheLimits` |
| raw preservation index | `PreservationIndex::new(archive, scratch, limits)` | `ArchiveLimits` |
| in-memory stream writer | `office::StreamingArchiveWriter::new(limits)` | `StreamingArchiveLimits` |
| custom stream writer | `office::StreamingArchiveWriter::with_writer(writer, limits)` | `StreamingArchiveLimits` |
| spooled stream writer | `with_writer_and_spool` and `with_writer_and_spool_and_name_plan` | `StreamingArchiveLimits` and existing spool limits |

Policy-bearing variants retain their policy argument but use the same required
limits position and canonical spelling: `from_reader_with_policy`,
`from_zip_archive_with_policy`, `LocatedArchive::index_with_policy`, and the
crate-private preservation policy helper. The writer's
`Default` implementation was also removed because it was a public
limit-less raw-constructor twin.

## Compatibility and semantic audit

Former limit-less calls pass the exact prior defaults explicitly:
`ArchiveLimits::default()` for readers/indexes/preservation and
`StreamingArchiveLimits::default()` for writers. Calls that already supplied
custom profiles retain those values and only lose the redundant method
spelling. Every former two-argument lazy-reader call now also passes
`LazyArchiveCacheLimits::default()` explicitly; the eight former
`new_with_cache_limits` calls use the canonical three-argument `new` with their
custom cache profiles. In the ODF
owner, the source-backed raw ZIP preservation call is bound to the retained
`SourceLimits::archive().into_zip_limits()` profile; the surrounding
`litchi-odf-common::package::ArchiveLimits` model type is a different
format-level value and is not substituted for the ZIP transport budget.

This distinction is intentional: `ArchiveLimits` and
`StreamingArchiveLimits` govern raw ZIP metadata/payload admission, while
format owners retain their `SourceLimits::zip_limits()` or equivalent source
profiles when they construct an archive. The migration does not broaden a
format policy into a transport default or move a custom owner profile to a
different API.

No ZIP parsing, preservation, validation, compression, output-layout,
streaming, spool, name-plan, or resource-accounting algorithm was changed.
The existing `SourceLimits::zip_limits()`/`archive().into_zip_limits()` bindings remain at OPC/ODF owners;
the ZIP transport does not replace those format-specific source profiles.
Error kinds, error precedence, policy checks, and progress-bearing streaming
failures remain unchanged.

The source-only inventory after migration recorded 577 borrowed-reader calls,
40 canonical lazy-reader calls, 56 positional-index calls, 121
preservation-index calls, 489 in-memory writer calls, and 19 custom-writer
calls in tracked Rust sources (excluding historical evidence snapshots).
These are syntactic call counts, not unique API or test counts. The inventory
also explicitly covered the required four-value lazy cache constructor, its
active-flight setter and `_value()` getters, the writer spool/name-plan
variants, policy variants, and the removed writer `Default` implementation;
those entries are beyond the eight initial heuristic rows. The cache follow-up
is recorded in `0362zaea`.

## Gate and validation receipt

The `soapberry-zip` API glossary profile (`tools/api_glossary.json`) bans both
the `_with_limits` suffix and the exact `with_limits` name, plus the lazy cache
constructor/accessor aliases removed by `0362zaea`. The shared glossary
records the required-limit raw ZIP methods and both records. A source scan over
tracked Rust consumers finds no old ZIP constructor spelling; unrelated
format-local names such as `litchi-odt::flat::from_reader_with_limits` remain
outside this substrate profile.

Before source preparation, the coordinator-pinned ignored lock was copied from
the budget-validation worktree with SHA-256
`52ed0cacac15b5e440349bfc0f4aca266bcf50d59d4321624e596bb410783c6`.

Validation used the coordinator-approved environment: the build shim was first
in `PATH`, `LITCHI_CAPPED` and `LITCHI_BLESS` were unset,
`LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`,
`CARGO_INCREMENTAL=0`, and the external target was
`(deleted-coord-artifact)`.
Both Cargo commands were offline and explicitly `--locked` against that
52ed… lock:

```text
cargo test --offline --locked -p soapberry-zip --all-targets
cargo check --offline --locked --workspace --all-targets
```

The corrected package test run passed 506 unit tests and 125 integration tests
(631 passed total), with one bounded multi-gigabyte test ignored; all four
examples compiled and had zero tests. The workspace all-targets consumer check
then completed successfully (`Finished dev profile`, exit 0). An initial
package compile attempt and the first workspace attempt were retained only as
terminal diagnostics: they exposed missed internal imports/arguments and an
ODF `ArchiveLimits` type-name collision, all corrected before the successful
runs. No command-log artifact was retained, so there are no log hashes to
claim. No rustdoc, dependency-graph comparison, or performance measurement is
claimed here.
