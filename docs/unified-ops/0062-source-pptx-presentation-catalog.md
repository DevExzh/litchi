# Source-backed PPTX presentation catalog

This wave implements the L-PPTX-R4 presentation-level read foundation. The
complete unified-operations goal, including the PPTX adapter, remains active.

## API and ownership

Source presentations acquire detached master and layout snapshots in authored
XML order. Positional selectors load only the selected master/layout payload;
explicit collection reads decode the requested collection. Missing master
positions and empty layout lists remain distinct. Masters expose authored name,
preserved state, positional identity and retained scenes. Layouts reuse the
related-part projection introduced in record 0061.

Presentation sections expose authored names and resolved slide positions.
Section GUIDs, native slide identifiers and inert extension bytes remain under
`presentation::sections::raw`. Lists and section handles share immutable data.
`presentation::raw` exposes the root payload already captured at open. Raw
access reads that retained snapshot; semantic source reads enforce current
source/execution fences.

## Validation

Source acquisition uses the existing bounded physical/MCE ingress and shared
leaf codecs. Master/layout lists require the correct namespace and structural
position, bounded reference counts and unambiguous identities. The normative
optional inert extension list inside an ID entry remains accepted. Slide
dimensions use the schema coordinate range. Section membership refuses native
slide IDs absent from the retained presentation catalog.

Collection acquisition validates all requested projections before publishing
scenes. A failed final fence rolls back each newly published cache entry by
exact snapshot identity. Bulk collections cap references at 4,096 and accounted
retained payload, scene and identity bytes at 64 MiB, in addition to per-part
limits. This is a logical retention bound, not a peak allocator or process-memory
bound: transient parsing allocations can coexist with retained snapshots.
No performance acceptance is claimed by this wave.

## Verification

Offline validation passed:

- `cargo test -p litchi-pptx --all-features`: 1,439 passed, 3 ignored,
  including doctests.
- `cargo test -p litchi-pptx --test source_catalog --test source_sections
  --test source_related`: 37 passed with default features.
- `cargo clippy -p litchi-pptx --all-features --all-targets -- -D warnings`.
- `RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx --all-features --no-deps`.
- `cargo check --workspace --lib`.
- PPTX minimal/full glossary profiles: 43 bans per profile; glossary checker
  unit tests: 21 passed.
- Offline crate-boundary audit: 66 workspace packages, 267 internal dependency
  declarations, 11 explicit existing debt items.

Regression coverage includes XML order, detached ownership, selective loading,
invalid relationships, source-change fences, dimensions, malformed section
membership, mixed core namespaces and undeclared prefixes. Existing tests of
slide-reference/size independence now use schema-valid dimensions rather than
one-EMU fixture values.

## Commit and cleanup

The local rewritten base `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an
ancestor of HEAD. No remote fetch or further rebase is claimed. `git add -A`
failed because `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
cannot be created on the read-only filesystem. No commit is claimed; this wave
and the earlier outstanding waves remain uncommitted.

Both required staged-artifact checks printed nothing, and the changed/untracked
file audit found no oversized evidence or compiled binaries. Removed 32.8 GiB
of generated Cargo output, the three owned LibreOffice fixture copies,
generated untracked lockfiles and Python caches. No temporary worktrees were
created. The existing unrelated worktrees were left intact.
