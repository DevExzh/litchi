# Follow-up after the bounded publication batch

Ordinary ODP Snapshot/Edit/Commit/Patch already exists in
`crates/litchi-odp/src/authoring/edit.rs`. Snapshot retains full source bytes,
an owned package, and parsed slides. Transaction creation rebuilds mutable
presentation state; commit serializes the package, reopens it, performs
readback, and retains before/after snapshots in a reversible Patch. Durable
patches and History also depend on these owned artifacts and exact source
bytes. This batch does not replace that contract.

The next directly comparable optimization experiment should attribute the
ordinary `odp_existing_append_lifecycle` control by phase, using the retained
64/4,096/8,192-slide corpus and the same normal/allocator repeats. Inspect
`MutablePresentation::to_bytes_bounded`, repeated owned-package opening and
indexing, semantic/compact-XML/media/RDF/chart validation, publication, and
before/after Patch retention. Select the next implementation only after that
attribution identifies avoidable work. Ordinary replace/remove/move operations
share much of this machinery and are useful follow-on scenarios.

The new source-tail proof binds source version and content hashes. It is not
the full-archive exact-byte authority used by ordinary Patch application.
Moving ordinary snapshots to lazy source capabilities would require explicit
answers for `bytes()`, semantic slide access, provider lifetime, durable
patches, History, inverse restoration, joins, and three-way merge. Do not
silently substitute the specialized publication report for those results.
Materializing before entering the ordinary API preserves its contract but
also restores the owned memory cost.

The wider goal still needs additional representative CRUD baselines,
input/output and cold/range-source matrices, native application roundtrips,
and measured bounded-worker scaling. The source-tail measurements and fixture
oracles do not close those obligations.
