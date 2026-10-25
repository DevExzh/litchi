# 0362ul157 — Public support aliases removed

- Date: 2026-10-25.
- Status: source-ready; Cargo/rustdoc validation remains with the coordinator.
- Base: `f8cd398a770f7363f628be40edb3ce4219566c16`.

This leaf closes the public support-name residue without retaining deprecated
forwarders. `litchi_ppt::RecordParser::find_records_owned()` keeps the original
recursive owned-clone capability under a name that states its allocation
behavior; all in-tree consumers already use the borrowed `find_records_ref`
view, and the old `find_records` spelling is deleted.

`litchi_iwa_graph::ReferenceGraph` now exposes only the typed identity API:
`add_object_reference`, `incoming`, `outgoing`, `all_object_ids`, `has_cycle`,
`reachable`, and named `statistics`. `ObjectId::try_from(0)` remains the
explicit invalid-wire-identity check, and tests cover that rejection before a
graph operation. The former raw-`u64` wrappers are deleted, so invalid zero
identities cannot be silently discarded.

`litchi_iwa_index::ObjectIndex::reference_snapshot` is the canonical borrowed
`ReferenceGraphSnapshot` projection. It preserves the complete immutable graph
and its cheap `Clone` capability while exposing no mutable state; the old
`reference_graph` compatibility bridge and its consumer are deleted/migrated.

The PPT minimal/full and new IWA graph/index glossary profiles ban the removed
spellings. The API glossary records the typed replacements and the invalid-ID
semantics. Source checks performed here:

- `PYTHONDONTWRITEBYTECODE=1 python3 -B -m json.tool tools/api_glossary.json`;
- focused searches for removed public names and deprecated attributes;
- `git diff --check`.

Cargo tests, rustdoc, clippy, native fixtures and performance measurements are
intentionally outstanding for the coordinator's independent build slot. No
binary, fixture, or large performance evidence is committed.
