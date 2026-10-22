# 0360adb — XLSX raw identity admission repair

Status: implemented in `ada9cd288d` on `uo2/w4-raw-capability-review`, with
the matched probe in `44c6d22be7`, following `2d44bc4c55` and the W4
source-backed work.

The XLSX slicer and timeline raw escapes keep native relationship IDs and OPC
part names while entering through the canonical edit families. `EditOptions`
can tighten the package's retained `SourceLimits`; it cannot widen them. Raw
XML parsing still requires an explicit `&SourceLimits`. The edit methods use
the retained policy, return the shared `Commit`, and leave publication behind
the family commit boundary.

Admission now checks ID grammar, archive member spelling, relative target
length, existing owner IDs and target names, part counts, batch relationship
counts, and aggregate relationship XML metrics before raw values are cloned.
Native part and worksheet names are charged to their member-name and
relationship-target limits; they are not incorrectly charged to the payload
limit. Prepared raw values use one fallible writer with a ceiling of
`min(max_part_bytes, max_total_part_bytes)`, so the aggregate policy is
applied before writer growth. Slicer and timeline storage reuses the prepared
bytes instead of serializing them again.

Relationship and payload admission now share one operation-local package
catalog. The catalog plans the root and every relationship owner while reading
each existing part payload through the fallible iterator, so the required
cross-owner relationship proof and aggregate part-byte proof use one package
pass. The planned owner metrics are passed into that pass instead of being
recomputed. The final check applies only the known replacements and additions;
it does not rescan the package. The catalog is discarded when the package can
change, so it cannot outlive its source graph or retained policy. Relationship
admission still replaces the affected owner's old metrics exactly, and the
candidate checks the resulting content-types manifest under the same limits
before publication.

The borrowed preflight also runs before the public typed raw-store helpers
clone or serialize values. The helpers are crate-private implementation
routes; callers reach them through the package edit families and the explicit
raw codec values, so there is no free public mutation path around the
canonical commit boundary. Batch admission rejects malformed or duplicate
native IDs before staging, and the aggregate catalog includes relationship
owners unrelated to the inserted cache.

The raw tests cover non-default IDs and paths, archive-member spelling versus
absolute OPC spelling, explicit-limit parsing, forward publication, patch
replay, inverse replay, no-op commits, untouched bytes, tight payload refusal,
identity names that exceed the payload limit while satisfying their own
limits, aggregate relationship refusal including an unrelated owner, and
batch identity admission. The tests pass without changing the physical
identity or fidelity of existing typed routes.

`Edit::apply` and `Edit::apply_raw` now invalidate a graph-proof flag before
staging and set it only after the candidate passes graph validation. `reset`
restores the proof for the original validated graph, and a failed operation
leaves the flag invalid conservatively. A successful raw operation therefore
does not run the same final graph validation again in `commit`, while the
proof cannot be reused across a changed candidate or policy.

The matched allocator/Ir probe is
`tools/perf-baseline/unified-xlsx-raw-identity/README.md`. It covers
`small_single`, `large_multipart_single`, `small_multi_cache`, and
`large_multipart_multi_cache`; it verifies output bytes and native IDs and
records the generated corpus and binary hashes. Baseline source is
`47c89d20da1594187c967e161b506d828aed1599`; candidate source is
`44c6d22be72a6351e48fb92e4c18017101657bcb`. The probe was prepared but not
run in this turn because the coordinator held new Cargo builds under disk
pressure.

Validation before the performance repair on this worktree:

* `cargo test -p litchi-xlsx --test xlsx_slicer_timeline_crud -- --nocapture`
  (9 passed);
* `cargo test -p litchi-xlsx --lib --all-features` (1,824 passed, 0 failed);
* `cargo clippy -p litchi-xlsx --all-targets --all-features -- -D warnings`
  passed during the repair before the final private-wrapper cleanup; and
* targeted `rustfmt --edition 2024` over all changed Rust files.

Static validation after the repair: targeted `rustfmt --edition 2024 --check`
and `git diff --check` pass. No new Cargo test, Clippy, allocator, or Ir run is
claimed while the build hold is active.

No ABBA, allocation, instruction-count, peak-heap, or throughput
qualification is claimed by this repair. The package-wide relationship and
part-byte catalog remains bounded by the retained source policy; its cost and
the graph-proof saving should be included in the W4 ABBA comparison. A second
Clippy run was deferred while the coordinator held new builds under disk
pressure; the full all-feature library suite above predates the final
performance-only source changes.
