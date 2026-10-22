# 0360adb — XLSX raw identity admission repair

Status: implemented in `8d3290cdb3` on `uo2/w4-raw-capability-review`,
following `2d44bc4c55` and the W4 source-backed work.

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
bytes instead of serializing them again. Relationship admission keeps one
package-wide aggregate ledger per insertion operation, replacing the affected
owner's old metrics exactly; it does not scan once per cache. Part-byte
admission accounts for replacements and additions, and the candidate checks
the resulting content-types manifest under the same limits before publication.

The borrowed preflight also runs before the public typed raw-store helpers
clone or serialize values. The helpers are crate-private implementation
routes; callers reach them through the package edit families and the explicit
raw codec values, so there is no free public mutation path around the
canonical commit boundary. Batch admission rejects malformed or duplicate
native IDs before staging, and the aggregate ledger includes relationship
owners unrelated to the inserted cache.

The raw tests cover non-default IDs and paths, archive-member spelling versus
absolute OPC spelling, explicit-limit parsing, forward publication, patch
replay, inverse replay, no-op commits, untouched bytes, tight payload refusal,
identity names that exceed the payload limit while satisfying their own
limits, aggregate relationship refusal including an unrelated owner, and
batch identity admission. The tests pass without changing the physical
identity or fidelity of existing typed routes.

Validation on this worktree:

* `cargo test -p litchi-xlsx --test xlsx_slicer_timeline_crud -- --nocapture`
  (9 passed);
* `cargo test -p litchi-xlsx --lib --all-features` (1,824 passed, 0 failed);
* `cargo clippy -p litchi-xlsx --all-targets --all-features -- -D warnings`
  passed during the repair before the final private-wrapper cleanup; and
* targeted `rustfmt --edition 2024` over all changed Rust files.

No ABBA, allocation, instruction-count, peak-heap, or throughput
qualification is claimed by this repair. The package-wide relationship and
part-byte ledgers are bounded by the retained source policy; their cost should
be included in the W4 ABBA comparison before coordinator acceptance. A second
Clippy run was deferred while the coordinator held new builds under disk
pressure; the full all-feature library suite above includes the final source.
