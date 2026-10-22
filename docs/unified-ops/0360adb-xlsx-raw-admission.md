# 0360adb — XLSX raw identity admission repair

Status: implemented in `d625ed9a25` on `uo2/w4-raw-capability-review`,
following `d0ac4d5ce4` and the W4 source-backed work.

The XLSX slicer and timeline raw escapes keep native relationship IDs and OPC
part names while entering through the canonical edit families. `EditOptions`
can tighten the package's retained `SourceLimits`; it cannot widen them. Raw
XML parsing still requires an explicit `&SourceLimits`. The edit methods use
the retained policy, return the shared `Commit`, and leave publication behind
the family commit boundary.

Admission now checks ID grammar, archive member spelling, relative target
length, existing owner IDs and target names, part counts, batch relationship
counts, and aggregate relationship XML metrics before raw values are cloned.
Prepared raw values use one fallible writer with a ceiling of
`min(max_part_bytes, max_total_part_bytes)`, so the aggregate policy is
applied before writer growth. Slicer and timeline storage reuses the prepared
bytes instead of serializing them again. Relationship admission keeps one
package-wide aggregate ledger per insertion operation, replacing the affected
owner's old metrics exactly; it does not scan once per cache. Part-byte
admission accounts for replacements and additions, and the candidate checks
the resulting content-types manifest under the same limits before publication.

The raw tests cover non-default IDs and paths, archive-member spelling versus
absolute OPC spelling, explicit-limit parsing, forward publication, patch
replay, inverse replay, no-op commits, untouched bytes, tight payload refusal,
aggregate relationship refusal, and batch identity admission. The tests pass
without changing the physical identity or fidelity of existing typed routes.

Validation on this worktree:

* `cargo check -p litchi-xlsx --tests`;
* `cargo test -p litchi-xlsx --test xlsx_slicer_timeline_crud` (7 passed); and
* targeted `rustfmt --edition 2024` over all changed Rust files.

No ABBA, timing, peak-heap, or throughput qualification is claimed by this
repair. The package-wide relationship and part-byte ledgers are bounded by
the retained source policy; their cost should be included in the W4 ABBA
comparison before coordinator acceptance.
