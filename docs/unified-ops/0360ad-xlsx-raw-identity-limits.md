# 0360ad — Bound XLSX native-identity insertions by source limits

Status: implemented in `d0ac4d5ce4` on `uo2/w4-raw-capability-review`, based
on `47c89d20da`. This closes the raw slicer and timeline identity repair that
was left after the typed insertion families were introduced.

The raw escape remains an owned edit-family operation. Callers construct
`slicer::raw::{Cache, Part}` or `timeline::raw::{Cache, Part}` with their
native relationship IDs and absolute OPC part names, parse XML with a
required `&SourceLimits`, pass the value to the corresponding package edit,
and publish through the shared `Commit`. No standalone store wrapper,
mutable package owner, alias, or direct staged publication was restored.

Raw XML is admitted with a borrowed event, depth, attribute, and part-size
pass before a format model is parsed. Authored cache and worksheet XML uses a
fallible bounded writer before a raw value is moved or cloned. Integration
extensions and nested timeline auto-filters use the same bounded sinks. The
package's retained open policy is intersected component by component with a
raw caller policy, so a raw call can tighten the policy but cannot widen it.

Relationship admission is scoped to each affected owner. The source
relationship member is loaded with the effective policy and
`OwnedRelationships::plan_edit` computes the final bytes, events, field
limits, and relationship count before the member is materialized. Native
part paths use an allocation-free `PackURI::relative_ref_len` check, which
admits the actual relative target spelling instead of rejecting a valid path
because its absolute spelling is longer. Untouched package parts are not
rescanned as a substitute for this owner-local admission.

The tests cover non-default IDs and paths, exact relative-target limits,
forward publication, patch replay, inverse replay, no-op commit behavior,
untouched bytes, and refusal of a tight output limit before timeline cache
values are cloned. A temporary counting allocator probe exercised four raw
timeline caches beside an 8 MiB unrelated part; it observed 6,741 allocation
calls and 930,766 requested bytes for the measured edit and commit. This is a
diagnostic of the current route only; it is not an ABBA or performance
qualification claim.

Validation completed on the isolated worktree:

* `cargo check -p litchi-xlsx --all-features`;
* `cargo clippy -p litchi-xlsx --all-targets --all-features -- -D warnings`;
* `cargo test -p litchi-opc --lib relative_ref_len_matches_the_allocating_writer`;
* `cargo test -p litchi-xlsx --test xlsx_slicer_timeline_crud --all-features`
  (6 passed); and
* `cargo test -p litchi-xlsx --lib --all-features` (1,822 passed, 0 failed).

No timing, peak-heap, or throughput qualification is claimed by this record.
