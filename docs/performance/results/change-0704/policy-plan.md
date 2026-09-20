# 0704 opened PPTX MCE retention policy and API plan

Status: read-only policy inventory at `d48523eec2` (the 0703 diagnostic
baseline). This file proposes policy plumbing only. It does not change
production Rust, patch bytes, or the existing six-argument constructor.

## Authority and measured sizing

ADR 0005's 2026-09-16 retained-state amendment requires a retention ceiling to
be a member of the operation's finite policy, to be intersected whenever two
policies meet, observable on the value that holds the bytes, and releasable
without changing that value's meaning. A ceiling exceeded by pure
recomputation must silently take the old path and must not become a typed
refusal.

The 0703 trace gives a useful first default. The real fixture has 14 distinct
owned MCE projections and 543,070 bytes of observed `Vec` capacity; 509,058
bytes are reused after one edit and 502,982 after two edits. The generated
control has only borrowed results and therefore adds no transformed-output
capacity. A tentative `max_retained_mce_bytes` default of 1 MiB admits the
measured real capture while remaining a separate, much smaller budget than the
existing 64 MiB serialized cross-slide candidate ceiling. The budget should
charge the retained transformed allocation's `capacity`, because that is the
memory kept alive; the documentation and observable counter must say whether
metadata is excluded. The entry table must still have an independent bound
derived from `max_parts`, and all table/output reservations must be fallible.

The 0703 numbers are logical capacity accounting, not a peak RSS result. They
exclude cache metadata, allocator rounding, source-owner overhead, admission
bookkeeping, and transient overlap. The implementation record must not present
the 1 MiB value as a measured peak-memory bound until that accounting is
closed.

## Current policy plumbing

| Concern | Current location | Consequence for the new member |
| --- | --- | --- |
| Policy definition and defaults | [`opened/model.rs:56-173`](../../../../crates/litchi-pptx/src/opened/model.rs#L56) | Add one private field and a default constant; preserve the six existing `Limits::new` arguments. `Debug`, `Copy`, equality, and `Default` then include the member automatically. |
| Public policy access | [`opened/model.rs:499-503`](../../../../crates/litchi-pptx/src/opened/model.rs#L499) | `Snapshot::limits()` already carries the complete policy. Add `Limits::max_retained_mce_bytes()`. |
| Initial capture | [`package/model.rs:179-200`](../../../../crates/litchi-pptx/src/package/model.rs#L179) and [`opened/model.rs:521-535`](../../../../crates/litchi-pptx/src/opened/model.rs#L521) | `opened_presentation_with_limits` already accepts and stores the policy. The default route gets the new default through `Limits::default()`. |
| Snapshot storage/rebind | [`opened/model.rs:270-297`](../../../../crates/litchi-pptx/src/opened/model.rs#L270) and [`opened/model.rs:471-496`](../../../../crates/litchi-pptx/src/opened/model.rs#L471) | `Snapshot::limits` is copied by the existing struct update. If the MCE table is added to `Snapshot`, `rebound_to` must project it onto the rebound package or drop entries whose source allocation is not held, just as `PartDigests::project` does. |
| Transaction handoff | [`opened/transaction.rs:30-37`](../../../../crates/litchi-pptx/src/opened/transaction.rs#L30) and [`opened/transaction.rs:1190-1241`](../../../../crates/litchi-pptx/src/opened/transaction.rs#L1190) | The source snapshot is cloned into the transaction; commit already passes `self.source.limits` to patch capture and the final capture. A cache handoff must retain only unchanged, source-authorized entries and must leave edited/replaced parts on the uncached path. |
| Patch-owned policy | [`opened/patch.rs:194-201`](../../../../crates/litchi-pptx/src/opened/patch.rs#L194) | `Patch` already carries the complete `Limits`, so the new member follows a patch through history, application, inverse, and decoded replay. |
| Patch policy meets | [`opened/patch.rs:412-428`](../../../../crates/litchi-pptx/src/opened/patch.rs#L412) and [`opened/patch.rs:436-450`](../../../../crates/litchi-pptx/src/opened/patch.rs#L436) | `Patch::join` and `Patch::three_way` each call `intersect_limits`; both must take the minimum new retention ceiling. |
| Cross-copy policy meet | [`opened/cross_copy_plan.rs:980-990`](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L980) and [`opened/cross_copy_plan.rs:1874-1885`](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L1874) | Add the new minimum even though cross-copy currently consumes only the candidate-archive member. Every policy member must be intersected when source and destination snapshots meet. |
| Removal wrapper | [`opened/remove_plan.rs:124-195`](../../../../crates/litchi-pptx/src/opened/remove_plan.rs#L124) and [`opened/remove_plan.rs:406-456`](../../../../crates/litchi-pptx/src/opened/remove_plan.rs#L406) | No separate intersection exists. The wrapper delegates to its inner `Patch`; the new member follows automatically. |
| Durable patch encoding | [`opened/patch.rs:492-515`](../../../../crates/litchi-pptx/src/opened/patch.rs#L492) and [`opened/patch.rs:576-610`](../../../../crates/litchi-pptx/src/opened/patch.rs#L576), plus the wrappers at [`opened/remove_plan.rs:90-121`](../../../../crates/litchi-pptx/src/opened/remove_plan.rs#L90) and [`opened/cross_copy_plan.rs:381-443`](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L381) | Limits are not serialized. `to_bytes` uses only the patch-byte member; `from_bytes_with_limits` receives a caller policy. No wire field, magic bump, or migration is needed for this retention member. Add a test that changing only the MCE budget leaves durable bytes unchanged. |
| History | [`opened/patch.rs:644-751`](../../../../crates/litchi-pptx/src/opened/patch.rs#L644) | History stores a `Limits` copy and uses only history members today. The new field must remain present in the copy but must not be charged as history bytes. |

The exhaustive current search found no other `Limits` intersection in
`crates/litchi-pptx/src/opened`: the two functions above are the only places
that reconstruct a policy from two policies. `Limits` struct literals are
private to `opened/model.rs`; callers use `Limits::new`, so adding the private
field does not create an external struct-literal break.

## Constructor compatibility and configuration

The current six-argument constructor is used in the opened unit tests, the
cross-copy retention/revision tests, the adversarial public test, and the
`tools/perf-baseline/src/pptx_slide_boundaries.rs:613-615` boundary probe. It
must remain callable with six arguments. Do not append a seventh parameter.

The preferred additive API is a checked fluent builder for the new member,
seeded from `Limits::DEFAULT`, for example:

```rust
let limits = Limits::builder()
    .max_retained_mce_bytes(512 * 1024)
    .build()
    .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
```

The builder should use the existing nonzero invariant: `0` is rejected, and
`1` is the documented off switch because no useful owned transformed output
fits it. The builder can mirror `litchi_opc::ReadLimitsBuilder` while exposing
all existing `Limits` members, or the minimal additive form can be paired with
a checked `Limits::with_max_retained_mce_bytes` setter. The API decision should
be made once by the design review; the invariant and six-argument compatibility
are the requirements. A setter returning an unchecked `Self` would permit a
policy value that `Limits::new` itself rejects and should not be introduced.

Whatever spelling is selected, `Limits::DEFAULT` must initialize
`max_retained_mce_bytes` to 1 MiB, `Limits::default()` must match it, and the
accessor must be `const`, `Copy`-friendly, and documented as an aggregate
retained transformed-output ceiling distinct from
`max_retained_candidate_bytes`.

No `Package` policy field is required. The existing
`Package::opened_presentation_with_limits` is the explicit ingress, and the
default `Package::opened_presentation` remains the default-policy route. A
separate `opened_presentation_transaction_with_limits` convenience method is
optional API polish, not required for propagation, because callers can already
call `opened_presentation_with_limits(limits)?.edit()`.

## Required intersection and propagation changes

The production policy edit should be mechanically limited to these data-flow
points before the cache implementation is wired:

1. Add `max_retained_mce_bytes` to `Limits`, initialize it in `DEFAULT`, keep
   `new` at six arguments, add the checked builder/setter and accessor.
2. Add `left.max_retained_mce_bytes().min(right.max_retained_mce_bytes())`
   to `opened/patch.rs:1061-1072`.
3. Add the same minimum to
   `opened/cross_copy_plan.rs:1874-1885`.
4. Leave all durable patch encoders/decoders byte-for-byte unchanged. The
   supplied decode policy controls future optional retention; a patch does not
   carry a retention promise.
5. Preserve the whole-policy comparison in
   `opened/patch.rs:821-826` (`committed.limits == limits`). A cache-bearing
   committed snapshot must not be rebound under a different MCE ceiling by the
   existing exact-candidate shortcut.

The new field is intentionally separate from
`max_retained_candidate_bytes`: the latter is the serialized archive retained
by `CrossSlideCopyPlan`, and its accessor/docs/tests at
`opened/model.rs:138-165` and `opened/cross_copy_plan/retention_tests.rs`
must continue to name only that archive. Do not combine, alias, or spend one
budget against the other.

## Snapshot retention surface for the cache implementation

The policy field alone does not satisfy ADR 0005. When the root/coder adds the
capture cache, the public `Snapshot` surface should include:

* `retained_mce_bytes(&self) -> usize` (or an explicitly documented optional
  form) reporting the aggregate retained transformed-vector capacities;
* `release_retained_mce(&mut self)` clearing the optional table while keeping
  slides, revision, package meaning, and policy unchanged; and
* a `Debug` summary that reports counts/bytes without printing XML.

The retained table must be capture/snapshot local, never global or ambient.
Its key must prove raw allocation aliasing to the visible `Part::blob()` and
retain the raw owner for the ABA argument; its value must bind the complete
default MCE profile and any relevant input/output limits. Errors and typed
refusals must never be cached. A miss, an admission failure, or a released
entry must recompute through the current path.

The retained byte counter should use checked additions and the actual output
capacity admitted. Admission must compare the candidate capacity against the
remaining aggregate budget before retaining it; a candidate over budget falls
back to recomputation and reports no retained bytes. The source package may
already own the raw `Arc`, but the implementation record must state whether
that owner is charged separately and why retaining it does not pin a payload
outside the snapshot's package.

`Snapshot::rebound_to` is a mandatory invalidation seam: a rebound package may
share built-in payload `Arc`s but a foreign `Part` can copy or return a
non-aliasing `blob_arc`. The cache must be projected with the same alias gate
as `PartDigests`, or discarded on any failed projection. `Transaction::new`
and `commit` must preserve the source cache only for allocations still held by
the staged package; edited, replaced, removed, or newly transferred parts must
miss. No-op commit already returns the source snapshot at
`transaction.rs:1222-1233`, so it must not manufacture a second cache.

## Tests and review gates to add

Policy-only tests should cover:

1. `Limits::default().max_retained_mce_bytes() == 1 << 20`, the builder's
   custom value, and rejection of zero (with `1` accepted as the off switch).
2. Six-argument `Limits::new` still compiles and receives the default MCE
   ceiling; all existing constructor call sites remain unchanged unless a
   test needs a deliberate custom MCE budget.
3. Both policy intersections choose the tighter MCE budget, including source
   and destination order, `Patch::join`, and `Patch::three_way`.
4. Durable `Patch`, `SlideRemovalPatch`, and `CrossSlideCopyPatch` bytes are
   identical when only `max_retained_mce_bytes` differs; decoding under a
   custom policy succeeds and carries that policy internally. No magic bump or
   policy bytes should appear.
5. A committed snapshot with a different MCE budget does not use the existing
   exact-candidate reuse shortcut merely because package content is equal.

Cache-specific tests belong with the root/coder implementation and must add:

* default-budget admission for the 0703 real fixture and a below-budget
  recomputation fallback;
* aggregate accounting at exact budget and one-byte-over-budget boundaries;
* borrowed generated XML retaining no transformed bytes;
* `retained_mce_bytes` and release preserving revision, slides, semantic
  values, patch bytes, and published bytes;
* one-edit/two-edit invalidation of changed slides while unchanged slides hit;
* snapshot clone/rebind and foreign non-aliasing `Part` fallback;
* profile/limit-key mismatch, malformed XML, typed error precedence, notes
  validation, and strict/refusal routes remaining uncached; and
* fallible table/output reservations under allocation pressure.

The existing cross-copy retention tests should remain focused on
`max_retained_candidate_bytes`; they should not silently begin asserting or
spending the MCE budget. The exact baseline/candidate and full production
Cargo/performance work remains the root agent's lane.
