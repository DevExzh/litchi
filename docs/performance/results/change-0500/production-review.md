# 0500 production review: managed paragraph batches

Status: **no semantic or atomicity blocker found**. This is a read-only review
of the frozen `transaction.rs` candidate; no production source was changed and
no build or test command was run by this reviewer. The candidate hash matches
`after-source.json`.

The review used the accepted snapshot/edit/patch, finite-budget, validation,
source-authority, and OOXML ownership contracts in ADR 0003, ADR 0005, ADR
0006, ADR 0008, ADR 0011, ADR 0017, and ADR 0024. The recorded gates provide
the execution evidence: the focused managed batch suite is 8/8, the existing
managed document suite is 36/36, formatting and warning-denied Clippy pass,
and the benchmark harness has equivalent repeated/batch semantic, output,
inverse, source-counter, and budget oracles. The completed after phase has 24
children (36 total with the repeated controls), 2,160 measured samples, and
216 warmups; all retained correctness, cleanup, source, and budget oracles
pass. Its adverse latency/RSS flags remain descriptive evidence, not a broad
performance claim.

## What the implementation proves

`Edit::replace_body_paragraph_texts` validates the canonical selector list and
dispatches managed snapshots to a dedicated planner. It does not call the
scalar method in a loop. The planner first admits every borrowed input length
and capacity, validates every authored string, and scans each selected
paragraph against both the projected snapshot and the immutable base. Existing
managed operations are updated or removed by base text; new operations retain
the base text. The resulting ledger is assembled once in stable prior-ledger
order followed by canonical new positions.

The shared reconstruction helper then checks the final ledger text bound,
plans every surviving operation against the immutable base, proves each text
range with one original `SourceXmlPart`, publishes one source splice, reopens
one managed candidate, and semantically reads back every planned paragraph.
`self` is assigned only after all proofs, publication, candidate parsing, and
readback succeed. Temporary admissions and candidate reservations are RAII
owners, so late refusal, cancellation, stale source, parse failure, and budget
failure leave the staged projection and operation ledger unchanged.

The no-op and revert branches have the required source behavior. If all
requested values already equal the projected values, the method returns
without source authorization or candidate publication. If the final ledger is
empty, it clears the ledger and restores `self.base.clone()`, retaining the
original source owner and exact bytes. A signed managed source therefore admits
an exact no-op/revert while a changed operation returns the typed unsafe-edit
refusal. Source authorization is obtained after those decisions but before
the final operation map/admissions and reconstruction; this is the necessary
ordering for the signed no-op capability.

## Resource-review resolutions

These points were checked against the final drop order and accepted budget
contract; none is a production blocker.

1. **Base scan lifetime is ordered correctly.** Each changed decision stores
   `_scan: ScanAdmission` with `base_text` while the base text is either still
   being planned or is transferred into the separately admitted operation
   ledger (`transaction.rs:2683-2736` and `4865-4872`). The whole decision
   vector, including these scan reservations, is explicitly dropped at
   `transaction.rs:2919`, before `reconstruct_managed_paragraph_operations`
   starts at line 2923. There is therefore no decision-scan/reconstruction-
   scan overlap. Dropping the scan at extraction would instead risk leaving
   the retained `base_text` temporarily uncharged; the current lifetime is a
   conservative, intentional owner boundary.

2. **Input-admission retention is intentional.** The generic
   `_temporary_admissions: impl Sized` parameter keeps `input_admissions` live
   through the shared helper (`transaction.rs:2170-2177` and 2923-2929), as
   the batch comment states. This matches the scalar route's conservative
   input-admission lifetime and covers the borrowed caller strings until
   source publication/readback either succeeds or fails. The parameter could
   be named more explicitly for readability, but its lifetime is safe and
   bounded.

3. **The ledger matching loop is a bounded metadata limitation, not a budget
   violation identified by the accepted ADRs.** For each selected replacement
   the planner scans the prior ledger (`transaction.rs:2693-2719`), giving a
   finite `O(selected * prior_ops)` comparison cost. ADR 0005 requires every
   operation to use a finite hierarchical budget, but does not require an
   exact `Work` unit for every metadata comparison; the accepted 0495 record
   explicitly describes conservative finite work as a safety ceiling rather
   than a CPU-instruction count (0495, lines 79-84). The existing scalar route
   performs the analogous ledger scans. The batch charges its owner scans and
   final reconstruction, and the measured rows correctly report that batch
   `Work` may differ from repeated scalar `Work`; no separate comparison charge
   is required for this change. A future tighter Work model may charge the
   linear metadata passes, but this is not a release blocker.

Selector validation precedes the managed execution-context check, but no
accepted contract requires cancellation to dominate malformed input. Both
outcomes remain typed and the valid-input cancellation path is covered.

## Coverage and remaining evidence

The focused tests cover exact XML equality with repeated scalar edits,
unknown/media member preservation, forward and complete-artifact inverse
publication, mutable-source refusal, scalar/batch composition in both orders,
exact no-op and revert pointer restoration, canonical-selector refusals,
late leaf failure atomicity, cancellation, Work and Memory failures, owner
release, and signed-source no-op/change behavior. The harness exercises
128/512 paragraphs, 1/8/32 replacements, owned and warm-file providers,
repeated and batch routes, and records lifecycle/edit timings plus semantic,
raw-member, inverse, source-version, and budget gauges.

No test currently isolates the object/depth boundary for a batch or an exact
metadata-comparison Work model. The after benchmark result and its independent
evidence verification remain the authority for performance claims; the
implementation review alone makes no broad CRUD claim.
