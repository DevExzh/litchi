# 0806 amendment protected native preflight

This packet is a bounded supplemental check for the warning failure found by
the 0806 full quality gate. The amended after leg reuses the existing
`BytesStartExt::unchecked_attributes` helper in each of the five helper
copies, so the helper is exercised by the constructor and no longer appears
unused under warnings denied. The OPC copy is public. This archived candidate
still contains the accidental OLE visibility narrowing later caught by
`quality-1/`; the separate `candidate-visibility-amendment/` restores its two
public declarations before workflow builds. The timed OPC implementation is
unchanged by that compatibility repair.

The comparison is bound to the original production bytes in
`../candidate/before` and to the amended after handoff in
`../candidate-quality-amendment/after`. The handoff's `before` directory is
excluded because it was captured while the combined 0806 candidate was already
applied. The packet's `source/before` and `source/after` copies are the
authoritative six-file comparison and are checked against those lineages.

The native lane reuses the sealed 0805 39-case probe (and retains the 0805
report schema/tool as a harness identity), including its semantic
oracle, 8 clone advances, two timing modes, six alternating paired blocks, 30
samples, three warmups, 4,096 iterations, and CPU 12. The new bootstrap seed is
806082. It has no Callgrind or heaptrack lane, no resource claim, no public
workflow claim, and no historical timing pooling. The result only decides
whether this amended candidate may return to the already frozen 0806 workflow
gates.

The source mirror quality driver runs the exact five helper copies and shared
OPC tests: 70 tests for the original before leg, 100 tests for the amended
after leg, and warning-denied all-targets Clippy for each. The native build is
separate and contains only the direct probe binary. Root executes Cargo and
native commands serially. Both mirror test counts and Clippy gates passed,
both native builds completed, and all 936 children / 28,080 samples were
captured. The independently reproduced decision advances the constructor
amendment to workflow trials, with no production adoption claim.

`analyze.py` is an offline fail-closed reader. Run it after capture (and again
after cleanup for replay) to write `analysis.json`; it validates source lineage,
the mechanical five-constructor rewrite, probe semantic oracles, all six
paired block receipts, binary identities, sample counts, and the protected
policy. The separate raw auditor must then write `root-native-audit.json` with
schema `litchi.performance.0806.amendment-root-native-audit.v1`, `passed: true`,
`matches_primary_analysis: true`, `native_reports: 936`, `native_samples: 28080`,
and matching policy fields. Re-running `analyze.py` after that witness creates
`decision.json`; without the witness it only writes a fail-closed
`decision-pending.json`. The final decision retains the 0805-compatible fields
`advance_to_workflow_trials`, `production_adoption`,
`protected_consume_regressions`, and `dominant_class_benefits`.
