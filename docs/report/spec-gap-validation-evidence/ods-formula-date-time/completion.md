# Date/time batch completion

The 24 OpenDocument 1.4 Part 4 §6.10 functions are implemented and validated
within [the explicit repository profile](contract.md). This completes this
function batch, not OpenFormula or the broader specification-gap program.
Production code remains the [frozen candidate](gates/freeze.json)
`d16039ce48cb441c35461318c8634a49ae0b2437`.

The final [root verification receipt](root-verification.json) reports
`verified: true`. Its [coverage manifest](coverage-requirements.json) is
SHA-256 `f1c0f1053f2394aa7b1463c6c2b374078df20e59552c98629148c28235c0788a`:
24 function requirement sets, 223 function bindings, 20 cross-cutting
requirements, and 54 cross-cutting bindings. The [coverage audit](coverage-audit.md)
distinguishes executed tests and oracle cases from reviewed frozen-source
proofs; supplementary proofs do not imply additional executed vectors.

Retained validation comprises seven successful isolated gates, 1,745 passing
ODS tests with no failures or ignored tests, 132 independent oracle vectors
with Rust replay, and 132 native observations with explicit compatibility
dispositions. Semantic and resource/cache reviews pass. Root also reran the
strict evidence verifier after the last coverage correction. No new Cargo
execution was needed for those documentation-only corrections.

The [performance disposition](performance-review.md) accepts 4,620 retained
samples descriptively. All 68 matched control groups preserve allocation,
work, read, and output accounting sets. The 4x4 SIN parse/evaluate median
increases by 5.228% (95% bootstrap interval 0% to 10.909%); 22 median RSS
groups increase by 184–240 KiB. Tail flags, host-load differences, and unpinned
CPU affinity remain disclosed. No overall speedup or causal attribution is
claimed.

Evaluation is explicit and read-only. NOW/TODAY and omitted-year EASTERSUNDAY
require a validated caller-supplied timestamp. Reference scans retain typed
provider/resource/cancellation/source failures, bounded storage, and complete
holiday/workweek sequence semantics. Scalar arguments remain position-sensitive
under projected evaluation. This batch does not add dependency recalculation,
result spilling, workbook cache publication, or ambient clock services.

[Owned date/time build and capture intermediates were removed](cleanup.json).
Raw final evidence and cited superseded gate histories remain for audit. The
root lockfile was preserved; isolated gates use their retained lockfile.
The separate financial development checkout belongs to the subsequent active
batch and is not part of this completed batch's temporary storage.
