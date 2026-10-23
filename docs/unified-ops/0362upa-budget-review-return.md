# 0362upa — Budget composition returned for correction

Parent: [0362up](0362up-budget-assembly.md). Status: returned; diagnostic
composition is not accepted into the source integration or trunk.

Review of exact committed shared proposal `158b7029f6` found correctness,
API and performance defects that must be repaired under the 0351 contract.
Compiler progress and successful unrelated family reviews do not waive them.

- Cell transfers must preserve pending typed styles and remove moved source
  content. Retained style maps must not recreate a removed source cell.
- Native no-op identity must include rich-text properties. Same visible text
  with changed formatting must remain an edit; genuinely canceled or identical
  edits must release their retention charge. Existing merges, cached values,
  absent cells and bounded/unbounded joins require the same effective-state
  accounting, including failure rollback.
- Metadata admission must charge actual source, target and native patch
  retention before capture. A `2 * source` or `3 * source` substitute can
  reject shrinking edits and miss expansion. The configured public entry
  point is `edit_metadata_with_options(EditOptions)` followed by `commit()`.
- Prepared metadata encoding must reuse admitted lengths/layouts rather than
  decode/scan/preflight the same source again. Default edits must preserve
  their prior path. Planning must release removed relationship IDs in the
  same order as final publication so tight limits remain exact.
- The combined sheet-property/auto-filter facade route must propagate limits
  and include its native patch charge through its intermediate owner.

The separate feature/chart proposal `5ae87e6e08` already projects native
source/target retention before patch capture and passed its scoped review.
It is not replaced by a broader claim based on the older shared leaf. DOCX
source admission and XLSB digest/census fixes also retain their scoped clear
reviews. Final combined compilation, meaningful boundary/fidelity tests and
performance qualification remain required for the complete composition.
