# 0781 results audit

This is an independent read-only audit of the retained final packet after
target cleanup. `validate.validate(require_final_seal=False)` passed, the
fresh `analyze.analyze()` result equals `analysis.json`, and `tables.py --check`
passed.

- The primary matrix has `6 × 10 × 2 = 120` native children, `2 × 10 × 2 =
  40` allocation children, and 10 before-only qualification children. Native
  has 20 groups with six processes each and six paired blocks; allocation has
  20 groups with two blocks of three samples each and 13 allocation fields.
- All 10 native and 10 allocation pair groups have the expected block and
  metric cardinalities. All 560 paired rows preserve the defined/undefined
  ratio partition. Every zero/zero row uses ratio 1 and zero change. Paired
  bootstrap metadata is median, 10,000 resamples, seed 781078, confidence
  0.95.
- Native flags are 42 within-group spread flags and 20 paired block-trigger
  flags. Allocation has zero spread or paired flags. Each reported flag has
  an underlying block or distribution spread above 5%; paired flags are
  emitted when any block crosses the threshold.
- The raw authored-text oracle checks pass for all 170 retained reports and
  3,730 samples, including raw bytes, digests, text-box counts, and match
  booleans.
- Candidate disposition is rejected. The restored source manifest equals the
  before-build manifest, the live production source hashes equal that
  manifest, and every before/candidate archive receipt hashes correctly.
  Cleanup records seven executable witnesses, including the two distinct
  initial-binary relocation witnesses; the reused `before-*` paths were not
  treated as initial binaries.

No arithmetic, ratio, raw-oracle, custody, or restoration discrepancy was
found.
