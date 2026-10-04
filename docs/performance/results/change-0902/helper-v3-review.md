# 0902 helper v3 amendment review

Status: **PASS (static diff and source-scope review)**

This review covers the qualification analyzer correction after the v2
qualification capture completed but analysis rejected the one-edit XLSX
control. No helper or workload was executed for this review.

## Bound artifacts

| Artifact | SHA-256 |
| --- | --- |
| `analyze_v2.py` | `157ab8b51fc10b2a7c1dd1448dc9dad3b03e9676d240babf25c99d1ced1e742e` |
| `analyze_v3.py` | `4d96e5ac9031a60c6a5b7024cf8aaefd05bd12280760b9f3a1cf6d01bd72829b` |
| `helper-v3-amendment.json` | `40c68a7d5e37d42080018dd101857d6cde54fe04c6f7f916c7e88e53b803b51a` |

## Diff review

The v3 diff is confined to `source_identity` in `analyze.py`. It removes
`unselected_worksheet_read_calls` and `unselected_worksheet_read_bytes` from
the common zero-valued source invariant list. It then branches on the existing
`one_edit` argument:

- one-edit controls require exactly three unselected worksheet compressed-range
  intersections and positive intersection bytes;
- all-selected primary cases retain the exact zero calls and zero bytes
  requirement.

The amendment changes no plan, case IDs, child schedule, sample counts,
statistical thresholds, source/binary bindings, or native admission rule.
The failed v2 receipt records `inputs_unchanged: true`; the v3 receipt is an
exit-0 qualification analysis with the same source hashes and helper command
environment.

## Raw workload-source scope

The workload source supports this distinction for the frozen corpus:

- `xlsx_cell_crud_updates_for_case` selects one coordinate for the one-edit
  case and 93/178 coordinates for the two primary percent cases.
- `xlsx_cell_values_tracked_ranges` derives the selected worksheet set from
  those coordinates. The medium corpus has four worksheets, so one-edit selects
  one and tracks the other three as unselected; the percent cases select all
  four.
- `InstrumentedSource::read_at` increments the unselected counter only for a
  nonzero overlap with those tracked compressed member ranges. The source
  comments and `tools/perf-baseline/README.md` classify this as raw compressed
  publication overlap, not semantic or decompressed worksheet reads.
- Existing workload tests assert one-edit selected-range cardinality 1 and
  unselected cardinality `sheet_count - 1`, while all-sheet batches have no
  unselected range.

The retained qualification reports match that source scope for both roles and
both samples: one-edit reports have
`unselected_worksheet_read_calls = [3, 3]` and
`unselected_worksheet_read_bytes = [20330, 20330]`; primary reports retain
`[0, 0]` for both fields. The v3 output contains eight passing rows, and the
baseline/candidate one-edit source identities remain equal.

## Review conclusion

The v3 change corrects an overgeneralized zero-read assumption using the
existing case distinction and the workload's documented compressed-member
accounting. It preserves the stronger all-selected invariant and all
cross-role identity checks. The corrected analyzer is suitable for the frozen
native run, subject to root's normal execution and unchanged gates.
