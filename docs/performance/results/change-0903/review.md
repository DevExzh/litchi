# Independent 0903 review

The retained offline evidence is internally consistent. This review used the
four frozen `callgrind.out.6` files and the retained raw reader only; it ran no
build, workload, profiler, or native measurement.

The independent raw parse reproduced the reducer's key invariants for all four
captures:

- each profile's `summary: Ir` equals `totals: Ir` and the sum of all function
  exclusive self rows;
- the measured publication owner has exactly one positive incoming edge from
  `run_xlsx_cell_values_edit_save`, with one call, and its self plus direct
  children equals both that edge and the profile summary;
- each selected target function's self plus positive-cost direct edges equals
  its positive-cost incoming edges, while zero-cost edge records remain
  retained as raw evidence; and
- source-file-classified XML self cost plus the disjoint non-XML remainder
  equals the publication summary. No nested inclusive row was added to that
  partition.

The hot-loop interpretation is appropriately bounded. In the first medium
profile, the `verify_source_replacement` edge to `Reader::read_event_impl`
records 46,502 calls and 13,448,131 Ir, while the corresponding
`verify_with_policy` edges record 116,296 calls and 13,448,055 Ir. The
`inspect_attributes` edges likewise record 18,642 calls / 6,504,477 Ir versus
46,631 calls / 6,507,004 Ir. The cause of these near-equal-cost, different-call
records is unproven. They cannot support a tag-cardinality census, branch-
coverage claim, or cost-per-call estimate. The source audit correctly requires
independent cardinality coverage before a candidate.

One source-audit wording issue was corrected during review: 0902 did not make
the observer stop; it added a branch on the existing `search.stopped` flag and
bypassed the observer path when that flag was already set. The corrected text
keeps the original-first and semantic-parity constraints intact.

The limit-order wording was also corrected: `check_attribute_layout` checks
before reading the next name and reports at `inner_offset + cursor`, while
`inspect_attributes` reports its later aggregate/per-element charge at the tag
offset.

Disposition: **pass with the stated limitations**. The zero/one-attribute
summary remains a future qualification hypothesis only. No production change
or performance claim is supported by this packet.
