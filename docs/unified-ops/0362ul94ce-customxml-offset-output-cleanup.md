# 0362ul94ce — Custom XML offset-output cleanup

This source-only follow-up fixes the compiler-facing issues found after the
94cd cleanup at `cfc2b62fe4d235617bd366cf368eb41022069e9b`.  It makes no
acceptance or qualification claim.

The Custom XML properties layout reader now converts each selected `u32`
source offset to `usize` before looking it up in the source-element table, in
both start-element and empty-element paths.  Conversion overflow remains an
internal invariant failure rather than a truncating lookup.

The legacy marker-based `active_offsets` implementation now returns its
selected offsets directly.  Its unreachable retained-XML/source-offset/marker
scaffolding and the corresponding `ActiveOffsetOutput` fields are removed;
the public legacy algorithm and results are unchanged.  The side-table helper
retains only the XML view and selected-offset data required by the Custom XML
layout projection.

The complete source audit covers all affected types, constructors, imports,
and consumers.  Only rustfmt, diff, and repository size/binary guards are run
for this return; Cargo, rustdoc, compiler, native, and performance execution
remain pending.
