# 0362ul94cd — Custom XML compiler-facing cleanup

This source-only return follows the 94cc typed-limit and provenance repair at
`032c8196cdcc866b4b344ef9324f9c4a6b827eb8`.  It makes no acceptance or
qualification claim.

The legacy marker-based active-offset helper is private to the MCE codec; its
unused sibling-module re-export is removed.  The Custom XML properties layout
reader uses the direct side-table `selected_offsets` contract exclusively, so
its unreachable marker-comment and `source_offsets` branches are deleted.
The legacy public active-offset behavior itself is unchanged.

The source audit covers all new helper call arities and compiler-facing uses.
Only rustfmt, diff, and repository size/binary guards are run for this return;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
