# 0362ul137b — DrawingML v2 profile-contract repair

This source-only follow-up preserves `b2bfc3a32a276040f23cb9b286367f25b18f69b1`
and restores the per-profile source and feature checks present in the reviewed
v1 coordinator.  It does not run Cargo, a native probe, an allocator capture,
or a performance workload.

The v2 master validator now requires each profile to carry a nonempty reviewed
parent list containing its selected baseline, validates optional
`route_baselines` (including the exact theme parent for theme/Ink routes), and
retains the theme-family caller-limit/current-feature compatibility rules.  All
four detached-owner routes explicitly select the complete six-feature current
DrawingML set; omitting that field now refuses instead of silently compiling
only `drawingml-owner-current`.

Regression coverage rejects empty or mismatched parent lists, malformed and
wrong theme route parents, incompatible theme-family baseline features, and
missing owner feature coverage.  The focused source-only validation passes 44
tests including these profile-contract cases.
