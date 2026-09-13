# Reference operator integration status

This checkpoint retains the private reference-operator module. The public
value evaluator and its worksheet adapter are still being integrated; this
checkpoint does not declare the value API production-ready. The parent module
must be included before this private child is reachable.

The module constructs finite local references, computes one global cuboid for
`:`, retains ordered list entries for `~`, and intersects records in left-major,
right-minor order for `!`. Record ownership follows contiguous physical-plane
offsets. Equal coordinates do not merge duplicate list entries or transfer
planes between overlapping 3-D entries. Empty intersections produce `#NULL!`.
Range construction uses the resolver's physical sheet order directly, avoiding
a repeated scan of operand areas for each sheet.

Independent source review is recorded in
[reference-operators-review.md](reference-operators-review.md), including
historical findings and their resolutions. The retained module SHA-256 is
`1f123b48b9c38b8878c0f26e58c841e08c1bbb019be5086eeda87f0bc5319fb3`.

Root diagnostic integration on 2026-09-13 compiled this module with the
in-progress parent VM. The 35-test value target passed 34 tests, including
duplicate/nested 3-D intersections, all-empty intersections, global range
geometry, explicit extents, cumulative read admission, and cancellation after
provider reads. Its remaining failure was repeated provider reads in nested
lazy shape planning. Earlier diagnostics also passed 390 library tests and
14 worksheet tests, but those runs used an earlier parent VM and are not a
final gate for this checkpoint.

The subsequent 99-case debug value preflight passed, including aggregate
read-count checks up to 4,096 cells. Debug preflight does not establish release
performance. The complete stable-source integration gates, public API review,
boundary checks, release captures, and comparison with the retained scalar
baseline remain required before declaring this feature complete.
