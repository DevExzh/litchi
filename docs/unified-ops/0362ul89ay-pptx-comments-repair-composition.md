# 0362ul89ay — Reviewed comments repair composition

This private merge joins independently reviewed comments repair
`43dbfabf2025d2913f99649b1bb170a8550645e3` (tree
`909eadd556814db3448360bae53dee0c22329ff0`) with the notes, downstream
consumer and measurement-harness candidate at `19c8f3d2bc`.
Both parent histories are retained; the merge has no conflicts.

The complete 89ar–89are chain restores physical admission before MCE,
checks selected fields before ownership, includes date-time, shares the
MCE source, and matches aggregate and physical writer limits to readers.
The final correction counts LF/CR numeric entities as five wire bytes;
exact/one-under writer and reader round-trip regressions cover the boundary.
The canonical 0362ul62 contract is updated with the reviewed source basis.

Independent source review is READY. The initial root formatting check finds
layout differences in four of five changed Rust paths. The composition commit
was created before that failed check was handled; its original passing claim
was incorrect. This follow-up applies only the repository formatter output
and reruns formatting successfully on all five paths; the diff check passes. No Cargo, runtime, native or
performance execution accompanies this merge. The previous 1,374-pass /
12-failure runtime capture remains historical; a fresh source-bound compiler
packet and full runtime qualification are required. The bounded physical
attribute scan still needs performance measurement, with no exception accepted.

The preceding 89ax formatter-only correction is retained. The isolated 89aw
baseline must contain the identical normalized harness before paired builds.
This is private composition, not final acceptance or publication.

The formatter-only delta is retained externally at
`.coord/composition-format-89ay-94u/uo2-pptx-combined-0362ul89.patch` (10775 bytes, SHA-256
`0f9e8f95027a36fb9f3acf072f1ef415bdb2d4df64ff2b561e3fcda3ed69be09`).
