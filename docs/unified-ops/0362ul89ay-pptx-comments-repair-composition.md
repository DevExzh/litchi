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

Independent source review is READY. Root checks formatting on all five
changed Rust paths and checks the staged diff. No Cargo, runtime, native or
performance execution accompanies this merge. The previous 1,374-pass /
12-failure runtime capture remains historical; a fresh source-bound compiler
packet and full runtime qualification are required. The bounded physical
attribute scan still needs performance measurement, with no exception accepted.

The preceding 89ax formatter-only correction is retained. The isolated 89aw
baseline must contain the identical normalized harness before paired builds.
This is private composition, not final acceptance or publication.
