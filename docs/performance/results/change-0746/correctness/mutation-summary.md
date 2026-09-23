# Mutation differential summary (change 0746)

`xls_mutation_corpus` (source: `../probe/mutate.rs`), `MUTATION_ROUNDS=200`, over the
126 fixtures in `fixtures.txt`; 93 fixtures have a Workbook stream the generator
can frame and repackage. Round 0 of each is the repackaged unmodified stream;
every fourth round stacks a second mutation. Each row is one package and four
owner outcomes. The full outputs of the base leg, the determinism-fix-only leg
and the final leg have one SHA-256 (see `output-sha256.txt`), so every row below
is identical in all three legs; the outputs themselves (9.4 MB each) are not
retained because the generator is deterministic.

| mutation | packages | `cell_values` refused | comments refused | visibility refused | public reader refused |
| --- | ---: | ---: | ---: | ---: | ---: |
| bit-flip | 2886 | 1837 | 1779 | 1767 | 320 |
| collide | 1490 | 636 | 650 | 644 | 188 |
| swap | 1479 | 822 | 834 | 813 | 163 |
| stray-string | 1439 | 1314 | 1336 | 1314 | 149 |
| duplicate-later | 1435 | 658 | 670 | 655 | 149 |
| xf-past-end | 1426 | 1278 | 1303 | 1278 | 171 |
| outside-grid | 1424 | 1309 | 652 | 643 | 179 |
| duplicate-in-place | 1415 | 606 | 617 | 598 | 188 |
| truncate | 1405 | 1163 | 1173 | 1161 | 152 |
| globals-bit-flip | 1399 | 826 | 841 | 835 | 532 |
| companion-repeat | 1367 | 618 | 621 | 624 | 159 |
| drop | 1342 | 697 | 704 | 690 | 149 |
| repackaged-unmodified | 93 | 35 | 37 | 36 | 11 |
| **total** | **18600** | **11799** | **11217** | **11058** | **2510** |
