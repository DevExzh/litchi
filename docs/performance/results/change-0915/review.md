# 0915 post-capture review

This review covers the retained 8-pair ABBA native run, its reduction, and the
retained Callgrind profile. The candidate remains limited to borrowing the
`quick_xml` namespace resolver in `ensure_source_section_inventory_xml`.

The native replay passed, and an independent read of all 16 compressed child
reports found 200 samples per child (3,200 samples total), with every report
bound to the frozen case, warmup/sample counts, binary identity, phase-sum,
source/cache/sink, preservation, and refusal gates. The eight matched
baseline-p50/candidate-p50 ratios were
`1.0148165152, 1.0137315074, 1.0112378768, 1.0145722326,
1.0155353703, 1.0132811882, 1.0139181775, 1.0132226796`.
The median ratio was `1.0138248425`; the fixed-seed 10,000-draw nearest-rank
95% interval was `[1.0132226796, 1.0148165152]`. The lower bound clears 1.0,
but the median is below the required 1.05, so the frozen timing gate is false
and the candidate is ineligible for adoption.

Independent pooled elapsed statistics were baseline p50/p95/p99
`4,813,090/4,851,850/4,876,500 ns` and candidate
`4,747,636/4,782,320/4,810,723 ns`; candidate-over-baseline changes were
`-1.3599%/-1.4331%/-1.3489%`. Whole-process maximum-RSS medians were
`37,676 KiB` baseline and `38,378 KiB` candidate (`+1.8633%`). These values
are descriptive only: the quality boundary checker occupied CPUs 14/15 during
the approximate `15:41:01.976978`–`15:41:41.871148 UTC` window, overlapping
seven early children (both legs of pairs 1–3 and the first leg of pair 4).
All rows were retained, with no post-hoc exclusion, rerun, or change to resampling;
there is no clean-isolation claim.

The profile replay passed with 11 artifacts checked. Three corruption tests
used temporary copies and left retained files untouched: changing the copied
profile line proof was rejected with `source clone line changed`; changing a
copied native reduction p50 was rejected with `retained reduction differs from
deterministic recomputation`; and flipping one byte in a copied child report
was rejected with `compressed report changed since freeze`.

The manifest descriptors still match the baseline (69,940,744 bytes),
candidate (69,938,512 bytes), frozen source archive (171,120,640 bytes),
protocol, comparison script, and validator. The final boundary receipt exited
zero (`65` workspace packages, `244` internal declarations, `11` explicit debt
items). The native gate failure and the host-overlap limitation together give
no basis for production adoption.
