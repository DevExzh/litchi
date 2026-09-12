# 0518 managed DOCX publication profile comparison

This is a matched six-arm × two-repeat Callgrind comparison of the same managed paragraph API routes on the frozen baseline and candidate binaries. All profiles use owned source, one measured sample, zero warmups, one repeat, and CPU 2.

The publication value is inclusive Ir for `Package::publish_document_commit_to_stream`; its direct owner is the `SourceBackedPackage::write_topology_to_stream` edge. Topology and XML-validator values are inclusive `callgrind_annotate --tree=both` function rows. The publication method ends before the caller drops the returned Snapshot, so these are method-scope instruction diagnostics and do not include that caller drop or RSS. The candidate snapshot-owner edge is optional: an optimized hit may inline the source-XML handoff, and the table uses an em dash when no direct semantic snapshot call remains.

Raw parsing and annotation checks pass for all 24 profiles: one positive publication edge with one call, publication incoming Ir equals the raw summary, self plus direct edges equals the publication total, annotation direct edges match raw edges, the direct owner equals topology inclusive Ir, and candidate hit profiles have no positive raw cost in either fresh DOCX scan/build helper. The candidate XML validator retains exactly one positive raw call and one positive final annotation row.

| Repeat | Route | Publication inclusive Ir (baseline → candidate) | Direct owner Ir | Snapshot owner Ir | Topology inclusive Ir | XML validator inclusive Ir |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| r1 | p128-k1-owned-batch | 5,467,227 → 2,520,473 (-53.90%) | 2,504,374 → 2,503,305 (-0.04%) | 2,952,989 → 9,336 (-99.68%) | 2,504,374 → 2,503,305 (-0.04%) | 2,386,179 → 1,193,281 (-49.99%) |
| r1 | p128-k1-owned-repeated | 5,468,556 → 2,519,698 (-53.92%) | 2,503,796 → 2,502,606 (-0.05%) | 2,955,183 → 9,336 (-99.68%) | 2,503,796 → 2,502,606 (-0.05%) | 2,387,027 → 1,192,871 (-50.03%) |
| r1 | p512-k1-owned-batch | 17,747,024 → 6,097,734 (-65.64%) | 6,079,869 → 6,077,740 (-0.04%) | 11,654,429 → 9,336 (-99.92%) | 6,079,869 → 6,077,740 (-0.04%) | 9,455,750 → 4,726,453 (-50.02%) |
| r1 | p512-k1-owned-repeated | 17,743,288 → 6,097,136 (-65.64%) | 6,077,878 → 6,077,200 (-0.01%) | 11,653,260 → 9,336 (-99.92%) | 6,077,878 → 6,077,200 (-0.01%) | 9,455,719 → 4,726,576 (-50.01%) |
| r1 | p512-k32-owned-batch | 17,752,919 → 6,098,980 (-65.65%) | 6,081,470 → 6,078,220 (-0.05%) | 11,658,823 → 9,891 (-99.92%) | 6,081,470 → 6,078,220 (-0.05%) | 9,456,690 → 4,725,193 (-50.03%) |
| r1 | p512-k32-owned-repeated | 17,755,856 → 6,099,746 (-65.65%) | 6,080,685 → 6,078,613 (-0.03%) | 11,662,253 → 9,873 (-99.92%) | 6,080,685 → 6,078,613 (-0.03%) | 9,461,537 → 4,725,873 (-50.05%) |
| r2 | p128-k1-owned-batch | 5,470,253 → 2,521,228 (-53.91%) | 2,504,592 → 2,504,060 (-0.02%) | 2,955,762 → 9,336 (-99.68%) | 2,504,592 → 2,504,060 (-0.02%) | 2,388,327 → 1,193,756 (-50.02%) |
| r2 | p128-k1-owned-repeated | 5,468,080 → 2,520,826 (-53.90%) | 2,504,260 → 2,503,661 (-0.02%) | 2,954,138 → 9,336 (-99.68%) | 2,504,260 → 2,503,661 (-0.02%) | 2,386,475 → 1,194,126 (-49.96%) |
| r2 | p512-k1-owned-batch | 17,746,149 → 6,098,491 (-65.63%) | 6,078,594 → 6,078,497 (-0.00%) | 11,654,851 → 9,336 (-99.92%) | 6,078,594 → 6,078,497 (-0.00%) | 9,454,795 → 4,726,841 (-50.01%) |
| r2 | p512-k1-owned-repeated | 17,743,774 → 6,096,605 (-65.64%) | 6,077,416 → 6,076,585 (-0.01%) | 11,654,362 → 9,336 (-99.92%) | 6,077,416 → 6,076,585 (-0.01%) | 9,456,930 → 4,726,138 (-50.02%) |
| r2 | p512-k32-owned-batch | 17,751,271 → 6,099,762 (-65.64%) | 6,078,719 → 6,078,981 (+0.00%) | 11,659,806 → 9,796 (-99.92%) | 6,078,719 → 6,078,981 (+0.00%) | 9,456,459 → 4,725,979 (-50.02%) |
| r2 | p512-k32-owned-repeated | 17,752,841 → 6,101,321 (-65.63%) | 6,078,682 → 6,080,783 (+0.03%) | 11,661,135 → 9,429 (-99.92%) | 6,078,682 → 6,080,783 (+0.03%) | 9,455,337 → 4,727,316 (-50.00%) |

Use the per-route deltas to attribute the candidate snapshot-reuse path. A Callgrind reduction is an instruction-count mechanism diagnostic; it does not replace native elapsed-time or RSS guards.

Baseline profiles: `profile-r1/` and `profile-r2/`. Candidate profiles: `profile-after-r1/` and `profile-after-r2/`. Full raw/direct-callee details remain in `profile-analysis.json` and `profile-after-analysis.json`.
