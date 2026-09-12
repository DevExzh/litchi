# 0517 managed DOCX publication profile comparison

This is a matched six-arm × two-repeat Callgrind comparison of the same managed paragraph API routes on the frozen baseline and candidate binaries. All profiles use owned source, one measured sample, zero warmups, one repeat, and CPU 2.

The publication value is inclusive Ir for `Package::publish_document_commit_to_stream`; its direct owner is the `SourceBackedPackage::write_topology_to_stream` edge. Topology and XML-validator values are inclusive `callgrind_annotate --tree=both` function rows. The publication method ends before the caller drops the returned Snapshot, so these are method-scope instruction diagnostics and do not include that caller drop or RSS.

Raw parsing and annotation checks pass for all 24 profiles: one positive publication edge with one call, publication incoming Ir equals the raw summary, self plus direct edges equals the publication total, annotation direct edges match raw edges, and the direct owner equals topology inclusive Ir.

| Repeat | Route | Publication inclusive Ir (baseline → candidate) | Direct owner Ir | Topology inclusive Ir | XML validator inclusive Ir |
| --- | --- | ---: | ---: | ---: | ---: |
| r1 | p128-k1-owned-batch | 6,660,412 → 5,467,082 (-17.92%) | 3,697,297 → 2,503,886 (-32.28%) | 3,697,297 → 2,503,886 (-32.28%) | 3,579,651 → 2,386,994 (-33.32%) |
| r1 | p128-k1-owned-repeated | 6,661,889 → 5,467,981 (-17.92%) | 3,699,308 → 2,503,148 (-32.33%) | 3,699,308 → 2,503,148 (-32.33%) | 3,581,129 → 2,386,603 (-33.36%) |
| r1 | p512-k1-owned-batch | 22,476,195 → 17,749,121 (-21.03%) | 10,805,246 → 6,079,188 (-43.74%) | 10,805,246 → 6,079,188 (-43.74%) | 14,181,738 → 9,455,917 (-33.32%) |
| r1 | p512-k1-owned-repeated | 22,472,859 → 17,743,953 (-21.04%) | 10,806,300 → 6,077,118 (-43.76%) | 10,806,300 → 6,077,118 (-43.76%) | 14,183,597 → 9,455,739 (-33.33%) |
| r1 | p512-k32-owned-batch | 22,485,900 → 17,749,977 (-21.06%) | 10,805,168 → 6,079,334 (-43.74%) | 10,805,168 → 6,079,334 (-43.74%) | 14,186,945 → 9,455,704 (-33.35%) |
| r1 | p512-k32-owned-repeated | 22,479,305 → 17,749,407 (-21.04%) | 10,803,895 → 6,078,607 (-43.74%) | 10,803,895 → 6,078,607 (-43.74%) | 14,180,807 → 9,454,859 (-33.33%) |
| r2 | p128-k1-owned-batch | 6,661,707 → 5,467,378 (-17.93%) | 3,698,264 → 2,503,468 (-32.31%) | 3,698,264 → 2,503,468 (-32.31%) | 3,579,277 → 2,386,227 (-33.33%) |
| r2 | p128-k1-owned-repeated | 6,660,110 → 5,466,770 (-17.92%) | 3,696,395 → 2,504,078 (-32.26%) | 3,696,395 → 2,504,078 (-32.26%) | 3,579,164 → 2,386,448 (-33.32%) |
| r2 | p512-k1-owned-batch | 22,473,282 → 17,749,056 (-21.02%) | 10,806,228 → 6,079,840 (-43.74%) | 10,806,228 → 6,079,840 (-43.74%) | 14,182,475 → 9,456,714 (-33.32%) |
| r2 | p512-k1-owned-repeated | 22,473,119 → 17,749,165 (-21.02%) | 10,807,082 → 6,081,801 (-43.72%) | 10,807,082 → 6,081,801 (-43.72%) | 14,183,801 → 9,457,268 (-33.32%) |
| r2 | p512-k32-owned-batch | 22,484,846 → 17,754,387 (-21.04%) | 10,808,579 → 6,081,221 (-43.74%) | 10,808,579 → 6,081,221 (-43.74%) | 14,190,157 → 9,457,106 (-33.35%) |
| r2 | p512-k32-owned-repeated | 22,473,401 → 17,748,828 (-21.02%) | 10,804,312 → 6,078,714 (-43.74%) | 10,804,312 → 6,078,714 (-43.74%) | 14,180,105 → 9,454,163 (-33.33%) |

The candidate reduces publication inclusive Ir by about 17.9% on p128 K=1 and about 21.0% on p512 K=1/K=32. The direct topology owner falls about 32.3% on p128 and 43.7% on p512, while inclusive `validate_source_xml` falls about 33.3% on every arm. These are consistent with removing one complete source XML validation pass; they do not replace native elapsed-time or RSS guards.

Baseline profiles: `profile-r1/` and `profile-r2/`. Candidate profiles: `profile-after-r1/` and `profile-after-r2/`. Full raw/direct-callee details remain in `profile-analysis.json` and `profile-after-analysis.json`.
