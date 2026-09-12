# 0519 managed DOCX publication profile comparison

This is a matched six-arm × two-repeat Callgrind comparison of the same managed paragraph API routes on the frozen baseline and candidate binaries. All profiles use owned source, one measured sample, zero warmups, one repeat, and CPU 2.

The publication value is inclusive Ir for `Package::publish_document_commit_to_stream`; its direct owner is the `SourceBackedPackage::write_topology_to_stream` edge. Topology and XML-validator values are inclusive `callgrind_annotate --tree=both` function rows. The publication method ends before the caller drops the returned Snapshot, so these are method-scope instruction diagnostics and do not include that caller drop or RSS. The candidate snapshot-owner edge is optional: an optimized hit may inline the source-XML handoff, and the table uses an em dash when no direct semantic snapshot call remains. The candidate XML-validator row is also intentionally absent when the retained source proof is accepted.

Raw parsing and annotation checks pass for all 24 profiles: one positive publication edge with one call, publication incoming Ir equals the raw summary, self plus direct edges equals the publication total, annotation direct edges match raw edges, the direct owner equals topology inclusive Ir, and candidate hit profiles have no positive raw cost in either fresh DOCX scan/build helper. The baseline XML-validator policy is present; the candidate source-proof-reuse policy is absent in both positive raw calls and annotation rows.

| Repeat | Route | Publication inclusive Ir (baseline → candidate) | Direct owner Ir | Snapshot owner Ir | Topology inclusive Ir | XML validator inclusive Ir |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| r1 | p128-k1-owned-batch | 2,520,209 → 1,327,360 (-47.33%) | 2,503,029 → 1,310,262 (-47.65%) | 9,336 → 9,306 (-0.32%) | 2,503,029 → 1,310,262 (-47.65%) | 1,192,927 → — (candidate direct call absent) |
| r1 | p128-k1-owned-repeated | 2,521,057 → 1,326,868 (-47.37%) | 2,503,957 → 1,309,831 (-47.69%) | 9,336 → 9,306 (-0.32%) | 2,503,957 → 1,309,831 (-47.69%) | 1,193,251 → — (candidate direct call absent) |
| r1 | p512-k1-owned-batch | 6,097,386 → 1,371,089 (-77.51%) | 6,077,380 → 1,351,155 (-77.77%) | 9,336 → 9,306 (-0.32%) | 6,077,380 → 1,351,155 (-77.77%) | 4,726,088 → — (candidate direct call absent) |
| r1 | p512-k1-owned-repeated | 6,096,996 → 1,370,389 (-77.52%) | 6,077,069 → 1,350,422 (-77.78%) | 9,336 → 9,306 (-0.32%) | 6,077,069 → 1,350,422 (-77.78%) | 4,726,753 → — (candidate direct call absent) |
| r1 | p512-k32-owned-batch | 6,100,460 → 1,374,167 (-77.47%) | 6,079,684 → 1,353,447 (-77.74%) | 9,790 → 9,760 (-0.31%) | 6,079,684 → 1,353,447 (-77.74%) | 4,726,846 → — (candidate direct call absent) |
| r1 | p512-k32-owned-repeated | 6,100,292 → 1,373,580 (-77.48%) | 6,079,848 → 1,353,278 (-77.74%) | 9,477 → 9,429 (-0.51%) | 6,079,848 → 1,353,278 (-77.74%) | 4,726,623 → — (candidate direct call absent) |
| r2 | p128-k1-owned-batch | 2,519,576 → 1,327,139 (-47.33%) | 2,502,464 → 1,310,114 (-47.65%) | 9,336 → 9,306 (-0.32%) | 2,502,464 → 1,310,114 (-47.65%) | 1,193,023 → — (candidate direct call absent) |
| r2 | p128-k1-owned-repeated | 2,520,010 → 1,327,168 (-47.33%) | 2,502,847 → 1,310,121 (-47.65%) | 9,336 → 9,306 (-0.32%) | 2,502,847 → 1,310,121 (-47.65%) | 1,192,947 → — (candidate direct call absent) |
| r2 | p512-k1-owned-batch | 6,097,474 → 1,371,244 (-77.51%) | 6,077,468 → 1,351,289 (-77.77%) | 9,336 → 9,312 (-0.26%) | 6,077,468 → 1,351,289 (-77.77%) | 4,726,215 → — (candidate direct call absent) |
| r2 | p512-k1-owned-repeated | 6,097,115 → 1,370,009 (-77.53%) | 6,077,026 → 1,350,046 (-77.78%) | 9,336 → 9,306 (-0.32%) | 6,077,026 → 1,350,046 (-77.78%) | 4,726,820 → — (candidate direct call absent) |
| r2 | p512-k32-owned-batch | 6,100,712 → 1,373,727 (-77.48%) | 6,079,905 → 1,352,982 (-77.75%) | 9,790 → 9,780 (-0.10%) | 6,079,905 → 1,352,982 (-77.75%) | 4,726,914 → — (candidate direct call absent) |
| r2 | p512-k32-owned-repeated | 6,100,278 → 1,374,228 (-77.47%) | 6,079,902 → 1,353,240 (-77.74%) | 9,453 → 9,857 (+4.27%) | 6,079,902 → 1,353,240 (-77.74%) | 4,726,640 → — (candidate direct call absent) |

Use the per-route deltas to attribute candidate SourceXmlPart proof reuse and XML-validator elision. A Callgrind reduction is an instruction-count mechanism diagnostic; it does not replace native elapsed-time or RSS guards.

Baseline profiles: `profile-r1/` and `profile-r2/`. Candidate profiles: `profile-after-r1/` and `profile-after-r2/`. Full raw/direct-callee details remain in `profile-analysis.json` and `profile-after-analysis.json`.
