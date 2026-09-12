# 0528 bounded scanner ceiling audit

This is a read-only audit of the four sealed 0525 candidate Callgrind dumps.
The analyzer is [analyze.py](./analyze.py), SHA-256
`3f11fe7d8e38089b4f55417e4302d848937d792ce053ff3ebc06699ec2346343`; its
replayed report is [profile-analysis.json](./profile-analysis.json), SHA-256
`46441db4b084576712e40f8c59aa729d5edce703e0b8d467220bea3a32dbbcbc`.
The 0525 550-entry seal, retained 0521/0519 helpers, 0526 decomposition, and
the 0528 source/dependency binding all pass.  The accepted 0525 release-path
source hashes remain unchanged; 0527's retained changes in
`cell_values/snapshot.rs` are test-only.

The selected commit totals 782,799,138 Ir across the four profiles.  The
scanner totals 416,939,582 Ir, or 53.262652% of commit Ir.  The direct raw
scanner → `quick_xml::name::NamespaceResolver::resolve_event` edge totals
31,526,144 Ir:

| repeat | shape | commit Ir | scanner Ir | resolver edge Ir | edge / scanner | edge / commit |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 1 | medium | 133,217,968 | 70,969,669 | 5,386,720 | 7.590172% | 4.043539% |
| 1 | dense-sparse | 258,175,362 | 137,497,129 | 10,376,352 | 7.546595% | 4.019110% |
| 2 | medium | 133,250,185 | 70,995,579 | 5,386,720 | 7.587402% | 4.042561% |
| 2 | dense-sparse | 258,155,623 | 137,477,205 | 10,376,352 | 7.547689% | 4.019417% |
| **aggregate** | | **782,799,138** | **416,939,582** | **31,526,144** | **7.561322%** | **4.027361%** |

The resolver number is the complete scanner edge and is therefore an upper
bound for an End-only resolver optimization.  It includes the required
Start and Empty event lookups; the retained profiles do not isolate a
removable End subset.  Call metadata is collection-off polluted, so the
analyzer omits call counts and performs no per-call split or estimate.

Within the retained scanner subtree, immediate children of `start_cell` rank
as follows.  The rows are disjoint only at that immediate parent; nested
inclusive costs are not added again to the scanner total.

| rank | immediate child | aggregate Ir | share of `start_cell` | share of commit |
| ---: | --- | ---: | ---: | ---: |
| 1 | `Scanner::cell_address` | 79,917,219 | 63.274575% | 10.209160% |
| 2 | `wire::cell_tag` | 43,468,178 | 34.415994% | 5.552916% |

`cell_address`'s own immediate children are `unqualified_attribute_value`
(64,582,435 Ir; 80.811665%), `parse_a1` (9,503,978 Ir; 11.892278%), and
`__rustc::__rust_dealloc` (1,995,670 Ir; 2.497171%).  These are subcosts of
`cell_address`, not additional scanner costs.  All direct annotation edges
for `scan_with_limit`, `start_cell`, and `cell_address` match their raw
Callgrind edges, and each inclusive = self + direct-child equation passes.

For publication prioritization, the 0527 sealed native primary vectors were
also reduced by summing every phase sample before division.  The denominator
is `Σ_i(open_i + plan_i + commit_i + publication_i)`, which reproduces the
reported elapsed mean; p50 values were not summed.  Reopen remains outside
this timed operation.

| repeat | shape | mean total ns | open | plan | commit | publication |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 1 | medium | 22,204,388.590 | 0.405673% | 32.552085% | 36.320072% | 30.722170% |
| 1 | dense-sparse | 42,792,348.210 | 0.236824% | 32.045814% | 35.611748% | 32.105614% |
| 2 | medium | 22,171,755.020 | 0.398873% | 32.237479% | 36.656785% | 30.706862% |
| 2 | dense-sparse | 42,310,837.445 | 0.219162% | 31.839927% | 35.550281% | 32.390630% |

The `start_cell` ranking is a next ROI only inside this retained scanner
subtree.  This audit does not rank the separate 0528 publication-owner
profile or turn any Callgrind value into a latency claim; that owner-level
review should decide the overall publication priority.
