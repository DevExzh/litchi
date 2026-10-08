# 0968 retained 0942 CFB commit attribution

This is an offline reduction of retained 0942 frame-pointer `cycles:u`
callchains. It makes no build, workload, latency, allocation, or adoption
claim. All outputs are outside the repository.

## Qualified scope

The exact physical root is `litchi_ole_common::object::editor::Editor::commit_candidate_with_rendered` nested below the exact physical
barrier `doc_retention_probe_0942_cfb_profile::profiled_measured_case` in leaf-first callchains. The FloatPictures/e32 focus
denominator is **1,477 samples / 6,302,069,415 period**
across repetitions r01 and r02. No selected record was truncated and no
lost record was reported. The all-formal denominator is
2,352 samples / 10,037,338,577 period.

The marker parser reproduces the retained formal reduction for all 16
children before the commit subset is applied. Source equivalence uses
byte SHA-256 comparisons of the retained runtime files against the current
checkout at HEAD `6493e26155ee05e554f1371dd2fad707cd4bca0e`; the retained
revision string is metadata only.

## FloatPictures/e32 focus

The percentages below use the commit denominator above. Inclusive rows can
overlap and must not be summed. The exclusive table assigns each sample to
the nearest observed physical target frame and conserves the denominator.

| inclusive physical target | samples | period | % of commit |
|---|---|---|---|
| Package::capture | 370 | 1578957556 | 25.054588 |
| Package::capture_with_budget | 0 | 0 | 0.0 |
| capture_container | 0 | 0 | 0.0 |
| OleFile::open_stream | 205 | 874651929 | 13.878805 |
| OleFile::open | 210 | 896459269 | 14.22484 |
| object::codec::open | 0 | 0 | 0.0 |
| discovery::from_package | 0 | 0 | 0.0 |
| OleWriter::write_to | 558 | 2381526914 | 37.789601 |

For this focus, Package::capture is 25.054588% inclusive; its
open_stream subset is 13.878805%, so adding those rows double counts
874,651,929 period (205 samples). The exact-target union is
1,138 samples / 4,856,943,739 period
(77.069030%); any physical
`litchi_cfb::` frame covers 1,124 samples /
4,796,517,962 period (76.110205%).

### Disjoint nearest-frame buckets

| exclusive bucket | samples | period | % of commit |
|---|---|---|---|
| cfb_ole_open | 210 | 896459269 | 14.22484 |
| cfb_open_stream | 205 | 874651929 | 13.878805 |
| cfb_writer_write_to | 558 | 2381526914 | 37.789601 |
| no_target_cfb_frame | 225 | 959277018 | 15.221619 |
| other_physical_cfb | 114 | 485848658 | 7.709351 |
| package_capture | 165 | 704305627 | 11.175783 |

The retained no-inline export has no physical frame named
`capture_container`, `object::codec::open`, or
`discovery::from_package` inside this commit subset. Their zero named
samples are not zero CPU proof: these helpers may be short, inlined, or
missed by sampling. The source path is confirmed separately by the
editor and codec source lines recorded in the JSON receipt.

## Replicate split

| child | commit samples/period | capture | open_stream | writer |
|---|---|---|---|---|
| profile-r01-docfloat-e32 | 731/3112476761 | 26.286683% | 13.821989% | 38.443945% |
| profile-r02-docfloat-e32 | 746/3189592654 | 23.852283% | 13.934247% | 37.151078% |

The earlier 0942 selected-stage CFB writer observations (36.8% and
38.5% for FloatPictures/e32) use the outer measured-case marker. The
rows here use the narrower commit root, so they are not interchangeable
denominators. The shared DOC/CFB source files are byte-identical to the
current 6493 checkout; no 0968 build or capture was performed.

## Limits

These are sampled callchain memberships. Inclusive ancestry is not self
cost, and exclusive nearest-frame buckets exclude callee work from the
leaf owner. Missing named frames do not establish absence. Any candidate
still requires a fresh bound ordinary build, correctness gates, and a
matched profile.
