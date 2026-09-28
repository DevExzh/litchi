# 0813 numerical results and validator review

The terminal 0813 packet contains 306 main reports and 6,714 main samples:
216 native reports, 72 allocation reports, and 18 before-only qualification
reports. The independent workflow replay and raw numerical audit both pass.

## Frozen policy result

The adoption guards are eligible: six capture or lifecycle rows improve by at
least three percent, each of those six intervals has a bootstrap high endpoint below one,
there are no latency vetoes, and all 144 allocation guard comparisons are
equal before and after for allocation calls, allocated bytes, net live bytes,
and peak above entry.

| case | paired p50 ratio | improvement | bootstrap 95% interval |
| --- | ---: | ---: | ---: |
| large/capture | 0.894555582 | 10.544442% | [0.882107453, 0.899554194] |
| large/lifecycle | 0.927312972 | 7.268703% | [0.920132154, 0.933996251] |
| medium/capture | 0.952519920 | 4.748008% | [0.947183557, 0.959293817] |
| vendor/capture | 0.961293756 | 3.870624% | [0.954688857, 0.970223695] |
| unicode-vendor/capture | 0.968801083 | 3.119892% | [0.962577353, 0.972042754] |
| valid-4attr/capture | 0.966499040 | 3.350096% | [0.959556158, 0.970461079] |

The result is a policy result for this 18-row PPTX probe. It does not imply a
universal workload speedup or a cross-format result.

## Tail and RSS diagnostics

The p50 spread review has no spread flag above five percent. There are nine
p99 spread flags above five percent: medium/lifecycle after (28.509%),
tiny/commit after (10.413%), tiny/commit before (9.360%),
unicode-vendor/commit before (7.258%), valid-4attr/lifecycle after (10.952%),
valid-4attr/lifecycle before (15.687%), vendor/capture before (31.265%),
vendor/commit after (7.672%), and vendor/commit before (5.406%). The three
groups with a p99 block regression flag are tiny/commit (median ratio
0.999073), medium/lifecycle (1.019671), and valid-4attr/lifecycle (0.992855);
their largest individual block increases are 10.328%, 26.331%, and 8.553%,
respectively. These are diagnostic tail observations and are not adoption
guards.

There are nine RSS spread flags above five percent: medium/capture after
(8.840%), medium/lifecycle before (7.482%) and after (7.566%), tiny/capture
before (6.939%) and after (5.475%), tiny/commit before (9.579%) and after
(8.237%), and tiny/lifecycle before (6.504%) and after (8.656%). RSS remains
review-only evidence; the packet makes no RSS reduction claim.

The five paired RSS increases above five percent are medium/lifecycle block 0
(5168 to 5452 KiB, 5.495%), medium/lifecycle block 3 (5132 to 5516 KiB,
7.482%), tiny/capture block 2 (4900 to 5240 KiB, 6.939%), tiny/commit block 4
(4856 to 5112 KiB, 5.272%), and tiny/lifecycle block 3 (4920 to 5176 KiB,
5.203%). No paired RSS median increase exceeds five percent, and the
allocation lane has no regression or spread flag.

## Validator repairs

The aggregate validator now follows the retained quality-summary schema. It
binds each probe's retained `inputs.json` and `receipts.json`, checks the
architecture and unrelated-file custody from `inputs.json`, and no longer
expects those fields at the probe summary level. It also replays
`analysis.py --qualification --check` and requires each codegen gate's
top-level `before` and `after` descriptors to equal the corresponding receipt
entries. The immutable quality-summary artifact and its generator were not
changed.

The offline checks completed after these repairs:

* `analysis.py --check`: 306 reports / 6,714 samples.
* `analysis.py --qualification --check`: 18 reports / 18 samples.
* `root_audit.py --check`: pass.
* `quality_summary.py --check`: pass.
