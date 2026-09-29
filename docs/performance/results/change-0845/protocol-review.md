# Independent main-matrix protocol review

The reviewer found no blocker. The frozen plan has 32 native cases, six matched
blocks and two legs (384 native reports), plus 64 qualifications and 64
Heaptrack reports. Tiny cases contribute 192,000 native samples and large cases
19,200; with 128 qualification/Heaptrack samples the total is 211,328. Every
case has three before-first and three after-first pairs. Heaptrack is a
separate publish-only population with two repeats per shape/count/leg.

Build, capture, qualification, analysis and admission bind the same plan from
the outset. Probe and independent reader/test sources are byte-identical to
0844. Derived-summary compaction occurs only after full reader validation;
raw reports, ZIPs and interpreted Heaptrack traces remain available. No check
is bypassed by omitting duplicate decoded manifests, samples and span tables
from summaries.

The inherited reader tests retain their original schema's total constant;
0845 admission and analysis independently enforce this packet's 512 reports /
211,328 samples. Baseline gates, 32 archive qualifications, observer preflight
and prior closure replay passed.

The selected topology-limit function executes twice; unrelated overlay-limit
scans remain out of scope. Candidate review must preserve physical order,
source/progress checks and all loop boundaries. The new iterator is hidden
ZIP-owner plumbing, with no physical ID exposed through ordinary semantic
CRUD APIs. Separate instruction profiles require their own frozen plan and
custody review before measurement.
