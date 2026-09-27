# 0791 — current PPTX capture attribution

The current generated large-PPTX capture still spends substantial sampled
work in notes XML scanning after the retained 0785 namespace optimization.
The exact public capture owner appears in 1,066 and 1,068 native sample stacks;
935 and 929 of those include the notes scanner. Checked attribute iteration
appears in 186 and 148, while the namespace helper appears in only 5 and 4.
These observed counts guide the next work-removal experiment; they are not
native phase fractions or an optimization result.

The next narrow candidate is to skip construction/advancement of the checked
attribute iterator for an exactly empty raw attribute tail, **after** the
existing namespace, local-name, and root checks. Whitespace tails and all
attribute-bearing tags retain the existing checked path. This packet supplies
a current baseline and source review; it does not implement or adopt that
candidate. The rejected 0787 cache candidate remains unchanged.

## Experiment and custody

The source revision is `10b1e59c0c`. The 0784 capture probe is reused on current
production source with its original wrapper and schema names. Its clock wraps
only `Package::opened_presentation`; borrowed package construction, fixture
construction, serialization, readback, and retained-owner destruction remain
outside that interval. The generated fixture and exact source/output/semantic
oracles are inherited from sealed earlier evidence, not historical timing.

Tiny is 3 slides × 4 shapes, medium 12×8, and large 100×100. The host remains
the 32-CPU EPYC 9R45 Linux environment; captures use CPU 12. The
[packet](results/change-0791/README.md) retains current toolchain and tool
versions, CPU/memory/affinity state, source hashes, commands, and raw outputs.
Root ran all builds and captures serially. Source review and offline analysis
were delegated separately.

There are three distinct lanes:

| Lane | Processes | Measured outputs | Purpose |
| --- | ---: | ---: | --- |
| Native control/profile wrapper, six alternating blocks × three shapes | 36 | 1,080 | Current capture latency and wrapper/code-generation perturbation |
| Callgrind, two ordered passes × three shapes | 6 | 6 | Exact wrapper-scoped guest instructions |
| Separate frame-pointer native profiles, large shape | 2 | 200 | Exact-owner sampled ancestry |

Native control processes use 30 samples after three warmups. Callgrind uses
one sample without warmup, collecting only the named wrapper and requiring one
positive numbered dump plus an empty termination dump. Native profiles use
100 samples after three warmups, `cycles:u`, 499 Hz, and frame-pointer stacks.
The frame-pointer binary is a separate diagnostic build, not a control latency
leg. Sampled stacks include warmup and measured captures.

All **44 reports and 1,286 measured publications** pass source/output/semantic
identity checks. These are warm in-memory generated shapes, not native Office
producer, source-backed, physical-cold, range-source, concurrent, or full CRUD
coverage. No new allocation, copied-byte, or decompression metric is claimed.

## Native controls and limitations

Displayed times are median per-process nearest-rank p50 values. Profile/control
changes use the median of six paired ratios; the percentile bootstrap uses
10,000 resamples with seed 791079 and indexes 250/9749.

| Shape | Control p50 ms | Wrapper p50 ms | Median paired ratio | Bootstrap 95% interval |
| --- | ---: | ---: | ---: | --- |
| tiny | 0.240061 | 0.239186 | 0.996554 | [0.992989, 0.999792] |
| medium | 0.474727 | 0.473787 | 1.000032 | [0.995542, 1.007944] |
| large | 19.871539 | 20.569372 | 1.035117 | [1.032398, 1.185124] |

The large profiling-wrapper process is 3.512% slower by median paired p50.
Its process-p50 spread is 28.760%; the wide interval remains visible. No
process is excluded or retried. This perturbation prevents treating the
profile binary as an interchangeable native baseline or claiming that its
instruction counts explain a precise production latency share.

All ten over-5% process-spread flags are retained: four whole-process RSS
flags (tiny and medium, both legs), two p99 flags (tiny and medium wrapper),
and the large wrapper's p50/mean/p95/p99 flags (mean spread 28.919%). Whole-child GNU-time RSS remains an
accounting observation with the 0789/0790 limitations; no residency reduction
or memory-regression disposition is inferred here.

## Current guest instruction profile

Every region's total equals the sum of parsed self costs. The owner has one
public-capture child, and owner self plus immediate-child inclusive cost
matches the region total. Descendant inclusive costs overlap and must not be
summed as independent shares.

| Shape | Pass 0 Ir | Pass 1 Ir |
| --- | ---: | ---: |
| tiny | 7,802,322 | 7,805,351 |
| medium | 13,943,830 | 13,949,381 |
| large | 568,290,472 | 568,319,784 |

In large pass 0, `inspect_element` is entered 181,678 times and calls the
checked attribute iterator 273,263 times. The iterator has 13,856,850 self
instructions across all incoming calls; its inclusive cost from the inspector
is 48,295,031. `inspect_element` itself has 34,965,460 self instructions.
The raw function census also retains remaining UTF-8 validation, namespace
resolution, XML event parsing, and the presentation-index path.

These counts establish that attribute iteration remains material. They do
not identify how many calls have exactly empty raw attribute tails, nor the
removable fraction of iterator cost. That fraction and native benefit still
require the candidate experiment. The checked path must not be replaced by
unchecked iteration on attribute-bearing or malformed input.

Software SHA remains large in guest instruction counts. The native profile
uses hardware SHA on this host, so the guest ranking alone is not a CPU-cycle
ranking or a reason to remove required fingerprint validation.

## Native sampled ancestry

The canonical no-inline decodes retain full Rust symbols. Counts below require
the exact `Package::opened_presentation_with_limits` owner on the stack; nested
rows overlap.

| Observation | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| Whole-process sample stacks | 3,143 | 3,153 |
| Exact capture-owner stacks | 1,066 | 1,068 |
| Notes scanner nested under owner | 935 | 929 |
| Element inspector nested under owner | 318 | 289 |
| Checked attribute iterator nested under owner | 186 | 148 |
| quick-xml attribute iterator nested under owner | 134 | 120 |
| Known namespace resolver helper nested under owner | 5 | 4 |
| Package fingerprint nested under owner | 96 | 92 |
| Notes presentation-index loader nested under owner | 8 | 5 |
| Qualified stacks with unresolved interior frame | 1 | 2 |

Unresolved interior frames violate the frozen prerequisite for native
phase-fraction claims. The packet reports exact observed counts only. Flat
leaf counts, event periods, all decoded samples, compressed original perf data,
and unqualified samples remain retained. No population confidence interval,
historical regression attribution, or native speedup follows.

The low presentation-index counts make its duplicate presentation scan a
lower-priority hypothesis than the per-slide scanner. The source review
records both routes and their error-masking constraints; it does not combine
them into one unmeasured rewrite.

## Verification and next production experiment

The [source review](results/change-0791/source-review.md) places the proposed
empty-tail branch after all existing element validation. The independent
buffered scanner oracle must remain unchanged. Required cases include bare
Start/Empty events, whitespace tails, namespace declarations, duplicates,
malformed names/values/entities, and node/depth/attribute limits. Vendor ASCII
and Unicode namespace controls must accompany a future native paired run.
A useful reduction must survive operation-level latency and allocation/memory
guards; no benefit is assumed from these profiles.

Fresh ordinary control/profile and frame-pointer release builds and probe
formatting pass. The inherited 16 unused-helper probe warnings remain in the
logs. No fresh production test-suite claim is made: all 9,196 production files
and all 35 architecture, goal, taxonomy, and ADR-index inputs are unchanged.
The reused probe adds no production dependency, public API, ambient runtime,
unsafe-policy change, or preservation/refusal change. ADR 0001/0002/0024
ownership, ADR 0003 publication semantics, ADR 0005 evidence/resource contracts,
ADR 0006/0008 preservation and verification, and ADR 0010/0011 archive ownership
remain intact; remaining accepted ADRs are unaffected.

Raw receipts, replay scripts, source hashes, review, final cleanup witnesses,
and the document seal are retained. Three exact executables are verified before
the owned target is removed; replay works after cleanup. Unrelated working
files and pre-existing worktrees are preserved.

This is supporting measurement evidence for existing PPTX/category-15 work,
with no capability or coverage-index promotion. OLE2/OOXML remain active,
ODF deferred, iWork excluded, and the broad performance goal remains open.
