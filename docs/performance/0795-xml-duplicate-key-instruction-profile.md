# 0795 — instruction and branch costs of the rejected XML iterator

**Diagnostic only; 0794 remains rejected and production is unchanged.**
The candidate reduces large-capture guest instructions by only 0.387–0.393%,
despite the previously measured 85.048% allocation-call reduction. Compiler-attributed
self costs rise in the checked iterator and element inspector. Simulated indirect
branch misses increase about 35%. These observations identify costs to investigate;
they do not establish the cause of the native timing regression.

## Protocol and source custody

The [packet](results/change-0795/README.md) starts at `65fdca1a04` and rebuilds
the exact baseline and rejected 0794 five-file candidate. The original candidate
archive, test results, source census, and frozen architecture inputs are hash-bound.
The unchanged 0793 probe wrapper encloses only `Package::opened_presentation`;
fixture construction, borrowed-package setup, readback, serialization, and result
destruction remain outside the counted operation. The source review distinguishes
the five modified copies; this PPTX owner exercises the OPC helper.

Root ran every build, capture, and decode serially on CPU 12 of the recorded
AMD EPYC 9R45 Linux x86-64 host with Rust 1.95.0 and LLVM 22.1.2. Ordinary
profile builds use opt-level 3, thin LTO, one codegen unit, debug level 1, and
unwind panics. Native sampling uses separately built frame-pointer variants.
Both source legs were built, then the exact baseline was restored before capture.

Twelve Callgrind processes cover tiny/medium/large capture in two opposite
shape and source-leg orders. Each uses one measured sample and no warmup.
Collection starts off, then zeros/toggles/dumps at the exact non-inlined wrapper.
Branch simulation records Ir, Bc, Bcm, Bi, and Bim. Four native perf processes
use the large shape, two alternating source-leg pairs, 100 samples after three
warmups, `cycles:u` at 499 Hz, and frame-pointer stacks. Canonical no-inline
decodes preserve full symbol names. Sixteen reports contain 412 measured outputs.

All source/output/semantic-readback oracles match sealed 0794 evidence. Each
Callgrind capture has one positive region dump and an empty termination dump.
All five counters conserve across raw self costs, summary, and totals; all twelve
owner partitions qualify and contain the public capture descendant. A second,
independent raw-cost reader checks all 24 dumps without the function-graph parser.
Native raw-file hashes are verified after decompression, and an independent
sample census matches the full parser.

## Guest counters

Ir is the guest instruction count; Bc/Bcm are conditional branches/simulated
mispredictions, and Bi/Bim are indirect branches/simulated mispredictions.
These are not native hardware counter measurements. Two repeats show observed
variation and do not supply a population confidence interval.

| Shape | Repeat | Before Ir | Candidate Ir | Ir change | Before Bim | Candidate Bim |
|---|---:|---:|---:|---:|---:|---:|
| tiny | 0 | 7,758,062 | 7,745,872 | -0.157% | 1,663 | 2,094 |
| tiny | 1 | 7,761,900 | 7,744,819 | -0.220% | 1,658 | 2,095 |
| medium | 0 | 13,724,056 | 13,706,121 | -0.131% | 4,361 | 5,460 |
| medium | 1 | 13,731,167 | 13,706,059 | -0.183% | 4,354 | 5,459 |
| large | 0 | 549,368,835 | 547,207,829 | -0.393% | 229,969 | 310,633 |
| large | 1 | 549,351,049 | 547,227,600 | -0.387% | 230,149 | 310,633 |

Large-capture conditional branches fall about 2.08%, simulated conditional
misses about 6.88–7.03%, and indirect branches about 1.63%. The indirect-miss
increase is a simulator observation; it is not evidence that the host CPU
suffered the same prediction failures. Predictor state and guest code generation
differ from native execution. Hardware SHA used natively also changes the cost
balance from Valgrind’s software path.

| Large capture, repeat 0 | Before self Ir | Candidate self Ir |
|---|---:|---:|
| Element inspector | 29,165,311 | 29,774,878 |
| Checked iterator | 8,660,171 | 8,936,993 |
| quick-xml iterator state | 42,805,826 | 38,818,090 |
| Allocation-named functions | 8,102,066 | 1,961,540 |

The inspector adds 609,567 self instructions and the checked iterator adds
276,822 in both large repeats. The local bookkeeping, end-position update,
inline-state dispatch, and larger iterator identified by the source review remain
specific hypotheses for these added costs. Function costs depend on compiler
inlining and attribution; incoming graph-call counts are not a census of source
method invocations. The allocation-named group is a lexical function selection,
not the allocation-call counter. Inclusive rows and selected groups can overlap.

## Native stack observations

| Repeat | Leg | Whole-process samples | Exact-owner samples | Nested inspector | Nested checked helper | Unresolved owner samples |
|---:|---|---:|---:|---:|---:|---:|
| 0 | before | 3,128 | 1,016 | 208 | 99 | 1 |
| 0 | after | 3,369 | 1,023 | 239 | 156 | 0 |
| 1 | after | 3,386 | 1,032 | 247 | 151 | 0 |
| 1 | before | 3,128 | 1,026 | 207 | 95 | 1 |

Two owner-descendant frames are unresolved, one in each baseline process. This
fails the frozen prerequisite for native phase fractions. The table retains exact
observed counts only; nested rows overlap and are not additive shares. Samples
include warmups, and whole-process samples include fixture and readback work.
All 4,097 exact-owner samples contain the public `opened_presentation` descendant.
No native latency, instruction-to-cycle conversion, fraction, or causal claim
is made. There are no fresh ordinary/wrapper timing controls in this batch.

## Retained build failure and validation

The first candidate build accidentally preserved archived source timestamps.
Cargo reused the baseline artifacts: both candidate binary hashes equaled the
respective baseline hash. Root detected this before profiling and retained the
failed source/build logs. Only source timestamps were refreshed; candidate bytes
remained exact. Both rebuilt candidate binaries differ from their baseline
counterparts. No captured data uses the stale artifacts.

Four source-qualified builds are retained, plus the two stale build attempts.
Production quality is inherited by exact source identity from 0794: six gates
per source leg, 12,850 baseline and 12,865 candidate passing tests, each with
89 ignored. These are not fresh test runs. Lightweight parser tests and both
independent raw audits validate the new evidence infrastructure. No production
dependency, public API, unsafe policy, preservation/refusal rule, limit, or
cancellation behavior changes.

## Next bounded work and limits

Allocation counts alone overstated the performance opportunity in this design.
The next experiment should isolate the short-tag iterator’s construction and
accepted-attribute bookkeeping, including zero-attribute and 1/4/5-name boundaries,
before choosing a smaller state layout. Its result must still survive the complete
PPTX workflow and cross-format timing/resource gates; these profiles cannot
authorize adoption. The 0792 empty-tail optimization and 0794 rejection remain
unchanged.

This is generated warm in-memory capture evidence, not native-producer, cold-cache,
range-source, concurrent, full-CRUD, RSS, or broad goal-completion evidence.
OLE2/OOXML remain active, ODF is deferred, and iWork is excluded. The owned
target and four final binaries are removed after recording exact identities.
The packet, this report, and five indexes are sealed for offline replay.

```sh
python3 -B docs/performance/results/change-0795/validate.py --require-final-seal
```
