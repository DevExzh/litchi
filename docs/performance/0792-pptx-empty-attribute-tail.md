# 0792 — exact empty attribute-tail experiment

**Retained.** Skipping the checked attribute iterator only for exactly empty
raw tails improves large generated PPTX capture by **5.211%** and full
lifecycle by **6.937%** in the frozen paired-p50 comparison. No latency or
allocation-resource guard fails. All 255 reports and 5,595 measured samples
pass exact source/output/semantic checks. The namespace pointer-test correction
is retained alongside the scanner change.

## Hypothesis and boundary

The current [0791 profile](0791-pptx-current-capture-profile.md) observed checked
attribute iteration beneath the PPTX notes scanner. The candidate returns early
when `BytesStart::attributes_raw()` is exactly empty, after existing namespace,
local-name, and root validation. Whitespace tails and attribute-bearing tags
retain checked iteration. This removes no validation applicable to actual
attributes and changes no XML limits, public API, runtime, ownership, or
publication policy. The independent buffered scanner oracle remains unchanged.

## Experimental scope

The source base is `4fa65522a0`. The exact 0785 probe and fixtures are reused
with their original tool/schema names, against freshly built before/after
production sources. Historical timing is not pooled. The unchanged 35
architecture/goal/taxonomy inputs are hash-bound. See the
[packet](results/change-0792/README.md) for commands, source manifests, toolchain,
host, frozen policy, and raw artifacts.

Fifteen cases cover capture, staged commit, and full public lifecycle for
tiny (3×4), medium (12×8), large (100×100), and medium ASCII/Unicode vendor
fixtures. Vendor text elements carry six unknown namespace attributes. Six
alternating native blocks use 30 samples after three warmups on CPU 12.
Separate allocation captures use two blocks of three samples without warmup.
The 15 baseline qualification children are outside the paired matrix.

The native clock measures each probe-defined operation, with construction and
readback outside the clock. Allocation counters are a separate instrumented
lane. Whole-child GNU-time RSS is an accounting observation with the retained
0789/0790 limitations, not a physical residency or operation peak measure.
No cold, range-source, concurrency, native-producer, or new CRUD coverage is
claimed. Guest instruction attribution is deferred pending native benefit.

The adoption policy was frozen before baseline builds: a capture or lifecycle
case must improve at least 3%, with its paired-p50 bootstrap upper bound below
one. A case exceeding 5% regression with lower bound above one rejects the
candidate. Allocation calls, allocated bytes, net live bytes, and peak above
entry may not increase. Bootstrap uses 10,000 resamples, seed 792079, and
explicit sorted endpoints 250/9749. All processes and spread flags remain.

## Paired native results

Times below are medians of per-process nearest-rank p50 values. Changes and
intervals use the six paired process ratios, not ratios of displayed medians.

| Shape | Operation | Before ms | After ms | Paired change | Ratio 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | capture | 0.237761 | 0.235631 | -0.986% | [0.988965, 0.992098] |
| tiny | commit | 0.212841 | 0.211546 | -0.672% | [0.992420, 0.998799] |
| tiny | lifecycle | 1.429182 | 1.408677 | -1.362% | [0.983237, 0.988907] |
| medium | capture | 0.469287 | 0.459647 | -1.892% | [0.969915, 0.997075] |
| medium | commit | 0.297222 | 0.295662 | -0.372% | [0.989468, 1.000922] |
| medium | lifecycle | 2.053304 | 2.012949 | -2.079% | [0.978613, 0.981792] |
| large | capture | 20.128215 | 19.091318 | -5.211% | [0.919250, 0.953105] |
| large | commit | 1.308701 | 1.294546 | -1.021% | [0.981030, 0.997316] |
| large | lifecycle | 30.768311 | 28.629127 | -6.937% | [0.925160, 0.933238] |
| vendor | capture | 0.550223 | 0.539892 | -1.810% | [0.977600, 0.986283] |
| vendor | commit | 0.326057 | 0.324597 | -0.511% | [0.993195, 0.997114] |
| vendor | lifecycle | 2.196005 | 2.163440 | -1.478% | [0.982057, 0.986709] |
| unicode-vendor | capture | 0.553533 | 0.542337 | -2.148% | [0.973081, 0.983740] |
| unicode-vendor | commit | 0.326497 | 0.323377 | -0.922% | [0.981377, 0.993791] |
| unicode-vendor | lifecycle | 2.206865 | 2.165075 | -1.878% | [0.960814, 0.982423] |

Large capture and lifecycle satisfy the frozen useful-benefit rule. No case
violates the latency guard. All 90 allocation sample pairs have identical
allocation calls, allocated bytes, net live bytes, and peak above entry;
there is no allocation-reduction claim. Whole-child RSS and tail/spread flags
remain separately visible in the machine-readable analysis.

The analysis retains **21 process-spread flags** (11 RSS, ten elapsed metrics)
and six metric families with an individual paired block exceeding +5%:

| Case/metric | Maximum block change | Median paired change |
| --- | ---: | ---: |
| medium capture RSS | +6.206% | +0.644% |
| medium commit p99 | +7.962% | +0.089% |
| medium lifecycle RSS | +7.621% | +0.127% |
| tiny capture RSS | +7.371% | +0.783% |
| tiny commit p99 | +9.946% | −1.391% |
| vendor capture p99 | +37.545% | −1.813% |

These are review flags, distinct from the frozen paired-p50 rejection rule.
Their median-ratio intervals all cross one. In particular, the vendor-capture
p99 interval is [0.861128, 1.183342], so no tail improvement is established.
The large baseline capture p50 spread is 7.257%; it remains included in the
reported uncertainty. Allocation captures have no spread or regression flags.

These are current native before/after results, not cumulative gains with 0785.
They support retaining this local change on the named generated workloads.
The larger lifecycle improvement does not prove a precise capture-only causal
share: code generation/layout and work outside capture can also affect timing.
No fresh guest-instruction count, phase fraction, or hardware-cycle attribution
is claimed for the candidate.

## Correctness and architecture

The probe's separate synthetic/default and real-allocator suites pass 28 and
7 tests. All 15 baseline cases match sealed 0785 source/output/semantic
identities. The candidate's focused tests cover exact-empty Start/Empty,
whitespace, attribute, namespace/root/name refusal, and existing limit paths;
the final all-feature suite passes 1,238 tests with three ignored across 85
suites, including doctests. Formatting, all-target checking, strict Clippy,
and warning-denied rustdoc also pass. The crate-boundary gate passes for 65
workspace packages and 244 internal declarations, retaining 11 explicit debt
items. The unchanged standalone probe retains 15 unused-helper warnings in
each native build and six in each allocator build; these are separate from
the passing warning-denied production checks.

The first quality attempt passed formatting and checking but its unit suite
reported 820 passed, one failed, and one ignored. The failure was an existing
namespace test comparing addresses of equal string constants. Its isolated
baseline diagnostic passed; after scanner/test layout changes the addresses
differed while text remained equal. The installed Rust reference explicitly
states that references to a constant need not have the same address. A test-only
correction retains exact text equality and checks that the known result does
not borrow the live input buffer. The production namespace helper is unchanged.
This second test file is an explicit pre-measurement source-allowlist amendment;
no adoption threshold changes. Both the failed attempt and passing baseline
diagnostic are retained. The final suite adds direct attribute-count/byte and
100,001-node Start/Empty boundary checks against the unchanged scanner oracle.

The local private branch preserves ADR 0001/0002/0024 ownership, ADR 0003
publication, ADR 0005 resource/evidence, ADR 0006/0008 refusal/preservation,
and ADR 0010/0011 archive boundaries. It introduces no new cache, retained
state, dependency, unsafe code, hidden executor, or public surface. Remaining
accepted ADRs are unaffected.

OLE2/OOXML remain active, ODF deferred, and iWork excluded. This experiment
does not establish completion of the broad performance goal.

## Disposition and cleanup

The source review, six quality gates, full artifact replay, and independent
raw numerical audit support retention under the original thresholds. The
small branch and its focused tests remain in production; the separate
namespace-test correction changes no runtime implementation. The original
formatting and quality failures remain visible, and no primary measured
process was retried, excluded, or overwritten.

Four exact executables are verified before removing the owned target. Replay
uses the cleanup witness after deletion. The final seal binds the evidence
and six report/index documents; unrelated files and pre-existing worktrees
remain unchanged. This is an optimization of existing PPTX capture/commit/
lifecycle scenarios, with no coverage-index promotion.
