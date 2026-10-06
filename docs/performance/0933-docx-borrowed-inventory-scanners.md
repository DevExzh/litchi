# 0933 — borrowed DOCX inventory scanner state

## Hypothesis and scope

[0932](0932-docx-allocation-attribution.md) attributed about 38.6% of
whole-process requested bytes to the nearest physical frames of the section
inventory and source-admission scanners. Source inspection found per-event
namespace-resolver copies, owned inventory events and copied topology names.
That percentage is not an estimate of removable allocations or a speedup bound.

This candidate keeps `NsReader<&[u8]>` events borrowed from the immutable XML
input and borrows the current resolver during each loop iteration. The inventory
topology stack retains borrowed `BytesStart` views, compares their complete
qualified names against closing tags, and keeps its fallible reservation.
The source-backed section inventory validator only loses its resolver clone.
No source-document admission scanner, fragment codec, cache, public API or
publication contract changes.

Both scanners still perform fresh validation in the same event order. Depth,
event, section and reference limits, namespace checks, exact source offsets,
MCE processing, source freshness, output fingerprints and inverse validation
remain in place. Borrowed event data does not outlive the parse call or become
part of the returned inventory.

## Semantic verification

Four regression tests cover namespace restoration after a nested rebind,
foreign-rebound section markup, mismatched qualified closing names, and exact
section-count limit admission/refusal. These tests describe public behavior;
they do not carry a duplicate scanner implementation. All 22 inventory tests,
including the new cases, pass on the baseline before the candidate is tested.
The candidate then passes 1,925 DOCX tests (one existing corpus test ignored),
78 doctests, warning-denied Clippy and warning-denied rustdoc. The changed
files pass rustfmt; workspace-wide formatting reports only unrelated existing
drift. The final full crate-boundary check passes for 65 packages and 244
internal dependency declarations after the isolated checker tree receives its
complete HEAD documentation inputs. Setup failures and their correction are
preserved in the [quality record](results/change-0933/new-quality.json).

The [static review](results/change-0933/static-review-v1.json) binds the accepted
ADR review, source identities and semantic constraints. A proposed change to
the separate general document scanner was removed before source freeze. The
final production diff is restricted to the two measured scanner paths.

## Measurement protocol

The normal-release comparison uses a fresh matched baseline and candidate,
six interleaved process pairs for each of four ingress routes, five warmups and
50 measured samples per child. Generic ReadAt is primary; owned bytes, owned
reader and short-range input are controls. The frozen primary gate requires a
median paired baseline/candidate p50 ratio of at least 1.05 and a bootstrap
95% lower bound above 1.0. Per-route tail and RSS regressions above 5% require
review. The earlier 0930 timings are historical context, not this baseline.

A separate counting-allocator lane measures 16 and 8,192 sections with eight
live commits, exact retention and shared/independent source ownership. Each
cell has three repetitions for each version. Construction deltas end at
`after_snapshot_drop`; later correctness work must not be included in that
interval. Allocation requests are distinct from live bytes, bytes copied,
RSS and normal-release latency. Before capture, a separate allocation adoption
criterion is fixed: at least 20% fewer construction-requested bytes and fewer
allocation calls in all four cells, passing correctness, and no unresolved
normal-route latency or RSS regression. The latency improvement claim retains
its separate paired-ratio gate.

All capture children use CPU 12, explicit environments and terminal-first
receipts. Failed children are retained without retries or selective replacement.
Full evidence and build artifacts stay outside the repository. All 72 children pass without retries. Independent allocator checks and
copied-packet replay pass; both packet copies produce identical numerical
results after normalizing only their root path.

## Results and adoption decision

[Adopt the candidate](results/change-0933/adoption-decision-v1.json) for its
measured allocation reduction. Each of the three
repetitions produces exactly the same construction byte and call counts in
every cell:

| Sections | Ownership | Baseline requested bytes | Candidate requested bytes | Reduction | Baseline calls | Candidate calls |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 16 | Shared | 2,970,668 | 2,108,662 | 29.0172% | 38,271 | 28,398 |
| 16 | Independent | 7,140,008 | 5,607,808 | 21.4594% | 65,480 | 47,928 |
| 8,192 | Shared | 1,251,354,390 | 800,608,673 | 36.0206% | 17,036,337 | 11,875,584 |
| 8,192 | Independent | 2,226,793,064 | 1,425,467,600 | 35.9856% | 30,284,264 | 21,109,592 |

These are cumulative construction requests, not retained/live bytes or bytes
copied. Exact-cap admission, shared ownership, output/inverse validation and
return-to-baseline checks all pass. This changes no retention policy, cache,
concurrency contract or source/publication checks.

The separate normal-release observations are:

| Route | Baseline pooled p50, ms | Candidate pooled p50, ms | Median paired baseline/candidate p50 |
| --- | ---: | ---: | ---: |
| Generic ReadAt | 4.151372 | 4.044596 | 1.026687 |
| Owned bytes | 3.382469 | 3.273008 | 1.033415 |
| Owned reader | 3.504685 | 3.411334 | 1.027159 |
| Short-range ReadAt | 4.179392 | 4.088437 | 1.000107 |

The primary paired ratio's 95% bootstrap interval is
[1.025509, 1.027245], but its median remains below the prospectively required
1.05. The latency improvement gate therefore fails; adoption relies on the
separate allocation gate. Pooled medians and paired ratios are different
statistics, particularly for the noisier reader and short-range controls;
their paired intervals cross 1.0. No broad latency or throughput claim follows.

All paired and pooled p95/p99 checks remain within the 5% regression guardrail.
Per-route maximum process RSS rises 0.33% (generic), 1.32% (owned bytes),
3.10% (owned reader), and 0.02% (short-range). One owned-reader pair rises from
37,904,384 to 41,029,632 bytes, an 8.2451% increase. This outlier is retained
and reviewed, not replaced. The route-wide maximum increases only 3.10%, its
median RSS rises 0.45%, and every allocation cell reduces requests without a
correctness failure. Accept this variation under the frozen per-route policy;
make no RSS-reduction claim. RSS is process-wide and includes setup.

The [independent audit](results/change-0933/independent-postcapture-audit-v1.json)
records both claim gates, the paired RSS flag, exact identity/cap checks,
counter conservation and replay equality. The frozen supplementary checker
rejects mutated corpus/output hashes, caps, logical retention, validation scope
and allocator observer state. Its identities were bound before capture.

## Evidence and reproduction

The [frozen manifest](results/change-0933/formal-manifest-authorized-v1.json)
binds the two normal binaries, allocator binaries, source lists, canonical
locks, schedule, environment, limits and mandatory independent checker.
[Full source comparison](results/change-0933/full-source-differential-v1.json)
checks all 10,778 files in each source tree: the retained 0931 allocator source
and baseline differ only in test-only code; baseline and candidate differ only
in the two reviewed production scanner files. Normal binaries are rebuilt
from these matched trees. The source origin revision, review HEAD and candidate
patch remain distinct identities.

The [reduction](results/change-0933/formal-reduction-v1.json) preserves individual
route metrics and allocation cells. The
[evidence bundle](results/change-0933/comparison-formal-evidence-v1.json)
records capture/build/reducer terminals, source and binary descriptors, and
regeneration commands. Original terminal records precede analysis. The canonical
packet is retained externally as
`/home/zhuhe/litchi-performance-0933/formal-packet-v1.tar.gz`, 427,218 bytes,
SHA-256 `a75a3826577077ed418bec11a4ed30643c7860535b372714d79abe47612634ee`.
Its deterministic archive script reproduces the same hash. Binaries and source
listings of 1 MiB or more remain outside the repository.

[Cleanup](results/change-0933/cleanup-v1.json) removes 18,470,776,166 logical
bytes of this batch's quality target and release intermediates. This is a
logical file-size sum, not measured physical disk recovery. All three newly built
executables remain hash-verified at their original paths; both source trees,
original receipts, checker inputs and canonical/copied packets are retained.
The shared checkout build cache and unrelated user work are untouched.
