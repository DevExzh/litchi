# Ordinary ODP lookup experiment and resolved phase profiles

The local-name-first attribute lookup experiment is **rejected and reverted**.
It fails the predeclared performance gate and changes no allocation metric.
The final production source is byte-identical to the starting revision
`36df468bd`. The measured patch, both source epochs and every raw result remain
available. The useful retained improvement is reliable diagnostic profiling and
a more precise ranking of remaining work.

## Experiment

A baseline and candidate each retain twelve reports: two repeats, normal and
allocator binaries, 64/4,096/8,192 source slides, three warmups and thirty samples
per lane. The 24-report matrix contains 720 retained operations in A1/B1/B2/A2
order. CPU affinity is 2 with one workload at a time. Both builds use Rust
1.98.1, the same release flags and unchanged harness; separate binary copies
preserve executable custody across rebuilds. Input/title/body/sink setup,
semantic preflight, row checks, report allocation and drops are outside the
lifecycle timer; actual publication sink hashing is inside it.

The one-line candidate checks cached local-name equality before namespace URI
equality. Both predicates are pure; iterator order, namespace resolution,
first-match semantics, decoding, and typed errors remain unchanged. Independent
assembly review confirms that generated comparison order changes. The full ODP
suite passes 368 tests, including differential attribute checks and native
fixture preservation. See [lookup review](lookup-review.md).

The frozen keep gate requires at least 5% normal p50 improvement on both medium
and large sources in both repeats, or a practically useful allocation reduction.
The candidate satisfies neither condition.

| Repeat | Shape | Baseline p50 ms | Candidate p50 ms | Candidate delta |
|---|---|---:|---:|---:|
| R1 | tiny | 1.905 | 1.907 | +0.078% |
| R1 | medium | 76.276 | 75.518 | -0.994% |
| R1 | large | 151.948 | 151.604 | -0.227% |
| R2 | large | 150.889 | 151.087 | +0.131% |
| R2 | medium | 75.366 | 85.410 | +13.327% |
| R2 | tiny | 1.921 | 1.897 | -1.291% |

R1 normal changes span -0.994% to +0.078%. R2 normal-medium p50 rises 13.327%,
and allocator-large p50 rises 43.039%; their p95/p99 changes produce six adverse
5% timing flags in total. All flags remain in [summary.json](summary.json).
The cause of the observed variation is not established. No confirmation rerun
is needed to justify rejecting the candidate. Every reported allocator metric
matches exactly, including volume, calls, regional peak above entry and retained
live delta. No process maximum-RSS comparison exceeds +5%.

Quantiles use midpoint median and nearest-rank p95/p99. Each paired median-ratio
interval uses 10,000 independent bootstrap resamples with a recorded seed; the
intervals do not remove run-order effects or correct for multiple comparisons.
Separate 100-sample large-normal whole-process counters show cycles -0.276%,
instructions -0.733%, branches -1.065%, branch misses +4.298% and cache misses
+5.837%. Setup, warmups and checks are included; these are not operation-only
causal fractions or grounds for retaining the failed candidate.

## Resolved diagnostic profile

A separate unchanged-source build enables `-C force-frame-pointers=yes` and
`-C force-unwind-tables=yes`. It is profiled at 99 Hz with both frame-pointer
and 32 KiB DWARF unwinding, on 100 large phase-instrumented samples per run.
The fp recording retains 1,654 samples and the DWARF recording 1,643, with zero
malformed parsed samples. Resolved phase markers account for 94.62% / 91.60% of
sampled periods, versus no resolved phase ancestry in 0458. Remaining periods
stay unattributed; this does not imply every unwind is complete.

Within the commit marker, candidate `Snapshot::from_owned_package` accounts for
59.36% / 60.01% of sampled periods, and serialization for 27.41% / 29.98%.
Within transaction construction, staging metadata accounts for 63.97% / 64.81%
and source-fragment parsing for 32.35% / 33.67%. Inclusive percentages overlap
and are sampled estimates; phase markers can include warmups. The diagnostic
build's timings are not the ordinary build's performance baseline.

Retaining the family package is rejected as an optimization direction: its
reopen occupies at most about 3% of transaction samples while keeping its large
XML payload would increase retention. The next coherent transaction experiment
is shared tokenization/namespace maintenance for staging metadata and source
fragments, preserving all validation state machines and historical error
priority. Initial and candidate slide parsing remain larger targets. See the
[source review](source-review.md) and [diagnostic summary](diagnostic-summary.json).

## Evidence and replay

The bundle retains build/source manifests, frozen protocol, exact binary hashes,
raw 720-sample matrix, four separate 100-sample diagnostic/counter runs, derived
summaries and rejection decision. Baseline R2 executes its bound original
binary while the worktree still contains the candidate; check receipts record
worktree custody separately from measured-binary custody. The disclosed [capture metadata amendment](capture-metadata-amendment.json)
binds a stale numeric batch field and omitted variant field in original capture
receipts; schema, directory, argv and binary/source bindings remain authoritative.
The original unstable
classifier tie order and its output are retained, then corrected before bundle
verification. Two Python hash seeds produce byte-identical final summaries.

The candidate ODP suite passes 368 tests. Post-restoration strict ODP Clippy,
warning-denied rustdoc and scoped formatting pass. Whole-workspace formatting
finds an existing Keynote difference outside this task; its failed receipt is
retained. Crate boundaries and final bundle verification are recorded in their
receipts. No native Office application was launched. See
[integration notes](integration-notes.md) for source epochs and gate limitations.

Run `python3 -B verify.py --portable` from a copied complete bundle to check
its seal, source/build/argv/row bindings and recompute both summaries. Before
cleanup, `--precleanup` additionally verifies retained executables and restored
source. `finalize.py` records a fresh-copy replay, then inventories and removes
only `/tmp/litchi-goal-0459`, including the diagnostic build tree. `seal.py`
regenerates the seal between receipt-producing steps. Shared build targets and
user-owned `docs/GOAL.md` remain unchanged.

This diagnostic command adds no registry selectors or CRUD promotion: counts
remain 439 selectors / 36 defaults. The full non-iWork goal remains open,
including wider selective CRUD/input/output matrices, native application
roundtrips, cold/range behavior and bounded-worker scaling.
