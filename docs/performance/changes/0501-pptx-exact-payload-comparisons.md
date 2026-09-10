# 0501: PPTX exact-payload comparison without redundant private hashing

0501 changes only the private `digest_touched` preparation path in
`source_cross_copy.rs`. Image and chart decoded payload bytes no longer feed a
redundant private SHA-256 update. The planner still compares source URI,
target URI, content type, declared size, and every decoded payload byte;
payload reuse still performs checked exact equality; and candidate validation
still rereads selected image and chart Parts and checks their exact bytes,
metadata, relationship closure, and chart XML. Graph/layout/master/theme
digests, XML and relationship metadata/order, source lineage and revision,
source freshness, compressed authorization, cancellation, resource admission,
and partial-output handling remain in force.

The replacement keeps the old 64 KiB cancellation observation cadence without
allocating or charging `Resource::Work`. The change adds no public API,
manifest dependency, package ownership edge, unsafe code, or publication
authorization shortcut. The private digest is not a public fingerprint and
cannot authorize a payload by itself; a changed payload with an unchanged
private digest remains rejected by exact vector comparison and candidate
reread.

## Before baseline and profile boundary

The matched source-backed PPTX lifecycle has plain and media-rich synthetic
corpora, owned bytes and warm-file providers, three warmups, thirty measured
samples, and two reversed repeats. The fresh before phase contains eight
reports and 240 measured samples. Timed API phases cover provider open,
destination open, planning, and consuming publication; setup, staged files,
diagnostics, correctness checks, report serialization, and final drops are
outside those phase clocks. Warm tmpfs files do not establish cold-storage,
native-producer, remote, or physical-I/O behavior.

| Corpus / provider | API p50 before → after (Δ), R1 / R2 (ms) | Plan p50 before (R1 / R2, ms) | Publication p50 before (R1 / R2, ms) |
| --- | ---: | ---: | ---: |
| plain owned | 2.274 → 2.229 (−1.977%) / 2.260 → 2.238 (−0.998%) | 0.797 / 0.790 | 1.032 / 1.027 |
| plain warm file | 2.562 → 2.516 (−1.796%) / 2.561 → 2.595 (+1.347%) | 0.822 / 0.821 | 1.263 / 1.260 |
| media-rich owned | 24.297 → 9.359 (−61.481%) / 24.337 → 9.378 (−61.467%) | 11.867 / 11.848 | 11.748 / 11.799 |
| media-rich warm file | 28.588 → 13.879 (−51.452%) / 28.910 → 13.914 (−51.870%) | 12.979 / 12.984 | 14.918 / 15.250 |

The media-rich plan p50 falls 64.672%/64.779% for owned R1/R2 and
59.171%/59.031% for warm-file R1/R2. Publication p50 falls
61.848%/61.704% owned and 47.061%/48.021% warm-file. Media-rich API p99
falls 61.199%/61.350% owned and 50.835%/51.122% warm-file. Plain lanes stay
within 2% at API p50 except the warm-file R2 increase of 1.347%; its API p99
increase is 1.632%, both below the review threshold.

A fresh whole-child profile reports `sha2::sha256::x86_sha::compress` at
31.01% for the owned lane and 30.08% for the warm-file lane; the matched
after profiles report 27.95% for owned and 26.93% for warm-file. These are
whole-child shares. The profiles include
corpus construction, preflight, correctness gates, diagnostics, and
serialization and have incomplete unwind/callgraph coverage; they prove
SHA-256 hotness in the child, not that the private touched digest caused that
hotness.
Historical targeted 0449 caller attribution places the production touched
digest at about 15.5–16.1% of lifecycle-frame period, while separately
assigning roughly half of SHA period to untimed harness output hashing. Neither
observation is a fresh timer-local speedup claim for 0501.

The first owned profile export was recovered by re-exporting the same recorded
`perf.data` with local symbols after its remote debuginfod subprocess was
terminated. The driver then reran the owned profiling workload twice; the first
rerun retained an old supplementary-dimension verifier rejection and the final
before invocation passed after that verifier correction. These supplementary
attempts do not change the formal 240-sample matrix.

The matched after phase has the same eight lanes and 240 measured samples.
All 208 comparison rows are retained. Every after report passed its lifecycle
oracle, and the comparison retains 52 review flags: 48 favorable timing flags
and four favorable throughput flags; no measured change above five percent is
adverse. No source archive, destination archive, exact output, source-read
histogram, cache, or resource-counter fingerprint changes. Whole-child RSS
changes span −3.914% to +1.123%, with media-rich lanes within ±0.042%. Each
media-rich fixture has eight 2 MiB image payloads; code inspection therefore
identifies 32 MiB of redundant payload hash input removed across the two
preparation passes in this workflow. This is not an allocation, physical-copy,
decompression, or I/O reduction claim. The two after supplementary profiles
also pass, but retain whole-child setup, gate, diagnostic, and serialization
scope. Final repository gate totals remain outside this comparison; the
completed default CRUD baseline is documented below. The production review's
proof-equivalence and cancellation conclusions remain acceptance scope.

## Coverage boundary

Existing synthetic source-backed PPTX lifecycle evidence in 0424, 0431, and
0448 is real retained evidence and is not relabeled as missing generic
lifecycle or ABBA coverage. 0501 does not add native notes or chart dependency
closure. Those producer-breadth and relationship-graph capabilities remain
separate work; the 0501 result stays scoped to the named source-backed
cross-copy workflow.

The separate default CRUD baseline refresh completed two serial matrices of 201
rows with 15 samples and three warmups each, or 6,030 measured samples, across
37 cases and 31 corpora. Both generated catalog/report validators passed; the
static coverage module ran 38 tests successfully. R1 took 47.42 seconds with
151,028 KiB maximum RSS and R2 took 46.57 seconds with 160,568 KiB. This closes
the current timing-report baseline gate for the checked 15-category,
33-selector index and its 48 measured mappings, but it does not promote
correctness-only mappings or establish native-producer coverage. The full
non-iWork `docs/GOAL.md` objective remains open.

All scoped production gates pass: the default all-target suite has 848 passing
tests with no failures, the all-features library suite has 552 passing with no
failures, doctests have 6 passing and 2 ignored, the focused suites have 58
passing, and private payload guards have 5 passing. Warnings-denied Clippy,
formatting, rustdoc, downstream, boundary, and the 38-test CRUD static check
also pass. Independent strict verification passes. Final cleanup removed
2,154,708,992 unique-inode allocated bytes and retained eight replay files in
local tmpfs: two binaries and six raw perf files, including two failed
attempts. The full non-iWork `docs/GOAL.md` objective remains open.
