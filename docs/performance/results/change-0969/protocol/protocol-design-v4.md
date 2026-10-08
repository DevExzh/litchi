# 0969 prospective CFB writerproof qualification protocol (v4 custody)

**Status:** fresh prospective v4 design, 2026-10-08. No source mutation, build,
benchmark, or capture is authorized; root GO follows sealed source, binaries,
controls, and plan.

The candidate is a private CFB writer model handoff that skips post-render
payload recapture. The handoff must retain structural/sector-layout
validation, stream readback proof, unknown-byte preservation, and atomic
publication failure behavior. A shorter internal path is not an adoption
claim without the ordinary DOC lifecycle and allocator evidence below.

## Custody and design gate

The immutable 0969 baseline and currently identical candidate exports are
`/home/zhuhe/litchi-performance-0969/source-{baseline,candidate}` (tree SHA
`df8a842b08fec1792c67921fb1ce3513d3c6cb3dd0149cf13672bad14090bfb0`). Bind
`manifests/source-manifest-v1.json` and a fresh production/test allowlist.
Before build, seal `model_handoff`: owner, validation boundary, model/source
lifetime, ordinary-writer fallback, and reopen/readback oracle. Unexpected
non-document bytes or paths stop the run. Source revision strings are
descriptive; documentation-only revision differences do not reject a matched
file manifest.

The current design-only memo is
`/home/zhuhe/litchi-performance-0969/design-v1.md` (SHA-256
`f043500923af2cb7348845d6f280429a76e11e9c36e744ebfa9bfd572bf9616c`). Its
sealed review is
`/home/zhuhe/litchi-performance-0969/review/design-review-v1.json` (SHA-256
`b83916d26aea5dc30710da60d285558987f498eb6d93211b4c8f97f03c0cad8a`). The
baseline reuse review is
`/home/zhuhe/litchi-performance-0969/review/baseline-reuse-v1.json` (SHA-256
`dbd362dae8e31dfe5ff54d6af5c37883df829bede3a5e5fa19e04337e0a8312e`). The
final plan must bind these receipts. The handoff is eligible only
after a source-layout reuse plan is validated and emitted. Copy-through,
rewrite, no-source, and declined/failing reuse routes retain full capture and
are reported separately rather than pooled as handoff benefit.

Reuse retained 0967 ordinary and allocator baseline binaries only after exact
runtime, `tools/perf-cfb-profile`, Cargo.lock, fixture, compiler/configuration,
and file-manifest equivalence. Build the 0969 candidate fresh against that
harness, or build a matched pair when equivalence fails. The retained 0968
diagnostic payload observations (16,121,304 bytes FloatingPictures/e32;
1,354,864 bytes NoHeadFoot/e32) are scope context only, not 0969 evidence.
The 0966 counter-only harness is not a native timing harness.

## Correctness and matrix

Baseline and candidate run the same `open_bounded` → `Edit` → repeated
paragraph-0 replacements → `commit` → output copy → reopen/readback lifecycle.
The sealed controls cover unchanged streams and directory metadata/CLSIDs,
regular FAT and MiniFAT streams, source-layout copy-through and fallback
writer paths, metadata/content edits, new-stream/removal paths where supplied,
no-op sharing, patch refusal and atomic limits, retention zero/default/release,
structural reopen, and final semantic readback. Expected output hashes and
semantic identities bind both variants. The handoff must not expose an
unvalidated stream or silently omit a payload.

The fixed real matrix is `NoHeadFoot.doc` and `FloatingPictures.doc` × edits
1/2/8/32 × retention `zero`/`default`/`release` (24 cells). Native uses 12
fresh counterbalanced pairs per cell (six each order), five warmups, and 80
samples. Report p50/p95 and descriptive p99; paired-bootstrap p50 and p95
upper bounds must be at most 1.05 in every cell. A median above 1.05 is a
review flag and p99 cannot waive a failure.

The allocator lane uses three fresh paired repeats per cell, warmup zero and
one sample, reporting lifecycle allocation calls, requested/allocated bytes,
peak live, retained bytes, and RSS. Safety limits are allocation-call median
≤1.01 per cell. The requested-byte metric is exactly the existing allocator
packet field `samples[i].allocation.allocated_bytes`; the retained harness
documents that field as total bytes requested from the system allocator. Do
not substitute `output_bytes`, `retained_bytes`, or a new `requested_bytes`
field. For each paired repeat, hard-fail if the candidate/baseline ratios for
that field exceed 1.01, `peak_live_bytes` exceed 1.05, or the runner's bound
process `rss.kb` exceed 1.05. No private writer counter substitutes
for lifecycle evidence.

Freeze A/A before any candidate formal child launch: the same baseline binary
in AB and BA over all 24 cells, 12 pairs per cell, same child settings, seed
`969004`, and 10,000 paired-bootstrap resamples. Require the p50 and p95
symmetric upper bounds `max(r,1/r)` to be at most 1.05 for every A/A cell
before launching timed controls or any real baseline/candidate child. A
complete A/A packet whose noise bound fails is `perf_decision` readiness-fail
and therefore uninterpretable; it is not an incomplete run and is not retried.
A missing, malformed, failed, duplicate, or otherwise incomplete A/A packet
is incomplete and also stops before candidate launch.

This A/A stage is baseline-only and may be prepared while candidate QA is in
progress, but it may launch only after root confirms process quiescence: no
Cargo, rustc, tests, benchmarks, or other workloads are active. It comprises
24 cells × 12 pairs × 2 positions = 576 baseline children. The qualification
plan binds the completed A/A receipt before binding or launching the candidate
cohort. The default candidate cohort is 576 native children and 144 allocator
children after that readiness gate.

The external runner represents this as two one-way plan states: an
`aa_authorized` plan can bind and launch only the baseline binary, while a
later `capture_authorized` plan must bind the exact passing A/A receipt and
packet manifest plus all four formal binaries. The formal invocation revalidates
every bound A/A packet before launching a candidate child; a missing or changed
receipt/manifest fails closed.

The frozen plan contains no GO path or mutable GO field. Each launch instead
requires an explicit `--go` receipt. The runner verifies that receipt binds the
exact plan path/SHA, both review receipts, every plan helper hash, source
manifest bindings, and the binary bindings used by that stage. The GO must also
record a passing quiescence admission: no Cargo, rustc, test, benchmark, or
other workload is active at launch or between A/A and formal stages. A missing
or mismatched quiescence admission stops before a child; no host-cause claim is
made from a failed readiness or performance decision.

## Qualification and failure

After custody, correctness, A/A readiness, native non-inferiority, and safety
pass, qualify through one predeclared branch:

1. The paired-repeat median of the exact `samples[i].allocation.allocated_bytes`
   ratio is ≤0.95 at both `NoHeadFoot/e32/default` and
   `FloatingPictures/e32/default`; or
2. one of those endpoints has native p50 median ≤0.95 with 95% bootstrap upper
   bound <1.00.

All cells retain the safety gates. Allocation-call changes alone are
descriptive. `perf_decision` is authoritative for complete qualification,
rejection, or incomplete evidence. Bind manifests, controls, harness,
fixtures, binaries, host/order schedule, A/A seed/resamples, metric schema,
and root GO content/hash. Use fresh exclusive output paths, retain terminal
stdout/stderr, reject missing/duplicate/failed/malformed/wrong-identity or
custody-mismatched packets, and perform no retries or selective replacement.
A complete A/A noise failure is `readiness_fail`/uninterpretable; only missing
or invalid packets are `incomplete`. External parser/runner preparation is
allowed under this v4 design, but no source mutation, build, or workload may
begin without the separate review and root GO.
