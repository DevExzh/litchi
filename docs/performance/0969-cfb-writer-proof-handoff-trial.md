# 0969 — CFB writer-proof handoff trial stops at A/A readiness

**Status: implementation correctness QA passed; performance qualification is
unqualified.** The external candidate remains retained for review, and the main
runtime is unchanged. The baseline-only A/A stage completed, failed its declared
readiness bound, and stopped before any candidate formal launch. No candidate
native or allocator workload, XLS performance workload, adoption, rejection,
speedup, slowdown, or host-cause claim follows from this record. The independent
raw audit of the bounded packet passes; the maintained decision and cleanup
receipt are recorded below. The maintained decision is uninterpretable and the
cleanup is complete; neither changes the no-adoption disposition.

## Candidate and handoff boundary

0969 evaluates an operation-local CFB writer handoff. A fresh source-layout
writer first validates and emits a `ReusePlan`, then may return an opaque private
receipt carrying the exact package model, source slice, limits, and owned output
bytes. After `write_to`, emission, and flush succeed, the fresh writer's layout
report is observed immediately; a reused-source-layout result gates construction
of that receipt. The report is not carried as receipt state. The source-bound
copy-through route, rewrite, no-source, plan decline, topology change, and every
other non-reused route remain receipt-free.

The editor still performs the ordinary final `OleFile::open` and `codec::open`,
including protection admission, directory validation, and limits. A receipt
route projects final directory metadata without reopening stream payloads: it
requires exact path, kind, size, and one-to-one coverage against the staged
model, takes physical SID/link/CLSID/state/timestamp/sector/size metadata from
the final CFB, then runs `discovery::from_package` in the existing internal
order. A projection mismatch or bounded index-allocation decline uses the
existing full `Package::capture` plus `reuse_stream_allocations` path. The
consumed candidate keeps the baseline internal editor-assignment sequence; the
public commit is published only after capture/projection or fallback and
discovery succeed, and an error leaves the original editor and snapshot state
unchanged.

The candidate is externally bound to source revision
`34fc85e23ccefcd8ad6ffd4d57d9e2671ef1a5a1`. Its candidate tree is
`d39b47d024035112c3f71dd06e3ff3d54829e908e9f612e90b2d25a41ffdc884`, and the
baseline tree is
`9da7843d8b8bce8fe7530a74bfffd2af1ee7d66c0439e93d3c0a57c24c336188`.
The external candidate changes only the codec, editor, and object-test paths;
no corresponding production change is present in this checkout.

## Correctness and owner QA

The independent implementation review is `pass` and approved for the next
stage. It passes receipt binding, final metadata projection, fallback/error
atomicity, source closure, and QA-claim consistency. The implementation seal
reports an external release build with no formal workload, core clippy with
warnings denied, 206 unit tests, 89 integration tests, and 9 focused handoff
tests. The focused tests cover metadata and `Arc` identity, sector-size
transitions, route absence, projection fallback, malformed-open and limit
errors, duplicate-index decline, and bounded index-allocation decline.

The shared XLS correctness tests remain green: `litchi-xls --lib` passed 1,210
tests and the selected `xls_comments_writer`,
`xls_comment_transactions`, `xls_user_names`, `xls_view`, and `xls_ole_objects`
suites passed 39 tests. These are library/integration guards, not an XLS
ignored tests; its remaining rustdoc diagnostic names four pre-existing private
`parse_recording` links in `litchi-doc`, outside the changed OLE codec/editor
paths. The build-equivalence audit found no observed compiler instrumentation or
dependency/profile confound, while withholding formal timing admission.

## A/A readiness result

The v4 baseline-only A/A matrix completed all 576 scheduled children: 24 cells,
12 pairs per cell, and both `AB`/`BA` positions, with no retries. It used seed
`969004`, 10,000 bootstrap resamples, and the declared symmetric p50/p95 upper
bound of `1.05`. The reduction has no packet errors but is
`readiness_fail`:

| Fixture | Cells | Cells with at least one failed gate | Failed gates | Worst symmetric upper bound |
| --- | ---: | ---: | ---: | ---: |
| `FloatingPictures` | 12 | 12 | 42 / 48 | `1.307340` (`e8/release`, `AB` p95; exact `1.3073395498987157`) |
| `NoHeadFoot` | 12 | 0 | 0 / 48 | all four gates pass in every cell |

The A/A failure is a readiness result for the baseline measurement system. It
does not measure the candidate, establish a host explanation, or authorize
selective retries. The stage receipt says `stopped_before_candidate_formal`,
with `formal_candidate_launch_permitted: false`, `timed_controls_started:
false`, and `real_lanes_started: false`. Consequently there are no candidate
native or allocator results to pool, compare, or qualify.

## Audit and disposition

The canonical A/A receipt is
`capture/aa-v4/aa-readiness-v4.json`, SHA-256
`6bb110145d0583be3ff12c3c6e1ada6611be6b5396f55f24698bb0358bb51efe`, and the
stage receipt is `capture/aa-v4/capture-stage-v4.json`, SHA-256
`59485d2c231f2c56e39300acfc5916c5a14ceea91918ebe0c2706fc7fe666bc3`.
The bounded packet is under `results/change-0969/`. Its [independent raw
audit](results/change-0969/review/independent-raw-audit-0969-v1.json) has
SHA-256
`f0d0d0b63a896be41dab84550ff39f59954c26538423be2a3d3d017df9f26775` and
passes custody, validates all 576 packets with zero errors, and reproduces the
exact 42/96 readiness-gate reduction. The read-only audit helper is bound by
SHA-256
`09a8749910a1da49579a8ea782c3a6a11b34a303cd56818ac30d1b500069441a`.
The maintained [qualification decision](results/change-0969/decision/perf-decision-output-v1.json)
has SHA-256
`125818db8bcf05306626230cbedd5864de53991b8a69e7ec7116299ae49ec0ba` and
reports `uninterpretable`: completeness, custody, and correctness pass;
readiness fails; formal performance and benefit are unmeasured/not admitted;
and `candidate_adoption` is `null`. The [cleanup receipt](results/change-0969/receipts/build-cache-cleanup-v1.json)
has SHA-256
`76611af155005a9f266be5c4a5267ef78ba55808cb76a9e1f7961ce75ab68c80`, records
13,264,674,816 freed allocated bytes, and preserves the four bound executables,
referenced metadata, source, raw captures, logs, and receipts. No new
measurement or formal candidate launch is implied.

The A/A readiness failure prevents formal candidate qualification. It neither
adopts nor rejects the implementation and does not justify a speedup,
slowdown, or host-noise explanation. Any future attempt needs a separately
reviewed readiness design or a new authorized measurement plan; it must retain
the private receipt boundary, final physical reopen/protection checks, metadata
projection, and full-capture fallback described above.

## Evidence references

| Evidence | SHA-256 |
| --- | --- |
| `review/implementation-review-v1.json` | `992652e728d24e16816edb20369a9444102d0fcb61e7caa972dbf21591675333` |
| `review/implementation-seal-v3.json` | `dbb9a85aca1459cfcf158b1d5d5cc3a3363e047539494501f5c3255e7c4ffe3f` |
| `review/doc-qa-receipt-v1.json` | `bc9f44519f771c89f097e9b82c4755b4840b9a0bfacf589ee75f034dde1ce11b` |
| `review/build-equivalence-audit-v1.json` | `4184809dd4cd564a7ce2d8ed498e00765a61121abfeaed4b7a78385fca0159b6` |
| `review/independent-raw-audit-0969-v1.json` | `f0d0d0b63a896be41dab84550ff39f59954c26538423be2a3d3d017df9f26775` |
| `analysis/perf-decision-output-v1.json` | `125818db8bcf05306626230cbedd5864de53991b8a69e7ec7116299ae49ec0ba` |
| `receipts/build-cache-cleanup-v1.json` | `76611af155005a9f266be5c4a5267ef78ba55808cb76a9e1f7961ce75ab68c80` |
| `protocol-design-v4.md` | `7766d1022bfef71ceca4fad7bb1edee9529178fb24b11eb083eb76b34fe94564` |
