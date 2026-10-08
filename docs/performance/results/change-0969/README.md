# 0969 bounded diagnostic evidence packet

**Status: sealed bounded diagnostic packet; baseline-only A/A readiness failed
and stopped before candidate formal launch. The independent A/A audit passed,
the maintained decision is `uninterpretable`, and cleanup completed.** This
packet records bounded source/correctness/build evidence and the authoritative
A/A stop evidence. It makes no timing, allocation, RSS, speedup, causality,
candidate-adoption, or XLS claim.

The candidate is the operation-local CFB writer handoff described by
[`design/design-v1.md`](design/design-v1.md). It is bound to source revision
`34fc85e23ccefcd8ad6ffd4d57d9e2671ef1a5a1`, changes exactly
`crates/litchi-ole-common/src/object/codec.rs`,
`crates/litchi-ole-common/src/object/editor.rs`, and
`crates/litchi-ole-common/tests/object.rs`, and is retained as the
[`repro/candidate-implementation-v2.patch`](repro/candidate-implementation-v2.patch)
with SHA-256
`219c799de930a1a73a54b0ef545a28cea5eb9a503f0e4e3e232d9753288bd04f`. The
candidate source tree, baseline tree, source-file manifest, and fixture trees
remain external; the bounded manifests in `provenance/` bind their sizes,
paths, tree digests, and changed-file hashes.

The implementation seal and reviews are copied for review:

- [`review/implementation-seal-v3.json`](review/implementation-seal-v3.json)
  records the sealed external candidate, 206 unit tests, 89 integration tests,
  9 focused handoff tests, 1,210 XLS library tests, and 39 selected XLS
  integration tests. It reports no formal workload and no main-checkout write.
- [`review/implementation-review-v1.json`](review/implementation-review-v1.json)
  passes receipt binding, projection correctness, fallback/error atomicity,
  source closure, and QA-claim consistency.
- [`review/build-receipt-v1.json`](review/build-receipt-v1.json) records the
  external release build and two candidate executable hashes. The executables
  stay outside the repository.
- [`review/doc-qa-receipt-v1.json`](review/doc-qa-receipt-v1.json) records the
  QA result `qa_pass_with_unrelated_rustdoc_lint_failure`; the passing clippy,
  test, and selected XLS checks are correctness evidence only.
- [`review/build-equivalence-audit-v1.json`](review/build-equivalence-audit-v1.json),
  [`review/design-review-v1.json`](review/design-review-v1.json), and
  [`review/baseline-reuse-v1.json`](review/baseline-reuse-v1.json) are bounded
  read-only reviews. They do not qualify a formal candidate run.

The used A/A protocol is [`protocol/protocol-design-v4.md`](protocol/protocol-design-v4.md)
with the sealed plan [`protocol/plan-0969-aa-v6.json`](protocol/plan-0969-aa-v6.json),
protocol seal [`protocol/protocol-seal-aa-v7.json`](protocol/protocol-seal-aa-v7.json),
and baseline-only root GO [`receipts/root-go-aa-v1.json`](receipts/root-go-aa-v1.json).
The v4 helper sources used by that plan are copied under `tools/` with their
source hashes. The plan has 24 cells, 12 balanced A/A pairs per cell, 80
retained samples, five warmups, AB/BA order, bootstrap seed 969004, 10,000
resamples, and a 1.05 symmetric upper bound.

The authoritative capture stopped after 576 baseline A/A children. The
readiness receipt [`capture/aa-readiness-v4.json`](capture/aa-readiness-v4.json)
has status `readiness_fail`; its errors list is empty, but one or more fixed
symmetric upper confidence bounds exceeded 1.05. The stage receipt
[`capture/capture-stage-v4.json`](capture/capture-stage-v4.json) records
`stopped_before_candidate_formal`, with no candidate native, candidate
allocator, or XLS children. The root terminal receipt
[`receipts/root-terminal-aa-v1.json`](receipts/root-terminal-aa-v1.json)
records exit code 2, no retry, 576 packets, and zero candidate/formal/XLS
children. The raw A/A capture remains external at
`/home/zhuhe/litchi-performance-0969/capture/aa-v4`: 2,306 files totaling
539,159,221 bytes; the largest is 697,438 bytes. No raw packet, stdout, RSS,
or stderr file is copied here.

No candidate formal audit exists because no candidate formal run was admitted.
The packet copies the authoritative independent A/A audit
[`review/independent-raw-audit-0969-v1.json`](review/independent-raw-audit-0969-v1.json),
which found 576 valid packets, zero errors, custody/oracle passes, and exact
agreement with the 42-of-96 readiness failure. The packet keeps
[`review/formal-admission-v5-independent-review-v1.json`](review/formal-admission-v5-independent-review-v1.json)
as a historical blocker review: v5 formal admission was not approved. The
v5 and v6 candidate-binding notes are described in
[`provenance/pending-artifacts-0969-v1.json`](provenance/pending-artifacts-0969-v1.json)
and remain unapproved historical material. The maintained decision is copied
under [`decision/perf-decision-output-v1.json`](decision/perf-decision-output-v1.json): its result is
`uninterpretable` with `candidate_adoption: null`, because formal candidate
lanes were not admitted or measured. The independent replay receipt is
[`review/perf-decision-replay-independent-v1.json`](review/perf-decision-replay-independent-v1.json).

The XLS owner guard was never launched. The original post-AA v1 status receipt
is preserved as historical incident evidence at
[`review/incident/xls-helper-status-post-aa-v1-invalidated.json`](review/incident/xls-helper-status-post-aa-v1-invalidated.json)
(SHA-256
`df02744a38c923fd0c6ab0c6fd69ab77776c950ae298c78e9a9c485638f23130`). Root
identified a factual design hash error in the v1 receipt and unexecuted
post-seal edits to the XLS guard/test drafts. The v2 design bytes were
unchanged; only the guard/tests changed after the old seal. That receipt and
its old seal are invalidated and provide no approval. The corrected v2 receipt is copied at
[`review/xls-helper-status-post-aa-v2-correction.json`](review/xls-helper-status-post-aa-v2-correction.json)
(SHA-256
`8ac0a66f07e715b77028c0df885688d3bde02a5f58bb98ab5c0e9296b315b845`). It
directly binds the retained A/A artifacts: 24 cells, 12 pairs per cell, 96
checks with 42 failed and 54 passed, bootstrap seed 969004, and no workload or
formal launch. Its status remains `unqualified_no_adoption`; it corrects the
receipt facts without approving XLS machinery. The external descriptors and
their non-authoritative statuses are recorded in
[`provenance/pending-artifacts-0969-v1.json`](provenance/pending-artifacts-0969-v1.json);
no XLS helper code is copied into this packet. The unexecuted XLS designs and
helpers provide no qualification for the DOC candidate.

Source exports, the 19,046,400-byte source archive, candidate and baseline
executables, build target directories, and raw captures remain external. Their
sizes, SHA-256 values, tree digests, file counts, and bounded regeneration
commands are recorded in
[`provenance/external-inventory-0969-v1.json`](provenance/external-inventory-0969-v1.json).
The source archive SHA-256 is
`297fadf1b743cb412a403f5dd58f33ccf0fbae546f5515814bb5cbedd51c4b2d` and it
is not copied. Every copied packet file is below 1 MiB and no compiled
executable is included. Cleanup authorization and completion are retained in
[`receipts/root-go-cleanup-v1.json`](receipts/root-go-cleanup-v1.json),
[`receipts/cleanup-plan-v1.json`](receipts/cleanup-plan-v1.json), and
[`receipts/build-cache-cleanup-v1.json`](receipts/build-cache-cleanup-v1.json);
the cleanup freed 13,264,674,816 allocated bytes while preserving the four
bound executables, referenced metadata, source trees, raw capture, logs, and
receipts.

For source-only reproduction, apply the patch with `/usr/bin/patch` in an
isolated scratch source tree and verify the three changed files against the
copied source-file manifest. The external build receipt gives the exact
locked offline Cargo command. The A/A helper self-tests were recorded by the
sealed plan and root terminal; this packet performs no build, workload, retry,
or capture rerun.

The coordinating draft record remains external at
`/home/zhuhe/litchi-performance-0969/draft-change-record-v1.md`. The packet
inventory is [`packet-manifest-0969-v1.json`](packet-manifest-0969-v1.json).
The manifest binds the sealed packet contents and excludes itself.
