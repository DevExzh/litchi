# 0525: retain XLSX unchanged-cell readback reconstruction (accepted)

This is the accepted per-change record for a source-backed OOXML/XLSX
performance experiment. The accepted decision in
[`decision.json`](../results/change-0525/decision.json) retains the candidate
for production. Baseline evidence, including the retained native A2 control,
is complete. The source-bound candidate preflight and native/profile admission
gates pass, and the supplemental four-child eager ABBA confirmation also passes
its guard. Quality has 12 passing gates and 1,293 successful executions.
Cleanup is complete and the strict post-cleanup verifier passes with exact
report replay and owned paths absent. Recursive seal verification also passes.
OLE2 and OOXML are
the active priority. ODF is deferred until that optimization goal is complete,
and iWork is outside scope.

## Motivation and evidence

The retained 0522 source-bound Callgrind attribution identifies the semantic
reconstruction after `MultiSourceEdit::commit` as a large part of the selected
XLSX path. Its aggregate instruction totals are:

| shape / repeat | commit Ir | reconstruction Ir | rewrite Ir |
| --- | ---: | ---: | ---: |
| medium / 1 | 197,179,569 | 121,438,255 | 75,197,769 |
| medium / 2 | 197,203,336 | 121,458,945 | 75,197,024 |
| dense-sparse / 1 | 377,116,324 | 229,721,741 | 146,606,245 |
| dense-sparse / 2 | 377,164,282 | 229,749,026 | 146,625,500 |

The reconstruction and rewrite rows are disjoint attribution branches; their
instruction costs must not be added as elapsed time. The 0522 source review
also records full worksheet parsing at 74,805,030 instructions and the full
XML validator at 46,624,585 for the medium baseline. These observations
motivate a work-elimination experiment; they do not predict a speedup.

The first row-only idea was rejected before implementation and capture. The
frozen source-derived closure report shows 9,216 medium cells and 17,792
dense-sparse cells. Row-only omission would skip 4,752 cells (51.56%) in
medium but only 1,023 (5.75%) in dense-sparse, because its touched rows contain
16,769 cells. Cell-span omission can skip 9,123 (98.99%) and 17,614 (99.00%)
cells respectively. These are deterministic corpus counts, not performance
measurements or allocation estimates.

## Candidate mechanism

The ordinary writer still emits the complete worksheet XML and that output is
the only published payload. The private value-only route records source-bound
cell spans while it writes. Each omission carries its zero-based row and
column bounds and output offsets, and the proof is tied to the exact immutable
source snapshot. A proof from another snapshot, or a source/version or
execution-fence mismatch, cannot authorize reuse.

`Snapshot` first validates the complete emitted XML, preserving the existing
validation boundary and error precedence. If the proof is eligible, it builds
a reduced document from spans of the actual output. The reduced document keeps
the XML/root context, worksheet metadata, columns, defaults, dimension,
sheet-data and row shells, and all non-cell gaps. It retains changed cells and
complete membership-changing rows, while omitting only proven unchanged cell
owners. The existing raw worksheet parser then reads this reduced actual
output, independently decoding changed cells and worksheet metadata.

A private fallible merge combines that parsed Store with the immutable source
Store. Parsed changed cells and metadata are authoritative; source records are
reused only for proven omitted owners. The merge checks address ownership and
duplicates, rebuilds sorted cell and row indexes and all extents, and preserves
style, formula, shared-string, inline-rich and other stored metadata. Staged
values remain expectations for the existing independent publication readback;
they do not populate the changed Store directly. The complete output bytes,
patch bytes, calculation-chain invalidation, byte/resource budgets and no-op
sharing contracts remain unchanged.

The effective action map determines the row closure:

* An untouched row keeps its row shell and can omit its cell owners.
* A replacement-only existing row keeps its shell and actual changed-cell
  spans. The writer emits explicit addresses for changed cells, so deleting
  preceding implicit unchanged cells does not change their parsed address.
* An existing row containing insertion or removal is retained and parsed in
  full, including unchanged followers whose implicit addresses may shift.
* A newly inserted row, shared-formula worksheet, unsupported provenance,
  metadata-changing action, proof mismatch, reduced parse failure, merge
  invariant failure, or checked reservation failure uses the complete parser.

The fallback is recoverable: speculative reduced XML, parsed Store and ranges
are dropped before parsing the complete output. Optional omission-recording
allocation failure discards the proof and keeps ordinary full output. The
candidate does not claim bounded constant memory, zero copying, selective
physical I/O, or support for a broadened worksheet grammar. Full validation
continues to reject unsupported foreign/MCE/extension forms and the existing
shared-string and shared-formula boundaries remain authoritative.

The source-review corrections in
[`preflight-corrections.md`](../results/change-0525/preflight-corrections.md)
recorded the two-dimensional `Rect` bound fix, monotonic omission membership,
the scratch lifetime and single-validation fix, rendering from the existing
layout, recoverable omission allocation, and removal of unnecessary marker
scans. The focused tests are retained in the evidence bundle, but passing
tests alone do not admit the candidate.

## Measured outcome

The completed native ABBA comparison retains 1,640 matched durations: 820 per
stage, covering the four 100-sample primary rows and fourteen 30-sample guard
rows. The original eager guard adds 240 samples across its four rows. Every
primary row passes the frozen 5% total and commit p50 gates:

| shape / repeat | total p50 baseline → candidate (ns) | reduction | commit p50 baseline → candidate (ns) | reduction |
| --- | ---: | ---: | ---: | ---: |
| medium / 1 | 25,504,253 → 22,637,341 | 11.2409% | 11,574,154 → 8,020,944 | 30.6995% |
| medium / 2 | 25,534,145 → 22,298,803 | 12.6706% | 11,590,438 → 8,071,393 | 30.3616% |
| dense-sparse / 1 | 49,082,428 → 42,600,330 | 13.2066% | 21,715,666 → 15,180,594 | 30.0938% |
| dense-sparse / 2 | 50,541,246 → 43,010,549 | 14.9001% | 22,583,895 → 15,634,883 | 30.7698% |

The separate allocator lane retains 40 matched samples. Its vectors are
identical across repeats: medium allocation calls fall from 118,744 to
91,391 (23.0353%), allocated bytes from 20,344,427 to 12,436,307 (38.8712%),
and incremental region peak from 2,984,983 to 2,344,427 (21.4593%). Dense-sparse
calls fall from 225,771 to 172,946 (23.3976%), bytes from 27,331,029 to
18,258,492 (33.1950%), and incremental peak from 7,335,225 to 4,467,508
(39.0951%). Deallocation and reallocation balances reconcile; these are
operation-region measurements, not whole-child RSS.

The four shape/repeat profile pairs (eight isolated profile children) pass the
15% independent instruction gate:

| shape / repeat | commit Ir baseline → candidate | reduction |
| --- | ---: | ---: |
| medium / 1 | 197,188,676 → 133,217,968 | 32.4414% |
| medium / 2 | 197,130,975 → 133,250,185 | 32.4053% |
| dense-sparse / 1 | 377,140,022 → 258,175,362 | 31.5439% |
| dense-sparse / 2 | 377,135,732 → 258,155,623 | 31.5484% |

The profile's raw worksheet parser inclusive Ir falls 96.0266–96.5004% while
the complete validator remains within 0.019% of its baseline. These are
Callgrind attribution results, not additional elapsed-time or allocation
counts. The main comparison retains 31 over-5% adverse rows: 7 open-phase,
19 publication-phase, 4 reopen-phase and 1 whole-child-RSS row. There are no
over-5% main commit or total-elapsed rows. The RSS row is the managed
dense-sparse repeat-2 guard, where whole-child max RSS rises from 85,648 to
90,244 KiB (+5.3661%). Reopen diagnostics are outside the measured edit/save
elapsed time by the harness. The accepted decision retains these observed
phase/RSS tradeoffs only alongside total p50 gains in all four primary rows;
the evidence makes no blanket no-regression or RSS-improvement claim. The
comparison also retains 76 same-build drift rows. The original eager guard has
one retained adverse metric: dense-sparse repeat 2 rises 4.5442% at p50,
4.4781% at the mean, 4.7339% at p95 and 5.8175% at p99; it makes no
primary-gain claim. The supplemental eager confirmation contains four
children (A1/B1/B2/A2), 20 warmups and 100 samples per child, or 400 samples
total. Its confirmation gate passes: A1/B1 candidate change is −2.543162% at
p50 and −2.622600% at the mean; A2/B2 is −3.799717% at p50 and −3.695763% at
the mean.
All p95, p99 and whole-child RSS comparisons remain below the adverse threshold,
with zero >5% matched or same-build drift flags. This bounded diagnostic does
not erase the original eager row shift or establish general eager acceleration
or host stability. The original eager rows remain retained.

The follow-up [`next-priority-review.md`](../results/change-0525/next-priority-review.md)
selects the `scan_with_limit` layout/provenance pass as the primary next source
audit at 53.26% of candidate commit Ir. `Store::merge_omitted_cells` is the
secondary audit at 5.97%. These are profile shares for source review, not a
measured opportunity or permission to edit source.

## Frozen measurement and acceptance plan

The frozen [plan](../results/change-0525/plan.json) compared a source-backed
one-percent edit/save on medium and dense-sparse shapes, with two repeats, 20
warmups and 100 samples. The primary gate required at least 5% p50 improvement
in both whole measured elapsed time and the commit phase for every shape and
repeat, with exact output, semantic, source/provider and resource oracles. It
also required at least 15% commit-Ir reduction in every shape/repeat as an
independent mechanism gate. Both gates pass. Any adverse latency, mean, RSS or
incremental allocation peak over 5% remains retained and reviewed by row.

The native schedule is serial ABBA: baseline native-r1 before candidate
application, candidate native-r1, candidate native-r2, then the retained
baseline native-r2 binary executed under the candidate checkout. The final
control's binary and working-source identities are checked separately. The
planned A2 control is now present and the primary native matrix is complete.

Separate allocator captures use a fresh allocator binary, two repeats and
five samples per primary shape around staged sets and `MultiSourceEdit::commit`;
publication, reopen and oracles are outside that scope. Callgrind uses two
repeats and one measured commit dump (`.4`) per shape after three preceding
retained lifecycle dumps (`.1`–`.3`), with dump-before/dump-after and exact
owner attribution. Hardware uses two repeats and 100 samples per primary shape
for cycles, instructions,
branches, branch misses, page faults, context switches and CPU migrations.
Those counters cover the whole child, including setup, publication, oracles
and drops, and are diagnostic rather than operation-local attribution.

The eager guard separately compares the ordinary rewrite consumer on the two
primary shapes with two repeats, ten warmups and 30 samples. It reviews every
over-5% p50, p95, p99, mean or peak-RSS adverse metric and cannot promote a
primary gain claim. The separate frozen supplemental eager plan compares four
children in A1/B1/B2/A2 order with 20 warmups and 100 samples per child; its
median/mean gate passes with no >5% matched or same-build drift flags. Neither
eager lane establishes general eager acceleration or host stability. ODF
measurement is intentionally absent from this plan.

## Custody and remaining work

The frozen plan and capture driver are hash-bound in the evidence bundle. The
baseline has 451 XLSX source files and source-manifest SHA-256
`4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03`, with an
empty source patch. Candidate source must be a nonempty exact diff entirely
under `crates/litchi-xlsx/`; harness sources remain unchanged. Every stage
will retain its source manifest, patch, binary hash, command receipt, raw
report and scratch-path identity. The baseline, candidate and retained-A2
bindings are replayed through a private Git index.

The supplemental eager plan SHA-256 is
`08bd39c750ef27ef4a337b603f67e432f40f59c8c3a5e424a1d0758d388b0eb9`, and its
capture wrapper SHA-256 is
`2485c9f70cdb475c1eed0bc25d266de3a8ab18739090d1d6c53ae4aaa2e36ca4`. The
confirmation comparison, adverse review and accepted decision bind their
receipts and report hashes in the evidence bundle.

The source-bound preflight is now complete: the first attempt stopped at
compilation because a new facade re-export was missing; after that fix, the
second compiled and ran 978 unit tests but two new tests failed while
constructing a fixture because formatting whitespace made the PackageWriter
reject its input as `NotCompact`. Root corrected the raw ZIP fixture
construction to preserve its indentation without changing production behavior.
The third attempt passed 1,288 test executions and XLSX Clippy. All failed
attempts remain retained evidence, and these are source-quality results rather
than performance results.

The quality lane has all 12 frozen gates passing, with 1,293 successful quality
test executions. This includes the source-bound XLSX and Clippy receipts reused
from preflight-3 under exact manifest equality, plus the other ten serial
checks. The supplemental eager confirmation gate also passes with zero >5%
matched or same-build drift flags.

Separate supplemental validator probes also pass: a valid retained 100-sample
control is accepted, while altered p50 and sink-vector probes are rejected.
These custody/schema probes are outside the 1,293 quality-execution count.

The canonical candidate manifest is now frozen at
`0fab9bafc238611761659bcafdd2e7b5df2543aecb194ac5477c168a2d6e6096`; its
normal release build has passed. Candidate allocator build,
allocator/profile/hardware comparisons, the original eager analysis and the
supplemental eager confirmation have passed their frozen analyzers. The
accepted [`decision.json`](../results/change-0525/decision.json) records
`disposition: accepted`, `production_change_retained: true`, and the passing
native, profile and eager-confirmation gates. Cleanup, the full verifier,
post-cleanup replay and owned-path check are complete, with strict verification
passing in [`verification.json`](../results/change-0525/verification.json) and
cleanup recorded in [`cleanup.json`](../results/change-0525/cleanup.json). The
recursive evidence seal covers 550 files, and `verify.py --sealed --strict`
passes. Acceptance, cleanup and evidence closure are complete; the overall
OLE2/OOXML optimization goal remains open.
