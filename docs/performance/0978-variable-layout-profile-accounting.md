# 0978 — variable-layout profile accounting

**Status: failed native qualification; the serial run stopped at the child-03
loss gate.**

0978 responds to the immutable 0977 child-02 admission failure. It evaluates
the retained child streams with a variable-layout raw scanner and a separate
lifecycle/accounting design. The current source is HEAD
`172d66458abacc224427bc9b86cf2d3bf3effe78`; no production source or runtime
decoder was changed. The pre-capture offline checks passed, then one reviewed
native Office run was bound to plan v3 and a root GO. It stopped serially at
child 03 after the loss gate failed; the 0977 failed cohort remains historical
evidence and is not revived by these replays.

## Purpose and scope

The retained event contract is the `perf evlist -v` contract for
`instructions:u`, period 1,000,000, CPU 31, and
`IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER`, with a 32 KiB user
stack cap. The scanner binds that contract to the primary
[Linux v6.12 perf UAPI](https://raw.githubusercontent.com/torvalds/linux/v6.12/include/uapi/linux/perf_event.h)
and accounts record size before reading optional fields. Its source is the
external [0978 raw scanner](results/change-0978/tools/raw_stack_0978.py).

The accounting policy distinguishes a valid variable-layout sample with no
usable user ancestry from a malformed or unaccounted record. A valid empty or
otherwise unassignable sample retains its full period in an explicit unknown
bucket. It is not silently dropped and it does not fail solely because the
whole process contains empty callchains. A record truncated before its
declared layout, a size or identity mismatch, invalid period or bounds, raw or
terminal loss, or a failed conservation/lifecycle gate remains a hard reject.

## Retained raw replays

The new scanner was run offline against immutable 0977 child inputs. Both
replays report complete sample framing, valid CPU/PID/TID and period fields,
sample conservation, and no scanner violations. The scanner's unconsumed count
is the other perf record types that it does not itself interpret; it is not data
loss. The reviewed lifecycle helper separately frames and checks those records
for the offline decoder/reducer path.

| Retained input | Sample records | Observed layout | Result |
| --- | ---: | --- | --- |
| 0977 child 01 | 7,200 | 7,200 records of 33,008 bytes; `REGS_USER` ABI 64 and nonzero stack | Complete variable-layout scan; sample conservation passed |
| 0977 child 02 | 7,201 | 7,200 records of 33,008 bytes plus one 72-byte record; ABI 64 for 7,200 and `NONE` for one; stack size 32,768 for 7,200 and zero for one | Complete variable-layout scan; sample conservation passed |

The 72-byte child-02 record is the sample that the 0977 fixed-offset scanner
classified as truncated. Under the primary ABI it is a valid framed sample
with `callchain_nr=0`, `REGS_USER` ABI `NONE`, and zero `STACK_USER` bytes.
The policy and temporary reducer therefore preserve it as explicit
valid-but-unassignable period. This does not establish a PMU, kernel, stack,
or source cause.

The focused scanner controls pass, and ten independent mutation controls pass
for variable `CALLCHAIN`/register layouts, zero-stack framing, missing or
duplicated pairing, raw/text mismatches, duplicate evlist fields, header
size/type errors, and related conservation cases. The final independent raw
review binds the scanner source and limits its conclusion to this offline
interface; it does not authorize a native adapter or workload.

## Child-02 diagnostic accounting

The direct corrected `perf script` run and forensic ledger are authoritative
for this diagnostic. An earlier focused-decoder receipt was terminated during
its redundant dump phase and is retained as process history; it is not treated
as a natural scanner rejection. The corrected offline decode conserves all
7,201 samples and all 7,201,000,000 instruction periods, with raw loss zero
and terminal `LOST_SAMPLES=0`.

It reports 76 empty callchains, all before the bound target `doc_retention_p`
`COMM_EXEC` while the PID is still the Python launcher. One is the 72-byte
kernel-context sample; the other 75 have full 33,008-byte records and raw
stack payloads but no decoded frame chain. These are explicit unknown or
unresolved accounting mass, not evidence that records were lost or malformed.
The target marker appears in 945 samples with 945,000,000 periods. Marker
physical-leaf conservation passes for that marker subset, but it does not
make the whole child source-attributable or qualify the failed 0977 cohort.

The marker wraps five warmups and 80 measured calls together. The native
sample stream has no measured-call index, so this record does not relabel the
marker as measured-only.

## Temporary offline reducer replay

The temporary offline reducer integration consumes the copied child-02 raw
inputs, corrected script ledger, and variable-layout scan without changing the
0977 capture. Its replay receipt reports `pass` for this one-child diagnostic:
`D=7,201,000,000`, `M=945,000,000`, and `R=6,256,000,000`; the target window
is `T=6,788,000,000`, consisting of `6,738,000,000` known physical-leaf
period and `U_target=50,000,000` explicit unknown period. The marker ledger
contains `U_marker=6,000,000` across six samples, with the remaining marker
period conserved to known physical leaves.

This pass demonstrates temporary offline accounting and conservation on the
retained child only. It does not admit a 12-child cohort, resolve source
ownership, or support an optimization claim. The final v4 reducer review and
packet-validation v4 bind the corrected offline path used by the native plan;
they do not turn the failed native run into a qualified cohort.

## Fresh offline packet integration

A fresh retained-child offline decode ran all five reader commands successfully
and produced a complete child-02 packet with 7,201 samples, 945 marker
samples, raw loss zero, and terminal `LOST_SAMPLES=0`. The input and bound
decoder/scanner/lifecycle bytes were unchanged before and after the replay;
this is a re-decode of the retained 0977 capture, not a native 0978 run. The
versioned packet-validation v4 correction passes its 23 scoped checks, and the
final reducer review accepts the v4 owner-bounds metadata correction. This
offline evidence remains separate from native cohort qualification.

That replay used the original decoder at SHA-256 `866723324881a57ce7c4282cdfd2dabe497397b1612e8d7c4f975cb7e9fe2134`.
The native plan instead binds decoder v3 at `9596e246c7c627fb3810e7c6c6d3dda12af58b62477f1f1233654c4256e2dac9`,
whose additional lifecycle-attribute guard passed separate controls and review.

## Accounting meanings and bounds

The policy keeps these quantities disjoint and visible:

| Symbol | Meaning in the diagnostic/policy | Value or disposition |
| --- | --- | --- |
| `D` | All valid `instructions:u` sample period in the process lifetime | `7,201,000,000` |
| `M` | Period of samples containing the preregistered marker | `945,000,000` |
| `R` | `D - M`, reported as marker ancestry not observed | `6,256,000,000`; it is not described as work outside the marker solely from subtraction |
| `T` | Target execution-window period after exact target `COMM_EXEC`/active map binding | `6,788,000,000` in the temporary child-02 replay; no cohort value exists |
| `K_j` | Exclusive physical-leaf period assigned to owner `j` after active DSO/map validation | Known physical total `6,738,000,000`; no current target-source owner claim |
| `U_target` | Target-window period with no defensible physical owner or ancestry | `50,000,000`; explicit and conserved, never discarded |
| `U_marker` | Marker period with no defensible physical leaf | `6,000,000` across six samples; explicit and conserved |

For a future admitted cohort, `T = sum(K_j) + U_target` and marker period
`M = sum(marker-exclusive-owner-period) + U_marker`. Report the bound
`[K_j/T, (K_j + U_target)/T]` for every owner; nominate a candidate or make a
source claim only when the pre-registered decision is invariant across the
unknown allocation. Inclusive ancestor tables overlap and are descriptive;
they cannot replace this disjoint physical-owner ledger. No current 0978
result identifies an owner, optimization, speedup, or adoption candidate.

## Pre-capture readiness and native outcome

The independent policy review v3 accepts the accounting direction: valid
unknown periods remain visible, zero empty callchains across the entire
lifetime is not a required gate, and target-source eligibility is distinct
from known libc/loader/Python/other-DSO physical ownership. The lifecycle
helper, decoder guard, and scoped reducer reviews then passed their bounded
offline checks. Those checks define the accounting boundary; they do not
themselves provide performance or source attribution.

The lifecycle-ledger review passes its scoped helper checks: raw-byte joins,
record offsets, loss binding, sample-ID handling, PF_X eligibility, marker
unknown conservation, and ABI controls pass within the reviewed offline
scope. The decoder guard, packet-validation v4, and reducer v4 reviews
separately bind the corrected decoder/reducer path. These reviews do not
establish a source or performance result.

The independent native-boundary review v4 and final readiness review pass
static admission: exact argv/affinity/cleanup constants, serial
stop-without-retry behavior, and the plan-bound decoder/reducer sources match
the retained boundary. Plan v3 and the root GO then bound one native Office
run. These reviews validate the execution contract; they do not qualify its
measurements.

The native run on 2026-10-08 started at 09:37:30 UTC and stopped at 09:47:18 UTC with runner
return code 2. Child 01 gated with 7,200 samples and 955 marker samples; child
02 gated with 7,201 samples and 950 marker samples. Both had raw loss zero and
terminal `LOST_SAMPLES=0`. Child 03's native command returned zero, but its
decoder loss gate failed before a packet was emitted. The independent audit
found two `PERF_RECORD_LOST` records with values 1 and 3 (raw sum 4), plus one
`PERF_RECORD_LOST_SAMPLES` terminal value of 4. These are two views of the
same loss and must not be added. The process group was cleaned up; children
04--12 were never launched, and the runner did not retry or replace child 03.
The independent post-capture audit is SHA-256
`3065afc51308416e01f303f4acf20d052e8e9809cf8c9f2fb6b8c846fd5dadb0`.

The v4 reducer was invoked once, returned 2, and wrote no cohort output after
rejecting the incomplete manifest before raw parsing. The two passing child
packets therefore remain individual custody evidence and cannot be promoted to
a 12-child reduction.

## Disposition and next step

0978 is a failed native qualification diagnostic with a bounded offline
accounting enabler. The retained 0977 failure, raw captures, and decoder inputs
are immutable. No 12-child reduction, timing comparison, source-owner ranking,
speedup, adoption, or production change follows from this run. The child-03
loss does not identify a host or source cause, and the child-01/02 pass does
not supply causal attribution.

The next diagnostic step, if pursued separately, is a reviewed capture-loss
and decoder-accounting design that preserves the raw/terminal loss gate and
serial stop behavior before another cohort is attempted. The current evidence
does not justify relaxing the gate, retrying child 03, or inferring a source or
performance cause.

Primary external references are the
[0978 protocol proposal](results/change-0978/design/protocol-design-v1.md),
[policy review v3](results/change-0978/review/independent-design-review-v3.json),
[raw-scanner control receipt](results/change-0978/receipts/raw-stack-scanner-controls-0978-v2.json),
[final independent raw tests](results/change-0978/review/independent-raw-stack-tests-v1-final.json),
[temporary child-02 reducer replay v6](results/change-0978/receipts/replay-receipt-v6.json),
[fresh offline decode](results/change-0978/analysis/fresh-offline-child02-result-v1.json),
[fresh packet validation](results/change-0978/analysis/fresh-offline-child02-packet-validation-v4.json),
[lifecycle-ledger review](results/change-0978/review/lifecycle-ledger-review-v4.json),
[native-boundary review](results/change-0978/review/independent-native-boundary-review-v4.json),
[scoped reducer approval](results/change-0978/review/independent-reducer-scoped-approval-v3.json),
the [child-02 offline diagnostic](results/change-0978/analysis/child-02-offline-diagnostic-v1.json),
[native capture summary](results/change-0978/analysis/native-capture-summary-v1.json),
[post-capture failed-cohort audit](results/change-0978/review/postcapture-failed-cohort-audit-v1.json),
and [reducer rejection receipt](results/change-0978/receipts/instruction-profile-reduction-v1-receipt.json).
