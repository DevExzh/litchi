# 0961 — XLSX edit-staging capacity trial

Status: rejected by the frozen screening rule. Production remains unchanged.

## Motivation and mechanism

The corrected [0959 owner ranking](results/change-0959/analysis/owner-ranking-summary-v3.json)
attributes 28,570,176 requested allocation bytes across 1,632 calls to
`MultiSourceEdit::apply_batch`. This attribution is cumulative allocation traffic,
not physical copies, resident memory or CPU time. Each scalar `set` calls
`apply_batch` with one request, but the current implementation reserves every
remaining slot under the 256-edit bound before inspecting the iterator.

The isolated candidate uses a capped iterator upper hint for initial pending
capacity, falling back to the existing bound when the hint is unknown. A lying
or underestimated hint must trigger fallible bounded growth before any push;
the actual request count continues to enforce the batch limit. Source parsing,
semantic materialization, publication and inverse validation are unchanged.

This removes a different allocation source from the rejected numeric sharing
and RawCell pre-scan candidates in 0959/0960. It adds no extra worksheet scan,
public API, dependency, cache, unsafe code or concurrency.

## Contract and controls

| Requirement | Candidate boundary |
| --- | --- |
| ADR 0001 / 0005 | Typed fallible reservation; measured selection before adoption. |
| ADR 0002 / 0024 | Private XLSX implementation; existing crate ownership. |
| ADR 0003 | Temporary pending state; existing atomic staging and rollback remain. |
| ADR 0006 | Same count ceiling, selector/duplicate checks and semantic validation. |
| ADR 0008 | Focused adversarial iterator tests, full XLSX library tests, lint and format checks before capture. |
| ADR 0031 | No new execution mechanism; measurements use explicit bounded serial contexts. |

Iterator hints are untrusted optimization inputs. Empty input, unknown or
inflated hints, understated hints, batch overflow, duplicates, invalid selectors,
cancellation and retained staged-state atomicity require controls. Reduced
initial allocation and a possible later fallback necessarily change allocation
failure opportunities under allocator exhaustion. Normal semantic validation
order remains the same; identical allocation-failure timing is not promised.
Calling `into_iter` and `size_hint` before reservation also changes the timing
of caller iterator code relative to the original reservation.

## Custody

The previous turn made progress: 0960's complete second screening epoch was
independently audited, rejected, committed as `4e69e7cda6`, and its two disposable
build caches were removed. All 35 normative input hashes are unchanged.
The new isolated source is exported from that commit, excluding performance
and report evidence, with the retained dependency lock overlay. Its 10,933
file contents match the prior baseline. Archive-extraction file modes initially
differed and were restored to the retained baseline modes before editing.
Sources, raw logs, build products and large evidence live outside the checkout
under `/home/zhuhe/litchi-performance-0961/`.

This record does not claim completion of the broader GOAL. iWork work and
unrelated worktree changes are excluded.


## Candidate quality and screening rule

The isolated implementation passed 1,403 XLSX library tests. After final
**test-only** strengthening, all 34 focused admission tests passed, followed by
warning-denied library Clippy and scoped rustfmt on both changed files. The
full library suite was not rerun after those test-only edits. An initial missing
test import caused a compile failure; the failed log is preserved, followed by
the corrected successful run. The final production source hash is
`7c12048110d604d3b487a46d2dfd6719d99bf0eb96bfa584e62233b3b6575f1d`.

Independent review checked clamping, unknown hints, zero and underreported
upper bounds, fallible bounded fallback, overflow refusal, cancellation after
`size_hint`, and late failure after fallback with preexisting staging retained.
No allocator-failure injection is claimed. Actual allocation recovery is
supported by the explicit `try_reserve_exact` paths, not a forced-OOM run.

The prospective screen uses the same three managed serial cases as 0960:
medium 1%, dense-sparse 1%, and medium one edit. Native capture uses three
ABBA/BAAB blocks, ten warmups and sixty samples per child; a separate allocator
lane uses three paired repeats, two warmups and ten samples. Every process is
fresh, pinned to CPUs 12–15, with one worker and no retries. Descriptive paired
medians cannot establish a statistically qualified effect.

A useful screen requires at least one native primary latency or native RSS or
allocator region-peak ratio at or below 0.95. Any native primary, native RSS,
allocator region peak or allocator RSS median above 1.05 rejects the screen.
Invalid metrics or custody/oracle failures also stop it. Allocation calls and
requested bytes alone do not qualify. A promising screen requires independent
review, fresh six-block confirmation and ordinary/fallback performance controls
before adoption. Exact successful 8-item and 256-item MultiSourceEdit controls
remain required before adoption; the screen does not claim those controls.


## Result and decision

All 36 native and 18 allocator children completed and validated, with 2,160
measured native iterations and 180 allocator rows. Source/output/semantic and
inverse oracles, sink behavior and resource-drain checks passed. The following
are candidate/baseline medians over six native paired legs or three allocator
pairs. They are descriptive screening statistics without confidence intervals,
not qualified performance effects. Allocation metrics are per-child means of
ten drained-core rows; live peaks and process RSS remain distinct metrics.

| Metric | Medium 1% | Dense-sparse 1% | Medium one edit |
| --- | ---: | ---: | ---: |
| Edit/commit p50 | 0.9844 | 0.9874 | 0.9896 |
| Edit/commit p95 | 0.9845 | 0.9934 | 0.9891 |
| Total p50 | 0.9893 | 0.9934 | 0.9931 |
| Total p95 | 0.9913 | 0.9949 | 0.9930 |
| Native peak RSS | 0.9995 | 0.9978 | 1.0037 |
| Allocator requested bytes | 0.9065 | 0.8972 | 0.9953 |
| Allocator allocation calls | 1.0000 | 1.0000 | 1.0000 |
| Allocator region peak live bytes | 1.0000 | 1.0000 | 1.0000 |
| Allocator-process peak RSS | 1.0779 | 1.0050 | 1.0050 |

The frozen selector returns `reject_screen` for two reasons: no eligible metric
reaches 0.95, and medium 1% allocator-process RSS is 1.0779, above the 1.05
safety threshold. The RSS flag is a descriptive review trigger, not proof of a
causal 7.8% regression. Requested traffic falls by 9.3–10.3% in the 1% cases,
but allocation calls and live peaks remain flat. Removing reservation traffic
alone does not establish a useful end-to-end or peak-memory benefit here.

All individual phase, tail and pair flags remain in the bounded reduction.
Dense-sparse ingress p95 has a 1.0594 median ratio and one-edit ingress p99 is
1.0532. The generic reducer field `followup_required_for_confirmation` means
that screened data cannot qualify an effect; it does not override the separately
frozen selector. The rejected candidate does not advance to confirmation or
production, and no speedup is claimed.

The independent audit recomputed raw custody, oracles, allocator conservation
and all metric summaries with zero mismatches. Its first replay stopped on an
inherited no-signal **message string** that differed from the 0961 selector;
decision comparison had passed. The v2 helper changes only that literal and
passes the exact comparison. Both versions and the failed replay receipt are
retained, with no benchmark rerun, result repair or threshold change.

The next investigation should prioritize native edit/commit CPU attribution,
including source semantic materialization and omitted-cell merge ownership,
before another allocation-traffic-only candidate. The 0959 requested-byte
ranking remains useful for memory traffic, but these screens do not establish
that its largest entries dominate end-to-end CPU or live memory.

## Closure

The two disposable 0961 build caches and empty owned temporary directories were
removed, releasing 2,806,525,952 allocated disk bytes. Retained executables still
match their build receipt hashes; source trees, raw evidence and replay tools
remain external. The [bounded packet](results/change-0961/README.md) binds them
and preserves the rejected patch. This batch changes documentation/evidence
only. No broader workspace, cold-cache, parallel-scaling or native Office
certification is inferred from the focused trial. The GOAL remains active.
