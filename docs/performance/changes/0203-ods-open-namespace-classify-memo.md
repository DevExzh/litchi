# Change 0203: ODS open namespace-classification memoization — WITHHELD, reverted

Date: 2026-08-19

## Verdict

**Not banked. The implementation was reverted; no test or code from it
remains.** The target was real in the profile — quick_xml
`NamespaceResolver::resolve_event` held 6.25% of post-0202 source-backed
open self samples (6.53% on the commit path) because `resolve_prefix`
reverse-linear-scans the whole binding stack for every Start/End/Empty
event, and ODS root elements declare dozens of namespaces — but two
successive implementations both failed the no-regression-pattern standard:
v1 (HashMap memo) measured a mechanism-confirmed regression on every
workload, and v2 (direct-mapped slot cache) measured the targeted
source-open phase neutral while reading adverse in both paired directions
across most withheld statistics, including source-identical phases. The
tree is restored byte-exact to the 0202 state (harness binary SHA-256
matches the frozen 0202 candidate
`475cf2898880363517eec9e0a9ac6b582eed1f78054f161f394bdd635bb19d7d`; the
`litchi-ods` suite is back to 356 tests, all passing).

## What was implemented (and then reverted)

Both versions replaced `read_resolved_event_into` in the fused open
driver (`crates/litchi-ods/src/open_parse.rs`) with `read_event_into` plus
a driver-local memo of the five per-pass namespace classifications,
invalidated by event-level mutation tracking: a shadow stack of
per-element flags recording whether the raw tag bytes contain `xmlns` (a
conservative superset of actually declaring namespaces), cleared on any
declaring Start/Empty push or the pending pop of a declaring element's
End. Attribute resolution inside the handlers kept using the reader's
resolver directly. A content-fingerprint alternative (bindings count +
resolver buffer bytes) was rejected during design with a concrete
counterexample: `xmlns:p="xy"` and `xmlns:px="y"` append identical bytes
with identical binding counts while resolving `p:e` differently, and the
split points are private to quick_xml.

- **v1**: `HashMap<Box<[u8]>, Classifications>` (std SipHash).
- **v2**: 64-slot direct-mapped array cache (~2.9 KB, struct-resident),
  fxhash-style multiply-fold index over 8-byte name chunks, generation
  counters for O(1) invalidation, 32-byte inline name cap with direct
  resolve beyond it, precomputed Unbound classifications for
  non-structural events. Hit path: ~2 hash chunk iterations + generation
  compare + ≤32-byte memcmp; no heap, no SipHash, no allocation.

Correctness evidence was strong in both versions: synthetic-document
tests against a reference loop covered prefix shadowing with scope
revert, `xmlns:p=""` unbinding (the zero-value pop case), valid and
invalid `xml` prefix rebinds (invalid mapped identically to the
standalone validator), default-namespace unset/restore, and slot-cap
bypass; the corpus-equivalence oracle and all precedence pins passed
unchanged (363 tests at v2).

## Measured outcome

Frozen cross-binary CPU-2 A/B/B/A (30 warmups, 500 samples, drift
ceilings 5%/5%/10%/15%); control is the banked 0202 binary in both runs.

v1 (candidate SHA `260f35890dbbf0a840ae1806e3327bb53a4777a2fafa1ce26f8691e7c480fa91`):
source-open 5.05%-15.22% slower in both directions with clean drifts;
one-edit lifecycle/commit, one-percent lifecycle/commit, repeated-edit
total/stage/commit, and even the source-identical eager open all adverse
in both directions. Profiling confirmed the mechanism: `resolve_event`
was eliminated but the SipHash + heap-indirected HashMap probe
(`lookup_or_resolve` 2.14% self + `hash_one` 1.57% self + key memcmp
growth) cost more than the 40-entry linear scan it replaced. Withheld on
the 0197 precedent (mechanism-plausible adverse pattern) and superseded
by v2.

v2 (candidate SHA `88d7c4c3d22c46d1c7fdb6a3cf6cbe9cd6567042aed07ef83ded19745772d24e`):
the targeted source-open read neutral — p50/mean 0.03%-0.87% slower in
both directions (sub-1%), p95 accepted 0.07%/2.26% lower, p99 mixed with
a control-drift violation — while the guardrails read broadly adverse in
both directions: eager-open p50/mean/p95 0.15%-0.79% slower
(source-identical phase), one-edit lifecycle p50-p99 0.32%-7.39% slower
and commit p50-p99 0.39%-6.74% slower (commit is source-identical),
repeated-edit total/stage/publication slower in both directions. The
one-percent lifecycle and commit p50/mean accepted 0.13%-1.23% lower —
the only accepts besides the two tail statistics above — insufficient
against the adverse pattern on the unchanged phases. No rerun was spent:
no statistic showed a win worth rescuing, and the adverse pattern was
consistent. Withheld and reverted.

## Lesson for the record

quick_xml's 40-entry bindings reverse-scan is cheaper on these workloads
than both a SipHash HashMap probe (v1) and a direct-mapped slot cache
with per-tag `xmlns` scanning and shadow-stack maintenance (v2); the
6.25% `resolve_event` self-time under dwarf call-graph attribution did
not convert to wall time once replaced. A future attempt at the remaining
namespace-machinery cost should target `NsReader::process_event`'s
per-tag attribute scan (4.78% self post-0202) — likely by replacing
`NsReader` with a litchi-owned incremental resolver — rather than
memoizing `resolve_event`, and should expect the same binary-layout
sensitivity this change hit twice.
