# Change 0212: ODP cached attribute namespace resolution

Date: 2026-08-19

## Decision

**Banked** (re-verdict after change 0213 calibrated the litchi-odp layout
noise floor). The executed-phase evidence is strong —
`odp_semantic_full_text` accepts all four statistics (20.82%-29.50%),
`odp_semantic_list_slides` p50/mean (19.15%-25.51%),
`odp_semantic_one_slide` p50/mean/p99 (17.48%-31.16%) — all far above the
calibrated floors. The byte-identical `odp_semantic_open` phase showed an
adverse both-directions p50 reading that reproduced in the single permitted
rerun (-1.21%/-1.37% → -1.51/-1.63%); under the pre-floor rule this
withheld the change. The 0213 floor for open p50 is 3.1%, so the reproduced
reading (max 1.63%) is a within-floor layout reading on a phase executing
no changed code and no longer blocks (0205 rule 1, extended to litchi-odp
by 0213). Claim scope: full-text p50/mean/p95/p99, list-slides p50/mean,
one-slide p50/mean/p99 at the magnitudes above.

## Mechanism and invariants

Post-0210/0211 profiling of `odp_semantic_full_text` attributed 39.66%
inclusive to `ElementAttrs::get`, dominated by resolution replay:
every lookup re-ran `resolver().resolve_attribute(key)` for every cached
attribute — `NamespaceResolver::resolve_prefix` was 24.56% inclusive /
10.62% self of the workload (the two `ElementAttrs::lookup` call paths
alone ~12% inclusive), plus `QName::decompose` ~3%.

The change resolves each attribute key ONCE when the incremental scan
first reaches it, storing a `ResolvedAttributeNamespace` snapshot
(`Bound(Vec<u8>)` / `Unbound` / `Unknown`) plus the borrowed local name
alongside the raw attribute (`validation.rs`). Lookup replay compares
the cached snapshot via `ResolvedAttributeNamespace::matches` —
semantically identical to `is_namespace` on the live `ResolveResult` —
with no resolver calls; value decoding still happens per-match at lookup
time, unchanged.

Exactness invariant (audited at all 12 call sites): the resolver's
namespace stack is constant between `ElementAttrs::new(element)` and its
last use, because every site is a straight-line block of lookups on one
element with no intervening event consumption (`drawing_attributes` uses
the resolver read-only). Under the invariant, resolve-at-scan ≡
resolve-at-lookup; resolution cannot error (unknown prefixes → `Unknown`
→ non-match, preserved). The 0210 lazy semantics are untouched:
malformed-attribute errors surface at the same reaching lookup,
duplicate detection stays in the iterator, first match in document order
wins.

Verification: the 7 existing `ElementAttrs` tests pass unmodified; two
new tests pin cached-vs-fresh resolution parity across an 8-target
matrix and prefix shadowing at nested scope (element-scope binding wins,
matching fresh one-shot scans). Full litchi-odp suite: 150 lib + 111
integration, 0 failed; 21 doc tests pass; fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` pass. No
public API change.

## Matched release timing

Two frozen release binaries differ only in the cached resolution; both
carry changes 0192-0196, 0198-0202, 0204, 0206, 0207, 0209, 0210, and
0211. Control SHA-256
`ceba155be185f1c213c4bf90200bb5e87bb697a5023e3a703d9cd7def6042922` (the
banked 0211 binary), candidate SHA-256
`246c6b1f916b2dc2fa8529edfdfed605fede29db0226f54334bd38bc5d4f8e13`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). Pre-floor acceptance applies
(litchi-odp, no calibrated floor): accepts require lower in both paired
directions with clean drifts; any adverse both-directions pattern blocks
unless cleared by the single permitted rerun of that workload.

Executed phases: `odp_semantic_list_slides`, `odp_semantic_one_slide`,
`odp_semantic_full_text`. `odp_semantic_open` executes no changed code.

### odp_semantic_open (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -1.21% | -1.37% | 0.38% | 0.54% | withheld; adverse both dirs → single permitted rerun |
| mean | -2.21% | -2.55% | 0.48% | 0.81% | withheld; adverse both dirs → single permitted rerun |
| p95 | -2.05% | -7.77% | 0.91% | 6.57% | withheld; adverse both dirs → single permitted rerun |
| p99 | -1.20% | -17.11% | 3.31% | 19.56% | withheld (candidate drift over ceiling) |

### odp_semantic_list_slides (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 22.17% | 23.37% | 0.90% | -0.66% | ACCEPTED |
| mean | 19.15% | 25.51% | 3.34% | -4.79% | ACCEPTED |
| p95 | -11.61% | 40.00% | 15.43% | -37.94% | withheld (drift over ceiling) |
| p99 | 3.19% | 28.96% | 20.57% | -11.52% | withheld (drift over ceiling) |

### odp_semantic_one_slide (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 22.48% | 22.78% | 0.32% | -0.08% | ACCEPTED |
| mean | 19.70% | 24.74% | 2.34% | -4.08% | ACCEPTED |
| p95 | 5.03% | 31.15% | 8.99% | -20.99% | withheld (drift over ceiling) |
| p99 | 17.48% | 31.16% | 5.32% | -12.15% | ACCEPTED |

### odp_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 20.82% | 23.51% | 2.41% | -1.08% | ACCEPTED |
| mean | 22.13% | 24.57% | 2.38% | -0.83% | ACCEPTED |
| p95 | 23.84% | 29.50% | 2.39% | -5.21% | ACCEPTED |
| p99 | 26.20% | 23.75% | -0.61% | 2.68% | ACCEPTED |

### odp_semantic_open — rule-2 rerun

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -1.51% | -1.63% | -0.35% | -0.23% | withheld; adverse both dirs REPRODUCED (clean drifts) |
| mean | 2.09% | -1.41% | -4.67% | -1.26% | withheld (disagreeing directions) — primary adverse NOT reproduced |
| p95 | 19.94% | 0.52% | -21.19% | -2.08% | withheld (control drift over ceiling) — primary adverse NOT reproduced |
| p99 | 4.86% | -0.97% | -20.37% | -15.49% | withheld (drift over ceiling) |

The p50 adverse both-directions reading **reproduced** at nearly
identical magnitude (-1.21%/-1.37% primary, -1.51%/-1.63% rerun) on a
phase executing zero changed code — a deterministic per-binary-pair
layout signature (layout noise is deterministic, so reruns reproduce
it; this is precisely why 0205 replaced rerun-adjudication with a
calibrated floor for litchi-ods). The mean/p95 primary adverse readings
did not reproduce. Under pre-floor rules the reproduced p50 reading
withholds the change; under the 0204 precedent the remedy is an
litchi-odp layout-noise floor calibration (0205-analog), after which
this change is eligible for re-verdict.

## Verdict

**Banked.** Re-verdicted under the 0213-calibrated litchi-odp floor: the
reproduced open p50 layout reading (max 1.63%) is within the 3.1% floor on
a non-executed phase and is recorded as a layout reading, not a regression;
the mean/p95 primary adverse readings did not reproduce in the rerun. The
change was re-applied bit-exact after calibration: the rebuilt harness
matches the original candidate SHA-256
`246c6b1f916b2dc2fa8529edfdfed605fede29db0226f54334bd38bc5d4f8e13`, and
the full litchi-odp suite (150 lib + 111 integration), fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass. The banked binary is the new
control for subsequent changes. Claim scope: `odp_semantic_full_text`
p50/mean/p95/p99 (20.82%-29.50% lower), `odp_semantic_list_slides` p50/mean
(19.15%-25.51% lower), `odp_semantic_one_slide` p50/mean/p99
(17.48%-31.16% lower). Raw artifacts:
`docs/performance/results/*-0212-*` and `*-0212r-*` (rerun).
