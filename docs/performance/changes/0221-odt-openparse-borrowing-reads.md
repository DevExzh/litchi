# Change 0221: ODT fused open-parse borrowing reads and depth-gated resolution

Date: 2026-08-19

## Decision

**Banked.** Both executed source-backed open workloads accept p50/mean/p95
in both paired directions with clean drifts: `odt_file_source_open`
43.28%-46.30% lower, `odt_file_source_open_full_text_lifecycle`
20.55%-23.22% lower. All five byte-identical guardrail phases are clean or
within-floor: the `odt_semantic_open` p99 primary adverse-both pattern
(max 32.81%, over the 28.2% floor) did not reproduce at magnitude in the
single permitted rerun (max 11.68%, within floor); the 0220 watch-listed
`ods_file_source_open` p95 reads within floor (max 2.30% vs 4.5%) under
this layout — the 0220 flag is CLEARED. Claim scope:
`odt_file_source_open` p50/mean/p95 and
`odt_file_source_open_full_text_lifecycle` p50/mean/p95 at the magnitudes
above.

## Mechanism and invariants

The post-0220 profiles (0221 targeting analysis) showed the fused
source-backed ODT open (`OpenParse::run`,
`crates/litchi-odt/src/document/open_parse.rs:80`) carrying the same two
costs 0220 removed from the odf-common validator: the per-event buffer
copy of `read_resolved_event_into` (~7.3% of process) and per-event
namespace resolution (`resolve_event` 8.23% of process), with the
resolved value consumed only at depth ≤ 2. Removable ≈ 35%-38% of the
timed `odt_file_source_open` region. The change replays the banked 0220
transformation on the identical loop shape:

- Borrowing `reader.read_event()` — events borrow `xml` directly; the
  scratch `Vec` and per-event `buffer.clear()` are gone. Binding push/pop
  still runs inside `read_event` for every event, so prefix-rebinding
  semantics and the tokenization error stream are unchanged.
- Namespace resolution via `reader.resolver().resolve_element()` only for
  `Start`/`Empty` events at depth ≤ 2, read from `validate.depth` BEFORE
  `on_event` mutates it. Consumer audit: `ValidateHandler::on_event`
  consumes `office` only in Start arms at depth 0/1/2 and Empty arms at
  depth 1/2 (exact replica of the 0220 validator arms);
  `StyleHandler::on_event` never resolves (raw-qname `style:style` byte
  match, pinned by the existing literal-match tests). Bindings declared
  at depth ≥ 3 scope only their own subtree, so no consumed resolution
  can change. No arm needed a fallback. The gate resolves at depth 0 for
  `Empty` events whose result is unused (the invalid-empty-root arm
  errors first) — a harmless superset keeping the gate a simple
  `depth <= 2`.

The pre-change loop body survives verbatim as the cfg(test) oracle
`run_buffered_oracle`; 2 new tests cross-check the gated borrowing loop
against it: 26 synthetic edge cases (rebinding on body/family/deep,
rebind-and-restore nesting, aliased/default/second-prefix office
bindings, unknown prefixes, forms ordering, mismatched ends,
comment/PI/CDATA interleavings, truncations, missing body/family/root,
style collection under a deep rebinding, style-error-before-malformed-XML
precedence) and all 69 ODT corpus fixtures — accept/reject parity,
byte-identical error messages, and identical sorted style projections.
The pre-existing fused-vs-sequential oracle tests now exercise the new
loop and still pass (litchi-odt suite 891 → 893).

Executed phases: `odt_file_source_open`,
`odt_file_source_open_full_text_lifecycle` (both time `Document::open`
through the source-backed fused parse). Byte-identical guardrails:
`odt_semantic_open`, `odt_file_eager_open` (owned path shares no executed
code with `OpenParse` — zero `open_parse` samples in the semantic
profile, 0.06% one-time setup in eager), `odt_semantic_full_text`,
`odp_semantic_open`, `ods_file_source_open`. Edit/save paths untouched.
No public API change.

Verification: 893 litchi-odt tests pass, 0 failed; fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the `OpenParse::run`
transformation; both carry the banked tranche through 0220. Control
SHA-256 `1971c3adf37536d1fa0963d111b79dade639ea5b624150b985faa6ab02bf8713`
(the banked 0220 binary), candidate SHA-256
`93c2279b9b5dff79bbfd58e028c5eedede38c3a917b9f3fc8edb86a4fb0641c7`.
Binary `.text` delta +2,336 bytes (different inlining of the leaner loop;
oracle code is cfg(test)-only). Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and 500
retained samples per leg, drift ceilings 5%/5%/10%/15%
(p50/mean/p95/p99). The 0218 litchi-odt and 0205/0213 ODS/ODP floors
apply per phase; the two executed phases have no calibrated floor, so
accepts there are claimed directly under the pre-floor rule.

### odt_file_source_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 45.72% | 45.32% | 0.13% | 0.87% | ACCEPTED — **claimed** |
| mean | 46.30% | 45.15% | -1.10% | 1.02% | ACCEPTED — **claimed** |
| p95 | 44.95% | 43.28% | -0.44% | 2.57% | ACCEPTED — **claimed** |
| p99 | 56.25% | 44.78% | -21.15% | -0.48% | rejected (control drift over 15% ceiling) |

### odt_file_source_open_full_text_lifecycle (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 23.13% | 20.55% | -3.01% | 0.25% | ACCEPTED — **claimed** |
| mean | 23.22% | 20.74% | -2.49% | 0.66% | ACCEPTED — **claimed** |
| p95 | 22.79% | 22.24% | -0.14% | 0.57% | ACCEPTED — **claimed** |
| p99 | 24.67% | 9.85% | -1.48% | 17.90% | rejected (candidate drift over 15% ceiling) |

### odt_semantic_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.47% | 0.97% | 2.24% | 0.78% | withheld (disagreeing directions) |
| mean | -6.15% | -1.84% | 1.23% | -2.87% | adverse both dirs, max 6.15% within the 7.2% floor — layout reading |
| p95 | -14.57% | -6.37% | -2.09% | -9.10% | adverse both dirs, max 14.57% within the 27.6% floor — layout reading |
| p99 | -32.81% | -3.14% | -0.83% | -22.98% | adverse both dirs, max 32.81% over the 28.2% floor → single permitted rerun |

Rerun (`*-0221r-*`): p99 adverse-both reproduced in sign only at
-10.09%/-11.68% — max 11.68% is within the 28.2% floor, so the primary
above-floor pattern did NOT reproduce; layout reading under rule 1.
Rerun p50 accepted favorable; mean/p95 disagree with drifts over ceiling
(this workload's tails remain the wildest in the ODT set).

### odt_file_eager_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.75% | -2.68% | 1.36% | 5.92% | withheld (disagreeing directions) |
| mean | 2.21% | -2.46% | 1.01% | 5.84% | withheld (disagreeing directions) |
| p95 | 3.58% | -2.98% | -2.47% | 4.17% | withheld (disagreeing directions) |
| p99 | 0.42% | -3.15% | -2.92% | 0.55% | withheld (disagreeing directions) |

Clean: no adverse both-directions statistic.

### odt_semantic_full_text (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 5.19% | 5.30% | -1.47% | -1.58% | accepted favorable — layout-favorable, not claimed |
| mean | 5.87% | 6.93% | -1.42% | -2.53% | accepted favorable — layout-favorable, not claimed |
| p95 | -18.39% | 4.35% | -2.56% | -21.27% | withheld (disagreeing directions) |
| p99 | 27.21% | 18.69% | -7.97% | 2.81% | accepted favorable — layout-favorable, not claimed |

### odp_semantic_open (byte-identical guardrail)

All four statistics disagree across paired directions (p50 -1.75/+0.33...
p99 +3.54/+38.52 with control drift +34.98%) — clean, no adverse
both-directions pattern.

### ods_file_source_open (byte-identical guardrail, 0220 watch-listed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 2.77% | -0.52% | -3.55% | -0.28% | withheld (disagreeing directions) |
| mean | 2.32% | -0.38% | -3.45% | -0.79% | withheld (disagreeing directions) |
| p95 | -0.99% | -2.30% | -3.70% | -2.44% | adverse both dirs, max 2.30% within the 4.5% floor — layout reading |
| p99 | -3.16% | 2.09% | -1.90% | -6.89% | withheld (disagreeing directions) |

0220 watch-list resolution: the 0220 above-floor p95 reading (max 4.90%
in its rerun vs the 4.5% floor) did NOT persist under the 0221 binary
layout — the phase reads within floor here, consistent with the
layout-displacement attribution. Flag cleared.

## Verdict

**Banked.** Claim scope: `odt_file_source_open` p50/mean/p95
43.28%-46.30% lower and `odt_file_source_open_full_text_lifecycle`
p50/mean/p95 20.55%-23.22% lower, both directions, clean drifts. Every
byte-identical guardrail is clean or within-floor; the one above-floor
primary reading (`odt_semantic_open` p99) did not reproduce at magnitude
in its rerun. Per-sample read evidence and semantic hashes are
bit-identical across control and candidate legs for all seven selectors.
The full litchi-odt suite (893 tests), fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` all pass.
The banked binary
`93c2279b9b5dff79bbfd58e028c5eedede38c3a917b9f3fc8edb86a4fb0641c7` is the
new control for subsequent changes. Raw artifacts:
`docs/performance/results/*-0221-*` and `*-0221r-*` (rerun).
