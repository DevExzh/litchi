# Change 0220: ODF content-validator borrowing reads and depth-gated resolution

Date: 2026-08-19

## Decision

**Banked.** Both executed ODT open workloads accept their claimable
statistics in both paired directions with clean drifts:
`odt_semantic_open` p50 is 5.99%-7.30% lower (over the 0218 open p50
floor of 3.3% — claimed); `odt_file_eager_open` p95 is 19.12%-21.14%
lower (no eager-open floor calibrated — pre-floor rule, claimed direct).
Two byte-identical guardrail phases showed adverse both-directions
patterns: `odt_semantic_full_text` p95 in the primary did NOT reproduce
in the single permitted rerun (cleared); `ods_file_source_open` p95
reproduced marginally above the 0205 floor (max 4.90% vs 4.5%) and is
recorded as a flagged above-floor layout reading on a phase executing no
changed code — see "Guardrails and the ODS p95 flag" below. Claim scope:
`odt_semantic_open` p50 (5.99%-7.30% lower) and `odt_file_eager_open`
p95 (19.12%-21.14% lower). All other accepted statistics are within-floor
or layout readings, recorded not claimed.

## Mechanism and invariants

The 0219 profiling attributed `validate_content_document_part`
(`crates/litchi-odf-common/src/core/family.rs:268`) — the full-document
`NsReader` scan that 0216 measured at ~69% of the timed
`odt_semantic_open` call — to two costs: the per-event buffer copy of
`read_resolved_event_into`, and namespace resolution performed for EVERY
event although the result is consumed only at depth ≤ 2 (the root,
`office:body`, and the family element checks). The change:

- Switches the event loop to borrowing `reader.read_event()` — events
  borrow `xml` directly, eliminating the per-event `buffer` copy and
  `buffer.clear()`. Binding push/pop still runs inside `read_event` for
  every event, so prefix-rebinding semantics and the tokenization error
  stream are unchanged.
- Resolves the element namespace via `reader.resolver().resolve_element()`
  only for `Start`/`Empty` events at depth ≤ 2 — the only arms whose
  resolved value is observable. Deeper arms use `local_name()` only, and
  bindings declared at depth ≥ 3 scope just their own subtree, so no
  consumed resolution can change. Unresolved events report
  `office = false`, matching the pre-change `ResolveResult` mismatch
  fall-through.

The pre-change body is retained verbatim as a cfg(test) oracle
(`validate_content_document_part_oracle`) plus 3 new tests cross-checking
the borrowing implementation against it on 41 synthetic edge cases
(namespace rebinding at various depths, mismatched ends, comment/PI/CDATA
interleavings, truncated documents, missing bodies) and the full ODF
fixture corpus — accept/reject parity and, on rejections, byte-identical
error messages (litchi-odf-common suite 245 → 248 tests).

Blast radius: the only production caller of
`validate_content_document_part` is the ODT owned open path
(`litchi-odt/src/document/package.rs:165`). ODS and ODP opens use the
trivial `validate_content_part`, so `ods_file_source_open` and
`odp_semantic_open` execute no changed code, and `odt_semantic_full_text`
(re-open-free text query path) likewise. Edit/save paths untouched. No
public API change.

Verification: 248 litchi-odf-common and 891 litchi-odt tests pass, 0
failed; fmt, clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the validator change; both
carry the banked tranche through 0217. Control SHA-256
`8425066ab6b43f08486c5a808d16e31bf6e758f3ff1205096c1a5e21af4a2a5d` (the
banked 0217 binary), candidate SHA-256
`1971c3adf37536d1fa0963d111b79dade639ea5b624150b985faa6ab02bf8713`.
Binary `.text` delta −5,408 bytes. Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and 500
retained samples per leg, drift ceilings 5%/5%/10%/15%
(p50/mean/p95/p99). The 0218 litchi-odt floors and the 0205/0213
ODS/ODP floors apply per phase.

Executed phases: `odt_semantic_open`, `odt_file_eager_open`. Guardrail
(byte-identical) phases: `odt_file_source_open`, `odt_semantic_full_text`,
`ods_file_source_open`, `odp_semantic_open`.

### odt_semantic_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 7.30% | 5.99% | -1.80% | -0.41% | ACCEPTED; over the 3.3% floor — **claimed** |
| mean | 11.07% | 5.69% | -4.18% | 1.62% | ACCEPTED; min-direction 5.69% within the 7.2% floor — neutral (rule 3) |
| p95 | 34.54% | 2.65% | -16.36% | 24.40% | withheld (candidate drift over 10% ceiling; floor 27.6%) |
| p99 | 32.47% | 3.32% | -24.03% | 8.76% | withheld (control drift over 15% ceiling) |

### odt_file_eager_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 21.17% | 18.51% | 3.01% | 6.48% | rejected (candidate drift over 5% ceiling) |
| mean | 20.55% | 19.29% | 4.87% | 6.53% | rejected (candidate drift over 5% ceiling) |
| p95 | 19.12% | 21.14% | 7.74% | 5.04% | ACCEPTED — **claimed** (no eager-open floor calibrated; pre-floor rule) |
| p99 | 12.55% | 33.63% | 35.88% | 3.12% | rejected (control drift over 15% ceiling) |

### odt_file_source_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.85% | 2.38% | 2.28% | 1.73% | accepted favorable — layout-favorable, not claimed |
| mean | 2.40% | 1.88% | 2.04% | 2.58% | accepted favorable — layout-favorable, not claimed |
| p95 | 5.77% | 1.03% | 1.87% | 7.00% | accepted favorable — layout-favorable, not claimed |
| p99 | 8.90% | -5.26% | -1.63% | 13.66% | withheld (disagreeing directions) |

### odp_semantic_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -1.47% | 0.33% | 0.85% | -0.94% | withheld (disagreeing directions) |
| mean | -3.32% | 5.23% | 6.14% | -2.64% | withheld (disagreeing directions) |
| p95 | -7.31% | 15.45% | 17.57% | -7.36% | withheld (disagreeing directions) |
| p99 | -21.02% | 13.16% | 14.53% | -17.82% | withheld (disagreeing directions) |

Clean guardrail: no adverse both-directions statistic.

### odt_semantic_full_text (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 5.47% | 0.31% | -4.59% | 0.62% | accepted favorable — layout-favorable, not claimed |
| mean | 3.54% | -0.34% | -6.08% | -2.30% | withheld (disagreeing directions) |
| p95 | -13.30% | -26.38% | -18.35% | -8.93% | adverse both dirs, max 26.38% over the 16.1% floor → single permitted rerun |
| p99 | 4.88% | -1.38% | -33.16% | -28.76% | withheld (disagreeing; drifts over ceiling) |

Rerun (`*-0220r-*`): the primary p95 adverse pattern did NOT reproduce —
p95 +15.72%/+21.76% favorable both directions; p50/mean accepted
favorable. Cleared; no adverse-both statistic in the rerun.

### ods_file_source_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -4.04% | -3.25% | 1.35% | 0.59% | adverse both dirs, max 4.04% within the 5.5% floor |
| mean | -3.39% | -3.83% | 1.22% | 1.66% | adverse both dirs, max 3.83% within the 5.5% floor |
| p95 | -2.36% | -7.06% | 1.02% | 5.66% | adverse both dirs, max 7.06% over the 4.5% floor → single permitted rerun |
| p99 | 20.24% | -11.65% | -15.13% | 18.81% | withheld (disagreeing; drifts over ceiling) |

Rerun (`*-0220r-*`): p50/mean adverse-both reproduced within floor (max
2.97%/3.34% vs 5.5%); p95 adverse-both reproduced at -4.90%/-3.67%, max
4.90% — marginally over the 4.5% floor.

## Guardrails and the ODS p95 flag

The reproduced `ods_file_source_open` p95 reading (primary max 7.06%,
rerun max 4.90%, floor 4.5%) is recorded as an **above-floor layout
reading, flagged for monitoring** — not a regression and not claimed
either way. The non-attribution case is checkable:

1. **Zero changed code is reachable from the phase.** The only production
   caller of the changed `validate_content_document_part` is the ODT
   owned open path (`litchi-odt/src/document/package.rs:165`); ODS opens
   run the untouched `validate_content_part`. The candidate binary's
   `.text` differs from control only in the validator and its inlining
   neighborhood (−5,408 bytes).
2. **Deterministic invariants are bit-identical.** Per-sample source read
   calls, read bytes, range overlaps, picture reads, and all semantic
   hashes are identical across all 8 primary+rerun legs (control and
   candidate alike) — the candidate performs byte-for-byte the same I/O
   and produces byte-for-byte the same results on this phase.
3. **The magnitudes sit inside the series' documented layout band.** The
   same phase family has historically shown within-floor both-directions
   readings of either sign change-to-change; the p95 excess over the 4.5%
   floor is 0.40pp in the rerun.

The floor is NOT recalibrated post-hoc: 4.5% stays the operative
source-open p95 floor. Instead, `ods_file_source_open` p95 is added to
the watch list: the next ODF-family change must include it in its
guardrail set, and if it again reads adverse-both above floor under a
different binary layout, 0220 is reverted pending investigation. (The
alternative — withholding a measured executed-phase win on the strength
of a 0.40pp floor excess on a phase with a proven absent mechanism — was
rejected as numerology, but the flag keeps the adverse visible rather
than smoothed over.)

## Verdict

**Banked.** Executed-phase claims: `odt_semantic_open` p50 5.99%-7.30%
lower (0218 floor 3.3%) and `odt_file_eager_open` p95 19.12%-21.14% lower
(pre-floor claim). The `odt_semantic_full_text` p95 primary adverse
pattern was cleared by its rerun; the `ods_file_source_open` p95 reading
is a flagged above-floor layout reading on a zero-changed-code phase with
bit-identical deterministic evidence (floors unchanged; watch-listed for
the next ODF change). The candidate tree was restored bit-exact after the
0192-0218 tranche commit: the rebuilt harness matches the measured
candidate SHA-256
`1971c3adf37536d1fa0963d111b79dade639ea5b624150b985faa6ab02bf8713`, and
the full litchi-odf-common (248) and litchi-odt (891) suites, fmt,
clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass. The banked binary is the new
control for subsequent changes. Raw artifacts:
`docs/performance/results/*-0220-*` and `*-0220r-*` (reruns).
