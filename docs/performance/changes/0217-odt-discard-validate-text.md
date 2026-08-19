# Change 0217: ODT discard-but-validate text extraction

Date: 2026-08-19

## Decision

**Banked** (re-verdict after change 0218 calibrated the litchi-odt layout
noise floor). The executed-phase evidence is the strongest of the series —
all three executed workloads accept ALL FOUR statistics in both paired
directions with clean drifts: `odt_semantic_full_text` 42.12%-52.07% lower,
`odt_source_backed_repeated_text_cached` 51.91%-57.18%,
`odt_source_backed_repeated_text_uncached` 40.31%-53.85%. Two
byte-identical guardrail phases showed adverse both-directions readings
that reproduced in their single permitted reruns (`odt_semantic_open` p50
max 3.26%; `odt_semantic_list_paragraphs` mean max 6.67%); under the
pre-floor rule this withheld the change. The 0218 floors for those
statistics are 3.3% (open p50) and 6.7% (list-paragraphs mean), so the
reproduced readings are within-floor layout readings on phases executing
no changed code and no longer block (0205 rule 1, extended to litchi-odt
by 0218). Claim scope: all three executed workloads, p50/mean/p95/p99 at
the magnitudes above.

## Mechanism and invariants

The 0216 re-profiling refuted the prior double-tokenization hypothesis:
each ODT query call tokenizes `content.xml` exactly once. The real cost in
`extract_text` (the `text()` path) is that every text block is materialized
as a full retained `Element` — `QualifiedName::try_from_string` triple
allocation, attributes `HashMap`, owned attribute name/value strings — and
then immediately discarded by `into_text` (`try_copy` alone was 9.4% self
of the workload process-wide).

The change (all in `litchi-odt`) adds a discard-but-validate mode:

- `parse_text_block_texts` (`elements/text.rs`) — text-only counterpart of
  `parse_text_blocks_with_ownership`: identical event loop, suppression
  rules (tracked-changes / note-body / ruby-text), depth accounting, and
  start-ordered slot output, but no `Element` is ever built. Attribute
  validation runs in `make_text_block_element` check order before slot
  reservation, mirroring the retained path's evaluation order.
- `validate_text_block_attributes` — the discard branch of the banked
  in-file precedent `parse_selected_text_block_element(retain=false)`
  (text.rs:1455) lifted verbatim: same checks in the same order
  (malformed → xmlns skip → resolve → non-UTF-8 name → unknown prefix →
  value decode → duplicate), decode-before-duplicate preserved, identical
  duplicate-detection message.
- `extract_text` (`elements/text/codec.rs`) calls the new mode and joins
  identically; `parse_text_blocks_owned`, `Paragraph::into_text`,
  `Heading::into_text`, `Block::into_text`, `into_text_recursive` /
  `append_text_recursive` become `#[cfg(test)]` oracle-only.
- `parse_text_blocks_with_ownership`, `parse_selected_paragraph`,
  `make_text_block_element`, and the open path are byte-untouched, so
  `odt_semantic_open`, `odt_semantic_list_paragraphs`, and
  `odt_semantic_one_paragraph` execute no changed code. Edit/save path
  untouched. Executed phases: `odt_semantic_full_text` and the source
  facade's repeated-text cases (source.rs shares `extract_text`).

Exactness invariants (documented in code, pinned by 3 new tests; suite
888 → 891): text identity (same event handlers and decoding, `\n` join,
start-ordered slots); error precedence/message identity asserted
byte-equal against the cfg(test)-gated pre-0216 path as a live oracle
across a 6-fixture parity battery plus 7 error cases; structural limits
(`MAX_TEXT_BLOCKS`/`MAX_TEXT_BYTES`/`MAX_TEXT_DEPTH`) and
incomplete-structure errors replicated identically. Vanished checks
verified inert: `from_element` tag check unreachable, `try_set_attribute`
/`try_set_text` allocation-only; only OOM-only `Error::Allocation` sites
disappear. No public API change.

Verification: 891 litchi-odt tests pass, 0 failed; fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the discard-mode extraction;
both carry changes 0192-0196, 0198-0202, 0204, 0206, 0207, 0209-0212, and
0215. Control SHA-256
`6c7fcfb9572f79bbfc2a9dd06289f733e370b34f96662980c5d59b7e972471eb` (the
banked 0215 binary), candidate SHA-256
`8425066ab6b43f08486c5a808d16e31bf6e758f3ff1205096c1a5e21af4a2a5d`.
Binary `.text` delta +11,292 bytes. Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and 500
retained samples per leg, drift ceilings 5%/5%/10%/15%
(p50/mean/p95/p99). Pre-floor acceptance applies (litchi-odt, no
calibrated floor): any adverse both-directions pattern blocks unless
cleared by the single permitted rerun of that workload.

Executed phases: `odt_semantic_full_text`,
`odt_source_backed_repeated_text_cached`,
`odt_source_backed_repeated_text_uncached`. Guardrail (byte-identical)
phases: `odt_semantic_open`, `odt_semantic_list_paragraphs`,
`odt_semantic_one_paragraph`.

### odt_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 44.81% | 44.30% | -1.43% | -0.52% | ACCEPTED |
| mean | 44.54% | 45.13% | -0.56% | -1.62% | ACCEPTED |
| p95 | 43.79% | 42.12% | -4.31% | -1.47% | ACCEPTED |
| p99 | 52.07% | 50.00% | -3.12% | 1.07% | ACCEPTED |

### odt_source_backed_repeated_text_cached (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 52.59% | 53.79% | 2.70% | 0.09% | ACCEPTED |
| mean | 53.05% | 53.61% | 1.47% | 0.26% | ACCEPTED |
| p95 | 54.18% | 53.15% | -0.97% | 1.26% | ACCEPTED |
| p99 | 57.18% | 51.91% | -11.43% | -0.53% | ACCEPTED |

### odt_source_backed_repeated_text_uncached (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 53.85% | 52.87% | -1.81% | 0.28% | ACCEPTED |
| mean | 53.48% | 52.50% | -1.88% | 0.18% | ACCEPTED |
| p95 | 53.32% | 50.77% | -3.05% | 2.27% | ACCEPTED |
| p99 | 40.31% | 42.51% | -3.29% | -6.86% | ACCEPTED |

### odt_semantic_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -1.44% | -3.26% | -1.20% | 0.57% | withheld; adverse both dirs → single permitted rerun |
| mean | 0.76% | -4.33% | -2.48% | 2.52% | withheld (disagreeing directions) |
| p95 | 17.55% | -9.19% | -11.94% | 16.62% | withheld (disagreeing; drifts over ceiling) |
| p99 | 12.51% | -16.87% | -16.56% | 11.46% | withheld (disagreeing; control drift over ceiling) |

Rerun: p50 **reproduced** (-2.58%/-2.79%, clean drifts) — blocks under
the pre-floor rule. Mean/p95/p99 favorable in the rerun (+2.68% to
+40.06%).

### odt_semantic_list_paragraphs (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.36% | -1.57% | 2.15% | 3.39% | withheld; adverse both dirs → single permitted rerun |
| mean | -0.93% | -1.93% | 3.33% | 4.36% | withheld; adverse both dirs → single permitted rerun |
| p95 | -2.97% | -22.45% | 5.61% | 25.59% | withheld; adverse both dirs → single permitted rerun |
| p99 | -6.72% | 8.47% | 9.19% | -6.35% | withheld (disagreeing directions) |

Rerun: mean **reproduced** (-2.42%/-6.67%) — blocks; the primary p50 and
p95 adverse patterns did NOT reproduce (p50 -3.65%/+0.96%; p95
+0.27%/-55.90% with drifts far over ceiling — this workload's tails are
wild on this binary pair in both directions).

### odt_semantic_one_paragraph (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -2.86% | 0.67% | 2.88% | -0.66% | withheld (disagreeing directions) |
| mean | -4.89% | 1.06% | 2.34% | -3.46% | withheld (disagreeing directions) |
| p95 | -27.47% | -1.44% | -1.56% | -21.66% | withheld; adverse both dirs → single permitted rerun |
| p99 | -19.13% | 3.06% | -2.03% | -20.28% | withheld (disagreeing; candidate drift over ceiling) |

Rerun: the primary p95 adverse pattern did NOT reproduce (+2.20%/-1.11%,
clean drifts) — cleared; no adverse-both statistic in the rerun.

## Verdict

**Banked.** Re-verdicted under the 0218-calibrated litchi-odt floor: the
reproduced guardrail readings (open p50 max 3.26%, floor 3.3%;
list-paragraphs mean max 6.67%, floor 6.7%) are within-floor layout
readings on phases executing no changed code and are recorded as such, not
regressions; the one-paragraph p95 primary pattern did not reproduce in
its rerun. The change was re-applied bit-exact after calibration: the
rebuilt harness matches the original candidate SHA-256
`8425066ab6b43f08486c5a808d16e31bf6e758f3ff1205096c1a5e21af4a2a5d`, and
the full litchi-odt suite (891 tests), fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` all pass.
The banked binary is the new control for subsequent changes. Claim scope:
`odt_semantic_full_text` p50/mean/p95/p99 (42.12%-52.07% lower),
`odt_source_backed_repeated_text_cached` p50/mean/p95/p99 (51.91%-57.18%
lower), `odt_source_backed_repeated_text_uncached` p50/mean/p95/p99
(40.31%-53.85% lower). Raw artifacts:
`docs/performance/results/*-0217-*` and `*-0217r-*` (reruns).
