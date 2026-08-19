# Change 0211: ODP per-query parse fuses transition-style collection into the slide scan

Date: 2026-08-19

## Decision

**Banked.** The fused per-query parse eliminates one of the two complete
`content.xml` tokenizations behind `slides()`/`text()`. Frozen
cross-binary CPU-2 A/B/B/A measurement accepts `odp_semantic_list_slides`
p50/mean (15.56%-17.84%), `odp_semantic_one_slide` p50/mean
(18.88%-19.50%), and `odp_semantic_full_text` p50 (15.85%-16.43%), all
lower in both paired directions with clean drifts — consistent with the
profile attribution (the removed scan was 14.95% inclusive). The only
adverse both-directions pattern (odp_semantic_open mean/p95, on a
byte-identical phase) did not reproduce in the single permitted rerun,
clearing the pre-floor block.

## Mechanism and invariants

Post-0210 profiling of `odp_semantic_full_text` attributed 90.65%
inclusive to `parse_pages_with_styles`, of which
`resolved_transition_styles` → `parse_transition_style_definitions(content)`
was a complete first `content.xml` tokenization at 14.95% inclusive,
immediately followed by the main slide-parse `NsReader` scan of the same
bytes.

The change fuses the two scans into one tokenization
(`codec/xml/codec.rs`): per event the driver feeds a
`TransitionStyleCollector` — an event-fed one-to-one transcription of
`parse_transition_style_definitions`' match arms that records its first
error instead of returning — and then the slide-scan closure. Error
precedence preserves the historical two-pass order exactly:

- tokenization errors surface as the transition scan's
  `"XML parsing error: {error}"` (that scan historically tokenized the
  same bytes first and read to EOF);
- a recorded collector error is returned before any slide-scan error;
- slide-scan read errors return immediately; slide-scan semantic errors
  are recorded and the loop keeps draining (still feeding the collector
  to EOF) so late collector errors keep their historical priority;
- post-pass: merge over styles.xml bases → `resolve_transition_styles`
  → deferred slide-scan error → per-slide transition assignment, in the
  historical order.

All six nested parsers feed every consumed event to the collector, so
collection sees the full stream inside subtrees. A `NsClass` Copy enum
frees the reader borrow without changing namespace classification (the
`Other` guards keep it exact for unknown URIs).
`parse_enhanced_geometry`'s read errors now use the common
`XML parsing error` mapping — the old message was historically
unreachable (the pre-scan tokenized first); pinned by an oracle test.
The standalone `parse_transition_style_definitions` shell is
byte-identical as the oracle anchor; `resolved_transition_styles` split
into a `#[cfg(test)]` wrapper plus merge/resolve helpers.

Verification: a new equivalence oracle (`codec/xml/oracle.rs`) runs the
verbatim-transcribed sequential two-pass reference against the fused
parse on 19 fixtures (11 `.odp`, 8 `.fodp`) across all four entry points
(`parse_slides_with_styles`, `parse_slide_with_styles_at` for
`0..=n+1`, `parse_drawing_pages`), plus 14 synthetic precedence pins
(late transition definitions, collector-vs-slide error priority,
tokenization-after-early-error, truncation in the transition region and
inside enhanced-geometry, duplicate attributes, custom prefixes, cyclic
parent-style in styles.xml, content default-style override, SELECT_ONE
across indices, sheet-mode `table:shapes`, empty content, undefined
references). The full litchi-odp suite passes (148 lib + 111
integration, +14 oracle tests), doc tests 21 pass; fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), `tools/check_crate_boundaries.py`,
and a facade `cargo check` pass. No public API change.

## Matched release timing

Two frozen release binaries differ only in the fused query parse; both
carry changes 0192-0196, 0198-0202, 0204, 0206, 0207, 0209, and 0210.
Control SHA-256 `c4b2b56897b17d41295c7286c07f1aff4745b34e0ade4cb8c1ef55d149d29371`
(the banked 0210 binary), candidate SHA-256
`ceba155be185f1c213c4bf90200bb5e87bb697a5023e3a703d9cd7def6042922`.
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
| p50 | 0.13% | 0.18% | 1.38% | 1.33% | ACCEPTED (mechanism-absent) |
| mean | -0.52% | -1.08% | -0.71% | -0.16% | withheld; adverse both dirs → single permitted rerun |
| p95 | -4.92% | -6.53% | -9.76% | -8.38% | withheld; adverse both dirs → single permitted rerun |
| p99 | 2.04% | -6.30% | -16.34% | -9.22% | withheld (disagreeing directions, drift over ceiling) |

### odp_semantic_list_slides (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 16.43% | 17.73% | 1.13% | -0.44% | ACCEPTED |
| mean | 17.84% | 15.56% | -1.35% | 1.39% | ACCEPTED |
| p95 | 25.63% | 4.83% | -12.83% | 11.55% | withheld (drift over ceiling) |
| p99 | 32.29% | 2.12% | -23.30% | 10.89% | withheld (drift over ceiling) |

### odp_semantic_one_slide (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 18.88% | 18.88% | 0.34% | 0.34% | ACCEPTED |
| mean | 19.29% | 19.50% | -2.96% | -3.21% | ACCEPTED |
| p95 | 20.31% | 28.36% | -11.64% | -20.57% | withheld (drift over ceiling) |
| p99 | 12.69% | 20.01% | -12.72% | -20.04% | withheld (drift over ceiling) |

### odp_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 16.43% | 15.85% | -1.74% | -1.06% | ACCEPTED |
| mean | 18.90% | 16.79% | -5.48% | -3.01% | withheld (control drift over ceiling) |
| p95 | 25.19% | 20.51% | -17.92% | -12.78% | withheld (drift over ceiling) |
| p99 | 19.47% | 11.48% | -16.62% | -8.35% | withheld (drift over ceiling) |

### odp_semantic_open — rule-2 rerun (clears the primary adverse reading)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.36% | 0.45% | -0.63% | 0.28% | ACCEPTED — primary adverse pattern NOT reproduced |
| mean | 4.80% | 4.79% | 0.90% | 0.91% | ACCEPTED — primary adverse pattern NOT reproduced |
| p95 | 14.58% | 18.87% | 10.89% | 5.31% | withheld (control drift over ceiling) |
| p99 | 16.48% | 7.69% | 1.41% | 12.08% | ACCEPTED |

The primary-run mean/p95 adverse both-directions reading on this
byte-identical phase did not reproduce in the single permitted rerun;
the block is cleared.

## Verdict

**Banked.** Claim scope, frozen cross-binary CPU-2 A/B/B/A (30 warmups,
500 samples), pre-floor acceptance:

- `odp_semantic_list_slides` p50/mean: **15.56%-17.84% lower** (p95/p99
  withheld on tail drift).
- `odp_semantic_one_slide` p50/mean: **18.88%-19.50% lower**.
- `odp_semantic_full_text` p50: **15.85%-16.43% lower** (mean/p95/p99
  withheld on control drift).
- `odp_semantic_open` p50/mean/p99 accepted at 0.45%-16.48% across the
  primary run and rerun (byte-identical phase, mechanism-absent,
  recorded for completeness).

No allocation/RSS, physical-I/O, cold-cache, producer, or broad-ODF claim
is made. Harness rebuild verified bit-exact to the measured candidate
(`ceba155b…`); this binary is the control for the next change. Raw
artifacts: `docs/performance/results/*-0211-*` and `*-0211r-*` (rerun).
