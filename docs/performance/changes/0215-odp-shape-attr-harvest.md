# Change 0215: ODP single-scan shape-attribute harvest

Date: 2026-08-19

## Decision

**Banked.** All three executed workloads accept p50/mean in both paired
directions with clean drifts, at magnitudes far above the 0213-calibrated
floor; the adverse both-directions p50 reading on `odp_semantic_open`
(executing no changed code) is within the calibrated floor and is recorded
as a layout reading (rule 1). Claim scope: `odp_semantic_list_slides`
p50/mean 6.53%-9.97% lower, `odp_semantic_one_slide` p50/mean
8.80%-15.10% lower, `odp_semantic_full_text` p50/mean 5.44%-11.32% lower.
p95/p99 on the executed workloads are withheld (same-implementation drift
over ceiling / disagreeing directions — the ODP corpus tails remain
drift-heavy).

## Mechanism and invariants

The 0214 re-profiling showed `Parser::drawing_attributes` at 12.7%-13.7%
inclusive on every ODP workload: after the 0210-0212 `ElementAttrs` work,
`shape_builder` still ran a FRESH `element.attributes()` re-scan per shape
element to harvest the non-modeled DRAW/SVG/DR3D/TABLE attributes, on top
of the shared incremental scan already maintained by `ElementAttrs`.

The change (`litchi-odp::codec::parser::codec::xml`): a new
`ElementAttrs::drawing_attributes(&mut self, reader)` replays the cached
attribute prefix and then continues the shared iterator to completion,
harvesting drawing attributes from both halves; `shape_builder` calls it
instead of the standalone fresh scan, which is deleted (single caller;
its pre-0215 body survives verbatim as the `oracle_drawing_attributes`
test oracle). Per-attribute classification/decoding is byte-identical to
the deleted body via a private `harvest_drawing_attribute` helper keyed
on the cached resolution snapshot (`ResolvedAttributeNamespace::matches`
≡ `is_namespace` on the live `ResolveResult`).

Exactness invariants (documented in code, pinned by 7 new tests; suite
150 → 157 lib tests):

- **Document order**: cached-prefix replay plus iterator continuation both
  emit in document order — the union is exactly the fresh scan's order.
- **Error-message identity by first reach**: a malformed attribute reached
  by a lookup's scan already returned `"invalid XML attribute: …"` from
  `get`, so `shape_builder`'s `?` never reaches the harvest; a malformed
  attribute first reached by the harvest continuation maps to
  `"invalid ODP shape attribute: …"` — identical to the fresh scan because
  quick-xml attribute errors carry byte offsets into the element.
- **Duplicates**: detection stays in the single underlying iterator; the
  pair is reported by whichever phase first reaches the second occurrence.
- **Decode positions**: modeled/foreign-namespace attributes are skipped
  undecoded; every harvested attribute is decoded at the same position
  with the same decoder settings; no attribute is double-decoded (lookup
  targets are all modeled or foreign to the harvest namespaces).

Full litchi-odp suite: 157 lib + 111 integration, 0 failed; fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass. No public API change.

## Matched release timing

Two frozen release binaries differ only in the harvest; both carry changes
0192-0196, 0198-0202, 0204, 0206, 0207, 0209, 0210, 0211, and 0212.
Control SHA-256
`246c6b1f916b2dc2fa8529edfdfed605fede29db0226f54334bd38bc5d4f8e13` (the
banked 0212 binary), candidate SHA-256
`6c7fcfb9572f79bbfc2a9dd06289f733e370b34f96662980c5d59b7e972471eb`.
Binary `.text` delta −2,132 bytes (candidate smaller). Fresh CPU-2-pinned
processes ran `A1 control, B1 candidate, B2 candidate, A2 control`, 30
warmups and 500 retained samples per leg, drift ceilings 5%/5%/10%/15%
(p50/mean/p95/p99). 0213-calibrated litchi-odp floor applies.

Floor-scope note: the 0213 probes calibrated positive text deltas of
+5.5KB to +14.5KB; this pair has a −2.1KB delta. The floor is invoked for
the open p50 reading on the reasoning that layout-noise magnitude is
driven by displacement size (|−2.1KB| is below the smallest calibrated
probe) and is sign-agnostic; this judgment is recorded here. The banking
does not depend on it beyond the single open p50 reading.

Executed phases: `odp_semantic_list_slides`, `odp_semantic_one_slide`,
`odp_semantic_full_text`. `odp_semantic_open` executes no changed code.

### odp_semantic_open (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -1.59% | -2.48% | -1.44% | -0.58% | layout reading (adverse both dirs, max 2.48% ≤ 3.1% floor) — does not block |
| mean | +5.91% | -3.08% | -8.26% | +0.51% | withheld (disagreeing; control drift over ceiling) |
| p95 | +27.13% | -4.77% | -26.95% | +5.04% | withheld (disagreeing; control drift over ceiling) |
| p99 | +32.81% | -8.68% | -27.19% | +17.78% | withheld (disagreeing; drifts over ceiling) |

### odp_semantic_list_slides (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 7.29% | 7.88% | 1.95% | 1.31% | ACCEPTED (floor 2.0%) |
| mean | 6.53% | 9.97% | 3.91% | 0.09% | ACCEPTED (floor 3.6%) |
| p95 | -0.43% | 10.91% | 8.04% | -4.16% | withheld (disagreeing directions) |
| p99 | -5.24% | 10.42% | 10.65% | -5.81% | withheld (disagreeing; control drift over ceiling) |

### odp_semantic_one_slide (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 8.80% | 10.29% | 0.81% | -0.83% | ACCEPTED (floor 2.5%) |
| mean | 15.10% | 10.48% | -4.81% | 0.36% | ACCEPTED (floor 3.2%) |
| p95 | 34.29% | 11.66% | -17.55% | 10.84% | withheld (drifts over ceiling) |
| p99 | 32.79% | 4.26% | -18.09% | 16.67% | withheld (drifts over ceiling) |

### odp_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 9.50% | 7.06% | -1.67% | 0.98% | ACCEPTED (floor 0.1%) |
| mean | 11.32% | 5.44% | -3.29% | 3.12% | ACCEPTED (floor 0.5%) |
| p95 | 27.83% | -7.51% | -17.43% | 22.99% | withheld (disagreeing; drifts over ceiling) |
| p99 | 3.45% | -3.27% | 3.08% | 10.26% | withheld (disagreeing directions) |

## Verdict

**Banked.** The change remains in the tree; the candidate binary is the
control for subsequent changes. Claim scope as in the Decision. No
allocation/RSS/physical-I/O/cold-cache claim. Raw artifacts:
`docs/performance/results/*-0215-*`.
