# Change 0224: ODT open-parse hand-rolled namespace binding tracker

Date: 2026-08-19

## Decision

**Banked** (v2, after a diagnosed-and-fixed fixed-overhead regression in
v1 — see "The v1 tiny-shape regression" below). Executed-phase claims:
`odt_file_eager_open` p50/mean/p99 9.41%-13.07% / 9.86%-14.48% /
14.01%-25.39% lower (over the 0223 eager floors 5.6/5.7/9.2),
`odt_file_source_open` p50 21.85%-23.05% lower (p50 uncalibrated on this
phase — pre-floor claim), `odt_file_source_open_full_text_lifecycle` mean
9.39%-11.77% lower (over the 0223 mean floor 2.5). `odt_semantic_open`
yields no claim this run (p50/mean rejected on control drift; p95/p99
accepted but within the 0218 floors) — the per-shape diagnosis shows the
win concentrates in medium/large documents (−11.6%/−24.1% p50) while the
harness's reported `results[0]` is the tiny 24-paragraph shape. All
guardrails clean or within-floor; no rerun was needed for v2.

## Mechanism and invariants

The post-0222 reprofile showed `process_event` — `NsReader`'s per-event
namespace binding maintenance — at 46.8% of timed `odt_semantic_open` and
54.8% of timed `odt_file_source_open`, the dominant residual after
0220-0222. The change (all in `crates/litchi-odt/src/document/open_parse.rs`)
replaces `NsReader` with a plain `Reader` plus a hand-rolled
`BindingTracker`:

- **Binding maintenance is NOT depth-gated**: push/error stream runs for
  every Start/Empty at any depth, byte-exactly as `NsReader` does.
  Reserved-prefix errors are constructed as real
  `quick_xml::name::NamespaceError` values, whose Display the
  `invalid ODT content.xml: {error}` wrapper forwards — messages are
  byte-identical BY CONSTRUCTION. Replicated: the `with_checks(false)`
  attribute scan with silent break on first malformed attribute (prior
  bindings kept), `TooManyDeclarations(256)`, the four reserved-prefix
  errors in attribute order, `xmlns:p=""` unbinding asymmetry (default →
  `Unbound`, prefixed → `Unknown(prefix.to_vec())`), deferred pop (End
  events resolve in their own scope), pre-bound xml/xmlns entries, and
  namespace-error preemption of the validation error for the same event.
- **`xmlns` prefilter**: `memmem(attributes_raw(), b"xmlns")` — absent
  ⇒ no key can start with `xmlns` ⇒ push is provably side-effect-free
  (`level += 1` only). v2 adds a length gate: shorter than 5 bytes ⇒ skip
  even the memmem call (`#[inline]` fast path; scan machinery out-of-line
  in `push_scanned`).
- **Flat-buffer layout** (v2): POD `Binding {start, prefix_len,
  value_len, level}` indexing one shared `Vec<u8>`, mirroring quick-xml
  0.41 `NamespaceResolver`'s allocation pattern — tracker init dropped
  from 5 small allocations to one buffer + one stack.
- The depth ≤ 2 resolution gating from 0221 is unchanged; the tokenizer
  (`Reader` with `check_end_names`/`check_comments`) is unchanged, so the
  tokenization error stream is unchanged.
- One documented divergence: quick-xml counts depth in u16 (wraps in
  release); the tracker uses u32 — unobservable at the consumed depths
  (≤ 3) and strictly removes a panic path.

Exactness proof: the pre-0221 buffered `NsReader` loop serves as the
differential oracle. New batteries: 24 adversarial namespace edge cases
(malformed-attribute breaks, 256/257 declarations root and deep, all four
reserved-prefix errors and both orderings, unbinding at consumed and deep
levels, duplicate last-wins, raw-entity URIs, quote/whitespace variants)
each with pinned accept/reject + byte-identical error strings; a
per-event resolution differential (tracker vs `read_resolved_event_into`)
at EVERY depth across 9 synthetic cases and the full ODT/FODT corpus. The
0221/0222 oracles pass unchanged. Suite 895 → 898.

Executed phases: `odt_semantic_open`, `odt_file_eager_open`,
`odt_file_source_open`, `odt_file_source_open_full_text_lifecycle` (all
funnel through `OpenParse::run`). Guardrails: `odt_semantic_full_text`,
`odp_semantic_open`, `ods_file_source_open`. No public API change.

## The v1 tiny-shape regression (diagnosed and fixed)

The v1 candidate (identical semantics, per-entry `Vec` bindings and an
ungated memmem call) measured a reproduced adverse both-directions p50 on
the EXECUTED `odt_semantic_open` phase (primary max 2.42%, rerun max
2.63%, clean drifts) — a withhold under rule 2. Differential diagnosis
(perf stat instruction counters, cross-shape fitting, VMA-corrected
histograms; `/tmp/0224-prof/sem-diagnosis.md`) found a REAL mechanism:
v1 added a fixed ~+9,000 instructions / ~1.5 µs per open (tracker
init/drop allocations + prefilter call cost) against a ~17 ns/paragraph
saving — crossover ≈ 65-75 paragraphs. The harness's `odt_semantic_open`
reports THREE corpus shapes and the series' summaries read `results[0]`,
the tiny 24-paragraph shape — the only shape below the crossover (medium
−6.3%, large −21.3% at v1). The v2 revision (length-gated push fast path
+ flat-buffer layout) eliminated the fixed overhead: tiny flipped to
−0.9%, medium deepened to −11.6%, large to −24.1% (sanity A/B; v2 then
ran the full protocol below).

## Matched release timing

Two frozen release binaries differ only in the `BindingTracker` change;
both carry the banked tranche through 0222. Control SHA-256
`f53a43f12c405238561bac89fddca861ce14a35bb90136b31a2bdfe528e04987` (the
banked 0222 binary), candidate (v2) SHA-256
`48bd4072fdc10f6be60c01fd3cc908c79f3cde07fa7768187ea3166ebd329ca2`.
Binary `.text` delta +848 bytes. Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and 500
retained samples per leg, drift ceilings 5%/5%/10%/15%. Floors: 0218
(semantic phases), 0223 (source/eager phases), 0205/0213 (ODS/ODP). The
v1 legs and rerun are superseded by this v2 run and are not archived.

### odt_semantic_open (executed; results[0] = tiny shape)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.56% | 13.45% | 16.83% | 0.56% | rejected (disagreeing; control drift over 5% ceiling) |
| mean | 2.69% | 14.57% | 14.87% | 0.85% | rejected (control drift over 5% ceiling) |
| p95 | 9.18% | 15.19% | 8.18% | 1.02% | ACCEPTED; within floor 27.6% — neutral |
| p99 | 18.50% | 19.41% | 10.37% | 9.14% | ACCEPTED; within floor 28.2% — neutral |

No adverse-both (v1's regression is gone). The tiny shape hides the win;
the diagnosed medium/large p50 wins (−11.6%/−24.1%) are recorded as
analysis evidence, not protocol claims.

### odt_file_eager_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 13.07% | 9.41% | -2.14% | 1.98% | ACCEPTED; over floor 5.6% — **claimed** |
| mean | 14.48% | 9.86% | -3.25% | 1.99% | ACCEPTED; over floor 5.7% — **claimed** |
| p95 | 22.57% | 11.19% | -11.18% | 1.87% | rejected (control drift over 10% ceiling) |
| p99 | 25.39% | 14.01% | -10.76% | 2.86% | ACCEPTED; over floor 9.2% — **claimed** |

### odt_file_source_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 23.05% | 21.85% | 1.99% | 3.59% | ACCEPTED — **claimed** (pre-floor) |
| mean | 23.89% | 20.64% | 0.75% | 5.05% | rejected (candidate drift 5.05% marginally over 5% ceiling) |
| p95 | 28.09% | 16.60% | -4.06% | 11.28% | rejected (candidate drift over 10% ceiling) |
| p99 | 31.79% | 4.80% | -14.70% | 19.05% | rejected (candidate drift over 15% ceiling) |

### odt_file_source_open_full_text_lifecycle (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 11.37% | 5.76% | -7.35% | -1.49% | rejected (control drift over 5% ceiling) |
| mean | 11.77% | 9.39% | -3.63% | -1.03% | ACCEPTED; over floor 2.5% — **claimed** |
| p95 | 16.47% | 35.83% | 25.34% | -3.71% | rejected (control drift over 10% ceiling) |
| p99 | 21.76% | 9.96% | 18.79% | 36.71% | rejected (drifts over 15% ceiling) |

### Guardrails (byte-identical)

- `odt_semantic_full_text`: p95 accepted favorable (+14.94%/+19.62%,
  layout-favorable, not claimed); p50/mean/p99 disagreeing — clean.
- `odp_semantic_open`: p50 adverse both dirs, max 0.97% within the 0213
  ODP open p50 floor of 3.1% — layout reading; others disagreeing — clean.
- `ods_file_source_open`: p50/mean accepted favorable (sub-floor, not
  claimed); p95/p99 disagreeing — clean; the 0220 watch-listed p95
  pattern remains absent.

## Verdict

**Banked.** Claim scope: `odt_file_eager_open` p50/mean/p99
(9.41%-13.07% / 9.86%-14.48% / 14.01%-25.39% lower),
`odt_file_source_open` p50 (21.85%-23.05% lower),
`odt_file_source_open_full_text_lifecycle` mean (9.39%-11.77% lower).
`odt_semantic_open` records no claim (drift/floor) but no adverse either;
its medium/large-shape wins are documented analysis evidence. The full
litchi-odt suite (898 tests) plus litchi-odf-common (372), fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass. The banked v2 binary
`48bd4072fdc10f6be60c01fd3cc908c79f3cde07fa7768187ea3166ebd329ca2` is the
new control for subsequent changes. Raw artifacts:
`docs/performance/results/*-0224-*`.
