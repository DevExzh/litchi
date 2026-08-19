# Change 0222: ODT owned open path promoted to the fused parse

Date: 2026-08-19

## Decision

**Banked** (re-verdict after change 0223 calibrated the litchi-odt
source-path/eager layout noise floors). Both executed workloads accept
their claimable statistics in both paired directions with clean drifts:
`odt_semantic_open` p50/mean/p95 are 6.33%-6.45% / 10.02%-11.55% /
31.02%-33.20% lower (all over the 0218 floors 3.3/7.2/27.6), and
`odt_file_eager_open` accepts ALL FOUR statistics — p50 12.05%-17.05%,
mean 12.29%-16.82%, p95 11.46%-18.86%, p99 13.40%-17.27% — all over the
0223 eager floors (5.6/5.7/9.3/9.2). One byte-identical guardrail reading
(`odt_file_source_open_full_text_lifecycle` p50, max 1.76%) reproduced in
its single permitted rerun (max 1.24%) and blocked under the pre-floor
rule; the 0223 calibration shows that magnitude is pure layout noise on
this phase (floor 3.8%; probe a reproduces adverse-both p50 at 1.70%
with zero changed code), so under rule 1 it no longer blocks. Claim
scope: `odt_semantic_open` p50/mean/p95 and `odt_file_eager_open`
p50/mean/p95/p99 at the magnitudes above.

## Mechanism and invariants

0219's candidate B, re-quantified post-0220/0221 at ~27% of timed
`odt_semantic_open` and ~12.5% of timed `odt_file_eager_open`. The owned
open path (`from_owned_package`,
`crates/litchi-odt/src/document/package.rs:154`) previously scanned
`content.xml` twice: the standalone `validate_content_document_part`
validation pass and the `StyleElements::parse_styles` content-styles
rescan. The change replaces both with one fused `OpenParse::run` pass
(the 0221 borrowing, depth-gated loop — no buffered reads resurrected)
plus `content_styles.finish()`.

Error precedence is preserved stage by stage and documented in the code
comment at package.rs:165-172:

1. Package open / mimetype / content.xml fetch / UTF-8 check — unchanged
   code, unchanged order.
2. Content validation (mid-stream and end-of-stream) errors return from
   `OpenParse::run` before `Content::from_vec` and before
   styles.xml/meta.xml are fetched — exactly where
   `validate_content_document_part` early-returned, with byte-identical
   messages (the 0220/0221 ValidateHandler replica).
3. styles.xml fetch/decode, meta.xml fetch/decode, styles.xml parse —
   unchanged statements in unchanged order.
4. Content-styles errors are recorded during the fused scan (first error
   wins, later events don't mutate state — matching the historical
   standalone scan's early return) and surface via `finish()` after the
   styles.xml parse and before `try_extend` — the historical position.
5. `try_extend` (allocation-only; duplicate names overwrite) still runs
   last.

The pre-change sequential owned path survives as the cfg(test) oracle
`from_owned_package_sequential_oracle`. Two new parity tests pin the
equivalence: 11 in-memory synthetic packages + 2 container cases with
distinguishable per-stage error kinds (validation vs styles.xml
tokenization vs content-styles attribute vs meta UTF-8 vs mimetype),
including cross-stage precedence (styles error beats content-styles
error; validation beats styles error; early-recorded style error loses to
a late validation error) and duplicate-style-name `try_extend` overwrite,
plus full-parity runs over every `.odt` fixture in `test-data/`
(accept/reject, byte-identical error messages, identical sorted style
projections and full document text). The pre-existing fused-vs-sequential
open_parse oracles and the 0221 gated-vs-buffered oracles pass unchanged.
Suite 893 → 895.

Executed phases: `odt_semantic_open`, `odt_file_eager_open` (both time
the owned `Document::from_bytes`/`open`). Byte-identical guardrails:
`odt_file_source_open`, `odt_file_source_open_full_text_lifecycle`
(source-backed path was already fused pre-0222), `odt_semantic_full_text`
(query path, open untimed), `odp_semantic_open`, `ods_file_source_open`.
Edit/save paths untouched. No public API change.

Verification: 895 litchi-odt tests pass, 0 failed; fmt, clippy
(`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the owned-path promotion; both
carry the banked tranche through 0221. Control SHA-256
`93c2279b9b5dff79bbfd58e028c5eedede38c3a917b9f3fc8edb86a4fb0641c7` (the
banked 0221 binary), candidate SHA-256
`f53a43f12c405238561bac89fddca861ce14a35bb90136b31a2bdfe528e04987`.
Binary `.text` delta −10,016 bytes (one full scan path eliminated). Fresh
CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate, A2
control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). Floors: 0218 for the semantic phases,
0223 for the source/eager phases, 0205/0213 for ODS/ODP.

### odt_semantic_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 6.33% | 6.45% | 1.61% | 1.47% | ACCEPTED; over floor 3.3% — **claimed** |
| mean | 11.55% | 10.02% | -0.60% | 1.11% | ACCEPTED; over floor 7.2% — **claimed** |
| p95 | 33.20% | 31.02% | -8.27% | -5.28% | ACCEPTED; over floor 27.6% — **claimed** |
| p99 | 32.59% | 18.56% | -17.86% | -0.77% | rejected (control drift over 15% ceiling) |

### odt_file_eager_open (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 17.05% | 12.05% | -4.57% | 1.18% | ACCEPTED; over floor 5.6% — **claimed** |
| mean | 16.82% | 12.29% | -4.47% | 0.73% | ACCEPTED; over floor 5.7% — **claimed** |
| p95 | 18.86% | 11.46% | -7.56% | 0.87% | ACCEPTED; over floor 9.3% — **claimed** |
| p99 | 17.27% | 13.40% | -7.59% | -3.27% | ACCEPTED; over floor 9.2% — **claimed** |

### odt_file_source_open (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 0.38% | -1.05% | -2.00% | -0.59% | withheld (disagreeing directions) |
| mean | 0.19% | -1.03% | -1.62% | -0.42% | withheld (disagreeing directions) |
| p95 | -0.78% | -0.60% | -0.32% | -0.51% | adverse both dirs, max 0.78% → rerun |
| p99 | -0.42% | -27.98% | 1.56% | 29.42% | adverse both dirs, max 27.98% → rerun |

Rerun (`*-0222r-*`): p95 flipped favorable (+7.22%/+8.35%) and the p99
adverse pattern is absent (disagreeing) — both cleared even pre-floor.
Under the 0223 source-open floors (p95 2.5%, p99 28.0%) the primary
readings are within-floor layout readings regardless.

### odt_file_source_open_full_text_lifecycle (byte-identical guardrail)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.11% | -1.76% | -0.99% | 0.65% | adverse both dirs, max 1.76% → rerun |
| mean | 0.50% | -1.45% | -1.41% | 0.52% | withheld (disagreeing directions) |
| p95 | 2.13% | -0.41% | -2.30% | 0.24% | withheld (disagreeing directions) |
| p99 | 24.17% | 1.06% | -22.83% | 0.68% | withheld (control drift over ceiling) |

Rerun (`*-0222r-*`): p50 adverse-both REPRODUCED (-0.39%/-1.24%, clean
drifts) — blocked under the pre-floor rule; 0223 then measured this
phase's pure layout noise at up to 3.79% adverse-both p50 (probe b) with
zero changed code, floor 3.8% — the reproduced reading is a within-floor
layout reading under rule 1 and no longer blocks.

### odt_semantic_full_text (byte-identical guardrail)

p50 adverse both dirs, max 1.78% within the 0218 full-text p50 floor of
4.1% — layout reading; mean/p95/p99 disagreeing. Clean.

### odp_semantic_open (byte-identical guardrail)

mean accepted favorable (+5.39%/+2.94%, sub-floor — not claimed);
p50/p95/p99 disagreeing or drift-failed. Clean.

### ods_file_source_open (byte-identical guardrail)

p50/mean accepted favorable (+0.95%/+2.22%, +0.62%/+3.24%, sub-floor —
not claimed); p95/p99 disagreeing. Clean; the 0220 watch-listed p95
pattern remains absent.

## Verdict

**Banked** (re-verdicted under the 0223 floors after the provisional
pre-floor withhold). The change was re-applied bit-exact after
calibration: the rebuilt harness matches the measured candidate SHA-256
`f53a43f12c405238561bac89fddca861ce14a35bb90136b31a2bdfe528e04987`, and
the full litchi-odt suite (895 tests), fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` all pass.
Claim scope: `odt_semantic_open` p50/mean/p95 (6.33%-6.45% /
10.02%-11.55% / 31.02%-33.20% lower) and `odt_file_eager_open`
p50/mean/p95/p99 (12.05%-17.05% / 12.29%-16.82% / 11.46%-18.86% /
13.40%-17.27% lower). The banked binary is the new control for
subsequent changes. Raw artifacts:
`docs/performance/results/*-0222-*` and `*-0222r-*` (reruns).
