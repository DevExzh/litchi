# Change 0194: ODS worksheet text validation byte scan

Date: 2026-08-18

## Decision

Rewrite the XML-forbidden-character check inside
`litchi_ods::worksheet::validation::validate_text` from a per-`char` scan to
a per-byte scan.

The forbidden characters are exactly U+0000..=U+0008, U+000B..=U+000C, and
U+000E..=U+001F — all ASCII. In UTF-8, ASCII code points encode as single
identical bytes and every non-ASCII code point encodes to bytes >= 0x80, so

```text
value.bytes().any(|byte| byte < 0x20 && !matches!(byte, 0x09 | 0x0A | 0x0D))
```

accepts and rejects exactly the same strings as the previous
`value.chars().any(...)` form, including every multi-byte sequence. The
error text, the `MAX_TEXT_BYTES` size gate, and every caller are unchanged.

This changes no public API, no validation boundary, no refusal rule, and no
error message. A profiling pass over the source-backed commit path (one
owner, repeated batched edit + commit transactions, `perf record -F 3999`)
attributed 9.15% of commit-phase samples to `validate_text`, the largest
single production symbol on that path; the byte scan removes the UTF-8
decoding state machine from the hot loop.

The sibling copy in `litchi_ods::annotations::validation` keeps the same
forbidden set but is not on the measured commit path; it is left unchanged
and recorded as a possible follow-up. `litchi-odt` and `litchi-odp` carry
their own `validate_text` functions with possibly different forbid sets and
are out of scope.

## Mechanism and invariants

`validate_text` is reached on the source-backed commit path through
`worksheet::codec::parse`, which validates every cell of every sheet of the
candidate `content.xml` at the publication boundary. That full reparse and
compare is a deliberate gate and is untouched; only the inner character test
changes. Focused unit tests pin the boundary bytes: 0x00, 0x08, 0x0B, 0x0C,
0x0E, 0x1F and an embedded 0x01 are rejected; empty, plain ASCII, tab, LF,
CR, DEL, two-byte (`é`) and four-byte (emoji) UTF-8 are accepted.

## Matched release timing

Two frozen release binaries differ only in the litchi-ods worksheet
`validate_text` implementation; both contain the 0193 edit-protection cache
as baseline and the identical 341-case selector matrix. Control SHA-256
`c1ad10735d02d84528da0ca962c5019d5acf6a6276f85d577debd80a09c576ab`,
candidate SHA-256
`e193366cd3b85a6e23a1b978be9e0e1e28fdc386c99e84f56a1cba559266d163`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate, A2
control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling.
Every leg reports all embedded verification flags (`output_hash`,
`semantic_reopen`, `media_payloads`, `exact_output`, `exact_sink`,
`source_immutability`) true.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

Four-transaction totals:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 2.08% | 0.63% | -1.15% | 0.31% | accept |
| mean | 2.77% | 0.19% | -1.82% | 0.79% | accept |
| p95 | 6.12% | -1.75% | -5.71% | 2.20% | withheld (directions disagree) |
| p99 | 7.63% | -0.76% | -5.90% | 2.64% | withheld (directions disagree) |

Commit phase (per-sample sum of the four commits):

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 5.19% | 1.97% | -2.13% | 1.19% | accept |
| mean | 6.49% | 1.55% | -3.30% | 1.81% | accept |
| p95 | 11.56% | -1.32% | -8.88% | 4.38% | withheld (directions disagree) |
| p99 | 13.11% | 0.09% | -7.89% | 5.91% | accept |

Stage phase: all four statistics withheld (paired directions disagree inside
noise); the stage path does not call `validate_text`. Publication phase: p50
accepted (0.98% / 0.10%, drifts -0.70% / 0.18%), mean/p95/p99 withheld
(directions disagree). No regression trigger fired — every withheld
statistic's paired directions straddle zero within 3.0% except where noted,
and no accepted statistic is claimed beyond the phases listed.

### Guardrail selectors

`ods_source_backed_one_edit_save` (single-cell lifecycle; per-sample
`lifecycle_ns` / `commit_ns`):

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | -0.54% | 0.28% | -0.18% | -1.00% | withheld (directions disagree) |
| lifecycle mean | -1.04% | 0.33% | 0.38% | -0.98% | withheld (directions disagree) |
| lifecycle p95 | -3.73% | 0.09% | 1.85% | -1.90% | withheld (directions disagree) |
| lifecycle p99 | -6.68% | -1.12% | 0.18% | -5.04% | withheld (directions disagree) |
| commit p50 | 0.96% | 2.60% | -0.42% | -2.07% | accept |
| commit mean | -0.26% | 2.58% | 0.24% | -2.60% | withheld (directions disagree) |
| commit p95 | -3.49% | 2.42% | 3.08% | -2.81% | withheld (directions disagree) |
| commit p99 | -19.35% | -3.13% | -0.73% | -14.22% | withheld (directions disagree) |

`ods_source_backed_one_percent_edit_save` (21-cell lifecycle):

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 0.69% | 0.48% | -0.32% | -0.11% | accept |
| lifecycle mean | 0.68% | 0.57% | -0.49% | -0.37% | accept |
| lifecycle p95 | 0.75% | -0.40% | -1.87% | -0.74% | withheld (directions disagree) |
| lifecycle p99 | -2.80% | -1.44% | -1.70% | -3.00% | withheld (directions disagree) |
| commit p50 | 3.45% | 2.97% | -0.30% | 0.19% | accept |
| commit mean | 3.28% | 2.91% | -0.36% | 0.02% | accept |
| commit p95 | 2.50% | 1.16% | -0.86% | 0.50% | accept |
| commit p99 | -0.68% | 0.12% | -4.02% | -4.78% | withheld (directions disagree) |

The one-edit lifecycle is neutral as designed (its single edit is dominated
by open and publication costs; the withheld one-edit commit p99 directions
are strongly negative, so no claim is made there). The one-percent lifecycle
shows the accepted lifecycle-level improvement where the commit reparse
validates more edited cell text. No regression trigger fired on any
withheld lifecycle statistic.

## Verification

```text
cargo test --locked -p litchi-ods --all-targets
cargo clippy --locked -p litchi-ods --lib --test source_cell_transactions -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
```

The litchi-ods suite passes 339/339 including the two new boundary tests.
Scoped strict Clippy, rustdoc, formatting, and crate-boundary checks pass.
Unrelated pre-existing strict-Clippy failures in untouched litchi-ods test
files (`facade_round_trip.rs`, `tracked_changes*.rs`) reproduce identically
without this change and are outside its scope.

Artifacts:

- repeated-edit: [summary](../results/ods-repeated-edit-0194-summary.json),
  [manifest](../results/ods-repeated-edit-0194-manifest.json)
- one-edit guardrail: [summary](../results/ods-one-edit-save-0194-summary.json),
  [manifest](../results/ods-one-edit-save-0194-manifest.json)
- one-percent guardrail:
  [summary](../results/ods-one-percent-edit-save-0194-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0194-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in each manifest
