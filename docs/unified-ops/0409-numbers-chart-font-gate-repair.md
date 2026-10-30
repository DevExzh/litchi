# 0409 — Numbers chart-font reset gate repair

- Base: coordinator source `3ce2a067742b9606e3c2558adee9e022c87f0420`.
- Scope: Numbers chart-font facade and focused parity/contract tests only;
  shared chart-font engine registry admission and the Numbers series-style
  paths remain separate ownership.
- Performance claim: none; no capture or timing was run.

## Root cause and bounded repair

The two failing workspace-parity cases were
`numbers_sheet_chart_font_reset_parity` and
`numbers_sheet_chart_font_size_reset_parity`. The reset source is host-mutated
and close to the default 100,000-field profile. `Edit::commit` performed a
second package-wide selector census and constructed a second graph workspace
before mutating the chart, while the immutable edit already retained the
resolved target. This left a residual field ceiling of 96,448 for a reset
operation that required about 108,000 fields.

The facade now reuses one admitted workspace for before-read, chart-font
mutation, and staged readback. The redundant source-side selector re-resolution
was removed; candidate publication still resolves the target and performs
strict semantic readback. The shared engine also no longer pre-admits a
registry mutation immediately after removing a style, before the actual
registry mutation admission; the subsequent exact inspection and mutation
admissions remain intact. The observed failure was a chart-font `Fields`
admission (mapped through the existing facade taxonomy as `WireWork`), not a
work-ceiling failure. No limit ceiling or headroom was raised.

## Verification

- `cargo fmt --all -- --check` — pass.
- Numbers chart-font parity set/size/reset/read vectors — pass (5 focused
  parity tests).
- `cargo test -p litchi-numbers --test sheet_chart_font --locked --offline` —
  pass (4 contract tests).
- Typed reset/readback/inverse/no-op behavior remains covered by the existing
  parity and contract vectors. Series-style failures are outside this slice.
