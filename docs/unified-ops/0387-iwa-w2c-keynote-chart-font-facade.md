# 0387 — W2-C Keynote slide-chart font facade

- Status: bounded facade slice on `uo2/iwa-w2c-chart-font`, cut from frozen
  root `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`. Coordinator integration,
  shared-engine admission review, native evidence and final integrated gates
  remain pending. The six exit-ledger rows intentionally remain `host-only`
  until E's shared contract and the scoped qualification gates are accepted;
  this record is not a provisional migration state.
- Date: 2026-10-28.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) W2-C; parity/evidence policy
  [0369](0369-iwa-exit-parity-harness-and-ledger.md); shared chart-font engine
  [0375a](0375a-iwa-w1e-style-variation-write-engine.md).
- Scope: the six Keynote host rows in
  `keynote/editor/slide_charts/font.rs`: font identity/face-trait read, set
  and reset, and point-size read, set and reset. No legend font, chart text,
  other chart styles, dispatch/schema exposure or host deletion is included.

## Public contract and implementation

The supported destinations are:

```text
litchi_keynote::Package::{slide_chart_font, edit_slide_chart_font,
                          apply_slide_chart_font}
litchi_keynote::Package::{slide_chart_font_size, edit_slide_chart_font_size,
                          apply_slide_chart_font_size}
```

The `chart::font` and `chart::font::size` modules expose the focused value,
`Edit::{set, reset}`, `Commit`, `Patch`, `Diagnostics`, `Error`, `LimitKind`
and `Refusal`. Selectors are slide/chart positions or exact semantic names;
native chart identifiers, IWA objects, component names, ZIP paths and
protobuf values do not cross the public boundary. Font and point-size values
reuse the validated `litchi_iwa_text::{font::Font, character::TextPointSize}`
types. Font and font-size are intentionally independent transactions, so a
uniform font remains readable when point sizes differ.

`slide_chart_font.rs` owns only Keynote selector resolution, source ownership,
exact package publication and candidate verification. It delegates paragraph
style graph work to `litchi_iwa_style::chart_font`; it does not duplicate the
variation algorithm. Successful commits retain the immutable source snapshot,
reopen the exact candidate, verify the selected chart and requested set value,
and carry an exact-source patch with `inverse()`. A set-to-current-value is a
byte-exact no-op. Patch application checks source authorization and preserves
the original touched-component diagnostic count.

The adapter currently awaits E's reviewed shared-engine admission/error seam.
The branch must not be accepted as final until the shared route charges all
pre-work before graph inspection/allocation and exposes typed archive, wire,
codec, registry, limit and allocation failures. No broad `InvalidSource`
mapping is evidence of that contract; the final adapter mapping is to be
updated against E's approved API before integration.

## Differential proof and evidence boundary

`crates/litchi-iwa/tests/parity/keynote_chart_font_cases.rs` owns the Keynote
facade vectors. The source is
`test-data/iwork/keynote/chart-data-native.key` with a host-added Column2d
chart. That fixture already has one native chart, so the host-added chart is
the second slide-0 chart and the leaf uses selector index 1; the host's native
drawable identifier remains test-only.

The six facade vectors are:

| Route | Case | Tier |
|---|---|---:|
| font set | `keynote_slide_chart_font_allocate_facade_parity` | B |
| font read | `keynote_slide_chart_font_read_facade_parity` | A |
| font reset | `keynote_slide_chart_font_reset_facade_parity` | A |
| point-size set | `keynote_slide_chart_font_size_allocate_facade_parity` | B |
| point-size read | `keynote_slide_chart_font_size_read_facade_parity` | A |
| point-size reset | `keynote_slide_chart_font_size_reset_facade_parity` | A |

Before each read proof, the host authors the source and the test directly
compares the host reader with the independent leaf reader on those same bytes;
the generic parity harness's same-reader checks remain additional coverage.
Set/reset vectors use the host differential, strict leaf reopen, inverse patch
restoration and untouched-member/object locality checks. Tier-B differences
are limited to the established fresh style UUID and save-token metadata
policy. `keynote_slide_chart_font_noop_inverse_and_refusal_controls` covers
exact no-op bytes, no-op inverse, and typed out-of-range selector refusals.

Native Apple acceptance is absent. The four focused Keynote facade queue rows
in [0363-mac-queue.md](0363-mac-queue.md) remain `pending`; the existing two
shared chart-font queue rows in that file also remain pending. No E3/E4 claim
is made. The exit ledger remains host-only until the shared admission/error
contract, focused gates and performance disposition are accepted.

## Performance and verification

`keynote_chart_font_facade_cost_abba` is an ignored full-facade probe covering
font/size set and reset in host/leaf/leaf/host order. It was added for the
coordinator's pinned release ABBA slot but was not run while timing was held;
no performance qualification or wall-clock claim is made here. Any later
receipt must name the final integrated source SHA, release binary SHA-256,
build command, exact test filter and route symbols.

Scoped evidence on the current branch:

| Check | Result |
|---|---|
| `cargo fmt --all` | pass after the focused edits |
| Keynote crate check | pass before the final diagnostics/test additions; rerun required after E's contract lands |
| six focused Keynote facade parity cases | 6 passed |
| no-op, inverse and typed-refusal controls | passed |
| native evidence / queue | pending; no claim |
| performance | held for coordinator; no claim |

The coordinator must rerun the affected strict Clippy, rustdoc/glossary,
ledger, crate-boundary and fuzz-lock checks after the reviewed shared-engine
contract is consumed, then rerun the complete 53-row gate on the final
integrated commit. This record does not authorize deleting the host or parity
harness.
