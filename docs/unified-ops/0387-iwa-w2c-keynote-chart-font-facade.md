# 0387 — W2-C Keynote slide-chart font facade

- Status: bounded facade slice on `uo2/iwa-w2c-chart-font`, cut from frozen
  root `9932c7d0691b2c14bbbd11a2dea7df41f66308b6`, with the reviewed shared
  substrate `faefe7303b` and frozen trunk `1bc8f22499` merged at branch tip
  `c3937337473bee2a985a27fb8ffa2407489c85f7`. The selector/locality and
  admission/cost-probe repair is `6ac2c0efcd`. The six exit-ledger rows are now
  migrated on this reviewed WP; coordinator integration, native evidence and
  final integrated gates remain pending.
- Date: 2026-10-29.
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

The adapter consumes E's reviewed shared-engine admission/error seam from
`faefe7303b`: it derives an explicit residual chart-font profile from one
package transaction budget, admits the graph index before `Workspace::new`,
charges engine reports, preflights reassembly and candidate ingress, and
keeps the same budget through candidate readback. It clamps only to the
codec's published hard ceilings; it does not call `Limits::for_source` or
replace a caller profile with a fresh source-sized default. Engine archive,
wire, codec, graph-limit, allocation and identifier failures retain typed
facade classifications. The shared substrate itself remains outside this
work package. Chart selection now enters the shared budget-aware selector,
output-free reassembly requirements are admitted before final ZIP execution,
and a source-relative locality verifier uses the patch's stored selected-graph
plan. It rejects extra/deleted members and objects, preserves untouched raw ZIP
records (allowing only the reassembler's central-directory offset rewrite),
checks static ZIP metadata on changed members, and preserves unknown fields in
the chart extension. Reset routes
read the staged inherited effective value before publication and reject mixed
or otherwise unmodelled reset results. Engine text-byte reports retain their
`TextBytes` classification and accumulate across the transaction.

Locality payload inspection uses a residual, non-allocating `RawWireFields`
preflight, debits its measured input/work/field pass, and only then admits the
single materializing field vector's cumulative fields, work, scratch and
allocation event. The unit controls cover zero, exact and one-under cumulative
field budgets plus exact and zero allocation residuals. The ignored release
probes use dedicated `#[inline(never)]` host/leaf wrappers: setup, inverse
application and output checks are outside the route, each read fetches one
property, and each ABBA route executes exactly two selected calls per side.

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
is made. The six ledger rows are migrated on Tier-A/B host differential proof,
but their pending native queue entries still block W3.

## Performance and verification

Six ignored full-facade probes isolate the coordinator's pinned release ABBA
routes: `keynote_chart_font_{set,size_set,reset,size_reset,read,size_read}_cost_abba`.
Each runs one host/leaf/leaf/host route; reset setup and read authoring happen
outside the captured route. On the clean frozen WP source, all 12 CPU-5
Callgrind legs passed exact one-test execution, positive selected Ir equal to
attribution Ir, and unchanged source/binary hashes. The release binary was
built with:

```text
CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0
CARGO_PROFILE_RELEASE_DEBUG=1 CARGO_BUILD_JOBS=4 CARGO_NET_OFFLINE=true
cargo test --release --locked --offline -p litchi-iwa \
  --features parity-host-oracle --test parity --no-run
```

Source `f875e5c3c9199930286c3295c67a42befd5e70b7` has tree
`25de453444c2d9791dacfe3fd699d21e393223db`; `Cargo.lock` is 139901 bytes with
SHA-256 `f7cca120269bcdaaad88f826299ccc51564bcc73985091446d8acc23b0b845bb`.
The archived/staged release binary is 214039104 bytes with SHA-256
`8871832870087b200d22336f8116662e7898c3c63e79aad3dd41476cb7cc78e5`.
The external run receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261029/perf/runs/keynote-chart-font-f875e5c3c9-cpu5/results.json`;
its SHA-256 is
`24740463740b65be29dd53a5df66c1ebbe88164de2b3869e00eb0347ada4a1c1` and the
artifact index SHA-256 is
`aede7d3abf82ecf9c5952e45a5cb3b98944f1bf10abe992844b40d542839e77c`;
the raw selected-call audit is
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261029/perf/keynote-chart-font-six-call-audit.json`
and its validation SHA-256 is
`2a1336821da846a6c0a808974660b9f98754b9610a0f4e97a6aee389ef670cd0`.

| Route | Host Ir | Leaf Ir | Leaf/host slowdown |
|---|---:|---:|---:|
| font read | 708614338 | 55828718 | -92.121% |
| font set | 1848774983 | 477377895 | -74.179% |
| font reset | 661952308 | 320762546 | -51.543% |
| point-size read | 596540266 | 55209214 | -90.745% |
| point-size set | 1504233570 | 469959135 | -68.758% |
| point-size reset | 661942039 | 315323675 | -52.364% |

Scoped evidence on the current branch (source repair `6ac2c0efcd`):

| Check | Result |
|---|---|
| `cargo fmt --package litchi-keynote --package litchi-iwa -- --check` | pass |
| `cargo check -p litchi-keynote --locked --offline` | pass |
| strict Clippy (`cargo clippy -p litchi-keynote --lib --tests --locked --offline -- -D warnings`) | pass; vendor warnings only |
| focused `keynote_chart_font` parity/control filter | 10 passed, 6 ignored isolated ABBA probes |
| eight Keynote facade parity/control cases within that filter | 8 passed |
| shared chart-font admission filter | 7 passed |
| `cargo test -p litchi-keynote --lib` | 335 passed, 1 ignored |
| production locality negative controls (unit) | 1 passed: extra member/object, non-chart unknown payload, ZIP metadata |
| low text-budget and stale-patch controls | passed |
| zero/exact/cumulative locality fields and allocation controls | passed |
| `python3 tools/check_crate_boundaries.py` | pass; 70 packages, 313 declarations, 12 explicit debt items |
| `python3 tools/iwa_exit_ledger.py --check` | pass; six chart-font rows migrated, native queue still pending |
| Keynote minimal/full API glossary profiles | pass; 267/266 banned names, 52 families and 52 result contracts each |
| `cargo rustdoc -p litchi-keynote --lib --locked --offline -- -D warnings` | pass |
| crate-boundary/non-iWork/no-`3rdparty`/ledger tooling tests | 1,221 passed; gate-runner fuzz-lock tests 5 passed |
| no-op, inverse and typed-refusal controls | passed |
| release parity build with `parity-host-oracle` | pass; exact source/binary receipt above |
| CPU-5 Callgrind six-route facade matrix | 12/12 legs passed; all six within 5%, no wall-clock claim |
| selected-call audit | pass; 6 routes, 12 legs, exactly 2 calls and 1 wrapper symbol per side |
| native evidence / queue | pending; no claim |
| performance | held for coordinator; no claim |

The coordinator must rerun the affected strict Clippy, rustdoc/glossary,
ledger, crate-boundary and fuzz-lock checks after the reviewed shared-engine
contract is consumed, then rerun the complete 53-row gate on the final
integrated commit. This record does not authorize deleting the host or parity
harness.
