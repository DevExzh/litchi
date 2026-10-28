# 0385 — W2-A Numbers chart-font facade

## Status and ownership

This record describes the isolated preparation slice on
`uo2/iwa-w2a-chart-font`, based on coordinator root
`9932c7d0691b2c14bbbd11a2dea7df41f66308b6`. It is not an integration or W3
claim. The six Numbers chart-font rows remain provisional until the shared
chart-font bounded-limit/report/error contract from the W1-E owner is
accepted, and until the coordinator runs final native, performance and
integrated-gate checks. The six canonical ledger rows intentionally remain
`host-only` until that contract and the final scoped gates are accepted. The
parity declarations and native queue references below are preparation
evidence, not a migration claim.

The slice owns only Numbers chart font identity/traits and point size:

* `sheet_chart_font` / `edit_sheet_chart_font` for font identity, bold and
  italic traits, and point-size reads and edits;
* semantic `Edit::set(Property)` and `Edit::reset(PropertyKind)` operations;
* exact-source `Patch`, `Commit`, `Diagnostics`, `Error`, `LimitKind` and
  `Refusal` plumbing for the focused route;
* six Numbers host rows: font read/set/reset and size read/set/reset.

Chart titles, captions and axis titles remain E's ownership. Drawable
comments remain G's ownership. This slice does not include chart arrangement,
series, axis style, image fill, endpoints, segments, presets, lifecycle,
catalog or any other chart/drawable family. It does not edit the shared
`litchi_iwa_style::chart_font` engine.

## Public contract

The public destination is `litchi_numbers::sheet::chart::font`. Selectors are
`SheetSelector` and `ChartSelector`; native object IDs, UUIDs, archive paths
and protobuf values do not cross the boundary. The package methods are:

```text
Package::sheet_chart_font(sheet, chart, PropertyKind)
Package::edit_sheet_chart_font(sheet, chart)
Package::apply_sheet_chart_font(&Patch)
Edit::set(Property)
Edit::reset(PropertyKind)
Edit::commit()
```

The adapter reuses `litchi_iwa_style::chart_font::{font, size, set}` and the
existing Numbers style-graph publisher. It resolves the selected chart before
reading; each read, staged baseline, candidate readback and patch application
inspects only the requested/touched family. Thus a mixed font/size chart does
not cause an unrelated family to refuse. Staged operations, selector failures,
source failures and patch conflicts are terminal typed results; the focused
route never falls back to the legacy host.

The current facade's admission and error mapping are explicitly provisional
pending E's chart-font-specific contract. In particular, the final contract
must define bounded pre-admission/reporting for payload, references, wire,
output, allocation, retained, scratch and identifier resources, and the
mapping of shared codec/text failures. The implementation does not broaden
those mappings or add fallback paths while that contract is pending. The
current finite staged-operation ceiling is covered by a public atomic-limit
test but is not itself the final resource-ceiling qualification.

Patch application is source-authorized and exact-source. A no-op returns the
original bytes and an inverse exchanges the family-scoped before/after
values. The `Patch::before/after` optional family projection is intentionally
provisional while E finalizes the shared contract. A changed candidate is
reopened through the strict Numbers ingress;
the touched property is read back before returning the commit. Untouched
members and unrelated graph objects are left to the existing exact-source
publisher policy.

## Fixtures and differential proof

The host differential uses the existing source-built Numbers chart fixture
from `chart_style_cases::input(Format::Numbers, Kind::Column2d)`, whose
`test-data/iwork/numbers/chart-data-native.numbers` source receives a host
Column2d chart. This keeps the host and leaf selectors aligned without
introducing a third-party fixture. The public contract tests use the tracked
Apple-authored `test-data/iwork/numbers/chart-grid-authored-native.numbers`
fixture, whose existing chart is readable by the shared chart-font engine.

`numbers_chart_font_cases.rs` contains four differential vectors:

* font set and point-size set are Tier B because fresh style identities and
  save-token/identity policy differences are intentional;
* font reset and point-size reset are Tier A where the host and leaf preserve
  the same graph bytes;
* every case retains the source through the read-only harness checks and
  directly compares the host reader with the leaf reader before and after the
  operation.

The direct reader proof has six populated vectors: Avenir font, Georgia font,
font reset after Avenir, 17 pt, 9 pt after Georgia, and size reset after 17 pt.
It also includes a negative control proving that mismatched reader values are
rejected by the same `reader_values_match` helper. The harness's independent
host/leaf reader callbacks remain supplementary; they are not substituted for
the direct cross-reader assertions.

Native acceptance is not claimed. The six routes are appended to
`docs/unified-ops/0363-mac-queue.md`, all with `pending` results and no
artifact hash. A native-authored read receipt would not prove that a
Litchi-mutated package survives native open/save/close/reopen, so E3/E4 are
not claimed here.

## Public tests and probes

The Numbers contract test covers populated reads, exact no-op bytes, exact
apply, inverse restoration, stale-patch conflict, missing sheet/chart typed
refusals, empty-transaction refusal and atomic finite staging-limit refusal.
The six-route ignored full-facade probe includes package open, selector
resolution, read or edit, candidate commit/readback for writes and
serialization. It performs no timing itself.

Commands used for the scoped preparation (build artifacts and temporary files
are outside the checkout):

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-iwa-w2a-chart-font
CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-target-20261028 \
  LITCHI_MEM_MAX=20G \
  cargo test --offline --locked -p litchi-numbers --test sheet_chart_font \
    --all-features -- --nocapture

CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-target-20261028 \
  LITCHI_MEM_MAX=20G \
  cargo test --offline --locked -p litchi-iwa --test parity --all-features \
    numbers_sheet_chart_font -- --nocapture

CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-target-20261028 \
  LITCHI_MEM_MAX=20G \
  cargo test --offline --locked -p litchi-iwa --test parity --all-features \
    numbers_chart_font_cost::numbers_chart_font_facade_cost_abba \
    -- --ignored --exact --nocapture
```

The focused runs passed: three Numbers public-contract tests, four
host-vs-leaf differential tests, the six-vector direct-reader test and the
ignored six-route probe. `cargo check --offline --locked -p litchi-numbers
--all-features` and the parity-test target check also passed. These are scoped
checks only; no full 53-row gate, native run, callgrind qualification or
wall-clock 5%/20-microsecond claim is made.

The post-preparation static gates also passed: `python3
tools/check_crate_boundaries.py` reported 70 workspace packages and 313 valid
internal declarations; `python3 -m unittest
tools.test_no_3rdparty_references` ran four tests successfully; and the
locked Numbers fuzz workspace compiled with:

```sh
CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-fuzz-target-20261028 \
  LITCHI_MEM_MAX=20G \
  cargo check --offline --locked \
    --manifest-path crates/litchi-numbers/fuzz/Cargo.toml --bins
```

The full-facade probe's six routes are `font_read`, `font_set`, `font_reset`,
`size_read`, `size_set` and `size_reset`, selected with `W2A_FONT_ROUTE`. The
coordinator must run the release-mode pinned-CPU ABBA measurement and report
engine/facade scope separately before accepting performance evidence. The
probe's two direct `#[inline(never)]` route symbols are
`numbers_chart_font_host_route` and `numbers_chart_font_leaf_route`; a receipt
must use their exact symbol rows rather than program totals or sums of inline
duplicates. Later integrations invalidate any earlier performance
qualification.

## Handoff conditions

Before integration, E must publish the chart-font-specific bounded limit and
error contract, and this adapter must be reconciled with it without adding a
legacy fallback. The coordinator must then review the six host-only ledger
rows and queued native routes, decide their final disposition, rerun affected
formatting/Clippy/docs, boundary/fuzz-lock checks and the final integrated
53-row gate, and schedule
native open/save/close/reopen verification. The branch must be synchronized
with current trunk only after those checks are requested; this record does not
authorize integration.
