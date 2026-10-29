# 0385 — W2-A Numbers chart-font facade

## Status and ownership

This record describes the isolated preparation slice on
`uo2/iwa-w2a-chart-font`, based on coordinator root
`9932c7d0691b2c14bbbd11a2dea7df41f66308b6`. It is not an integration or W3
claim. The six Numbers chart-font rows remain provisional until the
coordinator runs final native, performance and integrated-gate checks. The six
canonical ledger rows intentionally remain `host-only` until those final
checks are accepted. The parity declarations and native queue references
below are preparation evidence, not a migration claim.

The slice owns only Numbers chart font identity/traits and point size. The
shared bounded chart-font contract was adopted from E at
`faefe7303bacfa7c23d908cda73ff6384ca0ba0b`; this repair remains isolated and
is not an integrated-trunk or native-acceptance claim.

The slice owns:

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

The facade derives residual chart-font limits from the package transaction
ledger and maps the adopted engine/graph resource errors to typed facade
limits. Component/index/workspace collection capacity is admitted before
construction; nesting is tracked as maximum observed depth rather than a
consumed quota. Changed candidates are planned from output-free archive
lengths and bounded Snappy output before publication, then actual candidate
bytes are checked before readback. The finite staged-operation ceiling remains
covered by a public atomic-limit test; the resource ceiling is still subject
to final integrated qualification.

Patch application is source-authorized and exact-source. A no-op returns the
original bytes and an inverse exchanges the family-scoped before/after
values. The `Patch::before/after` optional family projection is family-scoped:
`None` means that family was untouched. A changed candidate is
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
refusals, empty-transaction refusal, atomic finite staging-limit refusal and
an expanded-candidate typed refusal that proves source atomicity.
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

The focused runs passed: four Numbers public-contract tests, four
host-vs-leaf differential tests, the six-vector direct-reader test and the
ignored six-route probe. `cargo check --offline --locked -p litchi-numbers
--all-features` and the parity-test target check also passed. These are scoped
checks only; no full 53-row gate, native run or wall-clock 5%/20-microsecond
claim is made. The current-branch release instruction qualification is
recorded below.

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
The release-mode pinned-CPU ABBA measurement below reports the full-facade
scope separately from any engine-only measurement. The probe's two direct
`#[inline(never)]` route symbols are
`numbers_chart_font_host_route` and `numbers_chart_font_leaf_route`; a receipt
must use their exact symbol rows rather than program totals or sums of inline
duplicates. Later integrations invalidate any earlier performance
qualification.

## Handoff conditions

Before integration, the adopted E bounded limit and error contract must remain
unchanged while the coordinator reviews this repair. The coordinator must
then review the six host-only ledger
rows and queued native routes, decide their final disposition, rerun affected
formatting/Clippy/docs, boundary/fuzz-lock checks and the final integrated
53-row gate, and schedule
native open/save/close/reopen verification. The branch must be synchronized
with current trunk only after those checks are requested; this record does not
authorize integration.

## Release instruction-count qualification

This is a qualification of the current isolated candidate only. The measured
source tip is `fbc7c51866f119e733521f4bc92467a849385c71`, after merging the
coordinator root `295b421bcd3d9ec0b0b9208e4039de22623685bc`; any later root
merge invalidates this receipt. The release parity binary was built with one
Cargo job, incremental compilation disabled, release debug information
disabled, and `RUST_TEST_THREADS=1`:

* binary: `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-perf-target-20261029/release/deps/parity-9f6fee151a57f4bc`
* binary size: 41,439,288 bytes; SHA-256 `258aaccd848ffde2069852ad8ed61f565170d9307d6046269c25f50474f3d9c3`
* immutable binary archive: `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-perf-20261029/parity-9f6fee151a57f4bc-release.tar`, mode `0444`, 41,441,280 bytes, SHA-256 `1325e220da1ab56c5b9548018298ab06906bf6a6183efd676e77ba14ebda2a05`

Each route used one common prepared input for its host/leaf pair. Setup and
reset seeding occurred before the measured forward route symbols; the measured
ABBA order was host, leaf, leaf, host, with no inverse operation in the route.
Each route was captured twice on pinned CPU 3, for 12 external Callgrind
captures total. Instruction counts use only the exact direct
`numbers_chart_font_host_route` and `numbers_chart_font_leaf_route` symbol
rows, never `PROGRAM TOTALS`.

| Route | Host Ir mean | Leaf Ir mean | Leaf vs host |
|---|---:|---:|---:|
| `font_read` | 285,250,784 | 38,891,554 | −86.365838% |
| `font_set` | 1,157,293,982 | 209,840,963 | −81.867964% |
| `font_reset` | 816,381,854 | 158,513,524 | −80.583409% |
| `size_read` | 175,305,496 | 36,847,437 | −78.981014% |
| `size_set` | 823,448,718 | 203,907,692 | −75.237354% |
| `size_reset` | 649,320,992 | 153,778,446 | −76.317038% |

All 12 Callgrind invocations exited 0, all route logs had an `ok` result and
empty stderr, and every route is below the 5% slower bar. The external raw
captures total 26,546,923 bytes. Their per-capture sizes and SHA-256 values,
the two-repetition means, and the positive-exit markers are in
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-chart-font-perf-20261029/summary.tsv`
(SHA-256 `a412373721c6f5ebabdfd079e204705c1225b651507ef8d016a918f66fdb26e1`).
The representative exact annotation was saved as
`font_read.rep1.annotate.txt` (SHA-256
`08dd00f6d2457c7c8a5d65a6e4a5440ca4822b240d99712a11858bca0133abef`); it
contains the separate host and leaf symbol rows alongside the program-total
context, which was not used for qualification.

The correctness positive run used the same release binary and exact parity
filter `numbers_sheet_chart_font`; it exited 0 with 5 passed tests. Its log is
`correctness-parity.log` (SHA-256
`ec534079c4ed3b01595fba3bd74b4b2176c148608bfe330672f994b8a733c5a7`). This
instruction-count result does not claim native acceptance, does not claim
wall-clock performance, and does not migrate the six ledger rows; the native
queue entries remain pending.
