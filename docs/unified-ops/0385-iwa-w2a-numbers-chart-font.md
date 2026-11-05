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
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  LITCHI_MEM_MAX=20G \
  cargo test --offline --locked -p litchi-numbers --test sheet_chart_font \
    --all-features -- --nocapture

CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  LITCHI_MEM_MAX=20G \
  cargo test --offline --locked -p litchi-iwa --test parity --all-features \
    numbers_sheet_chart_font -- --nocapture

CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
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
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  LITCHI_MEM_MAX=20G \
  cargo check --offline --locked \
    --manifest-path crates/litchi-numbers/fuzz/Cargo.toml --bins
```

The full-facade probe's six routes are `font_read`, `font_set`, `font_reset`,
`size_read`, `size_set` and `size_reset`, selected with `W2A_FONT_ROUTE`. The
corrected release qualification is recorded below; the historical instruction
and shared-helper sections later in this record are explicitly withdrawn.
The probe's two direct
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

## Corrected release ABBA requalification

The previous 18-route qualification was withdrawn because its reset setup
legs invoked the selected host wrapper and therefore recorded three host calls
instead of two. The coordinator-authorized probe repair `dc585d05cd` was
cherry-picked and resolved on this branch as source-only probe commit
`2a5073093c48b540e1f6fcf14fbd82a6101db3a8`; the chart-font reset setup was
then isolated in source-only probe commit
`59ffe48b3e0092840f68519571492a08b0fe5bfc`.

The corrected source was built with `parity-host-oracle`, release LTO/debug
settings, locked offline dependencies, four Cargo jobs and incremental
compilation disabled:

```text
source: 59ffe48b3e0092840f68519571492a08b0fe5bfc
binary: (deleted-coord-artifact)
size: 213,510,848 bytes
ELF SHA-256: 4d099883840c20abf82fac49df88e1de574c6a58f34bdfdc76a7a552a122125c
archive: (deleted-coord-artifact)/parity-20d17b646a6d1747-reset-isolated-release.tar
archive mode/size: 0444 / 213,514,240 bytes
archive SHA-256: 01d7c6b489205f5d23cbfc9576b5f7f511753933356b18c1ffdb60d31954abb9
tar member SHA-256: 4d099883840c20abf82fac49df88e1de574c6a58f34bdfdc76a7a552a122125c
```

The 24-route matrix combines the six chart-font routes and the 18 existing
style/layout routes. It used one fresh source per route, pinned CPU 7, and
true host/leaf/leaf/host ABBA with exactly two selected calls per side:

```text
manifest: (deleted-coord-artifact)/manifest-24-reset-isolated.json
results: (deleted-coord-artifact)/results.json
results SHA-256: 916982820dd0254d9b92f28142f62b81de0741413cdd21870700179e15d62fa8
captures: 48 Callgrind files / 48 annotations / 48 stdout logs / 48 stderr logs
raw Callgrind bytes: 22,224,570
full probe correctness (six font routes): `(deleted-coord-artifact)/correctness-font-full.log` (SHA-256 `00b7f561a20a061eebdb2893f9ea16e5c5a584eac7daf297db24d13980133bc8`), stderr empty (`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
full probe correctness (18 style/layout routes): `(deleted-coord-artifact)/correctness-style-full.log` (SHA-256 `1006202a0d367176831b9ac1fdb4e3318c91162a1da76e550dd081eae2bb9bb9`), stderr empty (`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
source-before SHA-256: df63cb852ca737c5e00325d166494d910392b38eb2b42b2514e78736c150f27e
source-after SHA-256: 54882a78211ffbc616cfa6eaee0b7a88846852117d612d9277a60d9cf7104d84
```

All 24 exact tests exited 0. Every leaf route was faster than its host
route; no optimization regression was found. The complete per-route
host/leaf Ir, exact symbol rows, command lines, hashes and positive-test
receipts are retained in `matrix24/results.json` and the external capture
directory above. The six chart-font rows measured as follows:

| Route | Host Ir | Leaf Ir | Leaf vs host |
|---|---:|---:|---:|
| `font_read` | 285,226,410 | 38,942,825 | −86.346697% |
| `font_set` | 1,156,251,052 | 209,895,725 | −81.846873% |
| `font_reset` | 237,782,057 | 158,574,727 | −33.310894% |
| `size_read` | 175,678,958 | 36,851,199 | −79.023556% |
| `size_set` | 823,495,243 | 203,858,328 | −75.244747% |
| `size_reset` | 237,717,303 | 153,833,491 | −35.287214% |

The 18 style/layout rows are the exact `ROUTES_3` set already listed in the
shared-helper section below; their full values remain in the external matrix
receipt rather than being copied into a second hand-maintained table.

The new strict parser receipt below is retained as a failed historical
diagnostic; it treated a smaller binary/file attribution row as the whole
function. The coordinator's DWARF-aware audit instead sums the selected
function's self and callee Ir across file contexts, without substituting or
normalizing any PROGRAM TOTALS. It verifies all 241 hashed artifacts/binaries
and exactly two selected calls on every leg:

```text
coordinator audit: (deleted-coord-artifact)/coordinator-derived-audit.json
audit SHA-256: 380557ce35908ea36299b7a65432a95a758bced5b49a39678b32a997e73fed55
coordinator review: (deleted-coord-artifact)/coordinator-review.json
review SHA-256: 57d6a6c23e73d73219fe7906e062f5899c889e568a01cd3c01035d98ac2fefaf
call-count receipt: (deleted-coord-artifact)/call-count-receipt.json (SHA-256 `e19bbdae0c9bfd1869d39b41d7b9d638af154b15573ee2873b6e2e9ffff38108`)
historical strict-parser failure: (deleted-coord-artifact)/call-audit-derived.failure.json (SHA-256 `226a8f4c7bdc28fc0c666a569d88c0965b2decbc001faa2ef44914ee38ff5740`)
synthetic validator log: (deleted-coord-artifact)/derived-validator-selftest.log
synthetic log SHA-256: a72c11d7f48cc0c1e4e659a05d8038b9d88d32ee36f9c71face0425df81f2914
```

The coordinator audit passes the corrected slice with maximum leaf/host ratio
`0.6915293508856026`. This is slice qualification only; native acceptance,
ledger migration and final integrated-gate claims remain absent, and the final
integrated candidate must be remeasured after trunk synchronization.

## Withdrawn historical instruction-count qualification

Everything in this section is retained for provenance only and is superseded
by the corrected CPU-7 matrix above. Its old totals must not be used for
acceptance.

This is a qualification of the current isolated candidate only. The measured
source tip is `fbc7c51866f119e733521f4bc92467a849385c71`, after merging the
coordinator root `295b421bcd3d9ec0b0b9208e4039de22623685bc`; any later root
merge invalidates this receipt. The release parity binary was built with one
Cargo job, incremental compilation disabled, release debug information
disabled, and `RUST_TEST_THREADS=1`:

* binary: `(deleted-coord-artifact)`
* binary size: 41,439,288 bytes; SHA-256 `258aaccd848ffde2069852ad8ed61f565170d9307d6046269c25f50474f3d9c3`
* immutable binary archive: `(deleted-coord-artifact)/parity-9f6fee151a57f4bc-release.tar`, mode `0444`, 41,441,280 bytes, SHA-256 `1325e220da1ab56c5b9548018298ab06906bf6a6183efd676e77ba14ebda2a05`

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
`(deleted-coord-artifact)/summary.tsv`
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

## Withdrawn historical shared-helper regression qualification

This section's old 18-route totals used the contaminated three-host-call
reset workload and are withdrawn. The corrected 18-route values are included
in the external 24-route matrix above.

The accepted facade also depends on the shared stylesheet registration path
used by the already integrated 18-route W2-A style/layout slice in record
0380. That dependency was requalified on the same measured code source
`fbc7c51866f119e733521f4bc92467a849385c71` and the same release parity binary
used above; the later documentation and ledger commits do not change that
binary. The audited runner is
`crates/litchi-iwa/tests/parity/numbers_w2a_cost.rs`, test
`numbers_w2a_cost::numbers_w2a_route_cost_abba_3`, with the exact 18-entry
`ROUTES_3` manifest:

```text
shape_style_fill_{set,reset,read}
shape_style_stroke_{set,reset,read}
shape_style_effects_{set,reset,read}
shape_style_layout_{set,reset,read}
text_box_style_columns_{set,reset,read}
text_box_style_layout_{set,reset,read}
```

Each selected route used one common prepared input for the host/leaf pair,
host/leaf/leaf/host ABBA, and no inverse operation in the measured forward
route. Two captures per route were taken on pinned CPU 7, 36 captures total.
Counts use only `w2a_host_route_3` and `w2a_leaf_route_3`, never program
totals.

The release binary was regenerated with this exact source/flag profile (the
target, temporary and evidence directories are outside the checkout):

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-iwa-w2a-chart-font
env -u LITCHI_BLESS -u LITCHI_CAPPED \
  PATH=(deleted-coord-artifact):$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_DEV_DEBUG=0 \
  CARGO_PROFILE_TEST_DEBUG=0 CARGO_NET_OFFLINE=true RUST_TEST_THREADS=1 \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  TMPDIR=(deleted-coord-artifact) \
  cargo test --offline -p litchi-iwa --all-features --test parity \
    --release --no-run
```

Each selected regression route used the following capture shape, with
`W2A_ROUTE` set to one exact `ROUTES_3` entry and `rep1`/`rep2` as the two
capture repetitions:

```sh
W2A_ROUTE="$route" taskset -c 7 valgrind --tool=callgrind --quiet \
  --dump-instr=yes --collect-jumps=yes \
  --callgrind-out-file="$evidence_dir/$route.$rep.callgrind" "$binary" \
  --ignored --exact numbers_w2a_cost::numbers_w2a_route_cost_abba_3 \
  --test-threads=1
```

| Route | Host Ir mean | Leaf Ir mean | Leaf vs host |
|---|---:|---:|---:|
| `shape_style_fill_set` | 257,969,442 | 152,237,222 | −40.986335% |
| `shape_style_fill_reset` | 346,429,422 | 137,515,542 | −60.304890% |
| `shape_style_fill_read` | 72,481,490 | 34,113,478 | −52.934911% |
| `shape_style_stroke_set` | 258,011,774 | 152,280,444 | −40.979265% |
| `shape_style_stroke_reset` | 346,371,434 | 137,620,428 | −60.267963% |
| `shape_style_stroke_read` | 72,460,233 | 34,133,341 | −52.893691% |
| `shape_style_effects_set` | 257,946,686 | 152,231,668 | −40.983282% |
| `shape_style_effects_reset` | 346,306,383 | 137,629,550 | −60.257865% |
| `shape_style_effects_read` | 72,463,821 | 34,131,404 | −52.898696% |
| `shape_style_layout_set` | 257,981,914 | 152,301,842 | −40.964140% |
| `shape_style_layout_reset` | 346,423,572 | 137,701,096 | −60.250656% |
| `shape_style_layout_read` | 72,472,922 | 34,133,098 | −52.902275% |
| `text_box_style_columns_set` | 221,812,232 | 152,208,272 | −31.379676% |
| `text_box_style_columns_reset` | 310,034,750 | 137,526,288 | −55.641654% |
| `text_box_style_columns_read` | 54,448,306 | 34,100,168 | −37.371479% |
| `text_box_style_layout_set` | 221,787,714 | 152,219,502 | −31.367027% |
| `text_box_style_layout_reset` | 310,003,414 | 137,654,109 | −55.595938% |
| `text_box_style_layout_read` | 54,451,990 | 34,128,222 | −37.324196% |

All 36 invocations exited 0; route stderr was empty and every selected test
reported `ok`. The external regression captures total 80,056,587 bytes, with
per-capture sizes and SHA-256 values in
`(deleted-coord-artifact)/summary.tsv`
(SHA-256 `6caab3d3f4d06a7c3aa3e3d5d19a6a4e0aaaf15214fd72a48484a4d2cdfe723d`).
The positive exact test log is `correctness-18.log` (SHA-256
`83c02acee0f9d1998b45327e2c21d96948ab3a856114c2e66cf77eef8e6f6aa2`), with
empty stderr (SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`).

The binary archive cited above is a mode-0444 tar container, so its container
SHA-256 intentionally differs from the ELF SHA-256. `tar -xOf` over its sole
member `parity-9f6fee151a57f4bc` produces the measured ELF SHA-256
`258aaccd848ffde2069852ad8ed61f565170d9307d6046269c25f50474f3d9c3`, exactly
matching the binary on disk. No wall-clock timing or native acceptance is
claimed; the six chart-font ledger rows remain migrated with native evidence
pending, and the shared-helper regression does not create E3/E4 evidence.
