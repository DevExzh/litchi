# 0271 — Dotted chart series keys and honest warn-mode failures

Date: 2026-10-13. Branch `uo2/chart-series`, cut from trunk
`feat/unified-ops-api` at `fd49666493`. Follows record 0270 (typed chart
CRUD) and record 0269 (command-channel "apply the rest and warn").

## 1. Why

The corpus table listed four bugs, all the same call: a PPTX chart `add`
whose series are given as OfficeCLI's dotted per-series keys
(`ppt/charts/charts-{area,bar,column,scatter}.sh`):

```sh
add deck.pptx /slide[8] --type chart --prop chartType=area ... \
  --prop series1.name=Web --prop series1.values=50,60,70,65,80 --prop series1.color=4472C4 \
  --prop series2.name=Mobile ...
```

Chart `add` knew only `series{N}=Name:1,2,3`, so every `series{N}.*` key was
`unsupported_property`. Under `--unsupported warn` (which the corpus harness
passes) all of them were skipped, and what was left failed
`missing_property data`: an error that told the caller it had omitted data it
had in fact supplied.

## 2. Dotted series keys on chart `add` (PPTX and XLSX)

Shared code in `crates/litchi-ops/src/format/chart.rs`; the registry tables
in `format/pptx/spec.rs` and `format/xlsx/chart.rs` gain 60 write-only
entries each, `series1.name`, `series1.values`, `series1.color`, ...,
`series20.color` (OfficeCLI reads numbers 1 to 20). The add admission's
duplicate-key bitset grows from `u64` to `u128` (124 entries in PPTX, 120 in
XLSX).

OfficeCLI's grammar (`Core/Chart/ChartHelper.cs`, `ParseLiteralSeriesData`,
`ParseSeriesDataExtended`, `ParseSeriesColors`), and what Litchi does:

| Input | OfficeCLI | Litchi |
|---|---|---|
| `series{N}.name` + `series{N}.values` | one series | same |
| `.values` without `.name` | named `Series N` | same |
| `.name` without `.values` (and no `series{N}=`) | an empty placeholder series | `missing_property` on `series{N}.values` (an empty series is not authored) |
| `series{N}=1,2,3` (no name) | named `Series N` | same (was `invalid_value`) |
| `series{N}=Name:1,2` | one series | same |
| `.name` + bare `series{N}=1,2` | the legacy key supplies the values | same |
| `.name` + `series{N}=Name:1,2` | `series{N}=` is parsed as bare values and fails | `invalid_argument` on `series{N}.name` (named twice) |
| `.values` + `series{N}=...` | the legacy key is ignored | `invalid_argument` on `series{N}.values` (values twice; never silently dropped) |
| numbers with gaps (`series1`, `series3`) | series in number order, gaps closed | same |
| `series{N}.color` | overrides position `N` of the `colors` palette; pie/doughnut: a series-level fill, while `colors` colors points | same, through the new leaf `Series::color` |
| `series{N}.color` with `data=` | colors the `N`th `data` series | same; beyond the series count is `invalid_value` (OfficeCLI ignores it) |
| `series{N}.color` for a number no other key defines | ignored | `invalid_value` on `series{N}.color` |
| `data=` with `series{N}` / `.name` / `.values` | `data` wins, the rest is ignored | `invalid_argument` on `data`, as before (record 0270) |
| a cell range (`Sheet1!B2:B13`, `B2:B13`) in `.values` or `series{N}` | a `numRef` into the workbook | `unsupported_capability` on that key (the key is supported, the reference form is not authored yet) |

A color is `RRGGBB` or `AARRGGBB` (`#` allowed), like `colors`. The keys are
add-only reads (`get: false`) like the other formatting keys.

**Leaf** (`litchi-drawingml::chart::authoring`): `Series` gains
`color: Option<String>` and `with_color`. The writer uses the series color in
place of that series' `Appearance::colors` entry; a single-series pie or
doughnut keeps its per-point `colors` and also gets the series-level fill.
`Spec::validate` refuses a color that is not `RRGGBB`/`AARRGGBB`. The writer
stays deterministic: the dotted form and `data=` plus `colors=` produce
identical bytes (tested on both hosts and three chart families).

**Chart-level `set`** (step 2 of the task, trivial enough to do):
`series{N}.name` and `series{N}.values` are also settable on the chart path
(`set /slide[1]/chart[1] --prop series2.name=Phone`), selecting the `N`th
series in reading order and reusing the source-preserving editor; the result
is byte-identical to the same edit through `/chart[1]/series[N]`. A number
past the last series is `out_of_range`. `series{N}.color` stays add-only
(`unsupported_property` on `set`): the chart editor has no series-fill edit.

## 3. Trunk regression fixed: XLSX chart `add`

On trunk `fd49666493`, every XLSX chart `add` failed with
`unsupported_element` ("only sheets can"): record 0269's kind-first check
(finding X1) ran before record 0270's chart branch in
`format/xlsx/structure.rs`, and the merge kept both. The test
`charts::xlsx_chart_add_get_set_remove_round_trip` failed on trunk, and the
corpus outcomes recorded 0 of 322 XLSX chart lines passing (record 0270
recorded 35). `new_kind` now admits `chart`, and the chart branch is chosen
from it. The 35 XLSX chart lines pass again.

## 4. Warn-mode failures carry what was skipped

`litchi_ops::Error` gains `warnings() -> &[Error]` (internal setter
`with_warnings`; `Inner` grows by one boxed slice, `Error` stays one
pointer). When an `Unsupported::Warn` command
(`File::apply_with`, `File::batch_with`) skips properties and the reduced
command then fails for another reason, the returned error is the reduced
command's own error with the skipped properties attached: each the
`unsupported_property` refusal, with property and item index, exactly as
`Applied::warnings` reports them on success.

Design choice. The alternatives were returning the first
`unsupported_property` refusal instead, or the reduced failure alone (the old
behaviour). The reduced failure alone hides the cause. The first refusal
alone hides a real, independent failure: in `set A1 sparkle=1 font.size=-3`
the invalid size is the caller's own error and must be reported. Keeping the
reduced command's code (so a caller can still act on the real failure) and
attaching the skipped properties (so nothing the caller sent is
misrepresented) is the only form that keeps both facts. Documented on
`Error::warnings`, `batch::Unsupported::Warn` and `File::apply_with`
(doctest: `sparkle` skipped, then `invalid_value` on `font.size`).

`Display` appends `(after skipping unsupported properties: a, b)`.

**CLI.** A failed write prints each skipped property as a
`warning[unsupported_property]` block before its `error[...]`; `--json`
puts them in the `ErrorResult`'s `warnings`; a failed batch item's JSON and
text line carry them too (`(after skipping: ...)`). The exit status is the
failure's own family (here 4). Test
`cli::unsupported_warn_failures_name_the_skipped_properties`.

**Corpus classification.** The harness records such a call with its code and
detail `<property>;skipped:<properties>` (the error's own property line is
read after the `error[` line, so a preceding warning cannot be mistaken for
it). `tools/officecli_corpus.py` classifies the call by its own failure
first (deltas, cascades, element refusals); a `missing_property` that no
other rule explains is the named gap **F3b** ("failed after skipping"), since
the skipped properties carried what is missing; every skipped property is
cross-checked against the registry like a warned one (refusal `skipped` in
the gap tables), so skipping a property the registry claims is still a bug,
and any other unexplained failure stays a bug. The table's bug rows show the
skipped list. Unit test
`test_failure_after_skipping_is_a_gap_not_a_bug`.

## 5. Corpus

`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, then `python3 tools/officecli_corpus.py` (and `--check`).
Script calls only, from the generated table:

| | Passed | Warned | Gap | Delta | Cascade | Bug |
|---|---:|---:|---:|---:|---:|---:|
| trunk `fd49666493` | 3,764 | 900 | 1,136 | 159 | 468 | 4 |
| this record | 3,803 | 1,105 | 891 | 168 | 464 | 0 |

Row by row (6,808 rows, identical order): no call that passed regressed.
Result changes: 4 PPTX dotted chart adds `missing_property` → ok; XLSX chart
adds (§3) `unsupported_element` → ok 35, warned 201, `unsupported_capability`
72 (69 extended or waterfall `chartType`, 3 `series1.values` cell ranges),
`invalid_value` 9 (delta D9, `labelPos`); 4 later `set` calls
`path_not_found` → warned (their chart now exists). 131 calls now fail after
skipping and list what they skipped (122 of them changed only their detail:
pictures whose `src` is missing, a `missing_property layout` batch, ...);
each is classified by its own failure as before. F3b matches 0 lines now that
the four chart adds pass. Remaining per-series gaps in the corpus:
cell-range values (3 XLSX lines, which also carry `.categories` ranges),
`.marker`, `.trendline*`, `.bubbleSize`, and series `color` on `set` (19
warned lines).

Registry and capability goldens re-blessed for `pptx`, `xlsx` and
`xlsx-encryption` (insertions only: the 60 keys per format), with
`--all-features` and with every feature but `encryption`;
`docs/unified-ops/officecli-capability-delta.md` regenerated (the keys are
`litchi-only`: OfficeCLI documents them in a note of `chart-series.json`,
not as chart properties).

## 6. Measurements

Harness `tools/perf-baseline/unified-chart-add`, new mode `--compare
dotted` (routes `pptx_chart_add_dotted_v1`, `xlsx_chart_add_dotted_v1`):
two unified legs on the same eager owner holding `N` charts, A = `data=` with
`colors=`, B = the same chart through the dotted keys; both legs publish
identical bytes (checked before timing). Release build
(`CARGO_NET_OFFLINE=true cargo build --release --manifest-path
tools/perf-baseline/unified-chart-add/Cargo.toml`, binary SHA-256
`fd8009fce3d4cce0de3daefcce2d1b43934cd008fc614570030d117735270115`, not
committed), `taskset -c 13`, 15 ABBA rounds (30 samples per leg), three
trials, load average about 8 on 32 cores. Raw:
[`results/0271/dotted-{1,2,3}.json`](results/0271/).

| Route | N | `data=` median (t1/t2/t3) | dotted median | dotted / data |
|---|---:|---|---|---|
| `pptx_chart_add_dotted_v1` | 0 | 884.9 / 891.5 / 880.6 µs | 886.9 / 892.9 / 883.4 µs | 1.000–1.003 |
| `pptx_chart_add_dotted_v1` | 10 | 1,528 / 1,534 / 1,532 µs | 1,528 / 1,534 / 1,533 µs | 1.000 |
| `pptx_chart_add_dotted_v1` | 100 | 6,239 / 6,267 / 6,239 µs | 6,237 / 6,270 / 6,227 µs | 0.998–1.000 |
| `xlsx_chart_add_dotted_v1` | 0 | 55.4 / 55.0 / 55.6 µs | 56.0 / 55.3 / 56.2 µs | 1.007–1.012 |
| `xlsx_chart_add_dotted_v1` | 10 | 123.9 / 122.8 / 124.5 µs | 124.9 / 123.9 / 125.9 µs | 1.008–1.011 |
| `xlsx_chart_add_dotted_v1` | 100 | 739.4 / 738.2 / 735.5 µs | 740.0 / 738.8 / 738.3 µs | 1.001–1.004 |

Within the 10% single-op budget everywhere (at most +1.2%, about 0.6 µs:
six keys admitted instead of two, and the per-number merge). The record 0270
route (unified vs native, one trial,
[`results/0271/native.json`](results/0271/native.json)) is unchanged:
PPTX 1.011 / 1.011 / 1.003, XLSX 1.055 / 1.021 / 1.007 at N = 0 / 10 / 100
(record 0270: 1.05 / 1.06 at XLSX N = 0).

## 7. Gaps and follow-ups

| Gap | Follow-up |
|---|---|
| Cell-range series (`series{N}.values=Sheet1!B2:B13`, `series{N}=<range>`, `series{N}.categories=<range>`), 3 corpus lines | per-series `Binding` references in `chart::authoring` (today one layout per chart); XLSX can read the cells as `dataRange` does |
| `series{N}.color` on `set`, and `color` on `/chart[N]/series[K]` (19 warned lines) | a series-fill verb on `chart::edit::Edit` |
| OfficeCLI's flat alias `series{N}Name` (undocumented, 0 corpus lines) | add as 20 registry aliases if a script needs it |
| `.marker`, `.trendline*`, `.bubbleSize` per series | extend `Series` alongside the record 0270 formatting gaps |
| A dotted name without values is refused; OfficeCLI writes an empty series | kept strict: an empty series is not a chart Litchi authors |
| Python does not surface `Error::warnings` (nor `Applied::warnings`) | record 0269 P3, the bindings stream |

## 8. Gates

Run with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-drawingml -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-drawingml --all-features` (410 passed) | 0 |
| `cargo test -p litchi-ops --all-features` (1,967 passed) | 0 |
| `cargo test -p litchi --all-features` (260 passed) | 0 |
| `cargo test -p litchi-cli --all-features` (27 passed, corpus included) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-drawingml -p litchi-ops -p litchi-cli` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (26 profiles, 1,201 bans, 38 families) | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_no_3rdparty_references tools.test_officecli_capability_delta` | 0 |

Python bindings are unchanged (they compile under the clippy gate).

## 9. Not claimed

- Rendering of the per-series colors in PowerPoint or Excel (the XML is the
  same `c:spPr` fill `colors` already writes).
- Any OfficeCLI behaviour for the per-series keys outside §2.
