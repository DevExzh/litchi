# 0431 — Diagnosis of the 15 macOS queue failures, two host defects fixed, control kit

- Status: merge-ready on branch `uo2/iwa-w5` (cut from trunk
  `feat/unified-ops-api` at `52f3d60714`). **No queue row changes.** A row
  passes only on a Mac result, and W3 stays held.
- Date: 2026-11-04.
- Input: queue run 1 (`results/0430/results.tsv`; Keynote, Pages and Numbers
  14.4 on macOS 26.5.2): 1,138 packages passed, 15 failed twice. Packages and raw
  dumps are in `.coord/mac-runbook/out/`, which this record did not change.
- `performance_claim`: the fixed routes stay well below the host (§5).

## 1. Result

| Group | Packages | Class | Cause | Fix | Status |
|---|---|---|---|---|---|
| A, pie charts | P0422 P0427 P0971 P0425 P0623 P0428 P0928 | (b) host defect the leaf copied | `series_non_styles.count` holds the number of entries, so an entry can sit beyond the array | `count` is one past the highest index | Isolated by evidence; **causation needs the control run** |
| B, Numbers chart captions | P1126 P1128 P1141 P1148 | (b) host defect the leaf copied | The caption storage names a paragraph style in another component with no `external_references` entry | The entry is written | Strong correlation only; **causation needs the control run** |
| C, synthetic pivot | P1096 P1142 P1143 P1144 | (c) invalid test input | P1096 has no Litchi edit at all; the other three start from the same bytes | None to the route; a reduced fixture for the control run | **P1096 proven.** Which part Numbers rejects, and whether the route works, need the Mac |

The leaf equals the host in every failing case of groups A and B, so the leaf
did not add either defect: it copied them faithfully, as parity required. Parity
could not have caught them. It compares the leaf with the host, and the host was
wrong.

## 2. Group A: the sparse array count

### Evidence

`results/0431/discriminators.tsv` lists every package with a populated array,
every package with an unlisted reference, and every failure. Across all 1,153
packages, `series_non_styles` populated arrays split cleanly:

| `count` | Entry indexes | Mac pass | Mac fail |
|---:|---|---:|---:|
| 1 | 0 | 48 | 0 |
| 1 | 1 | 11 | 0 |
| 2 | 0, 1 | 23 | 0 |
| 2 | 0, 2 | 13 | 0 |
| 3 | 0, 1, 2 | 12 | 0 |
| **1** | **2** | **0** | **7** |

These 7 are exactly Group A, on all three apps. They are the only packages with an
index greater than `count`. A wedge explosion and a leader line fail alike, so the
property does not matter.

What is special in `~3`: it writes one series only (`Some(2)`). `~1` writes
series 0 and 2, giving `count` 2 and indexes 0, 2, which opened. `~2` releases
everything. P0427, the partial release, starts from indexes 0, 2 and releases 0,
leaving index 2 alone.

The rule the table suggests is "index <= count opens". That would fit an array of
`count` slots filled by an append-tolerant setter, but this is an inference: no
Apple source was read.

Host and leaf output both hold `count` 1 and index 2 in all 7 cases. The inputs do
not carry the defect (`count` 0, or indexes 0, 2 for P0427).

### Native evidence, and its limit

**No tracked fixture holds a populated series array.** Across the 126 packages
under `test-data/iwork`, no object has type 5029 (`ChartSeriesNonStyleArchive`), a
pattern scan of object payloads finds no populated sparse array, and the 23 charts
all store `count` 0. `chart-arrangement-native.numbers` has 2 series and stores 0,
so native `count` is not the series count when no override exists.

So what Apple writes for a populated array is **unknown here**. One past the
highest index is the smallest value that satisfies the observed rule and agrees
with `count` 0 for an empty array. The control run returns Apple's own re-saves,
and the owner's hand-authored `native-pie-third-wedge.key` shows it directly.

### Fix (`8a8c15791a`)

- Writers: `chart_style_codec::rewrite_series_non_styles` (the leaf's shared
  engine) and the host's `series_non_style.rs` write `count` = highest index + 1,
  or 0 when empty.
- Readers: host and leaf refused any array whose `count` differed from its number
  of entries, with eight such checks across `series_non_styles`,
  `series_private_styles` and reference-line styles. These are removed. A native
  array with `count` above its entry count would have been refused by Litchi. This
  is a real read bug only if Apple writes such arrays, which is not yet known. The
  bounds check against the series count stays.
- Reference-line styles had the same writer shape (`count` = entries kept) on both
  sides and get the same change. **No package exercised this**, so that part rests
  on the array type being shared, not on a Mac result.
- Tests: `sparse_count_is_one_past_the_highest_series_index`
  (`litchi-iwa-style`) and `series_non_styles_match_the_generated_encoder`
  (`litchi-iwa-protos`).

### Consequence for packages that already passed

The fix changes the bytes of the 24 packages with `1:1` or `2:0,2`. They passed in
run 1 with the old `count`, and they back 11 queue rows, 4 of which already FAIL.
Run 1 does not cover the new bytes, so all 24 are in the kit as group D. The rows
are left as they are; the coordinator decides how to record them after the run.

## 3. Group B: the unlisted paragraph style

### Evidence

The chart caption graph is written beside the chart in the calculation engine's
component. Its storage (type 2001) references caption paragraph style 904465 (type
2022) in `DocumentStylesheet`. The route listed the new caption style in
`external_references` but not this one.

Auditing every package for references to a non-root object of another component
that the owner's component does not list:

| Unlisted reference type | Mac pass | Mac fail |
|---|---:|---:|
| `2001 -> 2022` (storage to paragraph style) | 0 | **4** |
| `3056 -> 212` (comment storage to author) | 8 | 0 |
| `6001 -> 6002/6005/6006` (duplicated table to its data) | 12 | 0 |

The audit covers object references recorded in archive metadata. It finds:

- all 4 Group B packages have the `2001 -> 2022` reference, and no other package does;
- host output carries it in all 4;
- the inputs of P1126, P1128 and P1148 carry it too, because they are host-captioned;
- the input of P1141 (no caption) does not.

Native files: across the 126 fixture packages, the only non-root references with no
entry are `3056 -> 212` (3 times) and `401 -> 2003` (2 times, in a source-built file).

### Field-by-field against the native chart caption

No Numbers fixture holds a chart caption, but `keynote/chart-caption-native.key` does,
saved by Keynote. Compared with P1141's graph (`(deleted-coord-artifact)`, `our.txt`):

| Part | Native Keynote | P1141 | Assessment |
|---|---|---|---|
| Caption info 633: version, three field infos, placement anchors | | identical | not the cause |
| Style 2025 in `DocumentStylesheet`, listed from the drawable's component | yes | yes | not the cause |
| Storage 2001 to paragraph style 2022 in `DocumentStylesheet` | **listed** in the slide's component | **not listed** | **the defect fixed here** |
| Storage field 7: a list style (2023), also listed | present | absent | not the cause by itself: see below |
| Style 2025 payload: override count 14 against 7, and some text properties | | differs | not the cause by itself: see below |

Checked for the storage row: the passing Numbers shape and image captions (P1130,
P1125 sources) have the same storage fields as P1141 (2, 3, 5, 6, 10, 14, 19, 24, 28,
so no field 7). The style row is the same authored payload (0427 §1), not separately
compared. Both are real differences from native and stay open; they do not separate
the failures from the passes.

One difference does separate them and is not excluded by that check: the failing
caption info uses version [10, 1, 0] with three field infos, and the passing Numbers
inline ones use [1, 0, 5] with none. The failing form matches the native Keynote
caption exactly, so it is not a defect by that evidence, but no Numbers file shows it
is accepted. This is the first thing to examine if the `7-min` control stays refused.

### Limit of this evidence

**A missing entry is not fatal in general**: 20 passing packages have other unlisted
references. The `2001 -> 2022` type is unique to the failures, but so is any other
property shared only by a Numbers caption on a chart in that component. No Numbers
native file holds one, and the Keynote one lives in a slide's component, so it shows
the listing rule but not that component's behaviour.

`B-*-7-min-only-reference-listed` decides this: it is the failing package with that
one metadata entry added. If it stays refused, this fix is correct but insufficient,
and the next evidence is the owner's `native-chart-caption.numbers`.

### Fix (`9a49714dda`)

`caption::write` (shared engine) and the host's
`insert_componentized_caption_style` list the paragraph style when it lives in
another component. Test `crates/litchi-numbers/tests/chart_caption_native_structure.rs`
runs the audit on `basic.numbers` and on `chart-data-native.numbers` as the native
baseline, then on set, replace and clear of a caption on the native chart. With the
fix disabled it fails with `[(907134, 904465)]`.

The inline layout has the same reference and is unchanged, because the audit finds
no unlisted case among the packages run. A component without that entry would hit
this; see §7.

## 4. Group C: the synthetic pivot fixture

### Evidence

- **P1096 is the input.** The case is read-only, and input, host output and leaf
  output are byte-identical (`8ea12ff94a38…`). Numbers dropped the AppleEvent
  connection (-609) on a file no Litchi write route produced. It was built by the
  test (`numbers_pivot_cases.rs`), by adding objects to `basic.numbers`.
- P1142, P1143 and P1144 start from those same bytes and fail the same way.
- A different cause for those three is not excluded, because they fail at the stage
  the input already fails at. **They say nothing about the route, good or bad.**

### Against a real pivot workbook

The owner authored `pivot-all-summary-functions-native.numbers` (Numbers 14.4; held in
`.coord/mac-runbook/native-authoring-in/`, ingested by record 0432, not committed here).
Its enabled group-bys 909202 and 909203 against the fixture's:

| Property | Native | Synthetic fixture |
|---|---|---|
| Who references the group-by | a `6372` owner, referenced by a table model | **nothing** references any of the five |
| `group_column` (field 2) | 2 and 1 | **absent** on all, though four are enabled |
| Fields 7, 9 to 13, 15, 16 (formula coordinates and related) | present | **absent** |
| Object references and field infos in archive metadata | recorded (2 and 2) | **none**, though the payloads hold references |
| Group node tree | a tree: no cycle, no repeated child | **`9100021` and `9100022` cycle; `9100022` listed twice**, both deliberate |
| `group_by_uid` | unique | **`9100010` and `9100040` repeat one**, deliberately |
| `owner_index` (field 14) | 205 and 206 enabled; 8 on the disabled category group-bys | 8 or 41, enabled |
| UUID map entries | none (0 of 21 pivot objects) | none: **not a difference** |
| Inline tree in field 3 alongside fields 17, 18 | yes | one or the other: inline tree alone is not native either |

So the fixture is structurally unlike anything Numbers saves, in many ways at once. That
supports (c), and P1096's bytes prove it needs no Litchi edit. **Which difference makes
Numbers quit is still unknown.** `last_object_identifier` alone is excluded: P1142 to
P1144 have it raised by the write and still fail. P1146, which also adds a synthetic
object but references it from the table model, opens, so a synthetic addition is not
fatal in itself.

Classification (c) is about the input. It does not prove the routes. The native workbook
opening shows nothing about Litchi's read or formula-author output; a package made by
those routes **from that workbook** is what can prove them (record 0432). The fixture
stays as it is for Linux reader tests, where its cycle is the point.

### Added

`numbers_pivot_formula_author_inline_only_parity`, on a fixture with only the inline
tree (no cycle, no repeated identifier), passes against the host. It exists to
separate the fixture from the route on a Mac. The native comparison above says this
reduced fixture is still far from native (no owner, no group columns), so expect it may
also be refused; a refusal there would not count against the route.

## 5. Performance

Callgrind, release, CPU 3, inclusive Ir of the probe's route functions, two calls
each (ABBA); binary SHA-256
`12de80267be250745336c83a114d9fe12d1f4b3511a7788ac14eb2f170e83215`; captures in
`(deleted-coord-artifact)`.

| Route | Host Ir | Leaf Ir | Leaf/host | 0427 leaf Ir |
|---|---:|---:|---:|---:|
| Numbers `chart_caption_create` | 400,158,927 | 142,666,387 | 0.357 | 133,027,390 |
| Series graph write, all formats (`series_graph_write_cost_abba`) | 15,084,031,325 | 2,837,880,887 | 0.188 | not compared |

The leaf route is about 7% above its 0427 figure and the host about 15%. That
comparison crosses other merges and no paired before-and-after run was made on one
commit, so the rise is **not attributed** to this fix, which adds one metadata
update. Leaf against host is well inside the bar (leaf <= 1.05 x host). The series
change is one maximum over the entries, with checks removed.

## 6. Control kit

`.coord/mac-runbook/out2/` (outside the repository). The coordinator tars it as
`out2.tgz` after record 0432's additions; that tar's SHA-256 is what binds the Mac run,
and it is not recorded here because it does not exist yet.

- 114 packages, in the `out/` layout and run by the same `run_on_mac.sh`
  (`--dry-run` verified all hashes on Linux). `RUN.md` is the owner's one page,
  `README.md` says what each outcome means.
- Each failing package has its input, host output and current leaf output, its fixed
  leaf and host output where a fix exists, and a passing sibling. It adds the
  one-change variants, the group C isolation set, and group D.
- `tools/mac_queue/control_kit.py` builds it (`build`), takes a second package list
  (`add`, for record 0432), audits a file (`audit`) and compares Mac results with expectations (`analyze`, which writes nothing). Its
  package writer is tested on tracked files. `C-7-none-of-the-objects` is the Mac
  control for that writer: if it does not open, every `7-*` result is void.
- Regenerate:
  `python3 tools/mac_queue/control_kit.py build --queue-out .coord/mac-runbook/out --fixed <parity dump> --native test-data --out <new dir>`,
  where the dump comes from
  `LITCHI_MAC_QUEUE_DUMP=<dir> cargo test -p litchi-iwa --all-features --test parity`.
  Host-built inputs draw random UUIDs, so a regenerated kit differs; the Mac result
  binds to the final tar's SHA-256.

## 7. Proven, and not yet proven

| Statement | Status |
|---|---|
| The 7 Group A packages are the only ones with an index above `count` | Proven on all 1,153 packages |
| `count` alone caused Group A | **Needs** `A-*-7-min-only-count-changed` |
| The fixed Group A output opens | **Needs** `A-*-4-leaf-fixed` |
| One past the highest index is what Apple writes | **Unknown**; needs a re-save or the hand-authored file |
| `2001 -> 2022` is unique to the 4 Group B packages | Proven on all 1,153 packages |
| That reference caused Group B | **Needs** `B-*-7-min-only-reference-listed` |
| The fixed Group B output opens | **Needs** `B-*-4-leaf-fixed` |
| Numbers rejected the synthetic fixture with no Litchi edit | Proven (P1096 bytes equal its input) |
| What in the fixture it rejects | **Needs** `C-7-*` and the `.ips` report |
| The pivot read and formula routes work natively | **Not proven by anything here**; needs route output from the native workbook (0432) |
| The apps' dialog text | **Unknown**; the runner cannot see dialogs. `RUN.md` §2 |
| The inline caption layout is safe in every component | **Not proven**: clean on the packages run, not fixed generally |
| The reference-line count change | **Not exercised** on a Mac |

Follow-ups after the control run:

1. Adjust the count writer if Apple writes something else.
2. Decide how the 11 rows behind group D are recorded.
3. Regenerate only the affected queue packages against a kept `out/`; the W3 input
   snapshot depends on those inputs.
4. Re-point the pivot rows at the native workbook (0432).

## 8. Gates

Logs are in `(deleted-coord-artifact)` (before merging trunk) and `gates2/` (after merging
`72ebb6e117`, record 0432). All exit 0.

| Gate | Before merge | After merge |
|---|---:|---:|
| `cargo fmt --all --check` | 0 | 0 |
| `cargo clippy` of the changed crates, `litchi-ops`, `litchi`, `litchi-py`, all targets and features | 0 | 0 |
| `cargo clippy -p litchi-ops --no-default-features`: none, numbers, keynote, pages | 0 | numbers only, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc` of the changed crates | 0 | not rerun |
| `cargo test --all-features`: `litchi-iwa-protos`, `litchi-keynote`, `litchi-pages` | 0 | not rerun |
| `cargo test --all-features`: `litchi-iwa-style`, `litchi-numbers` | 0 | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 | `--test parity` only, 0 (1192 passed, 210 ignored) |
| `cargo test -p litchi --all-features` | 0, see below | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, every `--test` in groups of 20 | 0, see below | `--lib`, `capabilities`, `registry`: 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | not run | 0 |
| `check_crate_boundaries`, `non_iwork_gate verify`, `check_iwork_leaf_public_api`, `iwa_exit_ledger --check`, `generate.py --check`, `unified_bench_matrix --check`, `officecli_corpus --check`, glossary | 0 | 0 |
| `officecli_capability_delta --check`, `check_feature_matrix` | 0 | not rerun |
| Tool unit tests | 0 | the mac, gate and 3rdparty modules, 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0, see below | not rerun |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 | not rerun |

The first run of the `litchi` tests, the `litchi-ops` doctests and the pyo3 tests failed
at the linker or `cp` with a `/tmp` disk-quota error, not in a test. All three passed
when rerun unchanged with `TMPDIR` on disk (`gates/rerun.txt`), which departs from the
`TMPDIR` the rules name. No capability or registry golden was re-blessed: no capability
changed, and those tests pass unblessed.

The merge changed no package of the kit. The fixed packages rebuilt on the merged code
differ only in the host's random UUIDs and metadata lengths (48 same form, 2
byte-identical), so `out2` was not rebuilt. 0432's 8 packages were not regenerated: they
edit cells and highlights, routes this record does not touch, and none has a populated
sparse array or an unlisted reference.
