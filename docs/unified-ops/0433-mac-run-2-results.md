# 0433 — macOS run 2: the failures explained, a read defect found, the queue refolded by bytes

- Status: merge-ready on branch `uo2/iwa-w6` (cut from trunk `feat/unified-ops-api`
  at `2c57b9a643`). **W3 stays held: 488 rows need run 3.**
- Date: 2026-11-04.
- Input: run 2 of the control kit (`.coord/mac-runbook/out2/`), 158 packages on Numbers,
  Pages and Keynote 14.4, macOS 26.5.2. It began at 23:55 +0800 on 2026-11-04 and crossed
  midnight, so the hand checks happened on 2026-11-05. The results are in
  `results/0433/run2-results.tsv`.
- `performance_claim`: each measured route stays within the +5% budget against the previous leaf; the largest change is the volatile rule on Numbers' family, +0.06% (§12).

## 1. Result

| Question | Answer | Status |
|---|---|---|
| Group A: is `count` the cause, and which rule is right? | Highest index + 1. The writer of 0431 is right; no change | **Proven** by the native file and the control run |
| Group B: is the missing listing the cause? | Yes. The fix stands | **Proven** by the control run. The other differences from a native caption remain |
| Group C: what does Numbers crash on? | The four referenced-tree objects, which hold a deliberate cycle | Object group **proven**; the cycle itself is **strongly indicated**, control in run 3 |
| Pivot routes on Numbers' own workbook | All four packages open, save and reopen; a body formula recalculates | **Proven** for those bytes; source recalculation and Refresh go to run 3 |
| First-rule highlight | The owner's note does not establish a failure or a pass | **Open.** One structural gap fixed; only the Mac can prove it |
| W1-F second native file | Delivered, read, and matched by the leaf | **Proven** |
| New: a by-column chart | The Keynote and Pages series routes refused Keynote's own default pie | **Defect fixed** (§2.1) |

## 2. The `count` rule (Group A)

`test-data/iwork/keynote/pie-third-wedge-native.key` decides it: **six wedges, one override
on the third, stored `count` 3 and index 2.** So `count` is neither the series count nor the
number of entries. It is the highest index + 1. `results/0433/sparse-arrays.tsv` lists every
populated array.

Across the 100 re-saved files (98 of run 2, P0420 and P0421 of run 1), populated arrays hold
(3, [2]) 20 times, (3, [0, 2]) 22 times and (2, [1]) 11 times; empty ones hold (0, []) 80 times.

| Input of a re-save | Apple's result | Count |
|---|---:|---:|
| (3, [2]) | preserved | 20 |
| (3, [0, 2]) | preserved | 13 |
| (2, [1]), two series | preserved | 8 |
| (2, [1]), **three series** | preserved: so not the series count | 3 |
| (2, [0, 2]), which run 1 accepted | **rewritten to (3, [0, 2])** | 7, and P0420 and P0421 too |

Every `A-*-7-min-only-count-changed`, `A-*-4-leaf-fixed` and `A-*-6-host-fixed` opens while
`A-*-2` and `-3` refuse. The owner's dialog for `A-P0422-3-leaf-current.key` reads "Keynote
couldn't read the file." The writer of 0431 stands, with
`leaf_writer_follows_the_native_count_rule` on Keynote's own chart. The reference-line change
of 0431 remains unexercised on a Mac: no package holds one.

### 2.1 The defect this file exposed

The leaf could not read this file. `slide_chart_series` returned `InvalidSource`. Keynote's
default pie is `series_direction` by column, one grid row with six columns, and the Keynote
and Pages routes counted series as grid rows, so index 2 exceeded a count of 1. The Numbers
route already followed the direction, as the host does. All 24 earlier charts in `test-data`
are by row, so no test could see it.

`ChartMetadataSnapshot::series_components()` gives the column labels of a by-column chart and
the grid rows otherwise; both routes use it. Tests: `leaf_reads_the_native_third_wedge`,
`by_column_series_are_the_column_labels`. Pages has no by-column fixture, so its side rests on
the shared codec. The public `series_count()` of chart metadata still reports grid rows, as
the host's does; that is a gap in §10.

## 3. Captions (Group B)

Every `B-*-7-min-only-reference-listed`, `-4`, `-6` and `-8` opens while the current ones
refuse: "B-P1141-3-leaf-current.numbers is damaged and can't be opened." The listing is the
cause. The leaf reads "Native caption" from `test-data/iwork/numbers/chart-caption-native.numbers`
(`native_chart_caption_reads_and_edits`), and replacing and clearing it keep references listed.

| Part | Numbers 14.4, authored | Leaf, fixed | After Numbers re-saves the leaf's |
|---|---|---|---|
| Caption 633, storage 2001, info 634: component | `CalculationEngine` | same | same |
| Caption info: version and field infos | [10, 1, 0], three | identical | same |
| Style 2025: identity | the template's `904468`, unedited | a **new** object with the same identifier string | both kept |
| Style 2025: overrides | 14, with an empty fill | 7, with a transparent fill | kept |
| Paragraph style | `904466`, "Caption Text", HelveticaNeue-Light | `904465`, **"Caption Title"**, HelveticaNeue | kept |
| Storage field 7, list style `903815` | present, listed | absent | **Numbers adds and lists it** |
| Payload | "Native caption" | the caller's text | same |
| External references | all listed, no unlisted | all listed, no unlisted | all listed |

So the version form that 0431 flagged is what Numbers writes, and the absent list style is
something Numbers repairs. **The caption's paragraph style is a real difference with a visible
effect:** the leaf's caption uses the title style, the host's choice, which parity copied. It
opens and survives, so it is not fixed here; correcting it breaks host parity on bytes every
caption row rests on. See §10.

What the control run says about the 19 run-1 packages with other unlisted references (`3056 ->
212` in eight, `6001 -> 6002/6005/6006` in twelve, one holding both; 0431 §3 counts 20 by kind): **nothing directly.** No package of run 2 holds
one, and only the `2001 -> 2022` kind was varied. It proves that kind is fatal; it does not prove
the others harmless beyond run 1's open, save and reopen. `3056 -> 212` also occurs in native
files; the `6001` kinds occur only in Litchi output and stay unexplained. No re-save of any
such package came back. Run 3 returns all re-saves, and these bytes are in it, so they can be
audited then. None of the 100 re-saves holds an unlisted reference.

## 4. The crash (Group C)

The `.ips` (not committed): `EXC_BAD_ACCESS`/`SIGBUS` writing the stack guard, on thread
`TSPFinalizeHandlerQueue.Access`, under `-[NSDocument initWithContentsOfURL:ofType:error:]`.
The report elides 1,652 frames: one `TSTables` function recurses to depth 1,658 of 1,702 on a
544 KiB stack. So it is unbounded recursion, not a validation failure. There are no symbols, so
the function's job is inferred.

| Package | Objects kept | Result |
|---|---|---|
| `C-P1096` | all 10 | crash |
| `C-7-without-inline-tree` | all but the inline group-by | crash |
| `C-7-no-cycle` | all but `9100020`–`23` | opens |
| `C-7-only-inline-tree` | the inline group-by | opens |
| `C-7-none-of-the-objects` | none: the kit's writer | opens, so the `7-*` results stand |
| `C-8-inline-only-{input,leaf,host}` | inline, with the formula route | opens |

**The crash follows the group `9100020`–`23`: aggregator, two nodes, group-by.** That group holds
the cycle, and recursion to the stack limit fits it. No package removes only the back edge, so
`W6-C-referenced-tree-no-cycle` in run 3 separates the cycle from the referenced tree.

The fixture came from a test, deliberately: `crates/litchi-iwa/tests/parity/numbers_pivot_cases.rs`,
"a referenced tree: aggregator and nodes as objects, with a cycle", written to pin the reader's
termination. No public route can write one: `child_ref` (field 10) is built nowhere outside
tests, in the leaf or the host, and Keynote and Pages refuse a populated group node.

The readers do not overflow: the leaf's walk and the host's are iterative with a visited set.
`pivot_group_tree_bounds.rs` proves it for the leaf on a 256 KiB stack, with a two-node cycle, a
self-reference and a repeated child, and with an inline tree 4,000 levels deep. The leaf is
safer on this input than Numbers.

Rows that rested on the fixture, traced by input hash through the run-1 manifest: **W4-H**
`pivot_categories` (P1096) and **W4-K** `set_formula`/`set_formula_cached` (P1142, P1143, P1144).
P0786 of W4-H is unaffected. Both rows now rest on the native workbook (§8).

## 5. The pivot on Numbers' own workbook

**No PDFs came back, so this is re-saved bytes only: nothing here shows what was displayed.**
`tools/mac_queue` regenerates all four packages byte for byte through
`crates/litchi-numbers/examples/mac_queue_native_routes.rs`. The leaf opens all four re-saves and
reads 169 categories from each.

| Package | Sent | After Numbers re-saved it | Reading |
|---|---|---|---|
| `W5N-pivot-native-control` | unedited | every tile identical | positive control |
| `W5N-pivot-source-cell-edit` | C2 = 55 | **C2 = 55**; hidden snapshot C2 = 10; pivot tile unchanged | the edit survives, and the pivot does not recompute on open or save (0432 §3.1) |
| `W5N-pivot-category-formula-on-source` | C2 = Grand Total Average / 2, cache 32.5 | formula bytes unchanged, **32.5**, cell format rewritten | accepted. **32.5 was the sent cache**, so this alone does not prove a recalculation |
| `W5N-pivot-category-formula-on-pivot-body` | B5 = West Maximum, cache 1 | formula bytes unchanged, **90** | accepted, not rewritten, and **recalculated**: 1 became the category's value |

**0432 §3.1:** the structure's prediction holds on open and save. Refresh was not exercised; it is
the hand step of run 3.

**0432 §3.2, resolved as far as bytes go.** Across all four re-saves, every owner's
`uuid_references`, read as a set, equals what was sent: **Numbers neither adds the new formula nor
removes the replaced one**, and the workbook's `uuid_reference_map` and `group_node_map` stay
empty. The body formula recalculated regardless, so open, recalculate and save do not need the
entries, and leaving them out is no defect there. What stays open:

- whether a source formula recalculates: `P1150` carries a stale cache of 1 and `P1148` none;
- what a Refresh or a regrouping does to a formula not listed. If those rewrites follow the list,
  an unlisted formula could go stale. Nothing here exercises that.

## 6. The first-rule highlight

**The owner's note does not establish what happened.** It reads "`W5N-volatile-first-rule-numbers-saved`
neither of them are highlighted". The file holds no dates: B2 holds text and the rule is on B2. Our
own RUN.md said all `W5N-volatile-*` files hold A2, A3 and A4 dates, and asked to retype them; the
row above asked "which of A2 and A4". If the owner typed dates in A2 and A4, no highlight is the right
answer, since those cells have no rule. If the owner typed today's date in B2, it is a failure. The
note does not name B2, and the kit's wording caused the doubt. It is treated as a failure until shown
otherwise: **the route has no proof.**

Field by field, on `basic.numbers`, against Numbers' files and the `V-*` packages the owner saw highlight:

| Item | `volatile-highlight-native` | `W5N-…-first-rule` | Its re-save | `V-P0389` (highlighted) |
|---|---|---|---|---|
| Rule predicate AST | | identical, including the table UUID | same | identical but for its own UUID |
| Rule style | preset "Red Fill" with a text colour | red fill, empty text style | same | **identical** |
| Cell binding | `cond` 1, applied 0 or 15 | `cond` 1, applied 15 | same | `cond` 1, applied 15 |
| Volatile set, tile, record edges | one range | same shape, one cell | same | same shape |
| The other owners' records | | byte-identical to `basic` | | host's shape |
| Rule's styles live in | `ObjectContainer` | the list's `DataList` | **left there** | `Document`, moved by Numbers |
| **`uuid_reference_map`** | **lists the rule's cells** | **empty** | **stays empty** | empty, in an engine with no other state |
| Table owner's tile | empty tile | none | **Numbers adds one** | none |

The checks asked for, one by one:

- **Did Numbers keep the rule?** Yes: the re-save reads the same rule on B2.
- **A style rendering nothing?** No: the rule and style equal `V-P0389`'s, which highlights.
- **B2's cell style and format:** B2 holds text with format kind 5 and no explicit-format marker, so it
  should parse a typed date; no bytes show it did. `V-*` were empty, so overwriting text is untested.
- **Volatile set, tile, owner shape:** the same as native.

**One real gap: Numbers lists a rule's cells in the engine's `uuid_reference_map`, under the table's UUID
and the conditional-style owner, and does not repair an empty map on re-save.** The leaf left it
empty. `V-*` show it may not matter, but their engines hold nothing else, so they may load differently.
This is a difference from native in bytes, not a proven cause.

The fix: `conditional_volatile` maintains the map when it adds or removes a cell, on engines that hold the
field and only for Numbers. Host-built documents and the Keynote and Pages routes keep the host's shape, so
parity is unaffected: the parity run (1192 passed) holds the same bytes. `W5N-volatile-native-add-tomorrow`
and `-clear-middle` change in that one entry, so they are new bytes too.

**This is not claimed as a success.** Nothing on Linux evaluates a rule. Run 3 decides it twice:

- **With nobody looking.** Numbers evaluates volatile rules on open and stores the applied rule in the cell.
  Run 2 shows it: the native control re-saved at 00:00:42 on 11-05 moved applied rule 0 from A2 to A4, the cell
  holding 11-05. `P0392` holds six weeks of dates, each with the first-rule path's rule, so its re-save names
  the cell Numbers took for today. `W6-first-rule-date-ladder-no-map` is the same without the map.
- **By hand, saved.** B2 with and without the map, an empty cell, and a native control, each typed and saved,
  with the fill reported before and after. The saved bytes hold the date and the applied rule, so the result
  no longer rests on a note.

The other two files, read from their re-saves on 11-05:

- **`W5N-volatile-native-clear-middle`** (owner: A2 highlighted): A3's rule is gone and its date remains; RUN.md's
  "A3 empty" was wrong. The re-save has applied rule 0 on A4, the date then equal to today, and 15 on A2. The owner's
  A2 fits retyped dates as RUN.md asked; no saved copy shows it.
- **`W5N-volatile-native-add-tomorrow`**: A4's leaf-written "is tomorrow" survives, with applied 15. A4 held today by
  then, so 15 is right, and it proves no evaluation.
- **`W5N-volatile-native-control` and `W5N-volatile-native-add-tomorrow`: the owner's note does not report them.** The
  control's evaluation above comes from its bytes.

The owner reports "Yes to V-* files": `V-P0386` to `V-P0389` highlighted as RUN.md expected.

## 7. W1-F, the second native file

`volatile-highlight-native-second.numbers`, re-saved by Numbers on 11-05:

| Cell | Date | Rules | Applied |
|---|---|---|---|
| A2 | 2026-11-06, tomorrow | "is today" Red Fill, then "is tomorrow" Orange Fill | second |
| A3 | 2026-11-03 | "is today" | none |
| A4 | 2026-11-05, today | "is today" | first |

The kit asked for the second rule on a different cell; the owner added it to A2. The shape: a new rule set,
its own four style objects, a new list key for A2, and the first set's use count down to 2. The volatile set,
the owner's records and the map do not change, because A2 was already registered. The formula count stays 8.

`leaf_second_rule_matches_what_numbers_wrote` writes the two rules on A2 of the first file: both predicates equal
Numbers' and the volatile state matches. The one difference is where the styles live (§6).

The route text of the W1-F row now says Numbers-saved tables are written, and a new row holds that route (§8).
The native-created row passes on the two files and the owner's "Yes to V-* files". It no longer stands for the
first rule on a Numbers-saved table.

## 8. Queue

Regenerated from `d3f8b897ac` with `tools/mac_queue/generate.py` (1192 passed, 210 ignored; 1,160 packages,
manifest SHA-256 `6d59fef38bb0900568c337743bcde98aad0169d4bbf55e8a6a3370f7b98d045a`), kept in
`.coord/iwa-w6/queue-out/`. `apply_results.py --evidence` now folds runs **by package SHA-256** and rewrites every
packaged row, with no carry-over from another byte string. The later commits change tests, documents and formatting; the example's packages rewrite to the same hashes.

| Status | Rows | Evidence |
|---|---:|---|
| pass, run 1 bytes | 271 | `results/0430/results.tsv` |
| pass, run 1 and run 2 bytes | 5 | and `results/0433/run2-results.tsv` |
| pass, run 2 bytes | 1 | W4-H, the native control |
| pass, native authoring | 2 | W4-K `Aggregate` (0432), the W1-F native-created row (§7) |
| **FAIL** | **0** | no regenerated package's bytes failed |
| **pending, run 3** | **488** | 757 packages |

**This is more open rows than run 1 left, and the reason is mostly not a defect.** Parity inputs that the host
builds draw random UUIDs on every generation, so 743 of the 757 packages differ from what ran. 0431 §8 compared
such rebuilds and found they differ only in those UUIDs, but this row-by-row fold does not assume it: the bytes
did not run, so they run again. Groups A, B and D are in here, since their fixed packages are host-built too.
The rows of Groups A and B are no longer FAIL because those bytes are gone, not because they have passed.

The pivot rows now say what they prove:

| Row | Now rests on | Status |
|---|---|---|
| W4-H `pivot_categories` | The native workbook's bytes, with the leaf's read tests. No Litchi write is involved | **pass**, run 2 |
| W4-K `set_formula`, `set_formula_cached` | Native: source cached, body (both passed), source stale cache, source no cache | **pending**, run 3 |
| W4-K cell edits with a pivot owner | Its earlier packages, plus the native source edit | **pass**, runs 1 and 2 |

`numbers_pivot_*_parity` on the synthetic fixture stay Linux proofs of the code paths. **A category on a tree stored
only by reference has no Mac proof:** no Numbers file holds one, and the row says so.

## 9. Run 3

`.coord/iwa-w6/out3/` (not committed), SHA-256 of its tar: `ae11ba82985ecd8b6b7e39a37885eb9258c13baa282205ae7c4d18257ff5daaa` (`.coord/iwa-w6/out3.tgz`). `RUN.md` is the owner's page; `README.md` says
what each outcome means. 762 packages, with every hash verified by `--dry-run` on Linux:

- 757 queue packages (Keynote 206, Numbers 277, Pages 274);
- `W6-first-rule-date-ladder-no-map`, `W6-C-referenced-tree-no-cycle`;
- `W6-first-rule-text-cell-no-map`, `W6-native-control`, `W6-pivot-source-cell-edit`, for the hand steps.

Hand steps, with the saved copies returned: type today's date in `P0394` B2, the no-map B2, `P0391` B4 and the
native control's A2, reporting the fill before and after; press Refresh on the pivot source edit and report the
value before and after. Fold with
`apply_results.py --manifest queue-out/manifest.json --evidence … --evidence … --evidence "run 3=…"`.

## 10. Gaps

- The first rule on a Numbers-saved table has no proof (§6). The map fix is unproven; it may not be the cause.
- The caption's paragraph style is the title's, as on the host (§3). Fixing it needs an owner's call on host parity.
- The leaf's new style objects share an identifier string with the template's (§3). Numbers keeps both.
- The rule's styles live in the list's `DataList`, not `ObjectContainer` (§6). Numbers leaves them.
- `series_count()` of chart metadata counts grid rows on a by-column chart (§2.1).
- Refresh and regrouping with unlisted category formulas (§5).
- `6001 -> 6002/6005/6006` unlisted references are unexplained (§3).
- Which part of the referenced tree Numbers recurses on (§4), and the crashing function's name.
- Keynote and Pages still refuse the first volatile rule, as the host does; only Numbers has native evidence.

## 11. Not verified

- **No displayed result was checked by me.** Everything Mac-side is bytes, `results.tsv`, or the owner's note.
- No run 3 result exists. The fix, the ladder idea, and the Refresh step's menu wording have not met a Mac.
- That the 743 regenerated packages differ only by host UUIDs: not checked here, and not relied on.
- The crash report has no symbols.
- Pages' by-column read rests on the shared codec.

## 12. Performance

Callgrind, release with debug lines, CPU 5, ABBA, inclusive Ir of a one-route probe over the same input. Base is the
four changed source files at `2c57b9a643` in this tree. Probes and binaries are in `.coord/iwa-w6/perf/`, not committed;
SHA-256 of the binaries: Numbers base `86b681b7…`, new `c903b487…`; Keynote base `9c16298b…`, new `a6437f57…`.

| Route | Base Ir (A, A) | Branch Ir (B, B) | Change |
|---|---:|---:|---:|
| `edit_table_cell_conditional_highlights`, a new cell in Numbers' family | 40,063,384; 40,058,653 | 40,088,026; 40,085,261 | +0.06% |
| the same, first rule on `basic.numbers` | 39,503,509; 39,509,947 | 39,516,487; 39,537,082 | +0.05% |
| Keynote `slide_chart_series`, Fills, by-row chart | 1,733,249; 1,733,073 | 1,733,236; 1,733,110 | +0.00% |

The budget is +5%. The Pages series route takes the same one-line change and was not measured. The caption route, and every
route whose packages only went to run 3 because of host UUIDs, has no code change. These are one-commit probes, not
the unified-against-leaf comparison: no unified route changed.

## 13. Gates

Logs are in `.coord/iwa-w6/gates/`. All exit 0. `TMPDIR` was on disk (`.coord/iwa-w6/tmp`), not the named `/tmp` path:
records 0431 and 0432 lost runs to its quota.

| Gate | Exit |
|---|---:|
| `cargo fmt --all --check` (first run found my formatting; rerun after `cargo fmt`) | 0 |
| `cargo clippy` of `litchi-iwa-protos`, `-keynote`, `-numbers-wire`, `-numbers`, `-pages`, `litchi-ops`, `litchi`, `litchi-py`, all targets and features (first run found one `clone`; rerun after the fix) | 0 |
| `cargo clippy -p litchi-ops --no-default-features`: none, numbers, pages, keynote | 0 |
| `cargo test --all-features`: `litchi-iwa-protos`, `-numbers-wire`, `-iwa-table`, `-iwa-style`, `-numbers`, `-keynote`, `-pages` | 0 |
| `cargo test -p litchi-iwa --all-features --test parity`, during generation | 0: 1192 passed, 210 ignored |
| `cargo test -p litchi-iwa --all-features --lib` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, every `--test` in chunks of 28 | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc` of the changed crates | 0 |
| `check_crate_boundaries`, `non_iwork_gate verify`, glossary (4 profiles), `check_iwork_leaf_public_api`, `iwa_exit_ledger --check`, `generate.py --check`, `check_example_targets` | 0 |
| Tool unit tests: the rules' modules and `test_mac_queue`, `test_mac_control_kit`, `test_gate_runner`, `test_iwa_exit_ledger`, `test_check_example_targets` | 0 |
| `officecli_capability_delta --check`, `officecli_corpus --check`, `check_feature_matrix` | 0 |
| `tools/unified_bench_matrix.py --check`, after regenerating for this record | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 |

Re-blessed for the three new fixtures only: `query_corpus` digests (Numbers and Keynote queries and no other format) and
`results/0289/fidelity.{md,tsv}` (three new rows). No capability changed, no capability golden was re-blessed, and no
dependency was added. The two clippy-and-format fixes after the first full pass touch a test and whitespace; the crate tests of
`litchi-numbers` touched by them were rerun (`pivot_group_tree_bounds`, `native_volatile_highlight`), and the example's
packages were rewritten and hash identical to the generated ones. The full suites were not rerun after those two fixes.
