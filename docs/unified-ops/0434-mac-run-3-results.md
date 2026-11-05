# 0434 — macOS run 3: the queue is folded, the first rule fires, three rows name an open part

- Status: merge-ready on branch `uo2/iwa-w7` (cut from trunk `feat/unified-ops-api` at `b02c3e8ae9`).
  **No row is pending and none failed. Three rows carry an open part, which is the coordinator's
  call before W3 (§10).**
- Date: 2026-11-05.
- Input: run 3 of the kit of record 0433, 762 packages on Numbers, Pages and Keynote 14.4,
  macOS 26.5.2, from 05:24Z to 05:51Z on 2026-11-05 (13:24 to 13:51 in Asia/Shanghai, the time zone in
  the files). Results are in `results/0434/run3-results.tsv`; the owner's copies are listed in
  `results/0434/hand-copies.tsv`. `out3-results.tgz` (117,484,769 bytes, SHA-256
  `a597f5515812a00deee3a5fd07a0fa50b3d7bb131b1fda278c308b08fe11898d`) and `hand3.tgz` (353,561 bytes, SHA-256
  `d67849541c5691dac9e79f7e37f2e258c850fdb296641d7e35f804e5cb590c50`) are not committed: they are in
  `.coord/mac-runbook/results3-in/`.
- `performance_claim`: none. No code changed and nothing was timed.

No Rust, generator or kit file changed. The only tracked changes are this record, its evidence, the
queue's Result and Artifact cells, and the bench-matrix exemption.

## 1. Result

| Question | Answer | Status |
|---|---|---|
| Does Numbers evaluate the leaf's first rule on a table it saved? | Yes: on open, with and without the map; and on a typed date, red and saved | **Proven** (§3) |
| Was `uuid_reference_map` the cause of the run-2 note? | No. The no-map copy, the same bytes as run 2's, turned red. Keeping the code is recommended; the coordinator decides (§3.3) | **Proven** |
| Does a source-table category formula recalculate? | Yes: stale 1 and none both became 32.5 | **Proven** (§4) |
| Does Refresh pick up a source edit? | Yes: 80 became 125. Every formula and `uuid_references` set is unchanged | **Proven**, on a package with no Litchi-written formula (§5) |
| Is the crash the cycle? | Yes. Removing one field, the back edge, turns the crash into an open | **Proven** (§6) |
| Does a re-save contradict the count rule? | No: 114 populated arrays, all highest index + 1 | **Proven** (§7) |
| What happens to the unlisted references? | `3056 -> 212`: Numbers keeps them unlisted, four of eight packages seen. `6001 -> 600x`: **not seen**, 0433 was wrong that run 3 held them | Part open (§7) |

## 2. The fold

```
apply_results.py --manifest .coord/iwa-w6/queue-out/manifest.json \
  --evidence "run 1=results/0430/results.tsv" --evidence "run 2=results/0433/run2-results.tsv" \
  --evidence "run 3=results/0434/run3-results.tsv" --record 0434 --next-run "run 4"
```

Manifest SHA-256 `6d59fef38bb0900568c337743bcde98aad0169d4bbf55e8a6a3370f7b98d045a`, as given. 762 rows
rechecked: 762 pass, 0 fail, no duplicate id, so no retry; every package hash equals `out3/queue.tsv` and the kit's
file; every re-saved file hashes to its row. The fold reports no conflict, updates 765 rows, and needs **0 packages**.

| Status | Rows | Evidence |
|---|---:|---|
| pass, run 1 bytes | 271 | `results/0430/results.tsv` |
| pass, run 3 bytes | 416 | `results/0434/run3-results.tsv` |
| pass, run 1 and run 3 bytes | 71 | both |
| pass, run 1 and run 2 bytes | 4 | |
| pass, run 2 bytes | 1 | W4-H `pivot_categories`, native control |
| pass, run 2 and run 3 bytes | 1 | W4-K formulas |
| pass, runs 1, 2 and 3 | 1 | W4-K cell edits with a pivot owner |
| pass, native authoring, not folded | 2 | W4-K `Aggregate` (0432), W1-F native-created (0433) |
| **FAIL** | **0** | |
| **pending** | **0** | |
| **Total** | **767** | 765 with packages, 2 without |

Every row not passing: none. **Three of the rows that say `pass` have a part that is open, and they say so (§10).**

Five rows have a note I added to the Result cell by hand, because the tool cannot express a hand
result. Those are the only hand edits; the Reason, Fixture and Artifact cells and the row ids are the tool's or
unchanged, so the manifest still matches:

| Row | Line | Note |
|---|---:|---|
| `q-e54e71bd4570` W1-F first rule | 262 | the rule fired, bytes and by hand |
| `q-9bb7c4ec065a` W1-B Pages identity | 304 | hand edit and save done in run 2, no conflict (correction below) |
| `q-c8ea770c712d` W4-H `pivot_categories` | 756 | **OPEN**: tree stored only by reference |
| `q-3080a227051a` W4-K formulas | 773 | 32.5 recalculated; **OPEN**: tree stored only by reference |
| `q-f9a5aaa1c1fa` W4-K cell edits | 774 | Refresh 80 to 125 |

Running `apply_results.py --evidence` again rewrites the Result cell of packaged rows and drops these notes; re-add
them from this table. No row cites "not reported": that wording was only in 0433 §6, and §3.4 closes it.

**Two Reason cells are now out of date.** W1-F's says "the owner's report did not show a highlight", and W4-K's says "the only such
fixture crashes Numbers", where §6 shows only the cycle crashes. They are identity cells, so changing them changes the row id and
needs a new manifest. I left them; this record supersedes both sentences.

## 3. The first rule

### 3.1 With nobody looking

Applied rule in the saved cell, 0 for the first rule, 15 for none. The kit wrote 15 everywhere in these packages.

| Package | Map sent | Cells | Applied after Numbers re-saved it | Map after |
|---|---:|---:|---|---:|
| `P0392` ladder, fixed | 1 entry | 42, C2 = 2026-11-05 to D22 = 2026-12-16 | **C2 = 0**, the other 41 = 15 | 1 entry |
| `W6-first-rule-date-ladder-no-map` | empty | same | **C2 = 0**, the other 41 = 15 | **empty** |

**Numbers took C2, 2026-11-05, as today, with and without the map.** The run date in Asia/Shanghai is 2026-11-05.
Numbers also rewrote the map entry of `P0392` from one range per cell to its own form, two column ranges of rows 2 to 22;
it kept all 42 cells and did not add one for the no-map copy.

The ladder has one weakness: it began on the run date, so C2 is both today and the first cell. Taken alone it does not
exclude an evaluator that marks the first cell. Other results narrow it. In `P0393` the leaf wrote "is tomorrow" on A4, and the
re-save has A4 = today with applied 15, so a leaf-written predicate that is false does not fire. In `P0390`, whose rules are Numbers' own, applied 0
moved from A2 (11-04) to A4 (11-05), so evaluation follows the date. The hand copies then show the leaf's rule firing on a typed date
(§3.2). No package shows a leaf-written "is today" staying off a cell of a different date while being first in its table.
Reopening a re-saved ladder on a later day would settle that, and is not needed for the row.

### 3.2 By hand, saved

| Copy | From | Owner | Saved bytes |
|---|---|---|---|
| `first-rule-B2-fixed` | `P0394` | red after today's date typed in B2 | B2 date 2026-11-05, **applied 0** |
| `first-rule-B2-no-map` | run 2's bytes `e636c14a7940…` | same | B2 date 2026-11-05, **applied 0** |
| `first-rule-B4-fixed` | `P0391` | red | B4 date 2026-11-05, **applied 0** |
| `native-control` | run 2's bytes `ab580f5b4dc7…` | as expected | A2 date 2026-11-05, **applied 0**; A4, also today, applied 0; A3 15 |

This is the table's first line in the kit's `README.md`: red, red, red, so the row can pass. It does. The owner reports
the fill after typing, not before, as the kit asked; the unattended re-saves of `P0394`, `P0391` and the no-map
copy hold B2 or B4 with applied 15, before anyone typed.

### 3.3 The map was not the cause

`first-rule-B2-no-map` is the file of the run-2 note, byte for byte. `W6-first-rule-text-cell-no-map` has SHA-256
`e636c14a7940…`, the same as `W5N-volatile-first-rule-numbers-saved`, and differs from `P0394` in one field: the engine's map. It
highlighted, so **the `uuid_reference_map` gap did not cause the run-2 note**, and the route was not failing there.

**Did Numbers add the map itself?** It depends on the save:

| Copy | Engine map |
|---|---|
| No-map copy, opened and saved by the runner | stays **empty**, also on the ladder |
| No-map copy after the date was typed in B2 and saved | **one entry, byte-identical to the leaf's in `P0394`** |

So Numbers does not repair the map on open or save; it writes it when the cell is edited. The two hand copies of B2 differ
only in the document seed, timestamps and the objects Numbers creates afresh on each save (view state and one dependency tile).

**What the run-2 note most likely meant (inference).** "Neither of them are highlighted" most likely refers to A2 and A4: the
kit's own wording named those cells, and the file has a rule on B2 only, so no highlight was the right answer. Another reading is
that B2 still held text, which Numbers leaves at applied 15. Nothing in the bytes separates these, because run 2 returned no saved
copy. Either way, the same bytes passed when the date went in B2.

**Recommendation: keep the map maintenance.** It is not required for evaluation or for the highlight, so 0433's fix did not fix
a cause. But it writes exactly what Numbers itself writes on a typed edit, Numbers accepts and normalizes it, and it costs +0.06% (0433 §12). Removing it
changes the bytes of `P0390` to `P0394`, whose passes would need a Mac again. No code changed here; **this is for the coordinator to
decide.**

### 3.4 The two files of run 2

The owner says `W5N-volatile-native-control` and `W5N-volatile-native-add-tomorrow` "works as expected". That closes 0433 §6's
"does not report them".

## 4. Pivot: a formula on the source table

Package ids checked against the manifest's sources: `P1150` is `numbers-native-pivot-formula-source-stale-cache`, `P1148` is
`numbers-native-pivot-formula-source-no-cache`. C2 = Grand Total Average / 2.

| Package | C2 sent | C2 after re-save | Formula bytes | `uuid_references` of both owners, as sets |
|---|---:|---:|---|---|
| `P1150` | 1 | **32.5** | unchanged | unchanged |
| `P1148` | none (0) | **32.5** | unchanged | unchanged |

**Numbers recalculated both, with no `uuid_references` entry added.** 32.5 comes from the pivot's own data: the hidden snapshot's
Amount has mean 65, with C2 still 10 there, while the live table's mean is 66.875. The pivot's values and its 78 body formulas do not
change. So 0433 §5's first open item is closed, and W4-K passes on Numbers' own workbook.

## 5. Pivot: Refresh

`pivot-refresh`, the owner's copy of `W6-pivot-source-cell-edit`, against its input:

| Item | Before | After Refresh |
|---|---|---|
| Owner's note, Amount (Sum), Widget and North | 80 | 125 |
| Pivot `B3` in the saved bytes | 80 | **125** |
| Average `C3`, Min `F3` | 40, 10 | 62.5, 55 |
| Columns I to N of row 3 | 40, 700, 42.43, 30, 1800, 900 | 62.5, 3850, 10.61, 7.5, 112.5, 56.25: each is its function of 55 and 70, by my arithmetic |
| Max `E3`, count `D3` | 70, 2 | unchanged |
| Hidden snapshot C2 | 10 | **55** |
| Live C2 | 55 | 55 |

Widget and North rows are C2 and the 70 in row 8: 10 + 70 = 80, 55 + 70 = 125. **Nine pivot cells and one snapshot cell change.
All are on the edited group.**

The unlisted-formula questions:

- **Category formulas: all kept.** The 78 pivot body formulas are byte-identical, cell by cell, to the input.
- **`uuid_references` did not change**: both owners equal the input, and even their bytes do, where a plain re-save reorders them.
- `uuid_reference_map` and `group_node_map` stay empty.
- The group-bys get new object ids, 909202 to 915917 and 909203 to 915918, with the same UIDs, the same inline root bytes and the same 15 node UIDs.
  `refresh_timestamp`, `refresh_uid` and the two rule-change UIDs of the pivot owner change. The leaf still reads 169 categories.

**This package holds no Litchi-written or unlisted formula: `Table 1` has none.** So none of this says what Refresh does to one. **What stays
open:** Refresh or regrouping over an unlisted category formula, such as `P1150`'s C2. There is a further reason to test it: C2 is
a source cell whose formula reads the pivot, so after a Refresh it would depend on itself. No package exercises that. See §10.

## 6. The crash

`W6-C-referenced-tree-no-cycle` against `C-P1096-1-input-host-leaf` of run 2: **one object differs, and in it one field: `9100022.child_ref -> 9100021` is
gone.** Same app and macOS versions.

| Package | Objects | Result |
|---|---|---|
| `C-P1096-1` | all 10, cycle | crash, run 2 |
| `C-7-without-inline-tree` | all but the inline group-by | crash, run 2 |
| `C-7-no-cycle` | all but `9100020`–`23` | opens, run 2 |
| **`W6-C-referenced-tree-no-cycle`** | **all 10, back edge removed** | **opens, saves, reopens, run 3** |

**Proven: the back edge is the cause of the crash, and the referenced tree without it is not.** It agrees with the report's unbounded recursion;
the crash report has no symbols, so what the function does remains inferred. No public route writes field 10 (0433 §4), so Litchi writes no cycle.

**What it does not prove: that Numbers supports a category on a tree stored only by reference.** All ten objects are orphans. A scan of every object
finds no reference to `9100023` or to the other group-bys; the category owner still lists only `904710`. Numbers reads them, and on re-save **drops all
ten**: the re-save holds none of them. The leaf reads 12 categories from the input and **0 from the re-save**: this synthetic data is not in Numbers' model at all.
See §10 and §11.

## 7. Audit of every re-save

`control_kit.py audit`'s functions over all 762 inputs and re-saves, and the five hand copies. No file failed to parse.

**Count rule.** 680 arrays: 566 empty, all (0, []); 114 populated, **all highest index + 1**, no entry beyond the count. No re-save contradicts the rule.
Each of the 114 came back as sent, so none needed a rewrite:

| Sent and re-saved | Arrays |
|---|---:|
| (1, [0]) | 48 |
| (2, [0, 1]) | 23 |
| (3, [0, 2]) | 13 |
| (3, [0, 1, 2]) | 12 |
| (2, [1]) | 11 |
| (3, [2]) | 7 |

**Unlisted references.** 0433 §3 said run 3 would return these bytes. **That was wrong for 15 of the 19:** run 3 held only what no earlier run had seen, and
those 15 are unchanged run-1 bytes. `results/0434/unlisted-references.tsv` lists all 19.

| Kind | Packages | In run 3 | After the re-save |
|---|---:|---|---|
| `3056 -> 212` | 8 | `P0903`, `P0914`, `P0915`, `P0916` | **kept and still unlisted**, all 8 references, same counts. Native files hold it too (`comment-reader-native-created`, `comment-compat-native-multiroot`) |
| `6001 -> 6002/6005/6006` | 12 | **none** | **no re-save exists.** Open, save and reopen passed in run 1 |

No re-save holds an unlisted reference of any kind other than `3056 -> 212`, and no hand copy holds one.

So `3056 -> 212` is Numbers' own habit, and harmless. `6001 -> 600x` occurs only in Litchi output, opens, and remains unexplained: Numbers has not shown what it
does to it. `P0858` holds both, in run 1.

## 8. The 743 regenerated packages

0433 §11 said nobody checked that they differ from run 1 only by host UUIDs. **That is moot:** they ran as they are, and the fold counts their own bytes.

## 9. Record 0433's gaps

| 0433 §10 gap | Now |
|---|---|
| First rule on a Numbers-saved table has no proof; the map fix is unproven | **Closed**; the map was not the cause (§3) |
| Refresh and regrouping with unlisted category formulas | Refresh on a package with no such formula changed no formula; with one, and regrouping, **open** (§5) |
| `6001 -> 6002/6005/6006` unexplained | **Open**, §7 |
| Which part of the referenced tree Numbers recurses on | **The cycle** (§6). The function's name stays unknown |
| Caption's paragraph style is the title's; new style objects share an identifier; rule's styles in the list's `DataList`; `series_count()` by column; Keynote and Pages refuse the first volatile rule | Unchanged, not reexamined |

## 10. Rows still open

**These three say `pass` because their packages passed. Each has a part with no Mac proof, named in the row.** Each is the coordinator's decision.

| Row | What is proven | What is not | Blocks W3 under D1? |
|---|---|---|---|
| `q-3080a227051a` W4-K `set_formula`, `set_formula_cached` with `pivot_category`, line 773 | Four native-workbook packages: open, save, reopen, formula kept, recalculated | **A category on a tree stored only by reference.** Only Linux proofs against the host (`numbers_pivot_formula_author_referenced_tree_parity`) | See below |
| `q-c8ea770c712d` W4-H `pivot_categories`, line 756 | Native workbook read; 169 categories | **The read of a tree stored only by reference**, by the same reasoning | See below |
| `q-9bb7c4ec065a` W1-B Pages `Identity::from_parts`, line 304 | Run 3: both open, save, reopen. Re-saved `versionUUID`, `documentUUID` and `privateUUID` differ between the two, and `stableDocumentUUID` keeps each input's. 0432 §5: the owner saw no prompt with both open, on earlier bytes | **Coordinator correction:** the hand edit and save was done. Run 2's `RUN.md` §4a asked for it on `I-P0512` and `I-P0513`, the same route's run-1 packages, and the owner reported "No conflict". The check was not repeated on run 3's bytes | Closed: D1's wording and the row's hand step are both met |

**The by-reference part.** What the Mac shows:

- Numbers stores every tree **both ways**: inline under field 3, and by object under fields 17 and 18, with field 10 children. A scan of the 331 Numbers files we hold (`test-data`, the 282 re-saves, the hand copies) finds 346 group-bys: **346 both, 0 inline only, 0 by reference only.** The leaf takes the inline tree first (`pivot_categories.rs`, line 276), **so it never walks the by-object tree of such a file.**
- No public route writes the by-reference-only form.
- The only such data sent to a Mac is unattached (§6), so it **does not test the route**, in either direction.

My reading: D1 asks a route to pass native open, save, close and reopen, and those two routes have. The by-reference tree is an input shape, which no writer makes and no native file is known to hold. So on D1's wording **I do not think it blocks W3**, but the row cannot claim the shape. After W3 the Linux proofs against the host go too, so deleting the host also removes the only evidence of that branch; decide with that in mind. If you want it settled, a run-4 package can be built without new leaf code: the native workbook with the inline root (field 3) and inline aggregators (field 4) removed from its two enabled group-bys. Numbers either keeps the shape and it becomes testable, or refuses or rebuilds it, which also answers the question.

**Open and not tied to a row** (no row's Fixture asks them):

- Refresh or regrouping over an unlisted category formula, §5. Hand step on `P1150`'s re-save: press Refresh, save, return. That is two minutes. It does not change any route's open, save and reopen.
- `6001 -> 600x`, §7. Run 4 could send those 12 packages to get re-saves. It would explain them; it is not needed for their pass.


## 11. Gaps found here

- **`pivot_categories` reads group-bys no owner lists.** It returns 12 categories from `W6-C`'s orphan objects, which Numbers discards. `numbers_pivot_categories_read_parity` holds it to the host's answer on that fixture. It gives no wrong answer on the files we hold: in the §10 scan, every group-by is listed by an owner. Not changed.
- **`apply_results.py --evidence` drops hand notes** on a later fold (§2).
- **The queue's Reason cells cannot be corrected without changing row ids** (§2), so a regenerated manifest would be needed to fix the W1-F text.
- **W3's snapshot must take host-built inputs from `.coord/iwa-w6/queue-out/inputs/`**, per `.coord/mac-runbook/w3-host-built-inputs.md`. Do not regenerate: passes on host-built inputs rest on bytes a fresh host run would not reproduce, as 0433 §8 found.

## 12. Not verified

- **No displayed result was checked by me.** PDFs were not returned; `results.tsv` holds their hashes only. All Mac-side evidence is bytes, `results.tsv`, or the owner's notes, which I took as written.
- The leaf opened all 282 Numbers re-saves and the five hand copies with a temporary probe, not committed. I did not open the Pages and Keynote re-saves with their leaf crates; the audit parses them at the wire level only.
- Fill colour is known only from the owner's note; saved bytes give the applied rule, not a rendered fill. The fill before typing was not reported.
- The rule's fired state after a date change on an unedited leaf-written ladder: one date only (§3.1).
- That a by-reference tree attached to an owner opens: never sent (§10).
- 0431's reference-line change: still no package holds one, per 0433 §2; not rechecked.
- The crash packages did not run again in run 3; §6 joins two runs on the same app and macOS versions.

## 13. Gates

No Rust changed, so `cargo fmt`, clippy, doc and the crate tests were not run; they have nothing new to check. A temporary reader probe was built with `cargo test -p litchi-numbers` (exit 0) and removed. Logs are in `(deleted-coord-artifact)`.

| Gate | Exit |
|---|---:|
| `python3 -m unittest tools.test_mac_queue tools.test_mac_control_kit` | 0 |
| `python3 tools/mac_queue/generate.py --check`: 767 rows, 765 mapped, 2 grouped, 0 errors | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/unified_bench_matrix.py --check`, after regenerating with this record present | 0 |
| The rules' tool tests: `test_officecli_corpus`, `test_officecli_capability_delta`, `test_no_3rdparty_references`, `test_check_crate_boundaries`, `test_non_iwork_gate`: 1276 tests | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check` | 0 |
| `check_feature_matrix.py`: 31 ok | 0 |
| `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0 |
| `apply_results.py --evidence … --dry-run` after the hand notes: only the five rows of §2 differ | 0 |
