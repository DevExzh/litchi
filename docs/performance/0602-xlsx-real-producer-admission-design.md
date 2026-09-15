# 0602: the XLSX value editor admits no real producer file, so change 0525's readback gate is not the blocker

Status: design, retained. `performance_claim: none`. No production code changed.
This record freezes a design, states what it would have to measure to be
admitted, and reports the sizing measurements that were taken to write it.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Item **XLSX-2** of change [0587](0587-remaining-opportunity-survey.md) (rank 12)
proposes extending change [0525](changes/0525-xlsx-unchanged-cell-readback.md)'s
omission readback and the source-backed value editor to real-producer
worksheets. It names two mechanisms: `stored_entry_is_supported`
(`crates/litchi-xlsx/src/cell_values/snapshot.rs:365-370`) disabling the reduced
readback "whenever *any* cell of the source or output is a shared string", and
planning refusing "any worksheet that carries a relationship" (`:441-443`). It
sizes the loss as "0525's 96% parser-instruction and 30% commit-p50 reduction,
currently zero on any Excel-produced file", and asks for a frozen design and an
SST-bearing harness shape.

This record was asked to size that opportunity on real files and freeze the
design. The sizing found that **XLSX-2's mechanism is wrong**, and the correction
changes the item from a readback widening into something much larger. Both named
gates are real, but neither is reached. The editor has eight admission gates;
XLSX-2 names the third and the fifth, every real fixture stops at the first, and
the fifth cannot fire at all on any input this module admits. The measurements
below establish that, size what admission would be worth, and the design states
what admitting a real producer would cost in contracts.

## What was measured, and what it corrects

### 1. The value editor admits none of the 95 real fixtures

The 95-fixture corpus is the one change 0587 counted: every `.xlsx` under
`test-data/ooxml/xlsx` and `test-data/office-interop`. This record reproduces
0587's structural counts on it exactly — 95 files, 77 with a `sharedStrings`
part — and then opens each one through the editor.

`SourceBackedEditor::open` followed by `snapshot(first sheet)` refuses **95 of
95**. Two fail in the read door before the editor is reached (a malformed
worksheet dimension and a shared-string count past `i32::MAX`). The other 93 all
refuse with the same typed error and the same message shape:

```
invalid XLSX structure: value-only edits refuse package relationship '<reltype>'
```

45 name `officeDocument/2006/relationships/extended-properties`
(`docProps/app.xml`), 41 name `package/2006/relationships/metadata/core-properties`
(`docProps/core.xml`), and the rest name thumbnails, custom properties or a
vendor metadata type. The gate is `validate_package_relationships`
(`snapshot.rs:1924-1952`), which admits only `officeDocument`, its Strict
spelling and `digital-signature/origin`. Every producer writes `docProps`, so
every producer file stops there.

### 2. The admission ladder, and where XLSX-2's two gates actually sit

Reimplementing each gate structurally over the raw package
(`census/ladder.py`, validated against the editor's own verdict on G1 for all
93 readable files) gives the whole ladder. "First refusal" is the gate a file
actually hits; "would refuse" counts each gate independently, as if every
earlier one were opened:

| | gate | site | first refusal | would refuse |
| --- | --- | --- | ---: | ---: |
| G1 | package-root relationship allow-list | `snapshot.rs:1924-1952` | **95** | 95 |
| G2 | workbook relationship allow-list (no `sharedStrings`) | `snapshot.rs:1704-1751` | 0 | 77 |
| G3 | worksheet carries any relationship | `snapshot.rs:441-443` | 0 | 45 |
| G4 | `t="s"` with a `<v>`, parsed with no table | `raw/worksheet/semantic.rs:134-146` | 0 | 64 |
| G5 | `stored_entry_is_supported` | `snapshot.rs:365-370` | 0 | **0** |

Three further gates are not in that table because they need a package that has
already cleared G1-G4 to be observed; section 4 reaches them.

XLSX-2 named G3 and G5. G3 is real but blocks 45, not "57 of 95"; G5, the gate
the item is built on, **blocks nothing**, and section 3 shows it cannot.

The shared-string story is also not the one XLSX-2 tells. A shared string does
not disable the reduced readback: it refuses the worksheet outright, twice over.
`sharedStrings` is absent from the workbook relationship allow-list
(`snapshot.rs:1718-1727` lists only worksheet, styles, theme and calcChain in
both dialects), so G2 refuses the package. And every worksheet parse inside
`cell_values` passes `|| Ok(None)` for the table — `validation.rs:203,214,219`
and `snapshot.rs:393,450,673,700,727,742`, exhaustively — so `parse_value`
refuses at `semantic.rs:136-138` with *"worksheet uses shared strings but the
workbook has no shared-string part"* before any store is built.

### 3. `stored_entry_is_supported` cannot fire

The guard has four clauses. Each was given a minimal synthetic package that is
inside every other allow-list and differs only in the feature that clause tests
(`fixtures/synth.py`, 256 x 24 cells, 6,144 cells each). Every one is refused
before the guard is consulted:

| clause | twin | refusal |
| --- | --- | --- |
| `shared_string.is_none()` | `sst` (`t="s"` + a `sharedStrings` part) | `value-only edits refuse workbook relationship '.../sharedStrings'` |
| `shared_string.is_none()` | `sstfree` (`t="s"`, no part) | `worksheet uses shared strings but the workbook has no shared-string part` |
| `!inline_rich` | `rich` (`<is><r><t>`) | `value-only edits refuse dependency-bearing or unknown element 'r'` |
| `cell_metadata.is_none()` | `cm` (`cm="1"` on every `<c>`) | `value-only edits refuse attribute 'cm' on 'c'` |
| `value_metadata.is_none()` | `vm` (`vm="1"` on every `<c>`) | `value-only edits refuse attribute 'vm' on 'c'` |
| — | `numeric`, `inline` | admitted; complete plan, commit and publication cycle |

The element and attribute allow-lists in `cell_values/validation.rs` are the
reason. The worksheet vocabulary is sixteen names — `worksheet dimension
sheetViews sheetView pane selection sheetFormatPr cols col sheetData row c f v
is t` (`validation.rs:387-400`) — with `(Some(b"is"), b"t")` as the only parent
rule under `is` (`:448`), so `<r>` is unreachable; and `<c>` admits exactly
`r|s|t` (`:603`), so `cm` and `vm` are unreachable.

So `stored_entry_is_supported` is, today, unreachable defence-in-depth. That is
not a defect: D4 of the design makes each of its clauses reachable, and it then
becomes the load-bearing guard it was written to be. But widening it is
not an optimization, because there is nothing on the other side of it.

### 4. What an admitted real-producer edit costs

Because no real file is admissible, sizing needs derived fixtures. Two scripted
passes, both retained, open the gates in order and change nothing else:

* `fixtures/degate.py` drops the package-root relationships outside G1's list
  and the parts they reach, drops the workbook relationships outside G2's list,
  folds `sharedStrings` back into the worksheets as inline strings so no cell
  content is lost, and drops the worksheet relationship parts with the
  `r:id`-bearing worksheet children that reach them.
* `fixtures/project.py` projects `workbook.xml` and each worksheet onto the
  element and attribute allow-lists verbatim, **leaving `<sheetData>` alone**:
  the same rows, the same `<c>` records in the same order, the same styles, the
  same numerals and the same string lengths.

The projection therefore keeps the producer's cell geometry, which is what
planning and commit cost is proportional to, and loses the envelope the editor
refuses. It is not the producer's file and nothing here is claimed for the
producer's file as shipped.

Deriving the five largest candidates surfaced three more gates, all of them
downstream of the five in section 2:

| | gate | site | observed on |
| --- | --- | --- | --- |
| G6 | `mc:Ignorable` on `workbook` | `validation.rs:476-481` | 4 of 5 |
| G7 | out-of-`sheetData` worksheet children (`sheetPr`, `mergeCells`, `pageMargins`, `pageSetup`, `autoFilter`, `conditionalFormatting`, `dataValidations`, `hyperlinks`, `extLst`) | `validation.rs:403-408` | 5 of 5 |
| G8 | `value-only edits currently require exactly one worksheet` | `snapshot.rs:510-513`, `:595-598` | 3 of 5, single-sheet door only |

G8 applies to `SourceBackedEditor::edit` and `snapshot` only; `edit_many` and
`edit_sheets` admit multi-sheet workbooks, and every measurement below uses that
door.

Four projections then run a complete one-cell plan and commit. Geometry is of
the edited part, `xl/worksheets/sheet1.xml`, in the file as shipped:

| fixture | cells | `t="s"` cells | worksheet rels | source sheet bytes |
| --- | ---: | ---: | :---: | ---: |
| `FormatConditionTests.xlsx` | 11 | 10 | no | 2,272 |
| `dataValidationTableRange.xlsx` | 217 | 55 | yes | 24,756 |
| `sheet-state-show.xlsx` | 4,108 | 76 | no | 91,187 |
| `no_drawing_patriarch.xlsx` | 75,770 | 66,935 | yes | 3,382,556 |

**Instruction attribution.** Callgrind isolation pairs at N=1 and N=4 samples
against one retained editor, differenced and divided by M=3, so the open, the
process start and the harness cancel and the residue is one complete plan and
commit. `--separate-callers=1` keeps `raw::worksheet::parse` attributable to its
caller. Inclusive Ir per operation, as a share of the operation:

| fixture | Ir / operation | planning | commit | complete worksheet parse (planning) |
| --- | ---: | ---: | ---: | ---: |
| `FormatConditionTests` | 20,890,294 | 73.18% | 26.70% | 32.34% |
| `dataValidationTableRange` | 95,355,352 | 82.64% | 17.23% | 49.33% |
| `sheet-state-show` | 463,879,806 | 87.53% | 12.42% | 74.43% |
| `no_drawing_patriarch` | 4,110,387,331 | 47.93% | 51.60% | 47.29% (fused) |

The first three worksheets declare `xmlns:mc` and `xmlns:x14ac`, so
`source_stream_eligible` (`raw/worksheet/mod.rs:60-68`) refuses the shared
traversal and planning runs `validation::worksheet_xml` and then
`raw::worksheet::parse` as two passes; their parse column is that separate
parse. `no_drawing_patriarch`'s projection carries no marker, so change 0546's
fused `worksheet_xml_and_parse_source` runs and validation and parse are not
separable — its column is the fused call.

With the reduced readback on, the commit's own readback is cheap:
`reduced_readback` is 0.00-0.01% of an operation on all four, and
`Store::merge_omitted_cells` 0.03-0.49%. `Snapshot::from_rewritten_value_source`
is 4.78-20.72%, almost all of it the mandatory complete validation of the output
(`snapshot.rs:723`), which this design does not touch.

**The readback's size, measured directly.** One route in the tree disables the
reduced readback without any patch: `rewrite_value_only_with_provenance` returns
an empty omission list when an edited address lands on a row the scanned layout
does not contain (`raw/worksheet/edit/package.rs:182-203`), and
`from_rewritten_value_source` then takes its `omitted.is_empty()` early-out
(`snapshot.rs:719`) into a complete parse of the candidate. So a one-cell
`insert` on an absent row is the same operation with change 0525 switched off,
and the same isolation pair run on both kinds measures the complete candidate
parse rather than modelling it.

In the `insert` profiles `reduced_readback` and `Store::merge_omitted_cells`
disappear entirely and `Snapshot::from_rewritten_value_source` absorbs the
difference. Inclusive Ir per operation:

| fixture | `from_rewritten_value_source` set → insert | difference | share of the commit that pays it |
| --- | ---: | ---: | ---: |
| `FormatConditionTests` | 3,563,043 → 7,559,035 | 3,995,992 | **41.7%** |
| `dataValidationTableRange` | 11,559,410 → 50,883,468 | 39,324,058 | **70.5%** |
| `sheet-state-show` | 22,191,537 → 355,750,004 | 333,558,467 | **85.3%** |
| `no_drawing_patriarch` | 851,538,753 → 1,892,367,855 | 1,040,829,102 | **32.9%** |

The whole operation grows 19.6%, 41.3%, 72.8% and 25.0% in instructions.

**Latency.** Paired A1/B1/B2/A2 legs, A = `set`, B = `insert`, 5 warmup and 40
samples per leg (30 for `no_drawing_patriarch`), pinned to CPU 18, one retained
editor per leg. p50 nanoseconds:

| fixture | A1 set | B1 insert | B2 insert | A2 set | A1→B1 | A2→B2 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `FormatConditionTests` | 1,048,245 | 1,239,097 | 1,239,197 | 1,042,376 | +18.21% | +18.88% |
| `dataValidationTableRange` | 4,646,153 | 6,470,642 | 6,490,703 | 4,630,383 | +39.27% | +40.18% |
| `sheet-state-show` | 22,019,400 | 38,145,932 | 38,327,523 | 21,662,399 | +73.24% | +76.93% |
| `no_drawing_patriarch` | 260,332,700 | 334,377,871 | 320,807,853 | 260,091,411 | +28.44% | +23.34% |

In the other direction the reduced readback saves 15.40%, 28.20%, 42.28% and
22.14% of a plan and commit at p50 (B1→A1), and 15.88%, 28.66%, 43.48% and
18.93% on the second pair (B2→A2). The latency deltas track the instruction
growth above — +18/+20%, +39/+41%, +73/+73%, +28/+25% — so on this path the two
metrics agree.

**A/A floor.** Four further legs per fixture, all `set`, interleaved in the same
window. S1→S2: +0.05%, −0.33%, +0.76%, **+4.03%**. S3→S4: +0.99%, +1.85%,
−1.35%, −0.35%. S1→S4: +0.96%, +0.87%, −0.70%, +3.65%. The floor is inside 2%
at p50 on the three small fixtures — better than the host's usual 4%, because
the editor is opened once per leg and the measured region is CPU-bound with no
I/O — but reaches 4.03% on `no_drawing_patriarch`, whose 260 ms operation runs
long enough to pick up host noise from the other agents. Its set/insert deltas
are five to seven times that floor, so they stand; but its numbers are the
weakest in this record and the instruction counts are what should be relied on.

An `insert` is not a pure isolation of the readback: it also writes a new row
and takes a different layout path. The callgrind pairs bound that difference.
`rewrite_value_only_with_provenance`, the part an insert would perturb, costs
424,560 → 428,637, 3,933,327 → 3,928,502, 34,482,199 → 34,471,435 and
1,269,227,906 → 1,268,230,188 inclusive Ir between the `set` and `insert` legs:
within 1% on every fixture, and within 0.15% on three. The whole of the growth
is in `from_rewritten_value_source`, so the difference attributed above is the
readback and not the rewrite.

**XLSX-2's falsification condition is not met.** It said the item would be
falsified "if on `ConditionalFormattingSamples.xlsx`-class inputs the complete
candidate parse is under 10% of commit Ir". Measured, it is 32.9%, 41.7%, 70.5%
and 85.3% of the commit that pays it. The mechanism is large. It is simply
already removed for every input this module admits.

### 5. Publication refuses 94 of the 95 anyway

`litchi-opc` audits **both** the original bytes and the replacement bytes of
every replaced XML part (`source_backed.rs:7696-7697`, `8121-8122`), and change
0528 requires both audits to stay. Running that exact audit —
`xml_minifier::audit::verify_authored` with default limits — over the
`xl/workbook.xml` and first worksheet of all 95 fixtures:

| part | compact | noncompact |
| --- | ---: | ---: |
| `xl/workbook.xml` | 1 | 94 |
| first worksheet | 1 | 94 |

93 of each fail with `FormattingWhitespace`, one with `WhitespaceBeforeClose`.
The single compact package is `test-data/office-interop/litchi-changed/date-autofilter-litchi.xlsx`,
which litchi wrote. Every third-party producer indents its XML.

This is visible in the derived fixtures too: `sheet-state-show`'s projection
plans and commits, and then fails publication with `XML publication rejected for
'/xl/workbook.xml': noncompact XML FormattingWhitespace at byte 55` — byte 55 is
the newline the producer wrote after its XML declaration. Change 0587's XLSX
section listed this as an item "to verify rather than a confirmed defect". It is
confirmed, and it is a harder blocker than any editor gate: **admitting real
producers to the value editor is worth nothing until the publication contract
admits a non-compact original.**

## The design

The design is stated as an ordered sequence, because the later parts are
worthless without the earlier ones. Each part says what changes, what proves it
sound, and what it costs in contracts. Nothing here is implemented.

### D0 (prerequisite, not this crate). The publication original-bytes audit

`litchi-opc` must stop requiring the *original* bytes of a replaced part to be
compact. The audit exists to keep litchi's own authored output compact; applying
it to bytes litchi did not author refuses every real producer. The narrowest
form: `verify_authored` stays on the replacement; the original is checked for
well-formedness and limits only, by a second policy that does not assert
authored compactness.

This moves where a refusal happens, so it needs its own record on `litchi-opc`
with change 0528's review, and it is the gate on everything below. This record
does not design it; it establishes that it is required and sizes it at 94 of 95
real fixtures.

### D1. The package-root relationship allow-list becomes a closure proof

Replace `validate_package_relationships`'s reltype allow-list with the property
it is a proxy for. Admit any package-root relationship provided it is not
external and there is exactly one `officeDocument` owner; refuse nothing else.

Soundness comes from the replacement set, not from the reltype. A value-only
commit's `SourceTopologyPlan` names exactly three things: the worksheet part's
bytes (`append_worksheet_replacement`, `snapshot.rs:1408-1428`), the workbook
part's bytes for recalculation invalidation (`append_owner_topology`,
`:1440-1448`), and the `calcChain` part with its workbook relationship
(`:1449-1470`). No other part is named, and `litchi-opc` transfers every
unnamed member verbatim (change 0578). `docProps/app.xml` and
`docProps/core.xml` are therefore preserved byte-for-byte, which is exactly what
ADR 0006's preservation default asks for. Refusing them refuses to preserve what
is already preserved.

Two things stay. External relationships stay refused, because their targets are
outside the package and no closure can be proved over them. And the signed-package
guard stays where it is: `Patch::apply_inner` already returns `Error::Signed`
(`cell_values/patch.rs`), so a signature-bearing package is refused at apply, not
by the reltype list.

Not claimed: that `docProps/core.xml`'s `dcterms:modified` is updated. It is
not, before or after. Preservation is the contract.

### D2. Shared strings: admit the relationship, and supply the table

Two changes, and they must land together.

*The relationship.* Add `SHARED_STRINGS` and `STRICT_SHARED_STRINGS` to
`validate_workbook_relationships`'s allow-list (`snapshot.rs:1718-1727`) and to
its cardinality check, with `shared_strings <= 1` alongside the existing
`styles`, `themes` and `calculation_chains` bounds.

*The table.* Capture the shared-string part in the snapshot the way
`capture_auxiliary` already captures styles and theme: retain its bytes, and
materialize `Box<[Text]>` lazily. Replace the `|| Ok(None)` closure at the
planning parse sites (`snapshot.rs:393`, `:450`, and
`validation.rs:203,214,219`) with one that materializes from the retained part.
The closure is invoked only when some cell actually carries `t="s"`
(`codec.rs:306-310`), so a workbook with a shared-string part and no
shared-string cells pays nothing.

**How the reduced readback of edited rows tolerates `t="s"` neighbours.** It
supplies the retained table, and change 0525's omission rule does not move. The
reasoning is that the table is not an extra cost at commit: planning cannot
build the source `Store` at all without resolving every `t="s"` cell into
`Value::Text` (`semantic.rs:134-146`), so for any worksheet that has one, the
table is already materialized and retained in the snapshot by the time commit
runs. The reduced document is parsed with that same table
(`snapshot.rs:742`), and so is the complete-candidate fallback (`:727`). Nothing
about the omission spans changes, and change 0525's proof shape is untouched.

The alternative the brief names — omitting `t="s"` neighbours from the proof so
they are reused from the source `Store` — is a later refinement, not a
prerequisite. It would work: `Store::merge_omitted_cells` clones already
materialized `Stored` values (`cell.rs:837`) and needs no table at all. But it
only helps replacement-only rows; change 0525 requires a row containing an
insertion or a removal to be parsed in full including its unchanged followers,
and those followers need the table anyway. It should be measured separately
against the per-neighbour `Arc<str>` clone it removes.

*Index identity is the preservation argument.* `Stored.shared_string` retains
the raw index (`cell.rs:637`). The rewrite copies `t="s"` cells verbatim — they
are inside `<sheetData>` but are not edited cells — so the candidate's indices
still address the unchanged `sharedStrings.xml`, which is never in the
replacement set. The merge takes omitted `Stored` with their indices from the
source store and parsed cells from the reduced document, and both index the same
table. This holds only while the shared-string part is preserved, so it becomes
an invariant: **if any action would add, remove or renumber a shared string,
refuse.** The module already "never creates ... shared strings"
(`cell_values/mod.rs:11`), so the invariant is a check, not a restriction.

*Resource bound.* Materializing `no_drawing_patriarch`'s table is 66,935
`Arc<str>` allocations. The existing `checked_multi_bytes` budget and the
`ReadLimits` fences must cover the shared-string part before it is materialized,
and the materialization must stay lazy so a worksheet with no `t="s"` cell never
pays it. This is the design's main resource risk and F3 below is its
falsification.

### D3. Relationship-bearing worksheets: a preservation proof from the rewrite grammar

Replace the refusal at `snapshot.rs:441-443` (and its two siblings at `:387-389`
and `:526-528`) with the proof that a value-only rewrite cannot disturb a
worksheet relationship.

`rewrite_value_only_with_provenance` (`raw/worksheet/edit/package.rs:163-258`)
writes exactly three things:

1. the `ref` attribute of `<dimension>`, rewritten in place (`:226-236`);
2. the cell records between `<sheetData>` and `</sheetData>`, via
   `write_sheet_data_with_provenance` (`:237-251`);
3. everything else, copied byte-verbatim — the head up to `<sheetData>` and the
   tail from `</sheetData>` to the end (`:252`).

In `CT_Worksheet`'s fixed ECMA-376 sequence, every child that can carry an
`r:id` — `hyperlinks`, `drawing`, `drawingHF`, `legacyDrawing`, `legacyDrawingHF`,
`picture`, `oleObjects`, `controls`, `webPublishItems`, `tableParts`,
`pageSetup` — follows `sheetData`. `CT_Row` and `CT_Cell` have no `r:id`
attribute in any dialect, so no relationship reference can occur inside the
rewritten span. A search of the whole `raw/worksheet/edit/` directory for `r:id`,
`relationship` or `rels` returns nothing: the code has no way to touch one.

Therefore the output's relationship references are byte-identical to the
source's, and the relationship part itself is never named by the topology plan,
so `litchi-opc` transfers it verbatim. This is ADR 0006's preservation default
satisfied by construction rather than by refusal.

One assertion must be restated rather than deleted. `Snapshot::apply_owned_target`
checks `part.rels().is_empty()` on the readback (`snapshot.rs:1145-1148`). It
becomes "the readback part's relationships equal the source part's
relationships" — a strictly stronger preservation statement than "there are
none", and the one that makes the byte-identity gate below checkable.

Two semantic questions were checked and are not problems. A hyperlink, table
part, data validation or conditional format whose range covers the edited cell
addresses the cell by *address*; a value change does not move it. And a
`calcChain` referencing the edited cell is already dropped atomically with its
workbook relationship (`:1449-1470`), which is the existing contract.

The adversarial case that must refuse is an `r:id` planted inside `<sheetData>`,
which the grammar forbids: it is not admissible SpreadsheetML and the validator
must keep refusing it. That is an explicit gate below.

### D4. The element and attribute vocabulary: split the allow-list

This is the expensive part and the reason this record is design-only.

*Inside `<sheetData>`, nothing changes.* That is the span the rewrite composes,
and it must stay fully understood: `row`, `c`, `f`, `v`, `is`, `t` with their
present attribute lists, and every parent rule as written.

*Outside `<sheetData>`, replace the name allow-list with a dependency rule.* An
element is admitted when it is well-formed SpreadsheetML in the transitional or
strict namespace and is copied verbatim by the rewrite. The soundness argument
is D3's: the head and tail are byte-verbatim, so admitting a name costs nothing
but the promise not to interpret it. `<extLst>` and vendor payloads are admitted
on the same terms — preserved, never interpreted — which is what
`process_ooxml` already does for markup-compatible content.

*Attributes: admit `cm` and `vm` on `<c>` as preserved-but-uninterpreted*, and
admit `<r>` under `<is>` on the same terms. `Stored` already has the fields
(`cell.rs:652-660`, `:641`). Authoring them stays refused: the module creates
only unstyled numeric cells.

This is the change that makes `stored_entry_is_supported` live. With `cm`, `vm`
and `<r>` admitted, a metadata-bearing or rich-inline cell reaches the readback,
and the guard correctly forces the complete candidate parse for that worksheet.
The guard is not to be widened — it is the safety net that makes D4 safe, and
D4 is what gives it something to catch.

*What does not change.* Validation still runs, complete, over the complete
output (`snapshot.rs:723`), under ADR 0005. Only the vocabulary it admits
widens; no traversal is skipped, no boundary is moved, and no refusal is traded
for a partial result. `mc:Ignorable` on `workbook` (G6) follows the same rule:
admitted as an attribute that is preserved verbatim, still never interpreted.

### D5. The single-worksheet requirement

`require_single_sheet` (`snapshot.rs:510-513`, `:595-598`) refuses a multi-sheet
workbook on the `edit`/`snapshot` door only. `edit_many` and `edit_sheets`
already admit one. The single-sheet door can be routed through `MultiSnapshot`
with one selector. This carries no contract change and is the only part of this
design that could be implemented on its own; it is listed here because three of
the five derived fixtures hit it, and a caller who uses the documented
single-sheet API on a real workbook will hit it first.

### Which refusals stay exactly where they are

Change [0541](changes/0541-xlsx-planning-error-order-guards.md) froze six
error-precedence properties. All six survive unchanged, because widening a
vocabulary changes *which inputs refuse*, never *which error wins when two
compete*:

* within a worksheet, complete value-only validation still wins over a parser
  error appearing earlier in the bytes;
* MCE preprocessing still precedes raw parsing, and a later validator refusal
  still overrides both;
* across worksheets, the loader still validates and parses each selected sheet
  in workbook order, and a later sheet's validation error still cannot replace
  the first sheet's raw error even when selectors arrive reversed;
* failed plans are still retried, source bytes still unchanged, valid unselected
  owners still usable;
* no partial multi-sheet transaction is returned.

Two orderings are added and must be frozen with the same kind of test. A
shared-string index past the table's end keeps its existing message
(`shared-string index {index} exceeds table length {len}`, `semantic.rs:141-145`)
and, being a parser error, still loses to a validation refusal in the same
worksheet. A missing shared-string part for a worksheet that uses one keeps
`worksheet uses shared strings but the workbook has no shared-string part` and
is reached only after G2 admits the relationship.

Every 0541 fixture must be re-run and must produce the identical typed error and
message. Fixtures that today refuse at a widened gate and will then refuse
somewhere later must be added as controls that name the new refusal point, so
the ordering is observable rather than assumed.

### The harness shape this needs

The existing XLSX corpora cannot measure any of this: every one has
`shared_strings: null`, no worksheet relationship, no out-of-`sheetData`
children and compact XML. The shape needed is a generator,
`litchi-xlsx-cell-values-producer-shape-v1`, with knobs for

* rows x columns and a shared-string fraction, with a real `sharedStrings` part;
* worksheet relationship count and kind (printerSettings, table, drawing);
* out-of-`sheetData` children (`sheetPr`, `mergeCells`, `autoFilter`,
  `conditionalFormatting`, `dataValidations`, `pageSetup`, `headerFooter`,
  `extLst`);
* namespace declarations (`mc`, `x14ac`), so both the fused and two-pass
  planning paths are exercised;
* package-root `docProps`;
* **and a non-compact variant**, because D0 is the real gate and no existing
  corpus can exercise it.

Selectors: `xlsx_source_backed_cell_values_one_edit_save` and its
`one_percent`, `batch` and `managed` siblings under
`--xlsx-cell-crud-shape producer-{small,medium,large}`. The derived fixtures in
this packet are the interim stand-in and are retained with the scripts that
made them.

### Predicted saving

Stated as what it is: not a speedup of an existing path, but the size of a path
that does not currently exist.

* Change 0525's mechanism already delivers, on every input this module admits,
  a saving **measured** here at 15.4-43.5% of a plan-and-commit p50, and the
  complete candidate parse it removes is **measured** at 32.9-85.3% of the
  commit that would otherwise pay it.
* The population that receives it is 0 of 95 real fixtures today. D0 through D4
  do not make the mechanism faster; they make it reachable.
* **Retained** from change 0525, for the shape of what becomes available on
  producer files: 30.1-30.8% commit-p50 and 96.0-96.5% raw-parser-Ir reductions
  on synthetic corpora. The absolute per-edit cost on the four derived fixtures
  is 1.05 ms (11 cells), 4.65 ms (217), 22.0 ms (4,108) and 260 ms (75,770) for
  plan and commit alone, **measured**.
* Nothing is predicted for publication, because D0 is unresolved and publication
  is where 94 of 95 real fixtures stop.

### Admission gates

No part of this design may land without all of these.

* **Byte-identical output over the corpus.** For every fixture the widened
  editor admits: an exact no-op commit reproduces the source worksheet bytes
  exactly; a one-cell edit differs from the source only in the edited `<c>`
  span, the `<dimension ref>` attribute, `workbook.xml`'s calculation
  properties and the removed `calcChain`; and every other part, including every
  `_rels` part, is byte-identical. Shown to fail without D3's restated
  assertion.
* **Relationship identity.** The readback part's relationships equal the source
  part's, per part, over the whole corpus — the restated form of
  `snapshot.rs:1145-1148`.
* **Error identity.** The complete 0541 matrix re-run with identical typed
  errors and messages, plus one control per widened gate naming the new refusal
  point, plus the two added shared-string orderings.
* **The differential oracle.** Parse each admitted output through the eager door
  and compare its semantic projection cell-for-cell against the source
  projection with the staged edit applied, including style indexes, shared-string
  identity and formula provenance.
* **Adversarial.** A malformed shared-string count; an index past the table's
  end; a `t="s"` cell whose part is absent; an `r:id` planted inside
  `<sheetData>` (must refuse); a worksheet whose relationship part is referenced
  but missing; a non-compact original once D0 exists.
* **Resource.** The shared-string table is materialized at most once per
  snapshot and never for a worksheet with no `t="s"` cell; the part is charged
  against the existing byte budget before materialization; incremental peak on
  the producer corpus stays inside the envelope the existing corpora set.
* **Performance.** Paired A1/B1/B2/A2 over the new producer corpus with an A/A
  floor measured in the same window, every scenario reported including the ones
  that get worse, and no scenario worse than 5% at p50 unreported.

### Falsification

* **F1.** If any out-of-`sheetData` element admitted by D4 turns out to be
  interpreted rather than copied — that is, if the byte-identity gate fails on
  any corpus fixture — the closure proof is void and D4 must revert to a name
  allow-list, extended one name at a time.
* **F2.** If the publication original-bytes contract (D0) cannot change, the
  whole item is void whatever the editor admits, because 94 of 95 real fixtures
  refuse at publication.
* **F3.** If materializing the shared-string table at planning costs more than
  the readback it enables — measured on a producer fixture with a large table,
  `no_drawing_patriarch`'s 66,935 items being the retained example — D2 must
  fall back to the omit-from-the-proof variant, or the item is void for
  large-table workbooks.
* **F4.** If admitting relationship-bearing worksheets changes any output byte on
  any corpus fixture, D3 is void.
* **F5.** XLSX-2's own condition — the complete candidate parse under 10% of
  commit Ir — is **not met**: measured at 32.9-85.3% of the commit that pays it.
  The item is not falsified on size.

## Measured

Every figure in section 4 and section 5 was taken in this batch, on this host,
from the outputs retained in the packet. Deterministic counts first: the
95-fixture census, the admission ladder, the seven synthetic gate twins and the
compactness audit are exact and reproducible from the retained scripts. The
callgrind isolation pairs are deterministic instruction counts. The timing legs
are last, with their A/A floor in the same window.

Instruction counts rank work, not latency. Callgrind counts `rep movsb` per byte,
so the rewrite's copy share is an upper bound; no SHA-256 runs on this path, so
the usual software-hashing caveat does not apply here.

## Correctness evidence

No production code changed, so there is no behavioural evidence to give. What
this record establishes about behaviour it establishes by observation: the
refusal messages in sections 1, 3 and 4 are the editor's own typed errors,
captured verbatim from the retained probe output.

Gates run in the worktree and retained in `gates.txt`: `cargo fmt --all
--check`; `cargo clippy -p litchi-xlsx --all-targets`; `cargo test -p
litchi-xlsx`; `cargo doc -p litchi-xlsx --no-deps`. They gate the record, not a
change: the tree is identical to `f22f93935` under `crates/`.

## Validation preserved

Nothing was changed, so nothing was weakened. The design's own commitment is
stated in D4: every traversal ADR 0005 mandates keeps running, complete, over
the complete output; only the vocabulary the value-only validator admits widens,
and every element it newly admits is one the rewrite already copies byte-verbatim
and never interprets.

## Limitations

* **No real producer file was measured as shipped.** Every timing and
  instruction figure in section 4 is from a *derived* fixture whose envelope was
  stripped to the admission surface. The cell geometry is the producer's; the
  package is not.
* The `insert` leg is a proxy for a disabled reduced readback, not a patch. It
  also writes a new row, so its timing delta is an upper bound on the readback's
  own share; the callgrind attribution is the tighter figure.
* `no_drawing_patriarch`'s A/A floor reached 4.03% at p50 in one pair, with
  eight agents building on the host. Its latency numbers are the weakest here
  and are reported, not relied on; its instruction counts are deterministic and
  are.
* Only four fixtures carry a working plan and commit, and one candidate
  (`MatrixFormulaEvalTestData.xlsx`) was dropped because all 64 of its stored
  value cells belong to range-scoped formulas.
* The admission ladder's G1-G5 counts are a structural reimplementation of the
  gates, validated against the editor's own verdict for G1 on 93 of 95 files;
  G2-G5 are not independently validated against the editor, because no file
  reaches them.
* This record's 56 "any worksheet has relationships" differs from change 0587's
  57, and its 65 marker-bearing first sheets from 0587's 63; the scopes are
  defined in `census/census.py` and the difference is in what counts as a
  relationship part and whether `mc:Ignorable` counts as a marker. The
  `sharedStrings` count, 77 of 95, reproduces exactly.
* No claim is made about publication cost, allocation, peak RSS, cold cache,
  range sources, concurrency or any platform other than this host. No claim is
  registered.
* Nothing here is authorized. D0 needs a record on `litchi-opc` with change
  0528's review; D1 through D4 need the producer corpus, the differential oracle
  and the gates above.

## Retained evidence

[`results/change-0602/README.md`](results/change-0602/README.md).
