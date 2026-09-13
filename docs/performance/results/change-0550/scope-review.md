# 0550 XLSX commit-attribution scope review

`performance_claim: none`

This independent read-only review covers the source-backed cell-values cases
selected by the 0550 plan. It identifies the actual transaction type and the
collection boundaries before any new profile or candidate result is
interpreted. It does not add a Rust or harness change and does not claim a
hotspot, latency result, allocation result, or speedup.

## Source and plan identity

The review was made at `HEAD` `090b15b64ae52da2f8bf765cbb745ef76122792e`.
The exact files read are bound by these SHA-256 values:

| File | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0550/plan.json` | `3eeb52cdad1312b3953006e538490d6523a707e9858354dcb6688b4a615471a7` |
| `docs/performance/results/change-0550/run.py` | `dcb20341f9558bd6de1befcfe0afe43d2045c90c6c7c5673529fe234d2e6d39f` |
| `tools/perf-baseline/src/lib.rs` | `b71922e53c507842e4d8fb4983f52872a690940e3f7e9bf1b0f2d03fdbcd525f` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/tests/source_backed_cell_values.rs` | `d3a30e7577bad6a0f2b832fe730b6734abbf1635d48704a96dea352f7b381cdb` |

The previous scope records read for continuity are
`docs/performance/results/change-0512/scope-review.md` (SHA-256
`25f0a25cab7722fd154a1f7b201705401fd36fe7ccebcea353f7333aa25f8ead`),
`docs/performance/changes/0512-xlsx-commit-attribution.md` (SHA-256
`a44718fb96cdcc4b3d633efcd29a87181bf429cac9cbc6df5ec43e0f60911dd2`), and
`docs/performance/results/change-0549/next-target.md` (SHA-256
`5b9e75255cfe7010d4ca1bb9426ff389e235a76034d093fb5d14d7c58b930ae4`).
The 0520 protocol remains useful for the owner toggle, but its measurements
are from an older source epoch and are not reused as 0550 evidence.

## The cases and their owners

The generic cases from 0512 and the source-backed cell-values cases are
different paths:

| Case family | Runner path | Transaction returned by the editor |
| --- | --- | --- |
| `xlsx_one_cell_commit`, `xlsx_one_percent_commit` | `run_xlsx_update_commit` (`lib.rs:40856`) calls `prepare_xlsx_updates`, which calls `Workbook::edit` (`lib.rs:22108`) | eager `litchi_xlsx::Edit`; neither `SourceEdit` nor `MultiSourceEdit` |
| `xlsx_source_backed_cell_values_one_edit_save`, `xlsx_source_backed_cell_values_one_percent_edit_save` | `run_xlsx_cell_values_edit_save` (`lib.rs:41771`), source-backed branch calls `SourceBackedEditor::edit_sheets` (`lib.rs:41906`) | `litchi_xlsx::cell_values::MultiSourceEdit` |

`SourceBackedEditor::edit` returns the single-worksheet `SourceEdit` at
`cell_values/source.rs:453-466`; it is not called by either selected 0550
case. `SourceBackedEditor::edit_sheets` returns `MultiSourceEdit` at
`cell_values/source.rs:469-476`. The selected owner is therefore exactly
`litchi_xlsx::cell_values::source::MultiSourceEdit::commit`, whose body starts
at `source.rs:1125` and ends at `source.rs:1233`.

The one-edit case does not become a `SourceEdit` merely because it selects one
worksheet. `xlsx_cell_crud_updates_for_case` chooses the first inventory cell
(`Sheet1!A1`) at `lib.rs:21540-21546`, and
`xlsx_update_sheet_selectors` produces one selector at `lib.rs:21577-21587`.
`edit_sheets([Sheet1])` still constructs a `MultiSourceEdit` and its commit
loops over the one selected snapshot.

The one-percent case returns the deterministic `spec.one_percent_updates`
vector at `lib.rs:21529-21538`. For the planned medium, noncompact, and
vendor-extension shapes, each of four sheets has a 48-by-48 inventory; for
dense-sparse the four inventories are respectively a dense 128-by-128 sheet,
every fourth row/column, every eighth row/column, and a diagonal. The
one-percent vector is flattened over those inventories and the deduplicated
selector set reaches all four sheets for these shapes. It therefore exercises
one `MultiSourceEdit` commit over four worksheet snapshots. The one-edit case
exercises one selected worksheet; neither case exercises the cell-values
`SourceEdit::commit` path.

## Exact timing and collection boundaries

The source-backed runner performs lifecycle gates and expected-output
construction before its measured loop:

* `run_xlsx_cell_value_lifecycle_gates` begins at `lib.rs:21952`. Its exact
  no-op commit is at `21961-21967`, its clear and remove commits are at
  `21979-21990`, and its vendor-extension partial-sink commit is at
  `22029-22040`.
* `xlsx_cell_crud_eager_output` is then used to build the expected output
  before the loop. That helper uses eager `Workbook::edit` and an eager
  `litchi_xlsx::Edit` commit; it is not the selected source-backed owner.
* Corpus construction also commits eagerly in `build_xlsx_workbook` at
  `lib.rs:21362-21422`, before the case runner is entered.

The source-backed per-iteration order is:

1. Create the instrumented source and sink, then open the source-backed
   editor (`lib.rs:41855-41901`).
2. Plan the selected worksheet closure with `editor.edit_sheets(selectors)`
   (`lib.rs:41903-41913`).
3. Stage every `edit.set` in the loop at `lib.rs:41924-41934`.
4. Call `edit.commit()` at `lib.rs:41936`, then publish with
   `publish_multi_commit_to_stream` at `lib.rs:41974-41985`.
5. Reopen and verify the output after the measured duration at
   `lib.rs:42100-42118`.

The native `commit_ns` clock starts before the staging loop and stops after
`MultiSourceEdit::commit` (`lib.rs:41924-41943`). Consequently, the existing
native phase includes `edit.set`/`apply_batch` work. The owner profile must
start at `MultiSourceEdit::commit` entry after all staging and stop on return;
its instruction total must not be presented as the native `commit_ns` total.
The complete native lifecycle interval additionally includes open, selector
planning, commit phase, and publication. Sink setup, expected-output
generation, lifecycle gates, reopen, semantic/package oracles, and later
destruction are outside that interval as applicable.

For a Callgrind owner profile, the exact controls are:

```text
--collect-atstart=no
--toggle-collect=litchi_xlsx::cell_values::source::MultiSourceEdit::commit
--zero-before=litchi_xlsx::cell_values::source::MultiSourceEdit::commit
--dump-after=litchi_xlsx::cell_values::source::MultiSourceEdit::commit
```

The reset and toggle are deliberately on the exact owner, rather than on
`edit_sheets`, `publish_multi_commit_to_stream`, or a wildcard containing
`SourceEdit::commit`. The collected region contains the owner body: effective
action counting, action-map construction, per-sheet rewrite, complete
validation and reduced readback/fallback, staged readback checks, calculation
invalidation, `MultiSnapshot` construction, and patch construction. It does
not contain source open, selector planning, staging, publication, or caller
oracles.

Every owner invocation must be retained and classified by its direct parent.
For medium, dense-sparse, and noncompact shapes there are three lifecycle
owner invocations before the selected operation: no-op, clear, and remove.
For vendor-extension there is a fourth lifecycle invocation for the
partial-sink gate. Thus the selected operation is the fourth owner dump for
the first three shapes and the fifth owner dump for vendor-extension. The
selected dump must have the positive incoming edge from
`litchi_perf_baseline::run_xlsx_cell_values_edit_save`; gate dumps have the
edge from `run_xlsx_cell_value_lifecycle_gates`. A final process dump with no
owner collection is not operation evidence. Descendant call metadata can
retain collection-off context, so descendant counts do not establish owner
invocation counts.

The current allocation regions have a related boundary limitation:
`plan_allocation_metrics` covers `edit_sheets` only (`lib.rs:41904-41909`),
while `commit_allocation_metrics` starts before `edit.set` and ends after
`edit.commit` (`lib.rs:41924-41941`). It is staging-plus-commit allocation
evidence, not owner-only allocation evidence. A future allocation claim about
the owner needs a region that begins after staging or must retain this wider
scope explicitly.

## Semantic coverage and limits

The selected timed edits are existing scalar numeric replacements. The corpus
builder fills the inventories with `xlsx_value(coordinate)` numbers
(`lib.rs:21362-21422`), and the runner changes each selected value to
`xlsx_value(coordinate) + 1` (`lib.rs:41926-41934`). The timed path therefore
does not cover insertion, clear, removal, formulas, shared-formula expansion,
styles, shared strings, rich values, row/column edits, or other cell content
types. It also does not cover the managed-budget execution variant, the
explicit `edit_many` route, or a single-worksheet `SourceEdit::commit`.

The four planned shapes vary the source envelope rather than the semantic edit
kind. Medium, noncompact, and vendor-extension use the four-sheet 48-by-48
inventory; dense-sparse combines the dense and sparse layouts described above.
The noncompact builder rewrites alternating cell tags with an `x` namespace
(`lib.rs:20608-20696`), and the vendor-extension builder adds opaque XML and
binary parts with an internal relationship (`lib.rs:20730-20740`). These
shapes are opt-in: `XlsxCellCrudShape::ALL` contains only medium and
dense-sparse (`lib.rs:505-514`), so the capture driver must explicitly request
noncompact and vendor-extension for the 0550 plan to cover them.

The pre-loop lifecycle gates still provide useful correctness evidence. They
exercise an exact no-op, clear/remove, stale and foreign source/patch refusal,
inverse restoration, and—only for vendor-extension—a partial sink refusal.
Those gates use `MultiSourceEdit` and are outside the selected operation's
native timing. They do not turn the measured numeric replacement into broad
semantic coverage. The source implementation retains the complete worksheet
validator and its fallback rules: `Snapshot::from_rewritten_value_source`
validates the complete output and tries reduced readback at
`cell_values/snapshot.rs:715-759`; the rewrite eligibility checks are at
`raw/worksheet/edit/package.rs:177-203`, and reduced-span validation is at
`package.rs:262-286`. Any optimization must preserve those validation,
provenance, error-order, source-lineage, calculation-invalidation, atomic
publication, limit, and untouched-member contracts.

## Actionable review conclusion

The 0550 owner selection is valid for attributing the source-backed
`MultiSourceEdit::commit` body, provided the shape-dependent lifecycle dump
count and direct parent edge are checked. The native phase should remain
described as open + planning + staging/commit + publication; owner Callgrind
Ir is a separate mechanism diagnostic. No 0550 result should be labeled
`SourceEdit` evidence, and the absent single-worksheet `SourceEdit` profile
should remain an explicit coverage gap rather than being inferred from the
one-edit case.
