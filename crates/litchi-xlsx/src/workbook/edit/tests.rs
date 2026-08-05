//! Regression tests for semantic edit planning, patching, and package rewrites.

use super::codec::calculation_chain_removal;
use super::*;
use crate::StyleState;
use crate::cell::{Number, Value};
use crate::column::Outline;
use crate::error::MergeEditBlock;
use crate::formula::Formula;
use litchi_opc::{BlobPart, TargetMode};

fn task_panes(app_ref: &str) -> common_web::Panes {
    let reference = common_web::Reference::new("test-add-in", "1.0", common_web::Store::Omex)
        .expect("reference");
    let add_in = common_web::AddIn::new("test-add-in", reference)
        .and_then(|add_in| add_in.bind(common_web::Binding::new("table", "table", app_ref)?))
        .expect("add-in");
    let mut panes = common_web::Panes::new();
    panes
        .push(common_web::Pane::new(add_in))
        .expect("task pane");
    panes
}

#[test]
fn web_bindings_and_task_panes_commit_join_and_reverse_atomically() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut pane_edit = source.edit().expect("pane edit");
    pane_edit
        .put_task_panes(
            task_panes("sheet-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage panes");
    let mut binding_edit = source.edit().expect("binding edit");
    binding_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!$A$1:$B$4").expect("binding"))
        .expect("stage binding");
    pane_edit.join(binding_edit).expect("dependent edits join");

    let committed = pane_edit.commit().expect("atomic web commit");
    assert_eq!(committed.patch().package_changes().len(), 1);
    assert!(committed.patch().changes().iter().any(|change| {
        matches!(change, Change::Web { sheet, .. } if sheet.as_ref() == "Sheet1")
    }));
    assert_eq!(
        committed
            .workbook()
            .task_panes()
            .expect("task panes")
            .expect("present")
            .len(),
        1
    );
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet");
    assert_eq!(
        sheet
            .web_bindings()
            .expect("bindings")
            .get("sheet-table")
            .map(crate::web::Binding::formula),
        Some("Sheet1!$A$1:$B$4")
    );

    let before_rename = committed.workbook().to_bytes().expect("pre-rename bytes");
    let mut rename = committed.workbook().edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("tab lookup")
        .expect("tab")
        .rename("Data")
        .expect("rename");
    let renamed = rename.commit().expect("dependency-aware rename");
    assert_eq!(
        renamed
            .workbook()
            .sheet("Data")
            .expect("renamed lookup")
            .expect("renamed worksheet")
            .web_bindings()
            .expect("renamed bindings")
            .get("sheet-table")
            .map(crate::web::Binding::formula),
        Some("Data!$A$1:$B$4")
    );
    assert_eq!(
        renamed
            .workbook()
            .apply(&renamed.patch().inverse())
            .expect("inverse rename")
            .workbook()
            .to_bytes()
            .expect("inverse rename bytes"),
        before_rename
    );

    let mut left = committed.workbook().edit().expect("left web edit");
    left.sheet("Sheet1")
        .expect("left lookup")
        .expect("left worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!A1").expect("left binding"))
        .expect("left stage");
    let mut right = committed.workbook().edit().expect("right web edit");
    right
        .sheet("Sheet1")
        .expect("right lookup")
        .expect("right worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!B2").expect("right binding"))
        .expect("right stage");
    let conflict = left.join(right).expect_err("whole binding sets overlap");
    assert!(
        conflict
            .conflicts()
            .expect("web conflict set")
            .conflicts()
            .iter()
            .any(Conflict::is_web)
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse web patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut unsafe_remove = committed.workbook().edit().expect("remove edit");
    unsafe_remove
        .remove_task_panes()
        .expect("stage unsafe removal");
    assert!(matches!(
        unsafe_remove.commit(),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));

    let mut safe_remove = committed.workbook().edit().expect("safe remove edit");
    safe_remove.remove_task_panes().expect("stage pane removal");
    assert!(
        safe_remove
            .sheet("Sheet1")
            .expect("lookup")
            .expect("worksheet")
            .clear_bindings()
            .expect("clear bindings")
    );
    let removed = safe_remove.commit().expect("combined removal");
    assert!(
        removed
            .workbook()
            .task_panes()
            .expect("task panes")
            .is_none()
    );
    assert!(
        removed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("worksheet")
            .web_bindings()
            .expect("bindings")
            .is_empty()
    );
}

#[test]
fn task_panes_and_sheet_removal_join_conflicts_are_symmetric() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("new worksheet");
    let source = create.commit().expect("source").into_workbook();

    for panes_first in [true, false] {
        let mut panes = source.edit().expect("pane edit");
        panes
            .put_task_panes(
                task_panes("sheet-table"),
                common_web::Conformance::Transitional,
            )
            .expect("stage panes");
        let mut removal = source.edit().expect("removal edit");
        removal
            .remove("Delete")
            .expect("lookup")
            .expect("worksheet");

        let error = if panes_first {
            panes.join(removal).expect_err("panes then removal")
        } else {
            removal.join(panes).expect_err("removal then panes")
        };
        assert!(matches!(
            error.failure(),
            JoinFailure::Overlap(conflicts)
                if conflicts.conflicts().iter().any(Conflict::is_remove)
        ));
    }
}

#[test]
fn web_patch_cross_graph_validation_is_atomic_on_alternate_snapshots() {
    fn with_bindings(workbook: &Workbook, bindings: WebBindings) -> Workbook {
        let mut package = workbook.inner.package.clone();
        let uri = workbook.inner.sheets[0].part_uri.clone();
        let source = package
            .get_part(&uri)
            .expect("worksheet part")
            .blob()
            .to_vec();
        let after = raw::web::replace(&source, &bindings).expect("write bindings");
        package
            .get_part_mut(&uri)
            .expect("worksheet part")
            .set_blob(after);
        Workbook::from_package(package).expect("reopen alternate workbook")
    }

    let baseline = Workbook::new().expect("baseline");
    let mut install = baseline.edit().expect("install edit");
    install
        .put_task_panes(
            task_panes("sheet-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage panes");
    let with_panes = install.commit().expect("install panes").into_workbook();

    let mut bind = with_panes.edit().expect("binding edit");
    bind.sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!A1").expect("binding"))
        .expect("stage binding");
    let binding_commit = bind.commit().expect("binding commit");
    let alternate = Workbook::new().expect("alternate without panes");
    let before = alternate.to_bytes().expect("alternate bytes");
    assert!(matches!(
        alternate.apply(binding_commit.patch()),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));
    assert_eq!(alternate.to_bytes().expect("unchanged bytes"), before);

    let old_bindings = WebBindings::try_from(vec![
        WebBinding::new("sheet-table", "Sheet1!A1").expect("old binding"),
    ])
    .expect("binding collection");
    let alternate = with_bindings(&with_panes, old_bindings);
    let before = alternate.to_bytes().expect("alternate bytes");
    let mut replace = with_panes.edit().expect("pane replacement");
    replace
        .put_task_panes(
            task_panes("replacement-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage replacement panes");
    let pane_commit = replace.commit().expect("pane replacement commit");
    assert!(matches!(
        alternate.apply(pane_commit.patch()),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));
    assert_eq!(alternate.to_bytes().expect("unchanged bytes"), before);
}

#[test]
fn cell_crud_is_atomic_reversible_and_source_preserving() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
        sheet
            .set("A1", "hello")
            .and_then(|sheet| sheet.set("B2", 42_i32))
            .and_then(|sheet| sheet.set("C3", Formula::new("B2*2").expect("formula")))
            .expect("cell changes");
    }
    let committed = edit.commit().expect("commit");
    assert_eq!(
        source.to_bytes().expect("source remains valid"),
        source_bytes
    );
    assert_eq!(committed.patch().len(), 3);

    let book = committed.workbook();
    let sheet = book.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "hello"
    ));
    assert!(matches!(
        sheet.cell("B2").expect("B2").stored(),
        Some(Cell::Value(Value::Number(number))) if number == &Number::new("42").expect("number")
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Formula(_))
    ));
    let extents = sheet.extents().expect("committed extents");
    assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("A1:C3"));
    assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("A1:C3"));

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(matches!(
        source.apply(committed.patch()),
        Ok(applied) if applied.workbook().sheet("Sheet1").expect("lookup").expect("sheet").cell("A1").expect("cell").stored().is_some()
    ));
    assert!(matches!(
        book.apply(committed.patch()),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn merged_range_crud_is_sparse_safe_reversible_and_composable() {
    let source = merged_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert_eq!(
        source_sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["A1:B2"]
    );
    assert!(matches!(
        source_sheet.cell("A1").expect("anchor"),
        crate::cell::View::Stored(Cell::Value(_))
    ));
    assert!(matches!(
        source_sheet.cell("B2").expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:B2").expect("range")
    ));

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .and_then(|sheet| sheet.set("B2", "revealed"))
        .and_then(|sheet| sheet.merge("C3:D4"))
        .expect("merged-range changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 3);
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert_eq!(
        sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["C3:D4"]
    );
    assert!(matches!(
        sheet.cell("B2").expect("uncovered cell").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "revealed"
    ));
    assert!(matches!(
        sheet.cell("D4").expect("covered cell"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("C3:D4").expect("range")
    ));
    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut blocked = source.edit().expect("blocked edit");
    blocked
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("plan merge");
    assert!(matches!(
        blocked.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::FollowerContent { address },
            ..
        }) if address == Address::from_a1("E2").expect("address")
    ));

    let mut cleared = source.edit().expect("clear edit");
    cleared
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .and_then(|sheet| sheet.merge("D1:E2"))
        .expect("clear then merge");
    let cleared = cleared.commit().expect("safe merge");
    assert!(matches!(
        cleared
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("D1:E2").expect("range")
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("merge");
    let mut clear = source.edit().expect("independent clear");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .expect("clear");
    merge.join(clear).expect("clear makes the merge safe");
    assert!(matches!(
        merge
            .commit()
            .expect("joined clear and merge")
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(_)
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("merge");
    let mut write = source.edit().expect("independent follower write");
    write
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("D4", "hidden")
        .expect("write");
    assert!(matches!(
        merge.join(write).expect_err("follower content must conflict").failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut overlap = source.edit().expect("overlap edit");
    overlap
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("B2:C3")
        .expect("plan overlap");
    assert!(matches!(
        overlap.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::Overlap { existing },
            ..
        }) if existing == Rect::from_a1("A1:B2").expect("existing")
    ));

    let mut created = source.edit().expect("create edit");
    created
        .add("Merged")
        .expect("new sheet")
        .set("A1", "anchor")
        .and_then(|sheet| sheet.merge("A1:C2"))
        .expect("new merged range");
    let created = created.commit().expect("create merged sheet");
    assert!(matches!(
        created
            .workbook()
            .sheet("Merged")
            .expect("lookup")
            .expect("sheet")
            .cell("C2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:C2").expect("range")
    ));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("F1:G2")
        .expect("right merge");
    left.join(right).expect("disjoint merge join");
    assert_eq!(
        left.commit()
            .expect("joined commit")
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .merges()
            .expect("merges")
            .count(),
        3
    );

    let mut left = source.edit().expect("left overlap");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right overlap");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D4:E5")
        .expect("right merge");
    let error = left.join(right).expect_err("overlapping joins must fail");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut left = source.edit().expect("left unmerge");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("A1")
        .expect("left unmerge");
    let mut right = source.edit().expect("right unmerge");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .expect("right unmerge");
    let error = left
        .join(right)
        .expect_err("two selectors for one merge must conflict");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| {
                conflict.merges().is_some_and(|ranges| {
                    ranges == [Rect::from_a1("A1:B2").expect("range")]
                })
            })
    ));

    let mut missing = source.edit().expect("missing unmerge");
    missing
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("H8")
        .expect("missing range is a no-op");
    assert!(missing.commit().expect("no-op commit").patch().is_empty());
}

#[test]
fn merged_range_reads_share_one_snapshot_without_public_locks() {
    let workbook = Arc::new(merged_workbook());
    std::thread::scope(|scope| {
        for _ in 0..8 {
            let workbook = Arc::clone(&workbook);
            scope.spawn(move || {
                let sheet = workbook.sheet(0usize).expect("lookup").expect("sheet");
                assert!(matches!(
                    sheet.cell("B2").expect("covered"),
                    crate::cell::View::Covered(range)
                        if range == Rect::from_a1("A1:B2").expect("range")
                ));
                assert_eq!(sheet.merges().expect("merges").count(), 1);
            });
        }
    });
}

#[test]
fn clear_and_remove_have_distinct_missing_and_empty_semantics() {
    let source = Workbook::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", "value")
        .expect("set");
    let first = edit.commit().expect("first commit").into_workbook();

    let mut edit = first.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.clear("A1").expect("clear");
    sheet.clear("B1").expect("clear missing");
    let cleared = edit.commit().expect("clear commit");
    assert_eq!(cleared.patch().len(), 1);
    let sheet = cleared
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("cell").stored(),
        Some(Cell::Empty)
    ));
    assert!(sheet.cell("B1").expect("cell").is_missing());

    let mut edit = cleared.workbook().edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .remove("A1")
        .expect("remove");
    let removed = edit.commit().expect("remove commit");
    assert!(
        removed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .is_missing()
    );
}

#[test]
fn row_visibility_is_checked_reversible_and_patch_visible() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
    sheet.set("A1", "visible").expect("cell");
    sheet.row(1).expect("row 2").hide();
    let committed = edit.commit().expect("commit");

    assert_eq!(committed.patch().len(), 2);
    assert!(matches!(
        &committed.patch().changes()[1],
        Change::Row {
            row,
            before: RowState::Missing,
            after: RowState::Stored(properties),
            ..
        } if row.get() == 1 && properties.hidden()
    ));
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    let row = sheet.row(1).expect("row 2");
    assert!(row.stored());
    assert!(row.hidden());

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut edit = committed.workbook().edit().expect("show edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(RowIndex::new(1).expect("row 2"))
        .expect("checked row")
        .show();
    let shown = edit.commit().expect("show commit");
    let shown_sheet = shown
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    let shown_row = shown_sheet.row(1).expect("row 2");
    assert!(shown_row.stored());
    assert!(!shown_row.hidden());

    let mut no_op = source.edit().expect("no-op edit");
    no_op
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(10)
        .expect("row 11")
        .show();
    assert!(no_op.commit().expect("no-op commit").patch().is_empty());
    let mut invalid = source.edit().expect("invalid row edit");
    let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.row(litchi_sheet::ROWS),
        Err(Error::Coordinate(_))
    ));
}

#[test]
fn row_layout_is_typed_reversible_and_facet_composable() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .row(1)
        .expect("row 2")
        .height(30)
        .expect("checked height")
        .outline(2)
        .expect("checked outline")
        .collapse()
        .thick_top()
        .thick_bottom()
        .show_phonetic();
    let committed = edit.commit().expect("layout commit");

    assert_eq!(committed.patch().len(), 1);
    let (_, before, after) = committed.patch().changes()[0].row().expect("row change");
    assert!(matches!(before, RowState::Missing));
    let RowState::Stored(properties) = after else {
        panic!("expected stored row properties")
    };
    assert_eq!(properties.height().map(crate::row::Height::get), Some(30.0));
    assert_eq!(properties.outline().get(), 2);
    assert!(properties.custom_height());
    assert!(properties.collapsed());
    assert!(properties.thick_top());
    assert!(properties.thick_bottom());
    assert!(properties.phonetic());
    assert!(!properties.hidden());
    assert!(!properties.custom_format());
    assert!(matches!(properties.style(), StyleState::Default));

    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    let row = sheet.row(1).expect("row 2");
    assert_eq!(row.height().map(crate::row::Height::get), Some(30.0));
    assert_eq!(row.outline().get(), 2);
    assert!(row.custom_height());
    assert!(row.collapsed());
    assert!(row.thick_top());
    assert!(row.thick_bottom());
    assert!(row.phonetic());
    assert!(matches!(
        sheet.row_style(1).expect("row style"),
        Some(crate::LocalStyle::Default)
    ));

    let mut reset = committed.workbook().edit().expect("reset edit");
    reset
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .reset_height()
        .outline(0)
        .expect("outline reset")
        .expand()
        .normal_top()
        .normal_bottom()
        .hide_phonetic();
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    let reset_row = reset_sheet.row(1).expect("row 2");
    assert_eq!(reset_row.height(), None);
    assert!(!reset_row.custom_height());
    assert_eq!(reset_row.outline(), Outline::NONE);
    assert!(!reset_row.collapsed());
    assert!(!reset_row.thick_top());
    assert!(!reset_row.thick_bottom());
    assert!(!reset_row.phonetic());

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut invalid = source.edit().expect("invalid edit");
    let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.row(1).expect("row 2").height(f64::NAN),
        Err(Error::RowHeight(_))
    ));
    assert!(matches!(
        sheet.row(1).expect("row 2").height(409.1),
        Err(Error::RowHeight(_))
    ));
    assert!(matches!(
        sheet.row(1).expect("row 2").outline(8),
        Err(Error::Outline(_))
    ));

    let mut height = source.edit().expect("height edit");
    height
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(2)
        .expect("row 3")
        .height(crate::row::Height::new(22.0).expect("prevalidated height"))
        .expect("height");
    let mut visibility = source.edit().expect("visibility edit");
    visibility
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(2)
        .expect("row 3")
        .hide();
    height.join(visibility).expect("disjoint facets on one row");
    let joined = height.commit().expect("joined commit");
    let joined_sheet = joined
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    let joined_row = joined_sheet.row(2).expect("row 3");
    assert!(joined_row.hidden());
    assert_eq!(joined_row.height().map(crate::row::Height::get), Some(22.0));

    let mut left = source.edit().expect("left height");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(3)
        .expect("row 4")
        .height(10)
        .expect("height");
    let mut right = source.edit().expect("right height");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(3)
        .expect("row 4")
        .reset_height();
    assert!(left.join(right).is_err());
}

#[test]
fn worksheet_defaults_are_typed_reversible_and_facet_composable() {
    let source = defaults_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let original = source.sheet("Sheet1").expect("lookup").expect("worksheet");
    let defaults = original
        .defaults()
        .expect("default lookup")
        .expect("stored defaults");
    assert_eq!(defaults.base_width(), 10);
    assert_eq!(defaults.width().map(layout::Width::get), Some(12.0));
    assert_eq!(defaults.height().get(), 15.0);
    assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.1));
    assert!(defaults.custom_height());
    assert!(defaults.hidden());
    assert!(defaults.thick_top());
    assert_eq!(
        original
            .row(1)
            .expect("row 2")
            .descent()
            .map(layout::Descent::get),
        Some(0.2)
    );

    let mut edit = source.edit().expect("defaults edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("worksheet");
        {
            let mut defaults = sheet.defaults();
            defaults
                .reset_base_width()
                .show()
                .normal_top()
                .thick_bottom();
            defaults.width(14.5).expect("checked width");
            defaults.height(20).expect("checked height");
            defaults.descent(0.25).expect("checked descent");
        }
        sheet
            .row(1)
            .expect("row 2")
            .reset_descent()
            .height(24)
            .expect("checked row height");
    }
    let committed = edit.commit().expect("defaults commit");
    assert_eq!(committed.patch().len(), 2);
    assert!(committed.patch().graph.is_empty());
    let (before, after) = committed.patch().changes()[0]
        .defaults()
        .expect("defaults change");
    assert!(before.is_some());
    let after = after.expect("updated defaults");
    assert_eq!(after.stored_base_width(), None);
    assert_eq!(after.base_width(), layout::DEFAULT_BASE_WIDTH);
    assert_eq!(after.width().map(layout::Width::get), Some(14.5));
    assert_eq!(after.height().get(), 20.0);
    assert_eq!(after.descent().map(layout::Descent::get), Some(0.25));
    assert!(!after.hidden());
    assert!(!after.thick_top());
    assert!(after.thick_bottom());

    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet");
    assert_eq!(sheet.defaults().expect("lookup"), Some(after));
    let row = sheet.row(1).expect("row 2");
    assert_eq!(row.descent(), None);
    assert_eq!(row.height().map(crate::row::Height::get), Some(24.0));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut width = source.edit().expect("width edit");
    width
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .defaults()
        .width(18)
        .expect("width");
    let mut hidden = source.edit().expect("hidden edit");
    hidden
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .defaults()
        .hide();
    width.join(hidden).expect("disjoint default facets");
    let joined = width.commit().expect("joined defaults");
    let joined_sheet = joined
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet");
    let defaults = joined_sheet.defaults().expect("lookup").expect("defaults");
    assert_eq!(defaults.width().map(layout::Width::get), Some(18.0));
    assert!(defaults.hidden());

    let mut left = source.edit().expect("left height");
    left.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .defaults()
        .height(16)
        .expect("height");
    let mut right = source.edit().expect("right height");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .defaults()
        .height(17)
        .expect("height");
    let error = left.join(right).expect_err("same default facet conflicts");
    let conflicts = error.conflicts().expect("default conflicts");
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].defaults(),
        Some(layout::Fields::HEIGHT)
    );

    let mut invalid = source.edit().expect("invalid defaults");
    let mut sheet = invalid.sheet(0usize).expect("lookup").expect("worksheet");
    assert!(matches!(
        sheet.defaults().height(f64::NAN),
        Err(Error::DefaultHeight(_))
    ));
    assert!(matches!(
        sheet.defaults().width(65_536.0),
        Err(Error::DefaultWidth(_))
    ));
    assert!(matches!(
        sheet.defaults().descent(-0.1),
        Err(Error::Descent(_))
    ));
    assert!(matches!(
        sheet.row(0).expect("row 1").descent(f64::INFINITY),
        Err(Error::Descent(_))
    ));
}

#[test]
fn new_sheet_defaults_require_height_and_commit_with_short_selectors() {
    let source = Workbook::new().expect("source workbook");
    let mut incomplete = source.edit().expect("incomplete edit");
    incomplete
        .add("Incomplete")
        .expect("new sheet")
        .defaults()
        .width(12)
        .expect("checked width");
    assert!(matches!(
        incomplete.commit(),
        Err(Error::DefaultsEditBlocked {
            reason: crate::DefaultsEditBlock::NeedsHeight,
            ..
        })
    ));

    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("new sheet edit");
    {
        let mut sheet = edit.add("Grid").expect("new sheet");
        sheet.set("A1", "ready").expect("cell");
        {
            let mut defaults = sheet.defaults();
            defaults.height(18).expect("height");
            defaults.width(13.5).expect("width");
            defaults.descent(0.2).expect("descent");
        }
        sheet
            .row(4)
            .expect("row 5")
            .descent(0.3)
            .expect("row descent");
    }
    let committed = edit.commit().expect("new sheet commit");
    let sheet = committed
        .workbook()
        .sheet("Grid")
        .expect("name lookup")
        .expect("new worksheet");
    let defaults = sheet
        .defaults()
        .expect("defaults lookup")
        .expect("stored defaults");
    assert_eq!(defaults.height().get(), 18.0);
    assert_eq!(defaults.width().map(layout::Width::get), Some(13.5));
    assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.2));
    assert_eq!(
        sheet
            .row(4)
            .expect("row 5")
            .descent()
            .map(layout::Descent::get),
        Some(0.3)
    );
    assert!(matches!(
        sheet.cell("A1").expect("cell lookup").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "ready"
    ));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse new sheet");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn row_layout_patch_guards_and_rebinds_hidden_shared_style_identity() {
    let source = styled_row_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .height(28)
        .expect("height");
    let committed = edit.commit().expect("commit");
    let (_, _, after) = committed.patch().changes()[0].row().expect("row change");
    let RowState::Stored(properties) = after else {
        panic!("expected stored properties")
    };
    assert!(properties.custom_format());
    assert!(matches!(properties.style(), StyleState::Shared(_)));

    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.row_style(1).expect("row style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
    let replayed = reopened
        .apply(committed.patch())
        .expect("source-checked replay");
    let (_, _, replayed_after) = replayed.patch().changes()[0]
        .row()
        .expect("replayed row change");
    let RowState::Stored(replayed_properties) = replayed_after else {
        panic!("expected replayed properties")
    };
    let StyleState::Shared(replayed_key) = replayed_properties.style() else {
        panic!("expected rebound shared style")
    };
    assert!(
        replayed
            .workbook()
            .styles()
            .expect("replayed styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        source
            .styles()
            .expect("source styles")
            .find(replayed_key)
            .is_none()
    );

    let mut changed_package = source.inner.package.clone();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let changed_xml = {
        let styles = changed_package.get_part(&styles_uri).expect("styles part");
        std::str::from_utf8(styles.blob())
            .expect("UTF-8 styles")
            .replace("FFFFFF00", "FFFF0000")
            .into_bytes()
    };
    changed_package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(changed_xml);
    let changed = Workbook::from_package(changed_package).expect("changed style table");
    assert!(matches!(
        changed.apply(committed.patch()),
        Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
    ));
}

#[test]
fn column_visibility_is_checked_reversible_and_composable() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
    sheet.set("A1", "left").expect("A1");
    sheet.set("C1", "right").expect("C1");
    sheet.column(1).expect("column B").hide();
    let committed = edit.commit().expect("commit");

    assert_eq!(committed.patch().len(), 3);
    assert!(committed.patch().changes().iter().any(|change| matches!(
        change,
        Change::Column {
            column,
            before: ColumnState::Missing,
            after: ColumnState::Stored(properties),
            ..
        } if column.get() == 1 && properties.hidden()
    )));
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    let column = sheet.column(1).expect("column B");
    assert!(column.stored());
    assert!(column.hidden());
    assert_eq!(sheet.columns().expect("columns").count(), 1);
    assert!(matches!(
        sheet.column_style(1).expect("column style"),
        Some(crate::LocalStyle::Default)
    ));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut show = committed.workbook().edit().expect("show edit");
    show.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column(ColumnIndex::new(1).expect("column B"))
        .expect("checked column")
        .show();
    let shown = show.commit().expect("show commit");
    let shown_sheet = shown
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    let shown_column = shown_sheet.column(1).expect("column B");
    assert!(shown_column.stored());
    assert!(!shown_column.hidden());

    let mut no_op = source.edit().expect("no-op edit");
    let mut sheet = no_op.sheet(0usize).expect("lookup").expect("sheet");
    sheet.column(10).expect("column K").show();
    assert!(matches!(
        sheet.column(litchi_sheet::COLUMNS),
        Err(Error::Coordinate(_))
    ));
    assert!(no_op.commit().expect("no-op commit").patch().is_empty());

    let mut cell = source.edit().expect("cell edit");
    cell.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("B1", "orthogonal")
        .expect("B1");
    let mut column = source.edit().expect("column edit");
    column
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column(1)
        .expect("column B")
        .hide();
    cell.join(column).expect("cell and column join");
    assert!(
        cell.commit()
            .expect("joined commit")
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(1)
            .expect("column B")
            .hidden()
    );

    let mut left = source.edit().expect("left");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column(4)
        .expect("column E")
        .hide();
    let mut right = source.edit().expect("right");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column(4)
        .expect("column E")
        .show();
    let error = left.join(right).expect_err("same column must conflict");
    assert_eq!(
        error.conflicts().expect("conflicts").conflicts()[0]
            .columns()
            .expect("column conflict"),
        &[ColumnIndex::new(4).expect("column E")]
    );
}

#[test]
fn column_layout_is_selector_first_typed_reversible_and_facet_composable() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .column("B")
        .expect("A1 column selector")
        .width(18.5)
        .expect("checked width")
        .outline(2)
        .expect("checked outline")
        .collapse()
        .best_fit()
        .show_phonetic();
    let committed = edit.commit().expect("layout commit");

    assert_eq!(committed.patch().len(), 1);
    let (_, before, after) = committed.patch().changes()[0]
        .column()
        .expect("column change");
    assert!(matches!(before, ColumnState::Missing));
    let ColumnState::Stored(properties) = after else {
        panic!("expected stored column properties")
    };
    assert_eq!(
        properties.width().map(crate::column::Width::get),
        Some(18.5)
    );
    assert_eq!(properties.outline().get(), 2);
    assert!(properties.collapsed());
    assert!(properties.best_fit());
    assert!(properties.custom_width());
    assert!(properties.phonetic());
    assert!(!properties.hidden());
    assert!(matches!(properties.style(), StyleState::Default));

    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet");
    let column = sheet.column("b").expect("case-insensitive A1 column");
    assert_eq!(column.index().get(), 1);
    assert_eq!(column.width().map(crate::column::Width::get), Some(18.5));
    assert_eq!(column.outline().get(), 2);
    assert!(column.collapsed());
    assert!(column.best_fit());
    assert!(column.custom_width());
    assert!(column.phonetic());

    let mut reset = committed.workbook().edit().expect("reset edit");
    reset
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .column("B")
        .expect("column B")
        .reset_width()
        .fixed()
        .outline(0)
        .expect("outline reset")
        .expand()
        .hide_phonetic();
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    let reset_column = reset_sheet.column("B").expect("column B");
    assert_eq!(reset_column.width(), None);
    assert!(!reset_column.custom_width());
    assert!(!reset_column.best_fit());
    assert_eq!(reset_column.outline(), Outline::NONE);
    assert!(!reset_column.collapsed());
    assert!(!reset_column.phonetic());

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut invalid = source.edit().expect("invalid edit");
    let mut sheet = invalid.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.column("XFE"),
        Err(Error::Coordinate(
            litchi_sheet::CoordinateError::ColumnA1 { .. }
        ))
    ));
    assert!(matches!(
        sheet.column("B").expect("B").width(f64::NAN),
        Err(Error::ColumnWidth(_))
    ));
    assert!(matches!(
        sheet.column("B").expect("B").outline(8),
        Err(Error::Outline(_))
    ));

    let mut width = source.edit().expect("width edit");
    width
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column("C")
        .expect("column C")
        .width(crate::column::Width::new(22.0).expect("prevalidated width"))
        .expect("width");
    let mut visibility = source.edit().expect("visibility edit");
    visibility
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column("C")
        .expect("column C")
        .hide();
    width
        .join(visibility)
        .expect("disjoint facets on one column");
    let joined = width.commit().expect("joined commit");
    let joined_sheet = joined
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    let column = joined_sheet.column("C").expect("column C");
    assert!(column.hidden());
    assert_eq!(column.width().map(crate::column::Width::get), Some(22.0));

    let mut left = source.edit().expect("left width");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column("D")
        .expect("column D")
        .width(10.0)
        .expect("width");
    let mut right = source.edit().expect("right width");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column("D")
        .expect("column D")
        .reset_width();
    assert!(left.join(right).is_err());
}

#[test]
fn column_layout_patch_guards_and_rebinds_hidden_shared_style_identity() {
    let source = styled_column_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .column("C")
        .expect("column C")
        .width(30.0)
        .expect("width");
    let committed = edit.commit().expect("commit");
    let (_, _, after) = committed.patch().changes()[0]
        .column()
        .expect("column change");
    let ColumnState::Stored(properties) = after else {
        panic!("expected stored properties")
    };
    assert!(matches!(properties.style(), StyleState::Shared(_)));

    let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
    let replayed = reopened
        .apply(committed.patch())
        .expect("source-checked replay");
    let (_, _, replayed_after) = replayed.patch().changes()[0]
        .column()
        .expect("replayed column change");
    let ColumnState::Stored(replayed_properties) = replayed_after else {
        panic!("expected replayed properties")
    };
    let StyleState::Shared(replayed_key) = replayed_properties.style() else {
        panic!("expected rebound shared style")
    };
    assert!(
        replayed
            .workbook()
            .styles()
            .expect("replayed styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        source
            .styles()
            .expect("source styles")
            .find(replayed_key)
            .is_none()
    );

    let mut changed_package = source.inner.package.clone();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let changed_xml = {
        let styles = changed_package.get_part(&styles_uri).expect("styles part");
        std::str::from_utf8(styles.blob())
            .expect("UTF-8 styles")
            .replace("FFFFFF00", "FFFF0000")
            .into_bytes()
    };
    changed_package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(changed_xml);
    let changed = Workbook::from_package(changed_package).expect("changed style table");
    assert!(matches!(
        changed.apply(committed.patch()),
        Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
    ));
}

#[test]
fn grid_default_styles_are_lineage_checked_reversible_and_facet_composable() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let accent = source
        .sheet("Sheet1")
        .expect("sheet lookup")
        .expect("worksheet")
        .style("A1")
        .expect("style lookup")
        .expect("accent style");

    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet
            .row(1)
            .expect("row 2")
            .style(&accent)
            .expect("row style")
            .height(24)
            .expect("row height");
        sheet
            .column("C")
            .expect("column C")
            .style(&accent)
            .expect("column style")
            .width(16)
            .expect("column width");
    }
    let committed = edit.commit().expect("style commit");
    assert_eq!(committed.patch().len(), 2);

    let row_change = committed
        .patch()
        .changes()
        .iter()
        .find_map(Change::row)
        .expect("row change");
    let RowState::Stored(row_after) = row_change.2 else {
        panic!("expected stored row")
    };
    assert!(row_after.custom_format());
    assert!(matches!(row_after.style(), StyleState::Shared(_)));

    let column_change = committed
        .patch()
        .changes()
        .iter()
        .find_map(Change::column)
        .expect("column change");
    let ColumnState::Stored(column_after) = column_change.2 else {
        panic!("expected stored column")
    };
    assert!(matches!(column_after.style(), StyleState::Shared(_)));

    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(sheet.row(1).expect("row 2").custom_format());
    assert!(matches!(
        sheet.row_style(1).expect("row style"),
        Some(crate::LocalStyle::Shared(style)) if style.same(&accent)
    ));
    assert!(matches!(
        sheet.column_style("C").expect("column style"),
        Some(crate::LocalStyle::Shared(style)) if style.same(&accent)
    ));

    let mut reset = committed.workbook().edit().expect("reset edit");
    {
        let mut sheet = reset.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet.row(1).expect("row 2").reset_style();
        sheet.column("C").expect("column C").reset_style();
    }
    let reset = reset.commit().expect("reset commit");
    let reset_sheet = reset
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(!reset_sheet.row(1).expect("row 2").custom_format());
    assert!(matches!(
        reset_sheet.row_style(1).expect("row style"),
        Some(crate::LocalStyle::Default)
    ));
    assert!(matches!(
        reset_sheet.column_style("C").expect("column style"),
        Some(crate::LocalStyle::Default)
    ));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let reopened = Workbook::from_bytes(source_bytes).expect("reopened source");
    let replayed = reopened
        .apply(committed.patch())
        .expect("source-checked replay");
    let (_, _, replayed_row) = replayed
        .patch()
        .changes()
        .iter()
        .find_map(Change::row)
        .expect("replayed row");
    let RowState::Stored(replayed_row) = replayed_row else {
        panic!("expected replayed row")
    };
    let StyleState::Shared(replayed_key) = replayed_row.style() else {
        panic!("expected rebound row style")
    };
    assert!(
        replayed
            .workbook()
            .styles()
            .expect("styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        source
            .styles()
            .expect("source styles")
            .find(replayed_key)
            .is_none()
    );

    let mut styles = source.edit().expect("styles edit");
    {
        let mut sheet = styles.sheet(0usize).expect("lookup").expect("sheet");
        sheet
            .row(2)
            .expect("row 3")
            .style(&accent)
            .expect("row style");
        sheet
            .column("D")
            .expect("column D")
            .style(&accent)
            .expect("column style");
    }
    let mut layout = source.edit().expect("layout edit");
    {
        let mut sheet = layout.sheet(0usize).expect("lookup").expect("sheet");
        sheet.row(2).expect("row 3").height(22).expect("height");
        sheet
            .column("D")
            .expect("column D")
            .width(18)
            .expect("width");
    }
    styles.join(layout).expect("disjoint grid facets");
    let joined = styles.commit().expect("joined commit");
    let joined_sheet = joined
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert_eq!(
        joined_sheet
            .row(2)
            .expect("row 3")
            .height()
            .map(crate::row::Height::get),
        Some(22.0)
    );
    assert!(matches!(
        joined_sheet.column_style("D").expect("column style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let mut left = source.edit().expect("left style");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(3)
        .expect("row 4")
        .style(&accent)
        .expect("style");
    let mut right = source.edit().expect("right style");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(3)
        .expect("row 4")
        .reset_style();
    assert!(left.join(right).is_err());

    let mut missing_width = source.edit().expect("missing-width edit");
    missing_width
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .column("E")
        .expect("column E")
        .style(&accent)
        .expect("lineage");
    assert!(matches!(
        missing_width.commit(),
        Err(Error::ColumnEditBlocked {
            reason: crate::error::ColumnEditBlock::StyleNeedsWidth,
            ..
        })
    ));

    let foreign = Workbook::new()
        .expect("foreign workbook")
        .styles()
        .expect("foreign styles")
        .base()
        .expect("foreign base style");
    let mut rejected = source.edit().expect("rejected edit");
    {
        let mut sheet = rejected.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.row(4).expect("row 5").style(&foreign),
            Err(Error::ForeignStyle)
        ));
        assert!(matches!(
            sheet.column("E").expect("column E").style(&foreign),
            Err(Error::ForeignStyle)
        ));
    }
    assert!(rejected.is_empty());

    let mut add = source.edit().expect("new sheet edit");
    {
        let mut sheet = add.add("Styled").expect("new sheet");
        sheet.set("A2", "row").expect("row cell");
        sheet.set("C1", "column").expect("column cell");
        sheet
            .row(1)
            .expect("row 2")
            .style(&accent)
            .expect("row style");
        sheet
            .column("C")
            .expect("column C")
            .width(12)
            .expect("column width")
            .style(&accent)
            .expect("column style");
    }
    let added = add.commit().expect("new sheet commit");
    let sheet = added
        .workbook()
        .sheet("Styled")
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.row_style(1).expect("row style"),
        Some(crate::LocalStyle::Shared(_))
    ));
    assert!(matches!(
        sheet.column_style("C").expect("column style"),
        Some(crate::LocalStyle::Shared(_))
    ));
}

#[test]
fn tab_visibility_is_selector_first_reversible_and_active_safe() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let source_bytes = source.to_bytes().expect("source bytes");
    assert_eq!(
        source.active_sheet().map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".to_owned())
    );

    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet1")
        .expect("name lookup")
        .expect("tab")
        .hide();
    let hidden = edit.commit().expect("hide active tab");
    assert_eq!(hidden.patch().len(), 2);
    let (before, after) = hidden.patch().changes()[0]
        .active()
        .expect("implicit active relocation");
    assert_eq!((before.name(), before.position()), ("Sheet1", 0));
    assert_eq!((after.name(), after.position()), ("Sheet2", 1));
    assert!(matches!(
        hidden.patch().changes()[1],
        Change::Visibility {
            position: 0,
            before: Visibility::Visible,
            after: Visibility::Hidden,
            ..
        }
    ));
    assert_eq!(
        hidden
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
    assert_eq!(
        hidden
            .workbook()
            .active_sheet()
            .map(|sheet| sheet.name().to_owned()),
        Some("Sheet2".to_owned())
    );
    assert!(
        hidden
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .is_active()
    );

    let restored = hidden
        .workbook()
        .apply(&hidden.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);

    let mut last = hidden.workbook().edit().expect("last visible edit");
    last.tab(1usize)
        .expect("position lookup")
        .expect("tab")
        .very_hide();
    assert!(matches!(
        last.commit(),
        Err(Error::TabEditBlocked {
            sheet,
            position: 1,
            reason: TabEditBlock::LastVisibleTab,
        }) if sheet == "Sheet2"
    ));

    let mut swap = hidden.workbook().edit().expect("swap visibility");
    swap.tab("Sheet1").expect("lookup").expect("tab").show();
    swap.tab(1usize).expect("lookup").expect("tab").very_hide();
    let swapped = swap.commit().expect("swap commit");
    assert_eq!(
        swapped
            .workbook()
            .sheet(1usize)
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::VeryHidden
    );
    assert_eq!(
        swapped
            .workbook()
            .active_sheet()
            .map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".to_owned())
    );

    let mut no_op = source.edit().expect("no-op edit");
    no_op.tab(0usize).expect("lookup").expect("tab").show();
    assert!(no_op.commit().expect("no-op commit").patch().is_empty());
}

#[test]
fn active_tab_is_selector_first_reversible_and_composable() {
    for kind in [WorksheetKind::Worksheet, WorksheetKind::Chart] {
        let source = two_sheet_workbook(kind);
        let source_bytes = source.to_bytes().expect("source bytes");
        assert!(
            source
                .sheet("Sheet1")
                .expect("lookup")
                .expect("first tab")
                .is_active()
        );
        assert!(
            !source
                .sheet(1usize)
                .expect("lookup")
                .expect("second tab")
                .is_active()
        );

        let mut edit = source.edit().expect("edit");
        edit.tab(1usize)
            .expect("position lookup")
            .expect("tab")
            .activate();
        assert_eq!(edit.len(), 1);
        let committed = edit.commit().expect("activate");
        assert_eq!(committed.patch().len(), 1);
        let (before, after) = committed.patch().changes()[0]
            .active()
            .expect("active change");
        assert_eq!((before.name(), before.position()), ("Sheet1", 0));
        assert_eq!((after.name(), after.position()), ("Sheet2", 1));
        let active = committed.workbook().active_sheet().expect("active sheet");
        assert_eq!(active.name(), "Sheet2");
        assert_eq!(active.kind(), kind);
        assert!(active.is_active());
        assert_eq!(
            committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );

        let mut no_op = committed.workbook().edit().expect("no-op edit");
        no_op
            .tab("Sheet2")
            .expect("lookup")
            .expect("tab")
            .activate();
        assert!(no_op.commit().expect("no-op commit").patch().is_empty());
    }

    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut cell = source.edit().expect("cell edit");
    cell.sheet("Sheet2")
        .expect("lookup")
        .expect("worksheet")
        .set("A1", "active payload")
        .expect("cell");
    let mut active = source.edit().expect("active edit");
    active
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    cell.join(active).expect("orthogonal join");
    let committed = cell.commit().expect("joined commit");
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.sheets[1].part_uri)
            .count(),
        1
    );
    let active = committed.workbook().active_sheet().expect("active sheet");
    assert_eq!(active.name(), "Sheet2");
    assert!(matches!(
        active.cell("A1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "active payload"
    ));

    let mut replaced = source.edit().expect("replacement edit");
    replaced
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    replaced
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .activate();
    assert_eq!(replaced.len(), 1);
    assert!(
        replaced
            .commit()
            .expect("last activation wins")
            .patch()
            .is_empty()
    );
}

#[test]
fn active_tab_requires_final_visibility_and_conflicts_globally() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut hide = source.edit().expect("hide edit");
    hide.tab("Sheet2").expect("lookup").expect("tab").hide();
    let hidden = hide.commit().expect("hidden source");

    let mut blocked = hidden.workbook().edit().expect("blocked edit");
    blocked
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    assert!(matches!(
        blocked.commit(),
        Err(Error::TabEditBlocked {
            sheet,
            position: 1,
            reason: TabEditBlock::NotVisible,
        }) if sheet == "Sheet2"
    ));

    let mut repaired = hidden.workbook().edit().expect("repair edit");
    repaired
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .show()
        .activate();
    let repaired = repaired.commit().expect("show and activate");
    let active = repaired.workbook().active_sheet().expect("active sheet");
    assert_eq!(active.name(), "Sheet2");
    assert!(active.visibility().is_visible());
    assert_eq!(repaired.patch().len(), 2);

    let mut contradictory = source.edit().expect("contradictory edit");
    contradictory
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate()
        .very_hide();
    assert!(matches!(
        contradictory.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::NotVisible,
            ..
        })
    ));

    let mut left = source.edit().expect("left");
    left.tab("Sheet2").expect("lookup").expect("tab").activate();
    let mut right = source.edit().expect("right");
    right
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .activate();
    let error = left.join(right).expect_err("active-tab intents are global");
    let conflicts = error.conflicts().expect("active conflict");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_active());
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");

    let mut active = source.edit().expect("active");
    active
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let mut visibility = source.edit().expect("visibility");
    visibility
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .hide();
    active.join(visibility).expect("orthogonal facets");
    let committed = active.commit().expect("joined commit");
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Sheet2"
    );
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
}

#[test]
fn tab_rename_is_typed_dependency_aware_move_first_and_reversible() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_part = source.inner.sheets[0].part_uri.clone();
    let source_id = source.inner.sheets[0].native_id;
    let source_relationship = source.inner.sheets[0].relationship_id.clone();

    let mut edit = source.edit().expect("edit");
    edit.tab("data")
        .expect("caseless lookup")
        .expect("tab")
        .rename(String::from("Input 2026"))
        .expect("checked rename");
    edit.sheet("Calc")
        .expect("Calc lookup")
        .expect("Calc sheet")
        .set("D1", "composed")
        .expect("same-part cell edit");
    let committed = edit.commit().expect("rename commit");
    assert!(
        committed
            .workbook()
            .sheet("Data")
            .expect("lookup")
            .is_none()
    );
    let renamed = committed
        .workbook()
        .sheet("INPUT 2026")
        .expect("Unicode caseless lookup")
        .expect("renamed sheet");
    assert_eq!(renamed.name(), "Input 2026");
    assert_eq!(renamed.data.part_uri, source_part);
    assert_eq!(renamed.data.native_id, source_id);
    assert_eq!(renamed.data.relationship_id, source_relationship);
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed.patch().changes()[0].renamed(),
        Some((0, "Data", "Input 2026"))
    );

    let calc = committed
        .workbook()
        .sheet("Calc")
        .expect("lookup")
        .expect("Calc");
    assert!(matches!(
        calc.cell("A1").expect("formula cell").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "'Input 2026'!A1"
    ));
    assert!(matches!(
        calc.cell("D1").expect("composed cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "composed"
    ));
    for uri in [
        "/xl/workbook.xml",
        "/xl/worksheets/sheet2.xml",
        "/xl/tables/table1.xml",
        "/xl/charts/chart1.xml",
        "/xl/pivotCache/pivotCacheDefinition1.xml",
        "/docProps/app.xml",
    ] {
        let text = part_text(committed.workbook(), uri);
        assert!(
            text.contains("Input 2026"),
            "expected renamed dependency in {uri}: {text}"
        );
    }
    assert!(
        part_text(committed.workbook(), "/xl/externalLinks/externalLink1.xml")
            .contains("[1]Data!A1")
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse rename");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
}

#[test]
fn worksheet_add_is_atomic_populatable_active_and_reversible() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.add(String::from("Summary")).expect("new sheet");
        assert_eq!(sheet.name(), "Summary");
        assert_eq!(sheet.position(), 1);
        sheet
            .set("A1", "created in one transaction")
            .and_then(|sheet| sheet.set("B2", Formula::new("1+1").expect("checked formula")))
            .expect("new cells");
        sheet.row(3u32).expect("row").hide();
        sheet.column(2u32).expect("column").hide();
        sheet.activate();
    }
    let committed = edit.commit().expect("create commit");
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Summary"]
    );
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Summary"
    );
    let summary = committed
        .workbook()
        .sheet("summary")
        .expect("caseless lookup")
        .expect("created sheet");
    assert_eq!(summary.data.native_id, 2);
    assert!(matches!(
        summary.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "created in one transaction"
    ));
    assert!(matches!(
        summary.cell("B2").expect("B2").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "1+1"
    ));
    assert!(summary.row(3u32).expect("row").hidden());
    assert!(summary.column(2u32).expect("column").hidden());
    assert!(
        committed
            .patch()
            .changes()
            .iter()
            .any(|change| { change.created() == Some((1, "Summary")) })
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse create");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    let replayed = source.apply(committed.patch()).expect("forward replay");
    assert!(
        replayed
            .workbook()
            .sheet("Summary")
            .expect("lookup")
            .is_some()
    );
}

#[test]
fn worksheet_insert_is_selector_first_order_aware_and_reversible() {
    let source = three_sheet_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("move lookup")
        .expect("both base tabs");
    assert!(
        edit.add_before("Never Added", "Absent")
            .expect("missing lookup")
            .is_none()
    );
    {
        let mut sheet = edit
            .add_before("Before A", "Sheet1")
            .expect("lookup")
            .expect("anchor");
        assert_eq!(sheet.position(), 1);
        sheet.set("A1", "before-a").expect("payload");
    }
    edit.add_before("Before B", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "before-b")
        .expect("payload");
    edit.add_after("After A", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "after-a")
        .expect("payload");
    edit.add_after("After B", "Sheet1")
        .expect("lookup")
        .expect("anchor")
        .set("A1", "after-b")
        .expect("payload")
        .activate();
    {
        let sheet = edit
            .add_before("Before Third", 2usize)
            .expect("numeric lookup")
            .expect("numeric anchor");
        assert_eq!(sheet.position(), 0);
    }
    edit.add("Tail").expect("tail");

    let committed = edit.commit().expect("insert commit");
    let names = committed
        .workbook()
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(
        names,
        [
            "Before Third",
            "Sheet3",
            "Before A",
            "Before B",
            "Sheet1",
            "After A",
            "After B",
            "Sheet2",
            "Tail",
        ]
    );
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("After B", 6));
    assert!(matches!(
        committed
            .workbook()
            .sheet("Before A")
            .expect("lookup")
            .expect("created")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "before-a"
    ));
    assert_eq!(
        committed
            .workbook()
            .defined_names()
            .iter()
            .map(|name| (name.name.as_str(), name.local_sheet_id))
            .collect::<Vec<_>>(),
        [
            ("FirstLocal", Some(4)),
            ("ThirdLocal", Some(1)),
            ("Global", None),
        ]
    );
    assert_eq!(
        committed
            .patch()
            .changes()
            .iter()
            .filter_map(Change::created)
            .collect::<Vec<_>>(),
        [
            (2, "Before A"),
            (3, "Before B"),
            (5, "After A"),
            (6, "After B"),
            (0, "Before Third"),
            (8, "Tail"),
        ]
    );

    let committed_bytes = committed.workbook().to_bytes().expect("committed bytes");
    assert_eq!(
        source
            .apply(committed.patch())
            .expect("forward replay")
            .workbook()
            .to_bytes()
            .expect("replayed bytes"),
        committed_bytes
    );
    assert_eq!(
        committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn worksheet_insert_join_preserves_explicit_edit_order() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut left = source.edit().expect("left");
    left.add_before("Left", "Sheet2")
        .expect("lookup")
        .expect("anchor");
    let mut right = source.edit().expect("right");
    right
        .add_before("Right", 1usize)
        .expect("lookup")
        .expect("anchor");
    left.join(right).expect("disjoint insertion join");
    assert_eq!(
        left.commit()
            .expect("joined insertion")
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Left", "Right", "Sheet2"]
    );
}

#[test]
fn worksheet_add_validates_names_visibility_and_parallel_joins() {
    let source = Workbook::new().expect("source workbook");

    let mut duplicate = source.edit().expect("duplicate edit");
    duplicate.add("sheet1").expect("checked spelling");
    assert!(matches!(
        duplicate.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut duplicate = source.edit().expect("anchored duplicate edit");
    duplicate
        .add_before("sheet1", "Sheet1")
        .expect("lookup")
        .expect("anchor");
    assert!(matches!(
        duplicate.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut hidden = source.edit().expect("hidden edit");
    {
        let mut sheet = hidden.add("Hidden").expect("new sheet");
        sheet.hide().activate();
    }
    assert!(matches!(
        hidden.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::NotVisible,
            ..
        })
    ));

    let mut replacement = source.edit().expect("replacement edit");
    replacement
        .tab("Sheet1")
        .expect("lookup")
        .expect("existing tab")
        .hide();
    replacement.add("Replacement").expect("visible replacement");
    let replaced = replacement.commit().expect("replacement commit");
    assert_eq!(
        replaced.workbook().active_sheet().expect("active").name(),
        "Replacement"
    );
    assert!(matches!(
        replaced
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("old tab")
            .visibility(),
        Visibility::Hidden
    ));

    let mut left = source.edit().expect("left");
    left.add("North")
        .expect("North")
        .set("A1", 1_i32)
        .expect("cell");
    let mut right = source.edit().expect("right");
    right
        .add("South")
        .expect("South")
        .set("A1", 2_i32)
        .expect("cell");
    left.join(right).expect("disjoint appends join");
    let joined = left.commit().expect("joined create");
    assert_eq!(
        joined
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "North", "South"]
    );

    let mut one = source.edit().expect("one");
    one.add("Résumé").expect("first name");
    let mut two = source.edit().expect("two");
    two.add("RE\u{301}SUME\u{301}").expect("equivalent name");
    let error = one.join(two).expect_err("equivalent creations conflict");
    assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());
}

#[test]
fn worksheet_add_closes_rename_formula_and_extended_property_dependencies() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.tab("Data")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("rename");
    edit.add("New & Sheet")
        .expect("new sheet")
        .set(
            "A1",
            Formula::new("Data!A1").expect("source-snapshot formula"),
        )
        .expect("formula");
    let committed = edit.commit().expect("composed structural commit");
    let created = committed
        .workbook()
        .sheet("New & Sheet")
        .expect("lookup")
        .expect("created sheet");
    assert!(matches!(
        created.cell("A1").expect("formula").stored(),
        Some(Cell::Formula(formula)) if formula.text() == "Input!A1"
    ));
    let properties = part_text(committed.workbook(), "/docProps/app.xml");
    assert!(properties.contains("size=\"4\""));
    assert!(properties.contains(
            "<vt:lpstr>Input</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>New &amp; Sheet</vt:lpstr><vt:lpstr>Input!Print_Area</vt:lpstr>"
        ));

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
}

#[test]
fn worksheet_add_synchronizes_optional_properties_during_a_simultaneous_reorder() {
    let source = rename_reference_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.move_before("Calc", "Data")
        .expect("move")
        .expect("both tabs");
    edit.add_before("Middle", "Data")
        .expect("lookup")
        .expect("anchor");
    edit.add("Tail").expect("new sheet");
    let committed = edit.commit().expect("composed commit");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Calc", "Middle", "Data", "Tail"]
    );
    let properties = part_text(committed.workbook(), "/docProps/app.xml");
    assert!(properties.contains("size=\"5\""));
    assert!(properties.contains(concat!(
        "<vt:lpstr>Calc</vt:lpstr>",
        "<vt:lpstr>Middle</vt:lpstr>",
        "<vt:lpstr>Data</vt:lpstr>",
        "<vt:lpstr>Tail</vt:lpstr>",
        "<vt:lpstr>Data!Print_Area</vt:lpstr>"
    )));
    assert_eq!(
        committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn worksheet_add_allocates_strict_graph_identity_without_exposing_it() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let main = package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part");
    main.set_blob(
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Strict1" sheetId="7" r:id="tab"/></sheets></workbook>"#.to_vec(),
        );
    main.rels_mut().remove("rId1").expect("old worksheet rel");
    main.rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::STRICT_WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "tab".to_owned(),
            TargetMode::Internal,
        )
        .expect("strict worksheet rel");
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("URI"))
            .expect("worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/></worksheet>"#.to_vec(),
        );
    let source = Workbook::from_package(package).expect("strict source");
    let mut edit = source.edit().expect("edit");
    edit.add_before("Strict2", "Strict1")
        .expect("lookup")
        .expect("strict anchor");
    let committed = edit.commit().expect("strict create");
    let sheet = committed
        .workbook()
        .sheet("Strict2")
        .expect("lookup")
        .expect("created");
    assert_eq!(sheet.position(), 0);
    assert_eq!(sheet.data.native_id, 1);
    assert_eq!(sheet.data.relationship_id, "rId1");
    assert_eq!(sheet.data.part_uri.as_str(), "/xl/worksheets/sheet2.xml");
    let main = committed
        .workbook()
        .inner
        .package
        .get_part(&committed.workbook().inner.workbook_uri)
        .expect("workbook part");
    assert_eq!(
        main.rels()
            .get(&sheet.data.relationship_id)
            .expect("new relationship")
            .reltype(),
        litchi_opc::constants::relationship_type::STRICT_WORKSHEET
    );
    assert!(
        part_text(committed.workbook(), sheet.data.part_uri.as_str())
            .contains("http://purl.oclc.org/ooxml/spreadsheetml/main")
    );
}

#[test]
fn worksheet_add_blocks_protected_and_compatibility_owned_catalogs() {
    for xml in [
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.as_slice(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><mc:AlternateContent><mc:Choice Requires="x"><x:payload/></mc:Choice><mc:Fallback/></mc:AlternateContent></sheets></workbook>"#.as_slice(),
        ] {
            let baseline = Workbook::new().expect("baseline");
            let mut package = baseline.inner.package.clone();
            package
                .get_part_mut(&baseline.inner.workbook_uri)
                .expect("workbook")
                .set_blob(xml.to_vec());
            let source = Workbook::from_package(package).expect("source");
            let mut edit = source.edit().expect("edit");
            edit.add_before("Blocked", "Sheet1")
                .expect("lookup")
                .expect("anchor");
            assert!(matches!(
                edit.commit(),
                Err(Error::TabEditBlocked {
                    reason: TabEditBlock::ProtectedWorkbook
                        | TabEditBlock::MarkupCompatibility,
                    ..
                })
            ));
        }
}

#[test]
fn worksheet_remove_is_selector_first_active_relocating_and_reversible() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create
        .add("Delete")
        .expect("Delete")
        .set("A1", "removed payload")
        .expect("payload")
        .activate();
    create
        .add("Keep")
        .expect("Keep")
        .set("A1", 42_i32)
        .expect("retained payload");
    let source = create.commit().expect("create source").into_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    assert_eq!(source.active_sheet().expect("active").name(), "Delete");

    let mut edit = source.edit().expect("remove edit");
    assert!(edit.remove("missing").expect("missing selector").is_none());
    edit.remove("delete")
        .expect("selector")
        .expect("Delete worksheet");
    assert_eq!(edit.len(), 1);
    let committed = edit.commit().expect("remove commit");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1", "Keep"]
    );
    assert_eq!(
        committed.workbook().active_sheet().expect("active").name(),
        "Keep"
    );
    assert!(
        committed
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );
    assert!(matches!(
        committed.patch().changes().first(),
        Some(Change::Remove {
            sheet,
            position: 1,
            ..
        }) if sheet.as_ref() == "Delete"
    ));
    assert!(
            committed
                .patch()
                .changes()
                .iter()
                .any(|change| matches!(change, Change::Active { after, .. } if after.name() == "Keep" && after.position() == 1))
        );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse remove");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    let replayed = source.apply(committed.patch()).expect("forward replay");
    assert!(
        replayed
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );

    let mut activate_last = source.edit().expect("activate last");
    activate_last
        .tab("Keep")
        .expect("lookup")
        .expect("Keep")
        .activate();
    let active_last = activate_last
        .commit()
        .expect("active-last source")
        .into_workbook();
    let mut remove_last = active_last.edit().expect("remove last");
    remove_last.remove("Keep").expect("lookup").expect("Keep");
    assert_eq!(
        remove_last
            .commit()
            .expect("remove last active")
            .workbook()
            .active_sheet()
            .expect("replacement active")
            .name(),
        "Delete"
    );
}

#[test]
fn worksheet_remove_blocks_live_formulas_last_sheet_and_mixed_edits() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    create
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set("A1", Formula::new("Delete!A1").expect("formula"))
        .expect("formula cell");
    let source = create.commit().expect("source").into_workbook();
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            sheet,
            reason: RemoveBlock::IncomingReference,
            ..
        }) if sheet == "Delete"
    ));

    let single = Workbook::new().expect("single");
    let mut last = single.edit().expect("last edit");
    last.remove(0usize).expect("lookup").expect("only sheet");
    assert!(matches!(
        last.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::LastSheet,
            ..
        })
    ));

    let baseline = Workbook::new().expect("visibility baseline");
    let mut create = baseline.edit().expect("visibility create");
    create
        .tab("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .hide();
    create.add("Delete").expect("Delete").activate();
    let visibility = create.commit().expect("visibility source").into_workbook();
    let mut last_visible = visibility.edit().expect("last-visible edit");
    last_visible
        .remove("Delete")
        .expect("lookup")
        .expect("Delete");
    assert!(matches!(
        last_visible.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::LastVisibleTab,
            ..
        })
    ));

    let mut mixed = source.edit().expect("mixed edit");
    mixed
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set("B1", 1_i32)
        .expect("cell edit");
    assert!(matches!(
        mixed.remove("Delete"),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::MixedEdit,
            ..
        })
    ));

    let baseline = Workbook::new().expect("dynamic baseline");
    let mut create = baseline.edit().expect("dynamic create");
    create.add("Delete").expect("Delete");
    create
        .sheet("Sheet1")
        .expect("lookup")
        .expect("Sheet1")
        .set(
            "A1",
            Formula::new(r#"INDIRECT("Delete!A1")"#).expect("dynamic formula"),
        )
        .expect("dynamic cell");
    let dynamic = create.commit().expect("dynamic source").into_workbook();
    let mut remove = dynamic.edit().expect("dynamic removal");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::UnmodeledReference,
            ..
        })
    ));
}

#[test]
fn worksheet_remove_joins_disjoint_plans_and_blocks_unknown_dependencies() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("North").expect("North");
    create.add("South").expect("South");
    let source = create.commit().expect("source").into_workbook();

    let mut north = source.edit().expect("north edit");
    north.remove("North").expect("lookup").expect("North");
    let mut south = source.edit().expect("south edit");
    south.remove(2usize).expect("lookup").expect("South");
    north.join(south).expect("disjoint removals join");
    let committed = north.commit().expect("joined removal");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet1"]
    );

    let mut package = source.inner.package.clone();
    let custom_uri = PackURI::new("/customXml/item1.xml").expect("custom URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            custom_uri,
            "application/xml".to_owned(),
            b"<root><futureFormulaCache>North!A1</futureFormulaCache></root>".to_vec(),
        )))
        .expect("custom part");
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-custom".to_owned(),
            "../customXml/item1.xml".to_owned(),
            "customRef".to_owned(),
            TargetMode::Internal,
        )
        .expect("custom relationship");
    let unknown = Workbook::from_package(package).expect("unknown producer workbook");
    let mut edit = unknown.edit().expect("remove edit");
    edit.remove("North").expect("lookup").expect("North");
    assert!(matches!(
        edit.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::UnmodeledReference,
            part,
            ..
        }) if part == "/customXml/item1.xml"
    ));
}

#[test]
fn worksheet_remove_blocks_macro_projects_and_extra_incoming_relationships() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();

    let mut macro_package = source.inner.package.clone();
    macro_package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/vbaProject.bin").expect("VBA URI"),
            litchi_opc::constants::content_type::OFC_VBA_PROJECT.to_owned(),
            vec![0, 1, 2, 3],
        )))
        .expect("VBA part");
    macro_package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::VBA_PROJECT.to_owned(),
            "vbaProject.bin".to_owned(),
            "vbaProject".to_owned(),
            TargetMode::Internal,
        )
        .expect("VBA relationship");
    let macro_book = Workbook::from_package(macro_package).expect("macro workbook");
    let mut remove = macro_book.edit().expect("macro remove");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::MacroProject,
            ..
        })
    ));

    let target = source
        .inner
        .sheets
        .get(1)
        .expect("Delete sheet")
        .part_uri
        .clone();
    let mut incoming_package = source.inner.package.clone();
    let mut referrer = BlobPart::new(
        PackURI::new("/xl/custom.xml").expect("custom URI"),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    referrer
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-incoming".to_owned(),
            target.relative_ref("/xl"),
            "sheetRef".to_owned(),
            TargetMode::Internal,
        )
        .expect("incoming relationship");
    incoming_package
        .try_add_part(Box::new(referrer))
        .expect("referrer part");
    let incoming = Workbook::from_package(incoming_package).expect("incoming workbook");
    let mut remove = incoming.edit().expect("incoming remove");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::IncomingRelationship,
            part,
            ..
        }) if part == "/xl/custom.xml"
    ));
}

#[test]
fn worksheet_remove_blocks_custom_workbook_view_identity() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();
    let native_id = source.inner.sheets[1].native_id;
    let mut package = source.inner.package.clone();
    let workbook = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook");
    let xml = std::str::from_utf8(workbook.blob())
            .expect("workbook UTF-8")
            .replace(
                "</workbook>",
                &format!(
                    r#"<customWorkbookViews><customWorkbookView name="Delete view" guid="{{00000000-0000-0000-0000-000000000001}}" activeSheetId="{native_id}"/></customWorkbookViews></workbook>"#
                ),
            );
    assert!(xml.contains("customWorkbookView"));
    workbook.set_blob(xml.into_bytes());
    let source = Workbook::from_package(package).expect("custom-view workbook");
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(matches!(
        remove.commit(),
        Err(Error::SheetRemoveBlocked {
            reason: RemoveBlock::IncomingReference,
            part,
            ..
        }) if part == "/xl/workbook.xml"
    ));
}

#[test]
fn worksheet_remove_accepts_case_equivalent_opc_targets() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("Delete");
    let source = create.commit().expect("source").into_workbook();
    let relationship_id = source.inner.sheets[1].relationship_id.clone();
    let mut package = source.inner.package.clone();
    let relationships = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut();
    let relationship = relationships
        .remove(&relationship_id)
        .expect("worksheet relationship");
    relationships
        .try_add_relationship(
            relationship.reltype().to_owned(),
            "worksheets/SHEET2.XML".to_owned(),
            relationship_id,
            TargetMode::Internal,
        )
        .expect("case-equivalent relationship");
    let source = Workbook::from_package(package).expect("case-equivalent workbook");
    let mut remove = source.edit().expect("remove edit");
    remove.remove("Delete").expect("lookup").expect("Delete");
    assert!(
        remove
            .commit()
            .expect("case-equivalent removal")
            .workbook()
            .sheet("Delete")
            .expect("lookup")
            .is_none()
    );
}

#[test]
fn tab_rename_validates_names_collisions_swaps_and_join_facets() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut invalid = source.edit().expect("invalid edit");
    let error = invalid
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("bad/name")
        .expect_err("invalid name");
    assert!(matches!(error, Error::SheetName(_)));

    let mut collision = source.edit().expect("collision edit");
    collision
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("sheet2")
        .expect("valid spelling");
    assert!(matches!(
        collision.commit(),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let mut swap = source.edit().expect("swap edit");
    swap.tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Sheet2")
        .expect("first swap");
    swap.tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .rename("Sheet1")
        .expect("second swap");
    let swapped = swap.commit().expect("simultaneous swap");
    assert_eq!(
        swapped
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet2", "Sheet1"]
    );

    let mut left = source.edit().expect("left");
    left.tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("First")
        .expect("left rename");
    let mut same = source.edit().expect("same");
    same.tab(0usize)
        .expect("lookup")
        .expect("tab")
        .rename("Other")
        .expect("same rename");
    let error = left.join(same).expect_err("same name facet conflicts");
    assert!(error.conflicts().expect("conflicts").conflicts()[0].is_name());

    let mut right = source.edit().expect("right");
    right
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .rename(Name::new("Second").expect("prevalidated"))
        .expect("moved typed name");
    left.join(right).expect("disjoint names join");
    let joined = left.commit().expect("joined renames");
    assert!(joined.workbook().sheet("first").expect("lookup").is_some());
    assert!(joined.workbook().sheet("second").expect("lookup").is_some());
}

#[test]
fn tab_reorder_is_selector_first_dependency_aware_and_reversible() {
    let source = three_sheet_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    assert!(
        edit.move_before("Sheet3", "Sheet1")
            .expect("move lookup")
            .is_some()
    );
    assert_eq!(edit.len(), 1);
    let committed = edit.commit().expect("reorder");
    assert_eq!(
        committed
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet3", "Sheet1", "Sheet2"]
    );
    assert_eq!(
        committed
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .name(),
        "Sheet2"
    );
    assert_eq!(
        committed
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .position(),
        2
    );
    assert!(matches!(
        committed
            .workbook()
            .sheet("Sheet3")
            .expect("lookup")
            .expect("Sheet3")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "three"
    ));
    assert_eq!(
        committed
            .workbook()
            .defined_names()
            .iter()
            .map(|name| (name.name.as_str(), name.local_sheet_id))
            .collect::<Vec<_>>(),
        [
            ("FirstLocal", Some(1)),
            ("ThirdLocal", Some(0)),
            ("Global", None),
        ]
    );
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(committed.patch().changes()[0].sheet(), "Sheet3");
    assert_eq!(committed.patch().changes()[0].moved(), Some((2, 0)));
    let (before, after) = committed.patch().changes()[1]
        .active()
        .expect("active position remap");
    assert_eq!((before.name(), before.position()), ("Sheet2", 1));
    assert_eq!((after.name(), after.position()), ("Sheet2", 2));

    let inverse = committed.patch().inverse();
    assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
    let restored = committed
        .workbook()
        .apply(&inverse)
        .expect("inverse reorder");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    assert_eq!(source.to_bytes().expect("source unchanged"), source_bytes);

    let chart_source = two_sheet_workbook(WorksheetKind::Chart);
    let chart_bytes = chart_source.to_bytes().expect("chart source bytes");
    let mut chart_edit = chart_source.edit().expect("chart edit");
    chart_edit
        .move_before("Sheet2", "Sheet1")
        .expect("chart lookup")
        .expect("chart tabs");
    let chart = chart_edit.commit().expect("chart reorder");
    let first = chart
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("tab");
    assert_eq!(
        (first.name(), first.kind()),
        ("Sheet2", WorksheetKind::Chart)
    );
    assert_eq!(
        chart
            .workbook()
            .apply(&chart.patch().inverse())
            .expect("chart inverse")
            .workbook()
            .to_bytes()
            .expect("chart bytes"),
        chart_bytes
    );
}

#[test]
fn tab_reorder_composes_with_other_facets_and_conflicts_globally() {
    let source = three_sheet_workbook();
    let mut order = source.edit().expect("order");
    order
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    let mut active = source.edit().expect("active");
    active
        .tab("Sheet3")
        .expect("lookup")
        .expect("tab")
        .activate();
    let mut visibility = source.edit().expect("visibility");
    visibility
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .hide();
    let mut cell = source.edit().expect("cell");
    cell.sheet("Sheet3")
        .expect("lookup")
        .expect("sheet")
        .set("B1", "moved payload")
        .expect("cell");
    order.join(active).expect("order and active");
    order.join(visibility).expect("order and visibility");
    order.join(cell).expect("order and cell");
    let committed = order.commit().expect("composed reorder");
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("Sheet3", 0));
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        &Visibility::Hidden
    );
    assert!(matches!(
        active.cell("B1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "moved payload"
    ));
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.workbook_uri)
            .count(),
        1
    );
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.sheets[2].part_uri)
            .count(),
        1
    );

    let mut left = source.edit().expect("left order");
    left.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    let mut right = source.edit().expect("right order");
    right
        .move_after("Sheet1", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let error = left.join(right).expect_err("order is one global facet");
    let conflicts = error.conflicts().expect("order conflict");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_order());

    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut same_position = source.edit().expect("same-position edit");
    same_position
        .move_before("Sheet2", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    same_position
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let committed = same_position
        .commit()
        .expect("active identity changes at the same position");
    let active = committed.workbook().active_sheet().expect("active");
    assert_eq!((active.name(), active.position()), ("Sheet2", 0));
    let active_change = committed
        .patch()
        .changes()
        .iter()
        .find_map(Change::active)
        .expect("semantic active change");
    assert_eq!(active_change.0.position(), 0);
    assert_eq!(active_change.1.position(), 0);
    assert_eq!(active_change.0.name(), "Sheet1");
    assert_eq!(active_change.1.name(), "Sheet2");
}

#[test]
fn numeric_tab_moves_are_checked_and_all_positions_round_trip() {
    for from in 0..3usize {
        for to in 0..3usize {
            let source = three_sheet_workbook();
            let source_bytes = source.to_bytes().expect("source bytes");
            let mut expected = vec!["Sheet1", "Sheet2", "Sheet3"];
            let moved = expected.remove(from);
            expected.insert(to, moved);
            let mut edit = source.edit().expect("edit");
            assert!(edit.move_to(from, to).expect("lookup").is_some());
            let committed = edit.commit().expect("move");
            assert_eq!(
                committed
                    .workbook()
                    .sheets()
                    .map(|sheet| sheet.name().to_owned())
                    .collect::<Vec<_>>(),
                expected.into_iter().map(str::to_owned).collect::<Vec<_>>()
            );
            let restored = committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse");
            assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
        }
    }

    let source = three_sheet_workbook();
    let mut missing = source.edit().expect("missing edit");
    assert!(
        missing
            .move_before("Absent", "Sheet1")
            .expect("lookup")
            .is_none()
    );
    assert!(missing.move_to("Sheet1", 3).expect("bounds").is_none());
    assert!(missing.commit().expect("no-op").patch().is_empty());

    let mut cancelled = source.edit().expect("cancelled edit");
    cancelled
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .move_after("Sheet3", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    assert!(cancelled.is_empty());
    assert_eq!(cancelled.len(), 0);
    assert!(cancelled.commit().expect("cancelled").patch().is_empty());

    let mut cancelled = source.edit().expect("cancelled join edit");
    cancelled
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .move_after("Sheet3", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let mut effective = source.edit().expect("effective join edit");
    effective
        .move_before("Sheet2", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    cancelled
        .join(effective)
        .expect("cancelled order has no conflict");
    let joined = cancelled.commit().expect("joined order");
    assert_eq!(
        joined
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet2", "Sheet1", "Sheet3"]
    );

    let source_bytes = source.to_bytes().expect("source bytes");
    let mut sequence = source.edit().expect("sequence edit");
    sequence
        .move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    sequence
        .move_after("Sheet1", "Sheet2")
        .expect("lookup")
        .expect("tabs");
    let sequence = sequence.commit().expect("move sequence");
    assert_eq!(
        sequence
            .workbook()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Sheet3", "Sheet2", "Sheet1"]
    );
    assert_eq!(sequence.patch().len(), 2);
    assert_eq!(sequence.patch().changes()[0].moved(), Some((2, 0)));
    assert_eq!(sequence.patch().changes()[1].moved(), Some((1, 2)));
    let inverse = sequence.patch().inverse();
    assert_eq!(inverse.changes()[0].moved(), Some((2, 1)));
    assert_eq!(inverse.changes()[1].moved(), Some((0, 2)));
    assert_eq!(
        sequence
            .workbook()
            .apply(&inverse)
            .expect("sequence inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn tab_reorder_blocks_protection_and_revision_tracking() {
    let source = three_sheet_workbook();
    let mut package = source.inner.package.clone();
    let workbook = package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook");
    let xml = std::str::from_utf8(workbook.blob())
        .expect("UTF-8")
        .replace(
            "<bookViews>",
            "<workbookProtection lockStructure=\"1\"/><bookViews>",
        );
    workbook.set_blob(xml.into_bytes());
    let protected = Workbook::from_package(package).expect("protected workbook");
    let mut edit = protected.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));
    let mut rename = protected.edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("checked name");
    assert!(matches!(
        rename.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));

    let mut package = source.inner.package.clone();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/revisions/revisionHeaders1.xml").expect("revision URI"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml"
                .to_owned(),
            br#"<headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
                .to_vec(),
        )))
        .expect("revision part");
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook")
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"
                .to_owned(),
            "revisions/revisionHeaders1.xml".to_owned(),
            "rIdRevisionHeaders".to_owned(),
            TargetMode::Internal,
        )
        .expect("revision relationship");
    let tracked = Workbook::from_package(package).expect("tracked workbook");
    let mut edit = tracked.edit().expect("edit");
    edit.move_before("Sheet3", "Sheet1")
        .expect("lookup")
        .expect("tabs");
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::TrackedWorkbook,
            ..
        })
    ));
    let mut rename = tracked.edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("lookup")
        .expect("tab")
        .rename("Input")
        .expect("checked name");
    assert!(matches!(
        rename.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::TrackedWorkbook,
            ..
        })
    ));
}

#[test]
fn tab_visibility_composes_with_cells_and_conflicts_by_facet() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut cell = source.edit().expect("cell edit");
    cell.sheet("Sheet2")
        .expect("lookup")
        .expect("worksheet")
        .set("A1", "preserved while hidden")
        .expect("cell");
    let mut tab = source.edit().expect("tab edit");
    tab.tab(1usize).expect("lookup").expect("tab").hide();
    cell.join(tab).expect("orthogonal join");
    let committed = cell.commit().expect("joined commit");
    assert_eq!(committed.patch().len(), 2);
    assert_eq!(
        committed
            .patch()
            .parts
            .iter()
            .filter(|part| part.uri == source.inner.workbook_uri)
            .count(),
        1
    );
    let sheet = committed
        .workbook()
        .sheet("Sheet2")
        .expect("lookup")
        .expect("sheet");
    assert_eq!(sheet.visibility(), &Visibility::Hidden);
    assert!(matches!(
        sheet.cell("A1").expect("cell").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "preserved while hidden"
    ));

    let mut left = source.edit().expect("left");
    left.tab("Sheet2").expect("lookup").expect("tab").hide();
    let mut right = source.edit().expect("right");
    right.tab(1usize).expect("lookup").expect("tab").very_hide();
    let error = left.join(right).expect_err("same tab facet must conflict");
    let conflicts = error.conflicts().expect("tab conflicts");
    assert_eq!(conflicts.len(), 1);
    assert!(conflicts.conflicts()[0].is_tab());
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet2");
}

#[test]
fn tab_visibility_applies_to_non_worksheet_sheet_kinds() {
    let source = two_sheet_workbook(WorksheetKind::Chart);
    assert_eq!(
        source
            .sheet("Sheet2")
            .expect("lookup")
            .expect("chart sheet")
            .kind(),
        WorksheetKind::Chart
    );
    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .very_hide();
    let committed = edit.commit().expect("chart tab commit");
    assert_eq!(
        committed
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("chart sheet")
            .visibility(),
        &Visibility::VeryHidden
    );
}

#[test]
fn active_relocation_synchronizes_worksheet_and_chart_view_selection() {
    for kind in [WorksheetKind::Worksheet, WorksheetKind::Chart] {
        let source = active_second_sheet_workbook(kind);
        let source_bytes = source.to_bytes().expect("source bytes");
        assert_eq!(
            source.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet2".to_owned())
        );
        let mut edit = source.edit().expect("edit");
        edit.tab("Sheet2").expect("lookup").expect("tab").hide();
        let committed = edit.commit().expect("active hide");
        assert_eq!(
            committed
                .workbook()
                .active_sheet()
                .map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".to_owned())
        );
        let new_active = committed
            .workbook()
            .inner
            .package
            .get_part(&committed.workbook().inner.sheets[0].part_uri)
            .expect("new active part")
            .blob();
        assert!(
            std::str::from_utf8(new_active)
                .expect("new active XML")
                .contains(r#"tabSelected="1""#)
        );
        let old_active = committed
            .workbook()
            .inner
            .package
            .get_part(&committed.workbook().inner.sheets[1].part_uri)
            .expect("old active part")
            .blob();
        assert!(
            !std::str::from_utf8(old_active)
                .expect("old active XML")
                .contains("tabSelected")
        );
        assert_eq!(
            committed
                .workbook()
                .apply(&committed.patch().inverse())
                .expect("inverse")
                .workbook()
                .to_bytes()
                .expect("restored bytes"),
            source_bytes
        );
    }
}

#[test]
fn tab_visibility_blocks_protected_workbook_structure() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookProtection lockStructure="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let protected = Workbook::from_package(package).expect("protected workbook");
    let mut edit = protected.edit().expect("edit");
    edit.tab("Sheet2").expect("lookup").expect("tab").hide();
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));

    let mut activation = protected.edit().expect("activation edit");
    activation
        .tab("Sheet2")
        .expect("lookup")
        .expect("tab")
        .activate();
    let activated = activation
        .commit()
        .expect("structure protection permits selection");
    assert_eq!(
        activated
            .workbook()
            .active_sheet()
            .expect("active sheet")
            .name(),
        "Sheet2"
    );
}

#[test]
fn showing_an_unknown_producer_state_repairs_it_explicitly() {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" state="show" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let source = Workbook::from_package(package).expect("producer workbook");
    assert!(matches!(
        source
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .visibility(),
        Visibility::Unknown(value) if value.as_ref() == "show"
    ));
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    edit.tab("Sheet2").expect("lookup").expect("tab").show();
    let committed = edit.commit().expect("repair commit");
    assert!(
        committed
            .workbook()
            .sheet("Sheet2")
            .expect("lookup")
            .expect("sheet")
            .visibility()
            .is_visible()
    );
    assert_eq!(
        committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse")
            .workbook()
            .to_bytes()
            .expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn clearing_a_cell_created_in_the_same_transaction_keeps_an_empty_record() {
    let source = Workbook::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.set("A1", "temporary").expect("set");
    sheet.clear("A1").expect("clear");
    let committed = edit.commit().expect("commit");
    assert!(matches!(
        committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Empty)
    ));
}

#[test]
fn independently_prepared_disjoint_edits_join_after_threaded_work() {
    fn assert_send<T: Send>() {}
    assert_send::<Edit>();

    let source = Workbook::new().expect("source workbook");
    let (mut left, right) = std::thread::scope(|scope| {
        let left_source = source.clone();
        let right_source = source.clone();
        let left = scope.spawn(move || {
            let mut edit = left_source.edit().expect("left edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .set("A1", "left")
                .expect("left cell");
            edit
        });
        let right = scope.spawn(move || {
            let mut edit = right_source.edit().expect("right edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("C3", 42_i32)
                .expect("right cell");
            edit
        });
        (
            left.join().expect("left worker"),
            right.join().expect("right worker"),
        )
    });

    left.join(right).expect("disjoint join");
    assert_eq!(left.len(), 2);
    let committed = left.commit().expect("joined commit");
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "left"
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ));

    let mut empty = source.edit().expect("empty edit");
    let mut incoming = source.edit().expect("incoming edit");
    incoming
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .set("D4", true)
        .expect("incoming cell");
    empty.join(incoming).expect("adopt incoming sheet map");
    assert_eq!(empty.len(), 1);
}

#[test]
fn join_conflicts_are_structured_and_return_the_rejected_edit() {
    let source = Workbook::new().expect("source workbook");
    let mut left = source.edit().expect("left edit");
    let mut left_sheet = left.sheet(0usize).expect("lookup").expect("sheet");
    left_sheet
        .set("C3", "left tail")
        .expect("left tail cell")
        .set("A1", "left")
        .expect("left first cell");
    let mut right = source.edit().expect("right edit");
    let mut right_sheet = right.sheet("Sheet1").expect("lookup").expect("sheet");
    right_sheet
        .set("A1", "right")
        .expect("right first cell")
        .set("C3", "right tail")
        .expect("right tail cell");

    let error = match left.join(right) {
        Ok(_) => panic!("overlapping edits must not join"),
        Err(error) => error,
    };
    assert_eq!(left.len(), 2);
    let conflicts = error.conflicts().expect("overlap details");
    assert_eq!(conflicts.len(), 2);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet1");
    assert_eq!(conflicts.conflicts()[0].position(), 0);
    assert_eq!(
        conflicts.conflicts()[0].cells().expect("cell conflicts"),
        &[
            Address::from_a1("A1").expect("first address"),
            Address::from_a1("C3").expect("tail address"),
        ]
    );
    let rejected = error.into_rejected();
    assert_eq!(rejected.len(), 2);

    let other_source = Workbook::new().expect("other source");
    let other = other_source.edit().expect("other edit");
    let error = match left.join(other) {
        Ok(_) => panic!("different snapshots must not join"),
        Err(error) => error,
    };
    assert!(matches!(error.failure(), JoinFailure::DifferentSnapshot));
    assert!(error.conflicts().is_none());
    assert_eq!(left.len(), 2);
}

#[test]
fn row_visibility_joins_with_cells_and_conflicts_by_row() {
    let source = Workbook::new().expect("source workbook");
    let mut cell = source.edit().expect("cell edit");
    cell.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A2", "same row")
        .expect("cell");
    let mut row = source.edit().expect("row edit");
    row.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    cell.join(row).expect("orthogonal row and cell effects");
    let committed = cell.commit().expect("joined commit");
    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(sheet.row(1).expect("row 2").hidden());
    assert!(matches!(
        sheet.cell("A2").expect("A2").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "same row"
    ));

    let mut left = source.edit().expect("left");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .hide();
    let mut right = source.edit().expect("right");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(4)
        .expect("row 5")
        .show();
    let error = left.join(right).expect_err("same row must conflict");
    let conflicts = error.conflicts().expect("row conflicts");
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts.conflicts().len(), 1);
    assert_eq!(
        conflicts.conflicts()[0].rows().expect("row conflict"),
        &[RowIndex::new(4).expect("row 5")]
    );
    assert!(conflicts.conflicts()[0].cells().is_none());
}

#[test]
fn calculation_chain_removal_is_atomic_and_reversible() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let chain_uri = PackURI::new("/xl/calcChain.xml").expect("chain URI");
    package
            .try_add_part(Box::new(BlobPart::new(
                chain_uri.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                    .to_owned(),
                br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#.to_vec(),
            )))
            .expect("chain part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::CALC_CHAIN.to_owned(),
            "calcChain.xml".to_owned(),
            "rId3".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain relationship");
    let source = Workbook::from_package(package).expect("workbook with chain");
    let source_bytes = source.to_bytes().expect("source bytes");
    let workbook_before = source
        .inner
        .package
        .get_part(&source.inner.workbook_uri)
        .expect("workbook part")
        .blob_arc();

    let mut visibility = source.edit().expect("visibility edit");
    visibility
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    let visibility = visibility.commit().expect("visibility commit");
    assert!(visibility.patch().graph.is_empty());
    assert!(
        visibility
            .workbook()
            .inner
            .package
            .get_part(&chain_uri)
            .is_ok()
    );
    assert_eq!(
        visibility
            .workbook()
            .inner
            .package
            .get_part(&source.inner.workbook_uri)
            .expect("unchanged workbook part")
            .blob(),
        workbook_before.as_slice()
    );

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    let committed = edit.commit().expect("commit");
    assert!(
        committed
            .workbook()
            .inner
            .package
            .get_part(&chain_uri)
            .is_err()
    );
    assert!(
        calculation_chain_removal(committed.workbook())
            .expect("chain query")
            .is_empty()
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(
        !calculation_chain_removal(restored.workbook())
            .expect("restored chain")
            .is_empty()
    );

    let mut shared_target = source.inner.package.clone();
    let mut referrer = BlobPart::new(
        PackURI::new("/xl/custom.xml").expect("custom URI"),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    referrer
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-reference".to_owned(),
            "calcChain.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("extra incoming relationship");
    shared_target
        .try_add_part(Box::new(referrer))
        .expect("referrer part");
    let shared_target = Workbook::from_package(shared_target).expect("shared target workbook");
    assert!(matches!(
        calculation_chain_removal(&shared_target),
        Err(Error::Invalid(message)) if message.contains("another incoming relationship")
    ));

    let mut outbound = source.inner.package.clone();
    outbound
        .get_part_mut(&chain_uri)
        .expect("chain part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain outbound relationship");
    let outbound = Workbook::from_package(outbound).expect("outbound chain workbook");
    let mut edit = outbound.edit().expect("outbound chain edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(message)) if message.contains("calculation-chain part cannot have relationships")
    ));
    assert!(outbound.inner.package.get_part(&chain_uri).is_ok());
}

#[test]
fn shared_style_crud_is_lineage_checked_reversible_and_exact() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let styles = source.styles().expect("styles");
    assert_eq!(styles.len(), 2);
    let base = styles.base().expect("base style");
    let accent = styles.get(1).expect("accent style");
    assert_eq!(base.fan_out().expect("base fan-out"), 1);
    assert_eq!(accent.fan_out().expect("accent fan-out"), 1);

    let sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.local_style("B1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    let Some(crate::LocalStyle::Shared(local)) = sheet.local_style("A1").expect("local style")
    else {
        panic!("A1 must have an explicit shared style")
    };
    assert!(local.same(&accent));
    assert!(
        sheet
            .style("B1")
            .expect("resolved style")
            .is_some_and(|style| style.same(&base))
    );

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
    sheet
        .set("C1", 42_i32)
        .and_then(|sheet| sheet.style("C1", &accent))
        .and_then(|sheet| sheet.style("D1", &accent))
        .expect("style changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 2);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, State::Missing, _))
    ));
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
        })) if number.as_str() == "42"
    ));

    let book = committed.workbook();
    let sheet = book.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("D1").expect("D1").stored(),
        Some(Cell::Empty)
    ));
    let styles = book.styles().expect("styles");
    let inherited = styles.find(&accent.key()).expect("inherited style key");
    assert!(inherited.same(&accent));
    assert!(!inherited.same_workbook(&accent));
    assert_eq!(accent.fan_out().expect("source fan-out"), 1);
    assert_eq!(inherited.fan_out().expect("descendant fan-out"), 3);
    assert!(
        sheet
            .style("C1")
            .expect("style")
            .is_some_and(|style| style.same(&inherited))
    );

    let mut descendant = book.edit().expect("descendant edit");
    descendant
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .style("E1", &accent)
        .expect("reuse inherited style lineage");
    let descendant = descendant.commit().expect("descendant commit");
    assert!(matches!(
        descendant
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .local_style("E1")
            .expect("local style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let reopened = Workbook::from_bytes(source_bytes.clone()).expect("reopened source");
    assert!(
        reopened
            .styles()
            .expect("reopened styles")
            .find(&accent.key())
            .is_none()
    );
    let replayed = reopened
        .apply(committed.patch())
        .expect("replay onto byte-identical source");
    let (
        _,
        _,
        State::Cell {
            style: StyleState::Shared(replayed_key),
            ..
        },
    ) = replayed.patch().changes()[0].cell().expect("cell change")
    else {
        panic!("replayed change must retain its shared style")
    };
    assert!(
        replayed
            .workbook()
            .styles()
            .expect("replayed styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        book.styles()
            .expect("original lineage")
            .find(replayed_key)
            .is_none()
    );

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let foreign = Workbook::new()
        .expect("other workbook")
        .styles()
        .expect("other styles")
        .base()
        .expect("other base style");
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &foreign),
        Err(Error::ForeignStyle)
    ));

    let mut changed_package = source.inner.package.clone();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let changed_xml = {
        let styles = changed_package.get_part(&styles_uri).expect("styles part");
        std::str::from_utf8(styles.blob())
            .expect("UTF-8 styles")
            .replace("FFFFFF00", "FFFF0000")
            .into_bytes()
    };
    changed_package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(changed_xml);
    let changed = Workbook::from_package_with_styles(changed_package, Some(&source))
        .expect("changed style table");
    assert!(
        changed
            .styles()
            .expect("changed styles")
            .find(&accent.key())
            .is_none()
    );
    let mut edit = changed.edit().expect("changed edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &accent),
        Err(Error::ForeignStyle)
    ));
    assert!(matches!(
        changed.apply(committed.patch()),
        Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
    ));
}

#[test]
fn payload_and_style_effects_on_one_cell_join_without_locks() {
    let source = styled_workbook();
    let accent = source.styles().expect("styles").get(1).expect("accent");
    let (mut payload, style) = std::thread::scope(|scope| {
        let payload_source = source.clone();
        let style_source = source.clone();
        let accent = accent.clone();
        let payload = scope.spawn(move || {
            let mut edit = payload_source.edit().expect("payload edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("B1", 9_i32)
                .expect("payload");
            edit
        });
        let style = scope.spawn(move || {
            let mut edit = style_source.edit().expect("style edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .style("B1", &accent)
                .expect("style");
            edit
        });
        (
            payload.join().expect("payload worker"),
            style.join().expect("style worker"),
        )
    });
    payload.join(style).expect("disjoint cell facets");
    assert_eq!(payload.len(), 1);
    let committed = payload.commit().expect("commit");
    assert_eq!(committed.patch().len(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
        })) if number.as_str() == "9"
    ));

    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("B1").expect("B1").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "9"
    ));
    assert!(matches!(
        sheet.local_style("B1").expect("style"),
        Some(crate::LocalStyle::Shared(_))
    ));
}

#[test]
fn resetting_style_is_distinct_from_removal_and_missing_is_a_no_op() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet
        .reset_style("A1")
        .and_then(|sheet| sheet.reset_style("Z99"))
        .expect("style resets");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((
            _,
            _,
            State::Cell {
                style: StyleState::Default,
                ..
            }
        ))
    ));
    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.local_style("A1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    assert!(sheet.cell("Z99").expect("missing").is_missing());

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
}

#[test]
fn signed_packages_refuse_edits_before_mutation() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    let signed = Workbook::from_package(package).expect("signed workbook snapshot");

    assert!(matches!(signed.edit(), Err(Error::Signed)));
    assert!(matches!(
        signed.apply(&Patch::default()),
        Err(Error::Signed)
    ));
}

fn rename_reference_workbook() -> Workbook {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rIdTab2"/></sheets><definedNames><definedName name="Source">Data!$A$1</definedName></definedNames></workbook>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("Data worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("Calc worksheet")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><f>Data!A1</f><v>7</v></c></row></sheetData><dataValidations count="1"><dataValidation type="custom" sqref="B1"><formula1>Data!A1&gt;0</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="C1" location="Data!$A$1"/></hyperlinks></worksheet>"#.to_vec(),
            );
    for (uri, content_type, content) in [
            (
                "/xl/tables/table1.xml",
                litchi_opc::constants::content_type::SML_TABLE,
                br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns count="1"><tableColumn id="1" name="Value"><calculatedColumnFormula>Data!A1</calculatedColumnFormula></tableColumn></tableColumns></table>"#.as_slice(),
            ),
            (
                "/xl/charts/chart1.xml",
                litchi_opc::constants::content_type::DML_CHART,
                br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
            ),
            (
                "/xl/pivotCache/pivotCacheDefinition1.xml",
                litchi_opc::constants::content_type::SML_PIVOT_CACHE_DEFINITION,
                br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1"/></cacheSource></pivotCacheDefinition>"#.as_slice(),
            ),
            (
                "/docProps/app.xml",
                litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#.as_slice(),
            ),
            (
                "/xl/externalLinks/externalLink1.xml",
                litchi_opc::constants::content_type::SML_EXTERNAL_LINK,
                br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook><definedNames><definedName name="External" refersTo="[1]Data!A1"/></definedNames></externalBook></externalLink>"#.as_slice(),
            ),
        ] {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(uri).expect("part URI"),
                    content_type.to_owned(),
                    content.to_vec(),
                )))
                .expect("reference part");
        }
    Workbook::from_package(package).expect("rename reference workbook")
}

fn part_text<'a>(workbook: &'a Workbook, uri: &str) -> &'a str {
    let uri = PackURI::new(uri).expect("part URI");
    let bytes = workbook
        .inner
        .package
        .get_part(&uri)
        .expect("package part")
        .blob();
    std::str::from_utf8(bytes).expect("XML part")
}

fn styled_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
            .expect("styles part")
            .set_blob(
                br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.to_vec(),
            );
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled workbook")
}

fn styled_column_workbook() -> Workbook {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="3" max="3" style="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled column workbook")
}

fn styled_row_workbook() -> Workbook {
    let baseline = styled_workbook();
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("styled row workbook")
}

fn defaults_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="ac"><sheetFormatPr baseColWidth="10" defaultColWidth="12" defaultRowHeight="15" customHeight="0" zeroHeight="1" thickTop="1" ac:dyDescent="0.1"/><sheetData><row r="2" customHeight="0" ac:dyDescent="0.2"/></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("defaults workbook")
}

fn merged_workbook() -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>anchor</t></is></c></row><row r="2"><c r="E2" t="inlineStr"><is><t>keep</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("merged workbook")
}

fn two_sheet_workbook(second_kind: WorksheetKind) -> Workbook {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    let (relationship_type, content_type, part_xml) = match second_kind {
            WorksheetKind::Worksheet => (
                litchi_opc::constants::relationship_type::WORKSHEET,
                litchi_opc::constants::content_type::SML_WORKSHEET,
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>"#.as_slice(),
            ),
            WorksheetKind::Chart => (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.as_slice(),
            ),
            _ => panic!("test helper only models worksheet and chart tabs"),
        };
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            match second_kind {
                WorksheetKind::Worksheet => "worksheets/sheet2.xml",
                WorksheetKind::Chart => "chartsheets/sheet2.xml",
                _ => unreachable!("guarded above"),
            }
            .to_owned(),
            "rIdTab2".to_owned(),
            TargetMode::Internal,
        )
        .expect("second sheet relationship");
    let part_uri = match second_kind {
        WorksheetKind::Worksheet => "/xl/worksheets/sheet2.xml",
        WorksheetKind::Chart => "/xl/chartsheets/sheet2.xml",
        _ => unreachable!("guarded above"),
    };
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(part_uri).expect("second sheet URI"),
            content_type.to_owned(),
            part_xml.to_vec(),
        )))
        .expect("second sheet part");
    Workbook::from_package(package).expect("two-sheet workbook")
}

fn three_sheet_workbook() -> Workbook {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1" firstSheet="0"/><workbookView activeTab="2" firstSheet="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/><sheet name="Sheet3" sheetId="3" r:id="rIdTab3"/></sheets><definedNames><definedName name="FirstLocal" localSheetId="0">Sheet1!$A$1</definedName><definedName name="ThirdLocal" localSheetId="2">Sheet3!$A$1</definedName><definedName name="Global">1</definedName></definedNames></workbook>"#.to_vec(),
            );
    package
        .get_part_mut(&source.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            "worksheets/sheet3.xml".to_owned(),
            "rIdTab3".to_owned(),
            TargetMode::Internal,
        )
        .expect("third sheet relationship");
    package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/worksheets/sheet3.xml").expect("third sheet URI"),
                litchi_opc::constants::content_type::SML_WORKSHEET.to_owned(),
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>three</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            )))
            .expect("third sheet part");
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[1].part_uri)
            .expect("second sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>two</t></is></c></row></sheetData></worksheet>"#.to_vec(),
            );
    Workbook::from_package(package).expect("three-sheet workbook")
}

fn active_second_sheet_workbook(second_kind: WorksheetKind) -> Workbook {
    let source = two_sheet_workbook(second_kind);
    let mut package = source.inner.package.clone();
    package
            .get_part_mut(&source.inner.workbook_uri)
            .expect("workbook part")
            .set_blob(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rIdTab2"/></sheets></workbook>"#.to_vec(),
            );
    package
            .get_part_mut(&source.inner.sheets[0].part_uri)
            .expect("first sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.to_vec(),
            );
    let second_xml = match second_kind {
            WorksheetKind::Worksheet => br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#.as_slice(),
            WorksheetKind::Chart => br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews></chartsheet>"#.as_slice(),
            _ => unreachable!("test helper only models worksheet and chart tabs"),
        };
    package
        .get_part_mut(&source.inner.sheets[1].part_uri)
        .expect("second sheet part")
        .set_blob(second_xml.to_vec());
    Workbook::from_package(package).expect("active second sheet workbook")
}
