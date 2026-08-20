//! Row, column, and worksheet-default edit tests.

use super::super::*;
use super::support::{
    defaults_workbook, styled_column_workbook, styled_row_workbook, styled_workbook,
};

use crate::column::Outline;
use crate::{StyleState, Value};
use litchi_opc::Part;

fn page_break_workbook() -> Workbook {
    let source = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = source.inner.package.clone();
    package
        .get_part_mut(&source.inner.sheets[0].part_uri)
        .expect("first worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks></worksheet>"#.to_vec(),
        );
    Workbook::from_package(package).expect("page-break workbook")
}

#[test]
fn worksheet_page_break_projection_reuses_successful_snapshot_parse() {
    let workbook = page_break_workbook();
    let sheet = workbook
        .sheet(0usize)
        .expect("sheet lookup")
        .expect("first worksheet");
    assert!(sheet.data.page_breaks.get().is_none());

    let first = sheet.page_breaks().expect("first page-break read");
    let cached = sheet.data.page_breaks.get().expect("cached projection");
    let cached_address = cached as *const _;
    assert_eq!(&first, cached);

    let second = sheet.page_breaks().expect("second page-break read");
    let cached_again = sheet.data.page_breaks.get().expect("cached projection");
    assert_eq!(first, second);
    assert_eq!(cached_address, cached_again as *const _);
}

#[test]
fn worksheet_page_break_projection_is_invalidated_by_publication() {
    let source = page_break_workbook();
    let source_sheet = source
        .sheet(0usize)
        .expect("source lookup")
        .expect("source worksheet");
    let original = source_sheet.page_breaks().expect("source page breaks");
    let original_cache = std::sync::Arc::clone(&source_sheet.data.page_breaks);

    let mut edit = source.edit().expect("edit");
    edit.move_page_breaks(0usize, 1usize)
        .expect("move page breaks")
        .expect("source and target sheets");
    let committed = edit.commit().expect("commit");
    let moved_source = committed
        .workbook()
        .sheet(0usize)
        .expect("source lookup")
        .expect("source worksheet");
    let moved_target = committed
        .workbook()
        .sheet(1usize)
        .expect("target lookup")
        .expect("target worksheet");

    assert!(!std::sync::Arc::ptr_eq(
        &original_cache,
        &moved_source.data.page_breaks
    ));
    assert!(!std::sync::Arc::ptr_eq(
        &original_cache,
        &moved_target.data.page_breaks
    ));
    assert!(moved_source.data.page_breaks.get().is_none());
    assert!(moved_target.data.page_breaks.get().is_none());
    let updated_source = moved_source
        .page_breaks()
        .expect("updated source page breaks");
    assert!(updated_source.horizontal().is_none());
    assert!(updated_source.vertical().is_none());
    assert_eq!(
        moved_target
            .page_breaks()
            .expect("updated target page breaks"),
        original
    );

    // The source snapshot remains immutable even after its descendant was
    // published and retains the original successful projection.
    assert_eq!(
        source_sheet.page_breaks().expect("source snapshot"),
        original
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
