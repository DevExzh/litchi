//! Independent selector-search coverage for the ODS sheet-metadata index.
//!
//! These tests keep the public selector contract separate from the metadata
//! lifecycle target.  They exercise logical coordinates over sparse physical
//! rows/cells, row containers, merge geometry, duplicate worksheet names, and
//! execution-policy boundaries.

use std::{
    fmt::Write as _,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Position,
    Resource,
};
use litchi_ods::{Merge, model::detective::Detective, sheet_metadata as metadata};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn document(spreadsheet: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet>{spreadsheet}</office:spreadsheet></office:body></office:document-content>"#
    )
}

fn selector_fixture() -> String {
    document(
        r#"
        <table:table table:name="Repeats">
          <table:table-row table:number-rows-repeated="2">
            <table:table-cell table:number-columns-repeated="3"><text:p>left</text:p></table:table-cell>
            <table:table-cell table:number-columns-repeated="2"><text:p>right</text:p></table:table-cell>
          </table:table-row>
          <table:table-row><table:table-cell><text:p>tail</text:p></table:table-cell></table:table-row>
        </table:table>
        <table:table table:name="Merged">
          <table:table-row><table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2"><text:p>anchor</text:p></table:table-cell></table:table-row>
        </table:table>
        <table:table table:name="ExplicitCovered">
          <table:table-row><table:covered-table-cell table:number-columns-repeated="2"/></table:table-row>
        </table:table>
        <table:table table:name="Grouped">
          <table:table-header-rows>
            <table:table-row table:number-rows-repeated="2"><table:table-cell><text:p>header</text:p></table:table-cell></table:table-row>
          </table:table-header-rows>
          <table:table-row-group>
            <table:table-row table:number-rows-repeated="2"><table:table-cell><text:p>group</text:p></table:table-cell></table:table-row>
          </table:table-row-group>
          <table:table-rows>
            <table:table-row><table:table-cell><text:p>rows</text:p></table:table-cell></table:table-row>
          </table:table-rows>
        </table:table>
        "#,
    )
}

fn duplicate_name_fixture() -> String {
    document(
        r#"
        <table:table table:name="Duplicate">
          <table:table-row><table:table-cell><text:p>first</text:p></table:table-cell></table:table-row>
        </table:table>
        <table:table table:name="Duplicate">
          <table:table-row table:number-rows-repeated="2"><table:table-cell><text:p>second</text:p></table:table-cell></table:table-row>
        </table:table>
        "#,
    )
}

fn dense_fixture(rows: usize, columns: usize) -> String {
    let mut spreadsheet = String::new();
    write!(spreadsheet, r#"<table:table table:name="Dense">"#).expect("table start");
    for _ in 0..rows {
        spreadsheet.push_str("<table:table-row>");
        for _ in 0..columns {
            spreadsheet.push_str("<table:table-cell/>");
        }
        spreadsheet.push_str("</table:table-row>");
    }
    spreadsheet.push_str("</table:table>");
    document(&spreadsheet)
}

fn context(scope: &str) -> (CancellationSource, ExecutionContext) {
    context_with_work(scope, 50_000_000_000)
}

fn context_with_work(scope: &str, work: u64) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        scope,
        CoreLimits::new(
            4 * 1024 * 1024 * 1024,
            64 * 1024 * 1024 * 1024,
            128 * 1024 * 1024 * 1024,
            250_000_000,
            1024,
            work,
        ),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1024 * 1024).expect("one MiB task budget"),
        0,
    )
    .expect("finite execution limits");
    (cancellation, ExecutionContext::new(budget, token, limits))
}

fn parse(source: &str) -> metadata::Snapshot {
    metadata::Snapshot::parse(source).expect("valid selector fixture")
}

#[test]
fn repeated_rows_and_cells_resolve_every_boundary_without_expansion() {
    let source = selector_fixture();
    let snapshot = parse(&source);

    for &(row, column) in &[(0, 0), (0, 2), (1, 0), (1, 2)] {
        let view = snapshot
            .cell_metadata(metadata::CellSelector::by_name("Repeats", row, column))
            .expect("repeated left lookup")
            .expect("repeated left cell");
        let metadata::CellMetadataView::Physical(view) = view else {
            panic!("repeated left coordinate must resolve to a physical cell")
        };
        let location = view.location();
        assert_eq!(location.row(), 0);
        assert_eq!(location.column(), 0);
        assert_eq!(location.row_repeat(), 2);
        assert_eq!(location.column_repeat(), 3);
        assert!(!view.is_covered(), "table-cell must remain stored");
    }

    for &(row, column) in &[(0, 3), (0, 4), (1, 3), (1, 4)] {
        let view = snapshot
            .cell_metadata(metadata::CellSelector::by_name("Repeats", row, column))
            .expect("repeated right lookup")
            .expect("repeated right cell");
        let metadata::CellMetadataView::Physical(view) = view else {
            panic!("repeated right coordinate must resolve to a physical cell")
        };
        let location = view.location();
        assert_eq!(location.row(), 0);
        assert_eq!(location.column(), 3);
        assert_eq!(location.row_repeat(), 2);
        assert_eq!(location.column_repeat(), 2);
    }

    let tail = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Repeats", 2, 0))
        .expect("tail lookup")
        .expect("tail cell")
        .location()
        .expect("tail location");
    assert_eq!((tail.row(), tail.column()), (2, 0));
    assert_eq!((tail.row_repeat(), tail.column_repeat()), (1, 1));

    for (row, column) in [(0, 5), (1, 5), (2, 1), (3, 0)] {
        assert!(
            snapshot
                .cell_metadata(metadata::CellSelector::by_name("Repeats", row, column))
                .expect("missing boundary lookup")
                .is_none()
        );
    }
}

#[test]
fn row_containers_preserve_logical_row_boundaries() {
    let snapshot = parse(&selector_fixture());

    for row in 0..5 {
        let view = snapshot
            .cell_metadata(metadata::CellSelector::by_name("Grouped", row, 0))
            .expect("row-container lookup")
            .expect("row-container cell");
        let metadata::CellMetadataView::Physical(view) = view else {
            panic!("row-container coordinate must resolve to a physical cell")
        };
        let location = view.location();
        assert_eq!(location.row(), row / 2 * 2);
        assert_eq!(location.column(), 0);
        assert_eq!(location.row_repeat(), if row < 4 { 2 } else { 1 });
    }
    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_name("Grouped", 5, 0))
            .expect("row-container missing lookup")
            .is_none()
    );
}

#[test]
fn implicit_merges_and_explicit_covered_runs_keep_distinct_views() {
    let snapshot = parse(&selector_fixture());

    let anchor = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Merged", 0, 0))
        .expect("merge anchor lookup")
        .expect("merge anchor");
    let metadata::CellMetadataView::Physical(anchor) = anchor else {
        panic!("merge anchor must resolve to a physical cell")
    };
    let location = anchor.location();
    assert_eq!(
        location.merge(),
        Merge::Span {
            rows: NonZeroUsize::new(2).expect("rows"),
            columns: NonZeroUsize::new(2).expect("columns"),
        }
    );
    assert!(!anchor.is_covered());

    let implicit = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Merged", 1, 1))
        .expect("implicit merge lookup")
        .expect("implicit covered view");
    assert!(matches!(
        implicit,
        metadata::CellMetadataView::ImplicitCovered {
            anchor: (0, 0),
            span: (2, 2),
        }
    ));
    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_name("Merged", 1, 2))
            .expect("outside merge lookup")
            .is_none()
    );

    for column in 0..2 {
        let covered = snapshot
            .cell_metadata(metadata::CellSelector::by_name(
                "ExplicitCovered",
                0,
                column,
            ))
            .expect("explicit covered lookup")
            .expect("explicit covered cell");
        let metadata::CellMetadataView::Physical(covered) = covered else {
            panic!("explicit covered coordinate must remain physical")
        };
        assert!(covered.is_covered());
        let location = covered.location();
        assert_eq!(location.merge(), Merge::Covered);
        assert_eq!(location.column_repeat(), 2);
    }
    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_name("ExplicitCovered", 0, 2))
            .expect("covered boundary lookup")
            .is_none()
    );
}

#[test]
fn duplicate_names_are_rejected_while_position_selectors_remain_exact() {
    let source = duplicate_name_fixture();
    let snapshot = parse(&source);

    let duplicate = snapshot.cell_metadata(metadata::CellSelector::by_name("Duplicate", 0, 0));
    assert!(
        duplicate.is_err(),
        "duplicate worksheet names are ambiguous"
    );

    let first = snapshot
        .cell_metadata(metadata::CellSelector::by_position(Position::new(0), 0, 0))
        .expect("first position lookup")
        .expect("first position cell");
    let metadata::CellMetadataView::Physical(first) = first else {
        panic!("first position must resolve to a physical cell")
    };
    assert_eq!(first.location().row_repeat(), 1);

    let second = snapshot
        .cell_metadata(metadata::CellSelector::by_position(Position::new(1), 1, 0))
        .expect("second position lookup")
        .expect("second position cell");
    let metadata::CellMetadataView::Physical(second) = second else {
        panic!("second position must resolve to a physical cell")
    };
    assert_eq!(second.location().row_repeat(), 2);
    assert_eq!(snapshot.source_xml(), source);
}

#[test]
fn high_cardinality_last_cell_fits_a_bounded_selector_work_budget() {
    const ROWS: usize = 128;
    const COLUMNS: usize = 128;
    const SELECTOR_COMPARISON_BUDGET: u64 = 64 * 8;
    let source = dense_fixture(ROWS, COLUMNS);

    let (_baseline_cancel, baseline_context) = context("ods-selector-search-baseline");
    let baseline = metadata::Snapshot::parse_with_context(
        &source,
        metadata::Limits::default(),
        &baseline_context,
    )
    .expect("dense baseline snapshot");
    let parse_work = baseline_context.budget().used(Resource::Work);
    drop(baseline);

    let (_limited_cancel, limited_context) = context_with_work(
        "ods-selector-search-bounded",
        parse_work + SELECTOR_COMPARISON_BUDGET,
    );
    let snapshot = metadata::Snapshot::parse_with_context(
        &source,
        metadata::Limits::default(),
        &limited_context,
    )
    .expect("dense bounded snapshot");
    let view = snapshot
        .cell_metadata(metadata::CellSelector::by_position(
            Position::new(0),
            ROWS - 1,
            COLUMNS - 1,
        ))
        .expect("bounded last-cell lookup")
        .expect("dense last cell");
    assert_eq!(view.location().expect("dense location").row(), ROWS - 1);
    assert_eq!(
        view.location().expect("dense location").column(),
        COLUMNS - 1
    );
}

#[test]
fn cancelled_lookup_and_failed_edit_leave_source_and_staging_unchanged() {
    let source = selector_fixture();
    let (cancellation, execution_context) = context("ods-selector-search-cancellation");
    let snapshot = metadata::Snapshot::parse_with_context(
        &source,
        metadata::Limits::default(),
        &execution_context,
    )
    .expect("cancellation fixture");
    let used_before_cancel = execution_context.budget().used(Resource::Work);
    cancellation.cancel();
    let error = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Repeats", 1, 4))
        .expect_err("cancelled lookup must fail");
    assert!(error.to_string().contains("cancelled"));
    assert_eq!(
        execution_context.budget().used(Resource::Work),
        used_before_cancel
    );
    assert_eq!(snapshot.source_xml(), source);

    let (edit_cancellation, edit_context) = context("ods-selector-search-edit-atomicity");
    let edit_snapshot =
        metadata::Snapshot::parse_with_context(&source, metadata::Limits::default(), &edit_context)
            .expect("edit fixture");
    let mut edit = edit_snapshot.edit();
    let missing = metadata::CellSelector::by_name("Repeats", 2, 1);
    assert!(
        edit.set_detective(missing, Detective::new()).is_err(),
        "missing edit target must be rejected"
    );
    assert!(
        edit.is_no_op(),
        "failed missing edit must not stage an operation"
    );
    edit_cancellation.cancel();
    let commit = edit.commit(&edit_context);
    assert!(commit.is_err(), "cancelled commit must fail");
    assert!(
        edit.is_no_op(),
        "cancelled commit must retain no partial staging"
    );
    assert_eq!(edit.before().source_xml(), source);
}

#[test]
fn position_boundaries_and_missing_sheet_positions_are_errors_without_mutation() {
    let source = selector_fixture();
    let snapshot = parse(&source);

    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_position(
                Position::new(0),
                usize::MAX,
                0
            ))
            .expect("missing row lookup")
            .is_none()
    );
    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_position(Position::new(99), 0, 0))
            .is_err()
    );
    assert_eq!(snapshot.source_xml(), source);
}
