//! Unified XLSB complete-catalog tab ordering through the public facade.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::Position;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Placement, Props, SourceMode};

const CHART: &str = "/sheet[1]";
const DATA: &str = "/sheet[2]";
const TAIL: &str = "/sheet[3]";

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::chart::{Anchor, Chart};
    use litchi_xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("build chart fixture");
    let mut builder = Builder::new();
    builder
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("add chart tab");
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "data");
    builder.add_worksheet(data);
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "tail");
    builder.add_worksheet(tail);

    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build mixed XLSB fixture")
        .write_to(&mut output)
        .expect("serialize mixed XLSB fixture");
    output.into_inner()
}

fn many_tab_fixture(count: usize) -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut builder = Builder::new();
    for index in 1..=count {
        let mut sheet = builder::Worksheet::new(format!("Tab {index}"));
        sheet.set_cell(0, 0, index.to_string());
        builder.add_worksheet(sheet);
    }

    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build many-tab XLSB fixture")
        .write_to(&mut output)
        .expect("serialize many-tab XLSB fixture");
    output.into_inner()
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .expect("open XLSB owner")
}

fn sheet_name(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == "name")
        .map(|property| property.value().to_string())
        .unwrap_or_else(|| panic!("missing sheet name at {path}"))
}

fn order(file: &File) -> [String; 3] {
    [
        sheet_name(file, "/sheet[1]"),
        sheet_name(file, "/sheet[2]"),
        sheet_name(file, "/sheet[3]"),
    ]
}

fn active_tab_name(bytes: &[u8]) -> String {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes)).expect("open native workbook");
    let position = workbook
        .active_catalog_position()
        .expect("workbook has an active tab");
    workbook
        .sheets()
        .nth(position)
        .expect("active tab position")
        .name()
        .to_owned()
}

fn move_batch(path: &'static str, anchor: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("move");
    item.path = Some(path.into());
    item.before = Some(anchor.into());
    item
}

fn swap_batch(left: &'static str, right: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("swap");
    item.path = Some(left.into());
    item.path2 = Some(right.into());
    item
}

#[test]
fn complete_catalog_move_supports_every_placement_for_source_and_eager_owners() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        assert_eq!(order(&source), ["Sales Chart", "Data", "Tail"]);

        let start = source
            .move_to(TAIL, Placement::Start)
            .expect("move to start");
        assert_eq!(order(&start), ["Tail", "Sales Chart", "Data"]);

        let end = source.move_to(CHART, Placement::End).expect("move to end");
        assert_eq!(order(&end), ["Data", "Tail", "Sales Chart"]);

        let position = source
            .move_to(CHART, Position::new(2))
            .expect("move to position");
        assert_eq!(order(&position), ["Data", "Tail", "Sales Chart"]);

        let before = source
            .move_to(CHART, Placement::before(TAIL).expect("before placement"))
            .expect("move before anchor");
        assert_eq!(order(&before), ["Data", "Sales Chart", "Tail"]);

        let after = source
            .move_to(CHART, Placement::after(DATA).expect("after placement"))
            .expect("move after anchor");
        assert_eq!(order(&after), ["Data", "Sales Chart", "Tail"]);

        let inside = source
            .move_to(
                CHART,
                Placement::inside_at("/", Position::new(1)).expect("inside"),
            )
            .expect("move into workbook");
        assert_eq!(order(&inside), ["Data", "Sales Chart", "Tail"]);
    }
}

#[test]
fn complete_catalog_swap_preserves_tab_identity_and_supports_noop() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let active = active_tab_name(&source.to_bytes().expect("publish source"));
        let swapped = source.swap(CHART, TAIL).expect("swap tabs");
        assert_eq!(order(&swapped), ["Tail", "Data", "Sales Chart"]);
        assert_eq!(
            active_tab_name(&swapped.to_bytes().expect("publish swap")),
            active,
            "active tab follows its identity"
        );

        let noop = source.swap(CHART, CHART).expect("same-tab swap is a no-op");
        assert_eq!(order(&noop), order(&source));
        assert_eq!(
            noop.to_bytes().expect("publish no-op"),
            source.to_bytes().expect("publish source")
        );

        let reopened = open_owner(&swapped.to_bytes().expect("publish swap"), eager);
        assert_eq!(order(&reopened), ["Tail", "Data", "Sales Chart"]);
    }
}

#[test]
fn ordering_is_atomic_across_set_move_and_swap_batches() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let before = source.to_bytes().expect("publish source");
        let mut rename = BatchItem::new("set");
        rename.path = Some(DATA.into());
        rename.props = Some(Props::new().with("name", "Input"));

        let result = source
            .batch(
                [
                    rename,
                    move_batch(CHART, "/sheet[3]"),
                    swap_batch("/sheet[1]", "/sheet[99]"),
                ],
                BatchMode::Atomic,
            )
            .expect_err("atomic ordering batch must roll back a missing tab");
        assert_eq!(result.item(), Some(2));
        assert_eq!(result.code(), Code::PathNotFound);
        assert_eq!(source.to_bytes().expect("atomic source unchanged"), before);

        let best_effort = source
            .batch(
                [
                    move_batch(CHART, "/sheet[3]"),
                    swap_batch("/sheet[1]", "/sheet[99]"),
                ],
                BatchMode::BestEffort,
            )
            .expect("best-effort ordering batch");
        assert!(matches!(
            &best_effort.value()[0],
            litchi_ops::ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            litchi_ops::ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::PathNotFound
        ));
        assert_eq!(order(best_effort.file()), ["Data", "Sales Chart", "Tail"]);
    }
}

#[test]
fn ordering_honors_selected_limits_and_rejects_non_workbook_destinations() {
    let source = open_owner(&fixture(), false);
    let limits = SemanticLimits::builder()
        .max_staged_bytes(1)
        .expect("staged limit")
        .build()
        .expect("limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.move_to(CHART, Placement::End),
        )
        .expect_err("staged target limit");
    assert_eq!(error.code(), Code::LimitExceeded);

    let error = source
        .move_to(CHART, Placement::inside(DATA).expect("inside worksheet"))
        .expect_err("worksheet is not a workbook parent");
    assert_eq!(error.code(), Code::UnsupportedType);
}

#[test]
fn ordering_charges_catalog_storage_separately_from_selected_targets() {
    let source = open_owner(&many_tab_fixture(32), false);
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .expect("one selected target")
        .max_staged_bytes(4096)
        .expect("catalog storage budget")
        .build()
        .expect("limits");
    let moved = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.move_to("/sheet[1]", Placement::End),
        )
        .expect("one move fits the selected-target quota")
        .into_file();
    assert_eq!(sheet_name(&moved, "/sheet[32]"), "Tab 1");

    let low_staged = SemanticLimits::builder()
        .max_expanded_targets(1)
        .expect("one selected target")
        .max_staged_bytes(64)
        .expect("small staged budget")
        .build()
        .expect("limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(low_staged),
            |edit| edit.move_to("/sheet[1]", Placement::End),
        )
        .expect_err("complete catalog storage must be admitted before allocation");
    assert_eq!(error.code(), Code::LimitExceeded);
}
