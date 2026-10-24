//! Unified XLSB worksheet-tab creation through the public facade.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::{FileFormat, Position};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, NewNode, OpenOptions, Placement, Props, SourceMode, create,
    props,
};

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

fn order(file: &File) -> [String; 4] {
    [
        sheet_name(file, "/sheet[1]"),
        sheet_name(file, "/sheet[2]"),
        sheet_name(file, "/sheet[3]"),
        sheet_name(file, "/sheet[4]"),
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

#[test]
fn add_supports_complete_catalog_placements_for_source_and_eager_owners() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        for (placement, expected) in [
            (
                Placement::Start,
                ["Inserted", "Sales Chart", "Data", "Tail"],
            ),
            (Placement::End, ["Sales Chart", "Data", "Tail", "Inserted"]),
            (
                Placement::Position(Position::new(1)),
                ["Sales Chart", "Inserted", "Data", "Tail"],
            ),
            (
                Placement::before("/sheet[2]").expect("before placement"),
                ["Sales Chart", "Inserted", "Data", "Tail"],
            ),
            (
                Placement::after("/sheet[1]").expect("after placement"),
                ["Sales Chart", "Inserted", "Data", "Tail"],
            ),
            (
                Placement::inside_at("/", Position::new(1)).expect("inside placement"),
                ["Sales Chart", "Inserted", "Data", "Tail"],
            ),
        ] {
            let added = source
                .add(
                    "/",
                    NewNode::new("sheet").prop("name", "Inserted").at(placement),
                )
                .expect("insert worksheet");
            assert_eq!(order(&added.file), expected);
            let path = added.path.to_string();
            assert_eq!(sheet_name(&added.file, &path), "Inserted");
        }
    }
}

#[test]
fn add_defaults_to_collision_free_name_and_preserves_active_tab_identity() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let active = active_tab_name(&source.to_bytes().expect("publish source"));
        let claimed = source
            .add("/", NewNode::new("sheet").prop("name", "Sheet1"))
            .expect("claim first default name");
        let added = claimed
            .file
            .add("/", NewNode::new("sheet"))
            .expect("insert default worksheet");
        let path = added.path.to_string();
        assert_eq!(sheet_name(&added.file, &path), "Sheet2");
        assert_eq!(
            active_tab_name(&added.file.to_bytes().expect("publish insertion")),
            active,
            "active tab follows its stable identity"
        );
        let reopened = open_owner(&added.file.to_bytes().expect("publish insertion"), eager);
        assert_eq!(sheet_name(&reopened, &path), "Sheet2");
    }
}

#[test]
fn add_admits_visibility_properties_and_rejects_invalid_names_positions_and_parents() {
    let source = open_owner(&fixture(), false);
    let hidden = source
        .add(
            "/",
            NewNode::new("sheet")
                .prop("name", "Hidden")
                .prop("hidden", true),
        )
        .expect("hidden worksheet");
    let hidden_path = hidden.path.to_string();
    assert_eq!(
        hidden
            .file
            .get(&hidden_path)
            .expect("hidden path")
            .view()
            .format()
            .expect("hidden properties")
            .find(|property| property.key() == "visibility")
            .map(|property| property.value().to_string()),
        Some("hidden".to_owned())
    );

    let very_hidden = source
        .add(
            "/",
            NewNode::new("sheet")
                .prop("name", "Very Hidden")
                .prop("visibility", "veryHidden"),
        )
        .expect("very hidden worksheet");
    let very_hidden_path = very_hidden.path.to_string();
    assert_eq!(
        very_hidden
            .file
            .get(&very_hidden_path)
            .expect("very hidden path")
            .view()
            .format()
            .expect("very hidden properties")
            .find(|property| property.key() == "visibility")
            .map(|property| property.value().to_string()),
        Some("veryHidden".to_owned())
    );

    let duplicate = source
        .add("/", NewNode::new("sheet").prop("name", "data"))
        .expect_err("duplicate sheet name");
    assert_eq!(duplicate.code(), Code::DuplicateName);
    let invalid = source
        .add("/", NewNode::new("sheet").prop("name", "bad/name"))
        .expect_err("invalid sheet name");
    assert_eq!(invalid.code(), Code::InvalidValue);
    let out_of_range = source
        .add(
            "/",
            NewNode::new("sheet")
                .prop("name", "Too Far")
                .at(Placement::Position(Position::new(4))),
        )
        .expect_err("position after catalog");
    assert_eq!(out_of_range.code(), Code::OutOfRange);
    let wrong_parent = source
        .add("/sheet[1]", NewNode::new("sheet"))
        .expect_err("worksheet parent");
    assert_eq!(wrong_parent.code(), Code::UnsupportedType);
}

#[test]
fn add_is_atomic_in_batches_and_can_be_followed_by_cell_edits() {
    let source = open_owner(&fixture(), false);
    let before = source.to_bytes().expect("publish source");
    let mut first = BatchItem::new("add");
    first.parent = Some("/".into());
    first.kind = Some("sheet".into());
    first.props = Some(Props::new().with("name", "Added"));
    let mut bad = BatchItem::new("add");
    bad.parent = Some("/".into());
    bad.kind = Some("not-a-sheet".into());
    let error = source
        .batch([first, bad], BatchMode::Atomic)
        .expect_err("atomic add batch");
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::UnsupportedElement);
    assert_eq!(source.to_bytes().expect("source unchanged"), before);

    let commit = source
        .edit(|edit| {
            let path = edit.add("/", NewNode::new("sheet").prop("name", "Added"))?;
            edit.set(&format!("{path}/A1"), props! { "value" => "after add" })?;
            Ok(path)
        })
        .expect("add then edit cell");
    assert_eq!(
        commit
            .file()
            .get(&format!("{}/A1", commit.value()))
            .expect("cell after add")
            .view()
            .text()
            .expect("cell text")
            .as_deref(),
        Some("after add")
    );
    let reopened = open_owner(&commit.file().to_bytes().expect("publish result"), false);
    assert_eq!(
        reopened
            .get(&format!("{}/A1", commit.value()))
            .expect("reopened cell")
            .view()
            .text()
            .expect("reopened cell text")
            .as_deref(),
        Some("after add")
    );
}

#[test]
fn mixed_set_add_move_batches_commit_sequentially_and_roll_back_atomically() {
    for eager in [false, true] {
        let source = open_owner(&fixture(), eager);
        let committed = source
            .batch(mixed_batch(false), BatchMode::Atomic)
            .expect("mixed set/add/move batch");
        assert_eq!(
            order(committed.file()),
            ["Renamed", "Tail", "Added", "Sales Chart"]
        );

        let before = source.to_bytes().expect("publish source");
        let error = source
            .batch(mixed_batch(true), BatchMode::Atomic)
            .expect_err("failing mixed batch");
        assert_eq!(error.item(), Some(3));
        assert_eq!(error.code(), Code::DuplicateName);
        assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
    }
}

fn mixed_batch(failing: bool) -> Vec<BatchItem<'static>> {
    let mut set = BatchItem::new("set");
    set.path = Some("/sheet[2]".into());
    set.props = Some(Props::new().with("name", "Renamed"));

    let mut add = BatchItem::new("add");
    add.parent = Some("/".into());
    add.kind = Some("sheet".into());
    add.props = Some(Props::new().with("name", "Added"));

    let mut move_to_end = BatchItem::new("move");
    move_to_end.path = Some("/sheet[1]".into());
    move_to_end.index = Some(3);

    let mut items = vec![set, add, move_to_end];
    if failing {
        let mut duplicate = BatchItem::new("set");
        duplicate.path = Some("/sheet[1]".into());
        duplicate.props = Some(Props::new().with("name", "Added"));
        items.push(duplicate);
    }
    items
}

#[test]
fn add_honors_semantic_limits_before_materializing_native_editor() {
    let source = open_owner(&fixture(), false);
    let limits = SemanticLimits::builder()
        .max_staged_bytes(1)
        .expect("staged limit")
        .build()
        .expect("limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.add("/", NewNode::new("sheet").prop("name", "Added")),
        )
        .expect_err("staged limit");
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn created_workbook_can_add_edit_save_and_reopen_the_new_sheet() {
    let created = create(FileFormat::Xlsb).expect("create XLSB workbook");
    let added = created
        .add("/", NewNode::new("sheet").prop("name", "Added"))
        .expect("add worksheet to created workbook");
    let path = added.path.to_string();
    let edited = added
        .file
        .set(&format!("{path}/A1"), props! { "value" => "created-add" })
        .expect("edit cell in added worksheet");
    let saved = edited.to_bytes().expect("save created workbook");
    let reopened = open_owner(&saved, true);
    assert_eq!(sheet_name(&reopened, &path), "Added");
    assert_eq!(
        reopened
            .get(&format!("{path}/A1"))
            .expect("reopened added cell")
            .view()
            .text()
            .expect("reopened added cell text")
            .as_deref(),
        Some("created-add")
    );
}
