//! Unified XLSB complete-catalog sheet removal through the public facade.
#![cfg(feature = "xlsb")]

use std::collections::HashSet;
use std::io::Cursor;
use std::mem::size_of;

use litchi_core::FileFormat;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_opc::OpcPackage;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, NewNode, OpenOptions, Props, SourceMode, create,
};
use litchi_xlsb::raw::{Encoder, Records, kind};

const FIRST: &str = "/sheet[1]";
const SECOND: &str = "/sheet[2]";
const THIRD: &str = "/sheet[3]";

fn mixed_fixture() -> Vec<u8> {
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
    let mut archive = builder::Worksheet::new("Archive");
    archive.set_cell(0, 0, "archive");
    builder.add_worksheet(archive);
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

fn two_sheet_fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut builder = Builder::new();
    for (name, value) in [("Visible", "visible"), ("Hidden", "hidden")] {
        let mut sheet = builder::Worksheet::new(name);
        sheet.set_cell(0, 0, value);
        builder.add_worksheet(sheet);
    }
    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build two-sheet fixture")
        .write_to(&mut output)
        .expect("serialize two-sheet fixture");
    output.into_inner()
}

fn many_sheet_fixture(count: usize) -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut builder = Builder::new();
    for index in 0..count {
        let mut sheet = builder::Worksheet::new(format!("S{index:02}"));
        sheet.set_cell(0, 0, format!("value-{index}"));
        builder.add_worksheet(sheet);
    }
    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build many-sheet fixture")
        .write_to(&mut output)
        .expect("serialize many-sheet fixture");
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

fn sheet_names(file: &File) -> Vec<String> {
    (1..=4)
        .filter_map(|position| {
            let path = format!("/sheet[{position}]");
            file.find(&path)
                .ok()
                .flatten()
                .map(|_| sheet_name(file, &path))
        })
        .collect()
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

fn scoped_sheet_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse sheet selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn sheet_parts(bytes: &[u8]) -> usize {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    package
        .iter_parts()
        .filter(|part| part.content_type() == "application/vnd.ms-excel.worksheet")
        .count()
}

fn chart_parts(bytes: &[u8]) -> usize {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default())
        .expect("parse XLSB package");
    package
        .iter_parts()
        .filter(|part| part.content_type() == "application/vnd.ms-excel.chartsheet")
        .count()
}

/// Keep the chart-sheet root but remove its outgoing chart/drawing edges.
/// This models the narrow native admission case where the selected root has
/// no dependent internal relationship to dispose of; the original fixture
/// remains the refusal case for a chart-backed tab.
fn dependency_free_chart_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&mixed_fixture(), litchi_opc::SourceLimits::default())
        .expect("parse chart package");
    let chart_sheet = package
        .iter_parts()
        .find(|part| part.content_type() == "application/vnd.ms-excel.chartsheet")
        .expect("chart sheet part")
        .partname()
        .clone();
    let mut pending = vec![chart_sheet.clone()];
    let mut owned = HashSet::new();
    while let Some(source) = pending.pop() {
        if !owned.insert(source.clone()) {
            continue;
        }
        let targets = package
            .get_part(&source)
            .expect("get chart dependency part")
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .map(|relationship| {
                relationship
                    .target_partname()
                    .expect("resolve chart dependency target")
            })
            .collect::<Vec<_>>();
        pending.extend(targets);
    }
    let (relationship_ids, original_blob) = {
        let part = package
            .get_part(&chart_sheet)
            .expect("get chart sheet part");
        (
            part.rels()
                .iter()
                .map(|relationship| relationship.r_id().to_owned())
                .collect::<Vec<_>>(),
            part.blob().to_vec(),
        )
    };
    let mut rewritten_blob = Vec::new();
    {
        let mut writer = Encoder::new(&mut rewritten_blob, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original_blob, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.expect("read chart sheet record");
            if record.kind() != kind::DRAWING {
                writer
                    .write_record(record.kind(), record.payload())
                    .expect("rewrite chart sheet record");
            }
        }
    }
    let part = package
        .get_part_mut(&chart_sheet)
        .expect("get mutable chart sheet part");
    part.set_blob(rewritten_blob);
    for relationship_id in relationship_ids {
        part.rels_mut().remove(&relationship_id);
    }
    owned.remove(&chart_sheet);
    for part in owned {
        package.remove_part(&part);
    }
    let mut output = Cursor::new(Vec::new());
    package
        .write_to(&mut output)
        .expect("serialize dependency-free chart package");
    output.into_inner()
}

fn remove_batch(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

#[test]
fn chart_tabs_with_children_are_refused_but_dependency_free_roots_remove() {
    for eager in [false, true] {
        let source_bytes = mixed_fixture();
        let source = open_owner(&source_bytes, eager);
        assert_eq!(
            sheet_names(&source),
            ["Sales Chart", "Data", "Archive", "Tail"]
        );
        let before = source.to_bytes().expect("publish source");

        let error = source
            .remove(FIRST)
            .expect_err("chart tab with outgoing dependencies");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(source.to_bytes().expect("publish refused source"), before);
        assert_eq!(sheet_parts(&source_bytes), 3);
        assert_eq!(chart_parts(&source_bytes), 1);

        let dependency_free = dependency_free_chart_fixture();
        let source = open_owner(&dependency_free, eager);
        let removed = source
            .remove(FIRST)
            .expect("remove dependency-free chart root");
        assert_eq!(sheet_names(&removed), ["Data", "Archive", "Tail"]);
        let removed_bytes = removed.to_bytes().expect("publish removal");
        assert_eq!(sheet_parts(&removed_bytes), 3);
        assert_eq!(chart_parts(&removed_bytes), 0);
    }
}

#[test]
fn homogeneous_sheet_query_removes_source_ordinals_in_one_commit() {
    let selector = scoped_sheet_query("sheet[name=Archive or name=Tail]");
    for eager in [false, true] {
        let source = open_owner(&mixed_fixture(), eager);
        let commit = source
            .edit(|edit| edit.remove(selector.clone()))
            .expect("remove query-selected sheets");
        assert_eq!(sheet_names(commit.file()), ["Sales Chart", "Data"]);
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.diagnostics().commits(), 1);
        let reopened = open_owner(
            &commit.file().to_bytes().expect("publish query removal"),
            eager,
        );
        assert_eq!(sheet_names(&reopened), ["Sales Chart", "Data"]);
    }
}

#[test]
fn sheet_removal_history_charge_counts_all_source_positions() {
    let selected_names = (0..17)
        .map(|index| format!("name=S{index:02}"))
        .collect::<Vec<_>>();
    let selector_text = format!("sheet[{}]", selected_names.join(" or "));
    let selector = scoped_sheet_query(&selector_text);
    let source = open_owner(&many_sheet_fixture(20), false);
    let before = source.to_bytes().expect("publish many-sheet source");
    let committed = source
        .edit(|edit| edit.remove(selector.clone()))
        .expect("remove seventeen source tabs");
    let after = committed
        .file()
        .to_bytes()
        .expect("publish many-sheet removal");
    let encoded_remove = 9usize + 17usize * size_of::<u64>();
    let exact_charge = before
        .len()
        .checked_add(after.len())
        .and_then(|bytes| bytes.checked_add(encoded_remove + 32))
        .expect("history charge");
    assert_eq!(committed.diagnostics().snapshot_bytes(), exact_charge);
    assert_eq!(committed.diagnostics().leaf_operations(), 1);

    let bound = exact_charge.checked_sub(1).expect("nonzero history charge");
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(bound)
        .expect("snapshot bound")
        .build()
        .expect("semantic limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.remove(selector),
        )
        .expect_err("exact history charge must enforce the bound");
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn removal_preserves_active_identity_and_rejects_last_visible_tab() {
    for eager in [false, true] {
        let source = open_owner(&dependency_free_chart_fixture(), eager);
        let active = active_tab_name(&source.to_bytes().expect("publish source"));
        let removed = source.remove(FIRST).expect("remove active chart tab");
        let next_active = active_tab_name(&removed.to_bytes().expect("publish result"));
        assert_ne!(next_active, "Sales Chart");
        if active == "Sales Chart" {
            assert_eq!(next_active, "Data");
        } else {
            assert_eq!(next_active, active);
        }

        let hidden = open_owner(&two_sheet_fixture(), eager)
            .set(SECOND, Props::new().with("visibility", "hidden"))
            .expect("hide second tab");
        let hidden_before = hidden.to_bytes().expect("publish hidden workbook");
        let error = hidden
            .remove(FIRST)
            .expect_err("cannot remove last visible tab");
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(
            hidden
                .to_bytes()
                .expect("publish unchanged hidden workbook"),
            hidden_before
        );
    }
}

#[test]
fn empty_sheet_query_and_mixed_queries_are_atomic_refusals() {
    let source = open_owner(&mixed_fixture(), false);
    let before = source.to_bytes().expect("publish source");
    let empty = source
        .remove(scoped_sheet_query("sheet[name=Missing]"))
        .expect_err("empty sheet query");
    assert_eq!(empty.code(), Code::ZeroMatches);
    assert_eq!(source.to_bytes().expect("source after empty query"), before);

    let mixed = source
        .remove(scoped_sheet_query("sheet, sheet > cell"))
        .expect_err("mixed sheet and cell query");
    assert_eq!(mixed.code(), Code::UnsupportedElement);

    let error = source
        .batch(
            [remove_batch(THIRD), remove_batch("/sheet[99]")],
            BatchMode::Atomic,
        )
        .expect_err("atomic removal batch");
    assert_eq!(error.item(), Some(1));
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(source.to_bytes().expect("atomic source unchanged"), before);
}

#[test]
fn mixed_set_add_remove_move_commits_or_rolls_back_as_one_transaction() {
    let source = open_owner(&mixed_fixture(), false);
    let committed = source
        .edit(|edit| {
            edit.set(SECOND, Props::new().with("name", "Input"))?;
            edit.add("/", NewNode::new("sheet").prop("name", "Added"))?;
            edit.remove(THIRD)?;
            edit.move_to("/sheet[4]", litchi_ops::Placement::End)
        })
        .expect("mixed set/add/remove/move");
    assert_eq!(
        sheet_names(committed.file()),
        ["Sales Chart", "Input", "Tail", "Added"]
    );

    let before = source.to_bytes().expect("publish source");
    let error = source
        .edit(|edit| {
            edit.remove(THIRD)?;
            edit.remove("/sheet[99]")
        })
        .expect_err("failing mixed removal transaction");
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(source.to_bytes().expect("source remains unchanged"), before);
}

#[test]
fn removal_honors_expanded_and_staged_limits_before_native_materialization() {
    let source = open_owner(&mixed_fixture(), false);
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .expect("expanded limit")
        .max_staged_bytes(1)
        .expect("staged limit")
        .build()
        .expect("limits");
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.remove(scoped_sheet_query("sheet[name=Data or name=Tail]")),
        )
        .expect_err("sheet removal limits");
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn created_workbook_can_add_remove_save_and_reopen() {
    let created = create(FileFormat::Xlsb).expect("create XLSB workbook");
    let added = created
        .add("/", NewNode::new("sheet").prop("name", "Added"))
        .expect("add worksheet");
    let removed = added
        .file
        .remove(added.path)
        .expect("remove added worksheet");
    let bytes = removed.to_bytes().expect("publish removal");
    let reopened = open_owner(&bytes, true);
    assert!(
        reopened
            .find("/sheet[1]")
            .expect("find surviving sheet")
            .is_some()
    );
}
