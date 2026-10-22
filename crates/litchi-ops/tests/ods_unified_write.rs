//! Unified writes of OpenDocument spreadsheets (record 0277): cells,
//! ranges, merges, sheets, rows, named ranges, annotations, metadata,
//! batches, creation and atomic save, each read back through the unified
//! reads and the leaf.

#![cfg(feature = "ods")]
#![allow(clippy::expect_used, clippy::unwrap_used, reason = "test assertions")]

use std::path::PathBuf;

use litchi_core::FileFormat;
use litchi_ops::view::Mode;
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Placement, SaveTo, props};

/// Unwrap with the whole error chain in the panic message.
trait Must<T> {
    fn must(self) -> T;
}

impl<T> Must<T> for litchi_ops::Result<T> {
    fn must(self) -> T {
        match self {
            Ok(value) => value,
            Err(error) => {
                let mut message = error.to_string();
                let mut source = std::error::Error::source(&error);
                while let Some(cause) = source {
                    message.push_str(" <- ");
                    message.push_str(&cause.to_string());
                    source = cause.source();
                }
                panic!("{message}")
            },
        }
    }
}

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn open_bytes(relative: &str) -> File {
    litchi_ops::open_bytes(std::fs::read(fixture(relative)).expect("read")).expect("open bytes")
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .expect("node")
        .view()
        .text()
        .must()
        .map(|text| text.into_owned())
        .unwrap_or_default()
}

fn json(file: &File, path: &str) -> serde_json::Value {
    serde_json::to_value(file.get(path).expect("node").view()).unwrap()
}

/// Reopen the published bytes through the unified opener and the leaf.
fn reopen(file: &File) -> File {
    let bytes = file.to_bytes().expect("publish");
    litchi_ods::Package::from_bytes(&bytes).expect("leaf reopens the package");
    litchi_ops::open_bytes(bytes).expect("unified reopen")
}

#[test]
fn cell_values_formulas_and_clears_round_trip() {
    let original = open_bytes("odf/corpus/calc-two-sheets.ods");
    let before = original.to_bytes().must();
    let file = original
        .set("/sheet1/A1", props! { "value" => "Renamed" })
        .must()
        .set("/sheet1/B2", props! { "value" => 42 })
        .must()
        .set("/sheet1/C2", props! { "value" => true })
        .must()
        .set("/sheet1/E5", props! { "text" => "new cell" })
        .must()
        .set("/sheet1/D2", props! { "formula" => "SUM([.A2:.C2])" })
        .must()
        .set("/sheet1/A3", props! { "clear" => true })
        .must();
    // The source handle is unchanged.
    assert_eq!(original.to_bytes().must(), before);
    let file = reopen(&file);
    assert_eq!(text(&file, "/sheet1/A1"), "Renamed");
    assert_eq!(text(&file, "/sheet1/B2"), "42");
    assert_eq!(json(&file, "/sheet1/B2")["format"]["type"], "Number");
    assert_eq!(json(&file, "/sheet1/C2")["format"]["type"], "Boolean");
    assert_eq!(text(&file, "/sheet1/E5"), "new cell");
    assert_eq!(
        json(&file, "/sheet1/D2")["format"]["formula"],
        "of:=SUM([.A2:.C2])"
    );
    assert_eq!(json(&file, "/sheet1/A3")["format"]["empty"], true);
    // A literal value replaces a formula.
    let file = file.set("/sheet1/D2", props! { "value" => 7 }).must();
    assert!(json(&file, "/sheet1/D2")["format"].get("formula").is_none());
    // The second sheet is untouched.
    assert_eq!(text(&file, "/sheet2/B3"), "1");
}

#[test]
fn an_unchanged_set_is_an_exact_no_op() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let before = file.to_bytes().must();
    let same = file.set("/sheet1/A1", props! { "value" => "col A" }).must();
    assert_eq!(same.to_bytes().must(), before);
    let commit = file
        .edit(|tx| tx.set("/sheet1/A2", props! { "value" => 1 }))
        .must();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.file().to_bytes().must(), before);
}

#[test]
fn ranges_broadcast_and_merge() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let file = file.set("/sheet1/F1:G2", props! { "value" => "x" }).must();
    assert_eq!(text(&file, "/sheet1/G2"), "x");
    let merged = file.set("/sheet1/F1:G2", props! { "merge" => true }).must();
    let merged = reopen(&merged);
    assert_eq!(json(&merged, "/sheet1/F1")["format"]["merge"], "F1:G2");
    assert_eq!(json(&merged, "/sheet1/F1:G2")["format"]["merge"], true);
    assert_eq!(merged.query("mergeCell").must().count(), 1);
    let unmerged = merged.remove("/sheet1/mergeCell[1]").must();
    assert_eq!(unmerged.query("mergeCell").must().count(), 0);
}

#[test]
fn styles_apply_by_name_and_unknown_styles_are_refused() {
    let file = open_bytes("odf/corpus/calc-cell-styles.ods");
    let styled = file.set("/Sheet1/D3", props! { "style" => "ce1" }).must();
    let styled = reopen(&styled);
    assert_eq!(
        styled
            .get("/Sheet1/D3")
            .must()
            .view()
            .style()
            .must()
            .as_deref(),
        Some("ce1")
    );
    let error = file
        .set("/Sheet1/D3", props! { "style" => "no-such-style" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(error.property(), Some("style"));
    let error = file
        .set("/Sheet1/D3", props! { "font.bold" => true })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn cells_and_ranges_are_removed_in_place() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let removed = file.remove("/sheet1/A1:B2").must();
    let removed = reopen(&removed);
    assert_eq!(text(&removed, "/sheet1/A1"), "");
    assert_eq!(text(&removed, "/sheet1/C1"), "col C");
    assert_eq!(removed.query("cell").must().count(), 12);
    // Removing an absent cell is an exact no-op.
    let bytes = file.to_bytes().must();
    assert_eq!(file.remove("/sheet1/Z9").must().to_bytes().must(), bytes);
}

#[test]
fn sheets_are_added_renamed_moved_swapped_and_removed() {
    // A LibreOffice workbook: sheets are appended and removed by exact
    // table-fragment splices.
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let added = file
        .add("/", NewNode::new("sheet").prop("name", "Extra"))
        .must();
    assert_eq!(added.path.to_string(), "/sheet[3]");
    let file = reopen(&added.file);
    assert_eq!(
        file.workbook().must().worksheet_names().must(),
        ["sheet1", "sheet2", "Extra"]
    );
    let file = file.set("/Extra/A1", props! { "value" => "hi" }).must();
    assert_eq!(text(&file, "/Extra/A1"), "hi");
    let error = file
        .add("/", NewNode::new("sheet").prop("name", "SHEET1"))
        .unwrap_err();
    assert_eq!(error.code(), Code::DuplicateName);
    // Its RDF metadata and view settings are order and name owners the
    // leaf's copy and move refuse to rewrite: typed refusals.
    let error = file
        .set("/Extra", props! { "name" => "Renamed" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let file = reopen(&file.remove("/Extra").must());
    assert_eq!(
        file.workbook().must().worksheet_names().must(),
        ["sheet1", "sheet2"]
    );

    // A created workbook has no such owners.
    let file = litchi_ops::create(FileFormat::Ods)
        .must()
        .add("/", NewNode::new("sheet").prop("name", "Extra"))
        .must()
        .file
        .add("/", NewNode::new("sheet"))
        .must()
        .file;
    assert_eq!(
        file.workbook().must().worksheet_names().must(),
        ["Sheet1", "Extra", "Sheet3"]
    );
    let file = file
        .set("/Extra/A1", props! { "value" => "hi" })
        .must()
        .set("/Extra", props! { "name" => "Renamed" })
        .must();
    assert_eq!(text(&file, "/Renamed/A1"), "hi");
    let file = file.move_to("/Renamed", Placement::Start).must();
    assert_eq!(
        file.workbook().must().worksheet_names().must(),
        ["Renamed", "Sheet1", "Sheet3"]
    );
    let file = file.swap("/sheet[1]", "/sheet[3]").must();
    assert_eq!(
        file.workbook().must().worksheet_names().must(),
        ["Sheet3", "Sheet1", "Renamed"]
    );
    let file = reopen(&file);
    assert_eq!(text(&file, "/Renamed/A1"), "hi");
}

#[test]
fn rows_are_inserted_moved_and_removed() {
    // LibreOffice files carry view settings (`settings.xml`) that name row
    // positions; the leaf's logical row editor refuses to shift rows under
    // them, a typed refusal.
    let producer = open_bytes("odf/corpus/calc-two-sheets.ods");
    let error = producer.remove("/sheet1/row[2]").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);

    let file = litchi_ops::create(FileFormat::Ods)
        .must()
        .set("/Sheet1/A1:A3", props! { "value" => 0 })
        .must()
        .set("/Sheet1/A1", props! { "value" => "col A" })
        .must()
        .set("/Sheet1/A2", props! { "value" => 1 })
        .must()
        .set("/Sheet1/A3", props! { "value" => "a text" })
        .must();
    let added = file
        .add(
            "/Sheet1",
            NewNode::new("row").at(Placement::Position(litchi_core::Position::new(1))),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/row[2]");
    let file = reopen(&added.file);
    assert_eq!(text(&file, "/Sheet1/A3"), "1");
    let file = file.remove("/Sheet1/row[2]").must();
    assert_eq!(text(&file, "/Sheet1/A2"), "1");
    let moved = file.move_to("/Sheet1/row[1]", Placement::End).must();
    assert_eq!(text(&moved, "/Sheet1/A3"), "col A");
}

#[test]
fn named_ranges_are_added_set_and_removed() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let added = file
        .add(
            "/",
            NewNode::new("namedrange")
                .prop("name", "Totals")
                .prop("ref", "sheet1!$A$1:$C$2"),
        )
        .must();
    assert_eq!(added.path.to_string(), "/namedrange[1]");
    let file = reopen(&added.file);
    assert_eq!(
        json(&file, "/namedrange[Totals]")["format"]["ref"],
        "$sheet1.$A$1:.$C$2"
    );
    let file = file
        .set("/namedrange[1]", props! { "ref" => "$sheet1.$B$2" })
        .must();
    assert_eq!(text(&file, "/namedrange[1]"), "$sheet1.$B$2");
    let file = file.remove("/namedrange[1]").must();
    assert_eq!(file.query("namedrange").must().count(), 0);
}

#[test]
fn annotations_are_added_set_and_removed() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let added = file
        .add(
            "/sheet1/B2",
            NewNode::new("comment")
                .prop("text", "check this")
                .prop("author", "Reviewer"),
        )
        .must();
    assert_eq!(added.path.to_string(), "/sheet[1]/comment[1]");
    let file = reopen(&added.file);
    let comment = json(&file, "/sheet1/comment[1]");
    assert_eq!(comment["format"]["ref"], "B2");
    assert_eq!(comment["format"]["author"], "Reviewer");
    assert_eq!(comment["text"], "check this");
    let second = file
        .add(
            "/sheet2",
            NewNode::new("comment")
                .prop("ref", "C3")
                .prop("text", "two"),
        )
        .must();
    assert_eq!(second.path.to_string(), "/sheet[2]/comment[1]");
    let file = second
        .file
        .set("/sheet1/comment[1]", props! { "text" => "done" })
        .must();
    assert_eq!(text(&file, "/sheet1/comment[1]"), "done");
    let file = file.remove("/sheet1/comment[1]").must();
    assert_eq!(file.query("comment").must().count(), 1);
}

#[test]
fn metadata_is_set_through_meta_xml() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let file = file
        .set(
            "/metadata",
            props! { "title" => "Budget", "author" => "Ada", "keywords" => "q3" },
        )
        .must();
    let file = reopen(&file);
    let metadata = file.metadata().must().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Budget"));
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    assert_eq!(json(&file, "/")["format"]["title"], "Budget");
    // The metadata family (record 0288) also writes lastModifiedBy and
    // language; producer-maintained dates stay refused.
    let file = file
        .set(
            "/metadata",
            props! { "lastModifiedBy" => "Bea", "language" => "en-GB" },
        )
        .must();
    let metadata = reopen(&file).metadata().must().unwrap();
    assert_eq!(metadata.last_modified_by.as_deref(), Some("Bea"));
    // `File::metadata()` does not report `dc:language` (the shared ODF
    // conversion omits it; record 0288 gap): check the written part.
    let meta = litchi_odf_common::OwnedPackage::from_bytes(file.to_bytes().must())
        .unwrap()
        .get_file("meta.xml")
        .unwrap();
    assert!(String::from_utf8_lossy(&meta).contains(">en-GB</dc:language>"));
    assert_eq!(metadata.author.as_deref(), Some("Ada"));
    let error = file
        .set("/metadata", props! { "created" => "2020-01-01T00:00:00Z" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.property(), Some("created"));
    let error = file
        .set("/metadata", props! { "nonsense" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn creation_authors_one_sheet_and_saves_atomically() {
    let file = litchi_ops::create(FileFormat::Ods).must();
    assert_eq!(file.workbook().must().worksheet_names().must(), ["Sheet1"]);
    let file = file.set("/Sheet1/A1", props! { "value" => "hello" }).must();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("created.ods");
    let saved = file.save_as(&path).must();
    assert_eq!(saved.path, path);
    let reopened = litchi_ops::open(&path).must();
    assert_eq!(text(&reopened, "/Sheet1/A1"), "hello");
    // Path-opened handles save back to their origin.
    let edited = reopened.set("/Sheet1/A2", props! { "value" => 2 }).must();
    edited.save().must();
    assert_eq!(text(&litchi_ops::open(&path).must(), "/Sheet1/A2"), "2");
    assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 1);
    // Overwrite refusal leaves the destination unchanged.
    let options = litchi_ops::SaveOptions::new().overwrite(litchi_ops::Overwrite::Refuse);
    let before = std::fs::read(&path).unwrap();
    assert!(edited.save_with(SaveTo::Path(&path), &options).is_err());
    assert_eq!(std::fs::read(&path).unwrap(), before);

    let metadata = litchi_core::Metadata {
        title: Some("Fresh".to_owned()),
        ..litchi_core::Metadata::default()
    };
    let created = litchi_ops::create_with(
        FileFormat::Ods,
        &litchi_ops::CreateOptions::new().metadata(metadata),
    )
    .must();
    assert_eq!(
        created.metadata().must().unwrap().title.as_deref(),
        Some("Fresh")
    );
    assert_eq!(
        litchi_ops::create(FileFormat::Ods).must().to_bytes().must(),
        litchi_ops::create(FileFormat::Ods).must().to_bytes().must()
    );
}

#[test]
fn batches_coalesce_cell_sets_into_one_leaf_commit() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let commit = file
        .edit(|tx| {
            for row in 1..=200 {
                tx.set(
                    format!("/sheet1/H{row}").as_str(),
                    props! { "value" => row },
                )?;
            }
            Ok(())
        })
        .must();
    assert_eq!(commit.diagnostics().operations(), 200);
    assert_eq!(commit.diagnostics().commits(), 1);
    let edited = reopen(commit.file());
    assert_eq!(text(&edited, "/sheet1/H200"), "200");
    // The reversible patch restores the input exactly.
    let restored = commit.patch().inverse().apply(commit.file()).must();
    assert_eq!(restored.to_bytes().must(), file.to_bytes().must());

    // A repeated coordinate closes the run and observes the earlier value.
    let commit = file
        .edit(|tx| {
            tx.set("/sheet1/H1", props! { "value" => 1 })?;
            tx.set("/sheet1/H1", props! { "formula" => "[.A2]*2" })?;
            tx.set("/sheet1/H2", props! { "value" => 3 })?;
            Ok(())
        })
        .must();
    assert_eq!(commit.diagnostics().commits(), 2);
    let edited = commit.file();
    assert_eq!(
        json(edited, "/sheet1/H1")["format"]["formula"],
        "of:=[.A2]*2"
    );
    assert_eq!(text(edited, "/sheet1/H1"), "1");
}

#[test]
fn atomic_batches_leave_the_input_untouched_on_failure() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods");
    let before = file.to_bytes().must();
    let set = |path: &'static str, props: litchi_ops::Props<'static>| {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(props);
        item
    };
    let items = vec![
        set(
            "/sheet1/A1",
            litchi_ops::Props::new().with("value", "changed"),
        ),
        set(
            "/sheet1/A2",
            litchi_ops::Props::new().with("style", "missing-style"),
        ),
    ];
    let error = file.batch(items.clone(), BatchMode::Atomic).unwrap_err();
    assert_eq!(error.item(), Some(1));
    assert_eq!(file.to_bytes().must(), before);
    let report = file.batch(items, BatchMode::BestEffort).must();
    assert_eq!(text(report.file(), "/sheet1/A1"), "changed");
}

#[test]
fn flat_spreadsheets_write_through_their_projection() {
    // Record 0292: a flat spreadsheet edits its packaged projection and
    // republishes the edit into its flat bytes. A rich-text cell the leaf
    // would flatten is refused, as in a packaged spreadsheet.
    let source = std::fs::read(fixture("odf/ods/rich-text-hyperlinks.fods")).unwrap();
    let file = litchi_ops::open(fixture("odf/ods/rich-text-hyperlinks.fods")).must();
    let error = file
        .set("/sheet[1]/A1", props! { "value" => 1 })
        .unwrap_err();
    assert_ne!(error.code(), Code::InternalError);
    assert_eq!(file.to_bytes().must(), source);
    assert!(
        file.capabilities()
            .verb(litchi_ops::capabilities::Verb::Set)
            .is_supported()
    );
    let edited = file.set("/sheet[1]/D9", props! { "value" => 7 }).must();
    let reopened = litchi_ops::open_bytes(edited.to_bytes().must()).must();
    assert_eq!(text(&reopened, "/sheet[1]/D9"), "7");
}

#[test]
fn source_backed_handles_edit_and_fence_changes() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.ods");
    std::fs::copy(fixture("odf/corpus/calc-formulas.ods"), &path).unwrap();
    let file = litchi_ops::open(&path).must();
    let edited = file.set("/Sheet1/B1", props! { "value" => 9 }).must();
    assert_eq!(text(&edited, "/Sheet1/B1"), "9");
    assert_eq!(text(&edited, "/Sheet1/A4"), "7");
    // A changed source is refused before any edit.
    let mut bytes = std::fs::read(&path).unwrap();
    bytes.extend_from_slice(b"tail");
    std::fs::write(&path, bytes).unwrap();
    let error = file.set("/Sheet1/B2", props! { "value" => 1 }).unwrap_err();
    assert_eq!(error.code(), Code::IoError);
}

#[test]
fn views_reflect_writes() {
    let file = open_bytes("odf/corpus/calc-two-sheets.ods")
        .set("/sheet1/A1", props! { "value" => "Head" })
        .must();
    assert!(file.view(Mode::Text).must().contains("A1=Head"));
}
