//! Unified XLS worksheet-tab and axis structure through the public [`File`]
//! API: add, remove, move and swap tabs, sheet-name identity, and row/column
//! insertion and deletion.  Every behavior runs through both the retained
//! positional owner and the eager byte owner; failed operations leave the
//! input handle publishable byte-for-byte.
#![cfg(feature = "xls")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "integration fixtures fail fast"
)]

use std::ops::Deref;

use litchi_core::Position;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, NewNode, OpenOptions, Placement, SourceMode, props,
};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_string(data, 0, 0, "seed").unwrap();
    builder.write_number(data, 1, 1, 2.5).unwrap();
    let archive = builder.add_worksheet("Archive").unwrap();
    builder.write_number(archive, 0, 0, 9.0).unwrap();
    let notes = builder.add_worksheet("Notes").unwrap();
    builder.write_string(notes, 2, 2, "note").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

fn test_data(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls")
            .join(name),
    )
    .expect("XLS fixture")
}

struct Owner {
    file: File,
    _source_guard: Option<tempfile::TempDir>,
}

impl Deref for Owner {
    type Target = File;

    fn deref(&self) -> &File {
        &self.file
    }
}

fn open_owner(bytes: &[u8], eager: bool) -> Owner {
    if eager {
        return Owner {
            file: litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .unwrap(),
            _source_guard: None,
        };
    }
    let guard = tempfile::tempdir().unwrap();
    let path = guard.path().join("source.xls");
    std::fs::write(&path, bytes).unwrap();
    Owner {
        file: litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .unwrap(),
        _source_guard: Some(guard),
    }
}

fn names(file: &File) -> Vec<String> {
    let mut names = Vec::new();
    let mut index = 1;
    while let Ok(node) = file.get(format!("/sheet[{index}]").as_str()) {
        names.push(node.view().text().unwrap().unwrap().into_owned());
        index += 1;
    }
    names
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

/// Reopen published bytes through the eager owner, checking the complete
/// native reader accepts the artifact.
fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[test]
fn sheets_are_added_with_default_and_explicit_names() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let added = owner
            .add("/", NewNode::new("sheet").prop("name", "Summary"))
            .unwrap();
        assert_eq!(added.path.to_string(), "/sheet[4]");
        assert_eq!(names(&added.file), ["Data", "Archive", "Notes", "Summary"]);
        let first = added
            .file
            .add("/", NewNode::new("sheet").at(Placement::Start))
            .unwrap();
        assert_eq!(first.path.to_string(), "/sheet[1]");
        assert_eq!(names(&first.file)[0], "Sheet1");
        let between = first
            .file
            .add(
                "/",
                NewNode::new("sheet")
                    .prop("name", "Middle")
                    .at(Placement::after("/Data").unwrap()),
            )
            .unwrap();
        assert_eq!(
            names(&between.file),
            ["Sheet1", "Data", "Middle", "Archive", "Notes", "Summary"]
        );
        // The new worksheet accepts cell writes and survives publication.
        let written = between
            .file
            .set("/Middle/B2", props! { "value" => 7 })
            .unwrap();
        let reopened = reopen(&written);
        assert_eq!(text(&reopened, "/Middle/B2"), "7");
        assert_eq!(text(&reopened, "/Data/A1"), "seed");
        assert_eq!(text(&reopened, "/Notes/C3"), "note");
        // The input handle is unchanged.
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn added_names_are_validated_with_caseless_identity() {
    let owner = open_owner(&fixture(), true);
    for (name, code) in [
        ("data", Code::DuplicateName),
        ("ARCHIVE", Code::DuplicateName),
        ("bad/name", Code::InvalidValue),
        ("", Code::InvalidValue),
        ("'quoted'", Code::InvalidValue),
    ] {
        let error = owner
            .add("/", NewNode::new("sheet").prop("name", name))
            .expect_err(name);
        assert_eq!(error.code(), code, "{name:?}");
    }
    let error = owner
        .add(
            "/",
            NewNode::new("sheet").at(Placement::Position(Position::new(9))),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::OutOfRange);
    let error = owner.add("/Data", NewNode::new("sheet")).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedType);
    let error = owner
        .add("/", NewNode::new("sheet").prop("hidden", true))
        .unwrap_err();
    assert_ne!(error.code(), Code::InternalError);
    assert_eq!(owner.to_bytes().unwrap(), fixture());
}

#[test]
fn sheet_names_resolve_and_rename_caselessly() {
    for eager in [false, true] {
        let owner = open_owner(&fixture(), eager);
        assert_eq!(text(&owner, "/DATA/A1"), "seed");
        assert_eq!(text(&owner, "/data/A1"), "seed");
        // Recasing a tab's own name is allowed; taking another's is not.
        let recased = owner.set("/Data", props! { "name" => "DATA" }).unwrap();
        assert_eq!(names(&recased)[0], "DATA");
        let error = owner
            .set("/Data", props! { "name" => "archive" })
            .unwrap_err();
        assert_eq!(error.code(), Code::DuplicateName);
        let error = owner.set("/Data", props! { "name" => "a:b" }).unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }
}

#[test]
fn sheets_are_removed_by_path_name_or_query() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let removed = owner.remove("/archive").unwrap();
        assert_eq!(names(&removed), ["Data", "Notes"]);
        assert_eq!(text(&reopen(&removed), "/Notes/C3"), "note");
        let removed = owner.remove("/sheet[3]").unwrap();
        assert_eq!(names(&removed), ["Data", "Archive"]);
        let error = owner
            .remove(
                litchi_core::selector::Query::parse("sheet")
                    .unwrap()
                    .with_scope(litchi_core::selector::Scope::Path(Box::new(
                        litchi_core::selector::Path::parse("/").unwrap(),
                    ))),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
    // The last tab cannot be removed.
    let single = litchi_ops::create(litchi_core::FileFormat::Xls).unwrap();
    assert_eq!(
        single.remove("/sheet[1]").unwrap_err().code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn removal_refuses_a_tab_that_owns_a_scoped_name() {
    let bytes = test_data("SimpleWithPrintArea.xls");
    let owner = open_owner(&bytes, true);
    let error = owner.remove("/Sheet1").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(owner.to_bytes().unwrap(), bytes);
    let removed = owner.remove("/Sheet2").unwrap();
    assert_eq!(names(&removed), ["Sheet1", "Sheet3"]);
}

#[test]
fn sheets_move_and_swap_through_the_complete_catalog() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let moved = owner.move_to("/Data", Placement::End).unwrap();
        assert_eq!(names(&moved), ["Archive", "Notes", "Data"]);
        let moved = owner
            .move_to("/Notes", Placement::before("/Data").unwrap())
            .unwrap();
        assert_eq!(names(&moved), ["Notes", "Data", "Archive"]);
        let moved = owner
            .move_to("/Data", Placement::after("/Archive").unwrap())
            .unwrap();
        assert_eq!(names(&moved), ["Archive", "Data", "Notes"]);
        let moved = owner
            .move_to("/Notes", Placement::Position(Position::new(1)))
            .unwrap();
        assert_eq!(names(&moved), ["Data", "Notes", "Archive"]);
        let swapped = owner.swap("/Data", "/Notes").unwrap();
        assert_eq!(names(&swapped), ["Notes", "Archive", "Data"]);
        let reopened = reopen(&swapped);
        assert_eq!(text(&reopened, "/Data/B2"), "2.5");
        assert_eq!(text(&reopened, "/sheet[1]/C3"), "note");
        // Exact no-ops keep the input bytes.
        let same = owner
            .move_to("/Archive", Placement::Position(Position::new(1)))
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        let same = owner.swap("/Data", "/data").unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
        let error = owner
            .move_to("/Data", Placement::Position(Position::new(3)))
            .unwrap_err();
        assert_eq!(error.code(), Code::OutOfRange);
        let error = owner.move_to("/Data/A1", Placement::End).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedType);
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn structure_edits_compose_in_transactions_with_exact_inverse() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let commit = owner
            .edit(|tx| {
                tx.set("/Data/A1", props! { "value" => "changed" })?;
                let path = tx.add("/", NewNode::new("sheet").prop("name", "Extra"))?;
                assert_eq!(path.to_string(), "/sheet[4]");
                tx.move_to("/Extra", Placement::Start)?;
                tx.swap("/Data", "/Notes")?;
                tx.remove("/Archive")?;
                tx.set("/Extra/A1", props! { "value" => 1 })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .unwrap();
        assert_eq!(names(commit.file()), ["Extra", "Notes", "Data"]);
        assert_eq!(text(commit.file(), "/Data/A1"), "changed");
        assert_eq!(
            commit.patch().apply(&owner).unwrap().to_bytes().unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn batches_carry_structure_items_with_attribution() {
    let owner = open_owner(&fixture(), true);
    let mut add = BatchItem::new("add");
    add.parent = Some("/".into());
    add.kind = Some("sheet".into());
    add.props = Some(litchi_ops::Props::from_props(props! { "name" => "Batch" }).unwrap());
    let mut swap = BatchItem::new("swap");
    swap.path = Some("/Data".into());
    swap.path2 = Some("/Batch".into());
    let mut bad = BatchItem::new("remove");
    bad.path = Some("/Missing".into());
    let result = owner
        .batch(vec![add.clone(), swap.clone()], BatchMode::Atomic)
        .unwrap();
    assert_eq!(names(result.file()), ["Batch", "Archive", "Notes", "Data"]);
    let error = owner
        .batch(vec![add, swap, bad], BatchMode::Atomic)
        .unwrap_err();
    assert_eq!(error.item(), Some(2));
    assert_eq!(owner.to_bytes().unwrap(), fixture());
}

#[test]
fn rows_and_columns_are_inserted_and_empty_ones_deleted() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let added = owner
            .add(
                "/Data",
                NewNode::new("row").at(Placement::Position(Position::new(0))),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/sheet[1]/row[1]");
        assert_eq!(text(&added.file, "/Data/A2"), "seed");
        assert_eq!(text(&added.file, "/Data/B3"), "2.5");
        let added = added
            .file
            .add(
                "/Data",
                NewNode::new("col").at(Placement::before("/Data/col[B]").unwrap()),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/sheet[1]/col[B]");
        assert_eq!(text(&added.file, "/Data/C3"), "2.5");
        let reopened = reopen(&added.file);
        assert_eq!(text(&reopened, "/Data/A2"), "seed");
        // The inserted empty row and column can be deleted again.
        let restored = added
            .file
            .remove("/Data/col[B]")
            .unwrap()
            .remove("/Data/row[1]")
            .unwrap();
        assert_eq!(text(&restored, "/Data/A1"), "seed");
        assert_eq!(text(&restored, "/Data/B2"), "2.5");
        // Deleting a row that holds cells is refused by the native owner.
        let error = owner.remove("/Data/row[1]").unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = owner.add("/Data", NewNode::new("row")).unwrap_err();
        assert_eq!(error.code(), Code::MissingArgument);
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn signed_and_encrypted_workbooks_refuse_structure_edits() {
    let source = fixture();
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(source),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    for path in ole.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = ole.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"synthetic")
        .unwrap();
    let mut signed = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut signed).unwrap();
    let signed = litchi_ops::open_bytes(signed.into_inner()).unwrap();
    assert!(signed.is_signed());
    for error in [
        signed.add("/", NewNode::new("sheet")).unwrap_err(),
        signed.remove("/Archive").unwrap_err(),
        signed.move_to("/Data", Placement::End).unwrap_err(),
        signed.swap("/Data", "/Notes").unwrap_err(),
    ] {
        assert_eq!(error.code(), Code::DocumentProtected);
    }
}

/// Record 0316: an array formula and a data table move whole with their
/// cells, on both owners, and removing the inserted row and column restores
/// the input bytes.
#[test]
fn array_formulas_and_data_tables_move_with_their_cells() {
    use litchi_xls::formula_metadata::Range;
    use litchi_xls::{DataTable, DataTableInputCell, DataTableRange};

    let mut builder = litchi_xls::Builder::new();
    let sheet = builder.add_worksheet("Model").unwrap();
    builder.write_number(sheet, 0, 0, 3.0).unwrap();
    builder.write_number(sheet, 0, 6, 1.0).unwrap();
    builder
        .write_array_formula(sheet, Range::try_new(6, 1, 7, 2).unwrap(), "A1+7")
        .unwrap();
    let table = DataTable::one_variable(
        DataTableRange::new(12, 16, 3, 5).unwrap(),
        true,
        DataTableInputCell::Present { row: 0, col: 6 },
    );
    builder.add_data_table(sheet, 11, 2, table).unwrap();
    let bytes = builder.build().unwrap().to_bytes().unwrap();

    for eager in [true, false] {
        let file = open_owner(&bytes, eager);
        for (kind, rows, columns) in [("row", 1_u16, 0_u16), ("column", 0, 1)] {
            let added = file
                .add("/sheet[1]", NewNode::new(kind).at(Placement::Start))
                .unwrap_or_else(|error| {
                    let mut text = error.to_string();
                    let mut source = std::error::Error::source(&error);
                    while let Some(cause) = source {
                        text.push_str(&format!(" <- {cause}"));
                        source = cause.source();
                    }
                    panic!("{kind} eager={eager}: {text}")
                });
            let shifted = added.file.to_bytes().unwrap();
            let workbook =
                litchi_xls::raw::Reader::new(std::io::Cursor::new(shifted.clone())).unwrap();
            let worksheet = workbook.xls_worksheet(0).unwrap();
            assert_eq!(
                worksheet.array_formulas().next().unwrap().range(),
                Range::try_new(
                    6 + u32::from(rows),
                    1 + columns,
                    7 + u32::from(rows),
                    2 + columns
                )
                .unwrap(),
                "{kind}"
            );
            let moved = worksheet.data_tables()[0];
            assert_eq!(
                moved.range(),
                DataTableRange::new(
                    12 + rows,
                    16 + rows,
                    3 + u8::try_from(columns).unwrap(),
                    5 + u8::try_from(columns).unwrap(),
                )
                .unwrap(),
                "{kind}"
            );
            let removed = added.file.remove(&added.path).unwrap();
            assert_eq!(removed.to_bytes().unwrap(), bytes, "{kind} eager={eager}");
        }
    }
}
