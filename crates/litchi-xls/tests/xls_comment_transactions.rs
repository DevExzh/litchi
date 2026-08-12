use litchi_xls::cell_values::{Reference, Selector};
use litchi_xls::comments::{Snapshot, Update, Value};
use litchi_xls::writer::Writer;
use litchi_xls::{Visibility, Workbook};
use std::collections::BTreeMap;
use std::io::Cursor;

fn authored(comments: usize) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Notes").unwrap();
    for index in 0..comments {
        writer
            .add_comment(
                sheet,
                u32::try_from(index).unwrap(),
                1,
                &format!("Author {index}"),
                &format!("text {index}"),
            )
            .unwrap();
    }
    let other = writer.add_worksheet("Untouched").unwrap();
    writer.write_number(other, 20, 4, 42.0).unwrap();
    writer
        .add_comment(other, 4, 3, "Other", "preserve me")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read_value(bytes: &[u8], sheet: usize, comment: usize) -> Value {
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    Value::from_comment(&workbook.xls_worksheet(sheet).unwrap().comments()[comment])
}

fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let paths = ole.list_streams();
    paths
        .into_iter()
        .map(|path| {
            let borrowed: Vec<_> = path.iter().map(String::as_str).collect();
            let value = ole.open_stream(&borrowed).unwrap();
            (path, value)
        })
        .collect()
}

fn add_stream(bytes: Vec<u8>, name: &str) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&[name], b"sentinel").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn replace_workbook_record_kind(bytes: Vec<u8>, from: u16, to: u16) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let mut workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut offset = 0_usize;
    loop {
        let kind = u16::from_le_bytes(workbook[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            workbook[offset + 2..offset + 4].try_into().unwrap(),
        ));
        if kind == from {
            workbook[offset..offset + 2].copy_from_slice(&to.to_le_bytes());
            break;
        }
        offset += 4 + length;
    }
    let mut writer = litchi_cfb::OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn updates_existing_comments_atomically_and_reversibly() {
    let bytes = authored(3);
    let before_workbook = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let before_comments = before_workbook
        .xls_worksheet(0)
        .unwrap()
        .comments()
        .to_vec();
    let before_streams = streams(&bytes);
    let snapshot = Snapshot::from_bytes(bytes.clone()).unwrap();
    assert_eq!(snapshot.worksheet_count(), 2);
    assert_eq!(
        snapshot
            .worksheet(Selector::Name("notes"))
            .unwrap()
            .unwrap()
            .comments()
            .len(),
        3
    );

    let first = Update::new(
        Reference::new(0, 1).unwrap(),
        Value::new("Renamed", "short").unwrap(),
    );
    let middle = Update::new(
        Reference::new(1, 1).unwrap(),
        Value::new("Unicode 作者", "wide 😀 text").unwrap(),
    );
    let last = Update::new(
        Reference::new(2, 1).unwrap(),
        Value::new("Last", "x".repeat(9_000)).unwrap(),
    );
    let mut edit = snapshot.edit();
    edit.replace_many(Selector::Name("Notes"), [first, middle, last])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_comments(), 3);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    assert_eq!(commit.patch().operations().len(), 3);
    assert_eq!(
        read_value(commit.snapshot().bytes(), 0, 0),
        Value::new("Renamed", "short").unwrap()
    );
    assert_eq!(
        read_value(commit.snapshot().bytes(), 0, 1),
        Value::new("Unicode 作者", "wide 😀 text").unwrap()
    );
    assert_eq!(
        read_value(commit.snapshot().bytes(), 0, 2),
        Value::new("Last", "x".repeat(9_000)).unwrap()
    );
    assert_eq!(
        read_value(commit.snapshot().bytes(), 1, 0),
        Value::new("Other", "preserve me").unwrap()
    );
    let after_workbook = Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    assert!(
        after_workbook
            .xls_worksheet(1)
            .unwrap()
            .row_block_index()
            .unwrap()
            .is_some()
    );
    assert!(
        after_workbook
            .xls_worksheet(0)
            .unwrap()
            .comments()
            .iter()
            .all(|comment| comment.visibility() == Visibility::Hidden)
    );
    for (before, after) in before_comments
        .iter()
        .zip(after_workbook.xls_worksheet(0).unwrap().comments())
    {
        assert_eq!(before.identity(), after.identity());
        assert_eq!(before.object_properties(), after.object_properties());
        assert_eq!(before.object_subrecords(), after.object_subrecords());
        assert_eq!(before.object_padding(), after.object_padding());
        assert_eq!(before.text_properties(), after.text_properties());
    }
    let after_streams = streams(commit.snapshot().bytes());
    assert_eq!(
        before_streams.keys().collect::<Vec<_>>(),
        after_streams.keys().collect::<Vec<_>>()
    );
    for (path, before) in before_streams {
        if path
            .last()
            .is_some_and(|name| name == "Workbook" || name == "Book")
        {
            continue;
        }
        assert_eq!(after_streams.get(&path), Some(&before));
    }

    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    let restored = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(restored.bytes(), bytes);
    assert!(commit.patch().apply(&applied).is_err());
}

#[test]
fn exact_noop_and_batch_failure_do_not_publish_bytes() {
    let snapshot = Snapshot::from_bytes(authored(3)).unwrap();
    let original = read_value(snapshot.bytes(), 0, 0);
    let mut edit = snapshot.edit();
    edit.replace(
        Selector::Position(0),
        Reference::new(0, 1).unwrap(),
        original,
    )
    .unwrap();
    edit.set_visibility(
        Selector::Position(0),
        Reference::new(0, 1).unwrap(),
        Visibility::Hidden,
    )
    .unwrap();
    assert!(
        edit.set_visibility(
            Selector::Position(0),
            Reference::new(0, 1).unwrap(),
            Visibility::Visible,
        )
        .is_err()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), snapshot.bytes());

    let good = Update::new(
        Reference::new(0, 1).unwrap(),
        Value::new("Good", "changed").unwrap(),
    );
    let duplicate = Update::new(
        Reference::new(0, 1).unwrap(),
        Value::new("Other", "late").unwrap(),
    );
    let mut edit = snapshot.edit();
    assert!(
        edit.replace_many(Selector::Position(0), [good, duplicate])
            .is_err()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
}

#[test]
fn refuses_unsupported_lifecycle_and_enforces_batch_bound() {
    let snapshot = Snapshot::from_bytes(authored(257)).unwrap();
    let mut edit = snapshot.edit();
    assert!(
        edit.remove(Selector::Position(0), Reference::new(0, 1).unwrap())
            .is_err()
    );
    let updates = (0..257).map(|row| {
        Update::new(
            Reference::new(row, 1).unwrap(),
            Value::new(format!("A{row}"), format!("v{row}")).unwrap(),
        )
    });
    assert!(edit.replace_many(Selector::Position(0), updates).is_err());
    assert!(edit.commit().unwrap().patch().is_empty());
}

#[test]
fn refuses_protected_signed_encrypted_and_invalid_values() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Protected").unwrap();
    writer.add_comment(sheet, 0, 0, "A", "text").unwrap();
    writer
        .protect_sheet(sheet, Some("password"), true, false)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let snapshot = Snapshot::from_bytes(output.into_inner()).unwrap();
    let mut edit = snapshot.edit();
    edit.replace(
        Selector::Position(0),
        Reference::new(0, 0).unwrap(),
        Value::new("A", "text").unwrap(),
    )
    .unwrap();
    edit.set_visibility(
        Selector::Position(0),
        Reference::new(0, 0).unwrap(),
        Visibility::Hidden,
    )
    .unwrap();
    assert!(
        edit.replace(
            Selector::Position(0),
            Reference::new(0, 0).unwrap(),
            Value::new("B", "changed").unwrap(),
        )
        .is_err()
    );
    assert!(edit.commit().unwrap().patch().is_empty());

    assert!(Value::new("", "text").is_err());
    assert!(Value::new("A".repeat(55), "text").is_err());
    assert!(Value::new("A", "\0").is_err());
    assert!(Snapshot::from_bytes(add_stream(authored(1), "DigitalSignature")).is_err());
    assert!(Snapshot::from_bytes(add_stream(authored(1), "EncryptedPackage")).is_err());
    assert!(
        Snapshot::from_bytes(replace_workbook_record_kind(authored(1), 0x0042, 0x002F)).is_err()
    );

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Empty").unwrap();
    writer.add_comment(sheet, 0, 0, "A", "").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let empty = Snapshot::from_bytes(output.into_inner()).unwrap();
    let mut edit = empty.edit();
    assert!(
        edit.replace(
            Selector::Position(0),
            Reference::new(0, 0).unwrap(),
            Value::new("A", "now nonempty").unwrap(),
        )
        .is_err()
    );
    assert!(edit.commit().unwrap().patch().is_empty());
}

#[test]
fn updates_poi_fixture_and_reopens() {
    let corpus = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet");
    let path = corpus.join("SimpleWithComments.xls");
    let bytes = std::fs::read(path).unwrap();
    let snapshot = Snapshot::from_bytes(bytes).unwrap();
    let worksheet = snapshot.worksheet(Selector::Position(0)).unwrap().unwrap();
    let source = worksheet.comments().next().unwrap();
    let reference = Reference::new(u32::from(source.row()), u32::from(source.column())).unwrap();
    let mut edit = snapshot.edit();
    edit.replace(
        Selector::Position(0),
        reference,
        Value::new("POI", "updated without normalizing OBJ").unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let workbook = Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    let updated = &workbook.xls_worksheet(0).unwrap().comments()[0];
    assert_eq!(updated.author(), "POI");
    assert_eq!(updated.text(), "updated without normalizing OBJ");

    let mixed = std::fs::read(corpus.join("DrawingAndComments.xls")).unwrap();
    let mixed = Snapshot::from_bytes(mixed).unwrap();
    let worksheet = mixed.worksheet(Selector::Position(0)).unwrap().unwrap();
    let source = worksheet.comments().next().unwrap();
    let reference = Reference::new(u32::from(source.row()), u32::from(source.column())).unwrap();
    let mut edit = mixed.edit();
    edit.replace(
        Selector::Position(0),
        reference,
        Value::new("Mixed", "comment beside unrelated drawing shapes").unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let workbook = Workbook::new(Cursor::new(commit.snapshot().bytes())).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().comments()[0].text(),
        "comment beside unrelated drawing shapes"
    );
}
