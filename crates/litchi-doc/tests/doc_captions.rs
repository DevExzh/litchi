use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::captions::{
    AutoEntry, AutoTable, Definition, Editor, Format, Heading, Info, LabelTable, Location,
    Numbering, Separator, Tables,
};
use std::io::Cursor;

#[test]
fn facade_exposes_contextual_caption_types_without_repeated_prefixes() {
    let info = Info::new(
        Location::Above,
        Some(Numbering::new(Heading::Level1, Separator::Period)),
        false,
        Format::Arabic,
    );
    let labels =
        LabelTable::try_new(vec![Definition::try_new("Figure".into(), info).unwrap()]).unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
    ])
    .unwrap();
    let tables = Tables::try_new(Some(labels.clone()), Some(auto.clone())).unwrap();

    assert_eq!(tables.labels(), Some(&labels));
    assert_eq!(tables.auto(), Some(&auto));
    assert_eq!(
        labels.to_bytes().unwrap(),
        LabelTable::parse_bytes(&labels.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
    assert_eq!(
        auto.to_bytes().unwrap(),
        AutoTable::parse_bytes(&auto.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
}

#[test]
fn package_editor_publishes_caption_crud_without_rewriting_unrelated_table_bytes() {
    let original_table = b"opaque table prefix\0";
    let original = base_template(original_table);
    let mut editor = Editor::open(original).expect("template package opens");
    assert!(!editor.captions().is_present());

    let tables = sample_tables();
    let committed = editor.set(tables.clone()).expect("caption tables publish");
    let bytes = committed
        .snapshot()
        .finish()
        .expect("package snapshot renders");
    let reopened = Editor::open(bytes.clone()).expect("published package reopens");
    assert_eq!(reopened.tables(), &tables);
    assert!(!reopened.is_changed());

    let mut ole = OleFile::open(Cursor::new(bytes.clone())).expect("published CFB opens");
    let table = ole.open_stream(&["0Table"]).expect("table stream exists");
    assert!(table.starts_with(original_table));
    assert!(committed.package_patch().inverse().apply(&bytes).is_ok());

    let mut cleared = Editor::open(bytes).expect("published package opens for clear");
    let cleared = cleared.clear().expect("caption ranges clear");
    let cleared_bytes = cleared.snapshot().finish().expect("clear renders");
    let mut ole = OleFile::open(Cursor::new(cleared_bytes)).expect("cleared CFB opens");
    let word = ole
        .open_stream(&["WordDocument"])
        .expect("WordDocument stream exists");
    let caption_pointer = 154 + 52 * 8;
    let auto_pointer = 154 + 53 * 8;
    assert_eq!(&word[caption_pointer..caption_pointer + 8], &[0; 8]);
    assert_eq!(&word[auto_pointer..auto_pointer + 8], &[0; 8]);
}

#[test]
fn package_editor_rejects_stale_transactions_and_keeps_noop_bytes_exact() {
    let mut editor = Editor::open(base_template(b"untouched")).expect("template opens");
    let labels = sample_tables().labels().cloned().expect("labels");
    editor
        .replace_labels(labels.clone())
        .expect("labels publish");
    let published = editor.snapshot().expect("snapshot captures");
    let exact = published.finish().expect("snapshot renders");

    let mut noop = Editor::open(exact.clone()).expect("published package reopens");
    let noop_commit = noop
        .set(noop.tables().clone())
        .expect("semantic no-op publishes");
    assert_eq!(noop_commit.snapshot().finish().unwrap(), exact);

    let stale = editor.edit();
    editor
        .replace_auto(
            AutoTable::try_new(vec![
                AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
            ])
            .unwrap(),
        )
        .expect("auto-caption rules publish");
    assert!(editor.apply(stale).is_err());
}

#[test]
fn package_editor_rejects_caption_ranges_on_ordinary_documents() {
    assert!(Editor::open(base_document(b"table", false)).is_err());
}

fn sample_tables() -> Tables {
    let info = Info::new(Location::Above, None, false, Format::Arabic);
    let labels =
        LabelTable::try_new(vec![Definition::try_new("Figure".into(), info).unwrap()]).unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
    ])
    .unwrap();
    Tables::try_new(Some(labels), Some(auto)).unwrap()
}

fn base_template(table: &[u8]) -> Vec<u8> {
    base_document(table, true)
}

fn base_document(table: &[u8], template: bool) -> Vec<u8> {
    let pointer_count = 117usize;
    let word_len = 154 + pointer_count * 8;
    let mut word = vec![0; word_len];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    word[10..12].copy_from_slice(&u16::from(template).to_le_bytes());
    word[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}
