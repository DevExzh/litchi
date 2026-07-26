use litchi_ole::doc::{AssociatedStringSlot, DocWriter, Package, SavedByEntry, SavedByTable};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

const ASSOCIATED_STRINGS_FIB_INDEX: usize = 32;
const SAVED_BY_FIB_INDEX: usize = 71;

static NEXT_FILE: AtomicUsize = AtomicUsize::new(0);

fn temporary_doc_path() -> std::path::PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-auxiliary-strings-{}-{}.doc",
        std::process::id(),
        NEXT_FILE.fetch_add(1, Ordering::Relaxed)
    ))
}

fn configured_writer() -> (DocWriter, SavedByTable) {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Body").unwrap();
    writer
        .set_associated_string(AssociatedStringSlot::TemplatePath, "C:\\模板\\Normal.dot")
        .unwrap();
    writer
        .set_associated_string(AssociatedStringSlot::Title, "Quarterly 😀")
        .unwrap();
    writer
        .set_associated_string(AssociatedStringSlot::Author, "张三")
        .unwrap();
    writer
        .set_associated_string(AssociatedStringSlot::LastRevisedBy, "Alice")
        .unwrap();
    writer
        .set_associated_string(
            AssociatedStringSlot::MailMergeDataSourcePath,
            "D:\\data\\customers.csv",
        )
        .unwrap();
    writer
        .set_associated_string(AssociatedStringSlot::WriteReservationPassword, "reserve")
        .unwrap();

    let saved_by = SavedByTable::try_new(vec![
        SavedByEntry::new("Alice", "C:\\draft.doc"),
        SavedByEntry::new("张三", "D:\\最终.doc"),
    ])
    .unwrap();
    writer.set_saved_by_table(saved_by.clone());
    (writer, saved_by)
}

fn assert_configured_document(
    package: &mut Package<Cursor<Vec<u8>>>,
    expected_saved_by: &SavedByTable,
) {
    let document = package.document().unwrap();
    let associated = document.associated_strings().unwrap();
    assert_eq!(associated.template_path(), "C:\\模板\\Normal.dot");
    assert_eq!(associated.title(), "Quarterly 😀");
    assert_eq!(associated.author(), "张三");
    assert_eq!(associated.last_revised_by(), "Alice");
    assert_eq!(
        associated.mail_merge_data_source_path(),
        "D:\\data\\customers.csv"
    );
    assert_eq!(associated.write_reservation_password(), "reserve");
    assert_eq!(document.saved_by_table().unwrap(), expected_saved_by);

    let (_, associated_length) = document
        .fib()
        .get_table_pointer(ASSOCIATED_STRINGS_FIB_INDEX)
        .unwrap();
    let (_, saved_by_length) = document
        .fib()
        .get_table_pointer(SAVED_BY_FIB_INDEX)
        .unwrap();
    assert_eq!(
        associated_length as usize,
        associated.to_bytes().unwrap().len()
    );
    assert_eq!(
        saved_by_length as usize,
        expected_saved_by.to_bytes().unwrap().len()
    );
}

#[test]
fn write_to_emits_the_mandatory_default_associated_string_table() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Body").unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();
    let associated = document.associated_strings().unwrap();
    assert!(associated.iter().all(|(_, value)| value.is_empty()));
    assert!(
        document
            .fib()
            .get_table_pointer(ASSOCIATED_STRINGS_FIB_INDEX)
            .is_some_and(|(_, length)| length > 0)
    );
    assert_eq!(
        document.saved_by_table().unwrap().entries(),
        &[] as &[SavedByEntry]
    );
    assert!(
        document
            .fib()
            .get_table_pointer(SAVED_BY_FIB_INDEX)
            .is_none_or(|(_, length)| length == 0)
    );
}

#[test]
fn associated_strings_and_saved_by_round_trip_through_write_to() {
    let (mut writer, expected_saved_by) = configured_writer();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_configured_document(&mut package, &expected_saved_by);
}

#[test]
fn associated_strings_and_saved_by_round_trip_through_file_save() {
    let (mut writer, expected_saved_by) = configured_writer();
    let path = temporary_doc_path();
    writer.save(&path).unwrap();
    let bytes = std::fs::read(&path).unwrap();
    std::fs::remove_file(path).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    assert_configured_document(&mut package, &expected_saved_by);
}

#[test]
fn writer_mutations_are_atomic_and_clearable() {
    let mut writer = DocWriter::new();
    writer
        .set_associated_string(AssociatedStringSlot::Title, "kept")
        .unwrap();
    assert!(
        writer
            .set_associated_string(AssociatedStringSlot::Title, "x".repeat(256))
            .is_err()
    );
    assert_eq!(writer.associated_strings().title(), "kept");
    assert!(
        writer
            .set_associated_string(
                AssociatedStringSlot::WriteReservationPassword,
                "x".repeat(16),
            )
            .is_err()
    );
    assert_eq!(writer.associated_strings().write_reservation_password(), "");

    let saved_by =
        SavedByTable::try_new(vec![SavedByEntry::new("Alice", "C:\\draft.doc")]).unwrap();
    assert!(writer.set_saved_by_table(saved_by.clone()).is_none());
    assert_eq!(writer.saved_by_table(), Some(&saved_by));
    assert_eq!(writer.clear_saved_by_table(), Some(saved_by));
    assert!(writer.saved_by_table().is_none());

    writer.reset_associated_strings();
    assert!(
        writer
            .associated_strings()
            .iter()
            .all(|(_, value)| value.is_empty())
    );
}
