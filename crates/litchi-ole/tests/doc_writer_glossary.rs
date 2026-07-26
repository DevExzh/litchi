use litchi_ole::doc::{
    DocWriter, GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle, Package,
};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

const STTBF_GLSY_FIB_INDEX: usize = 9;
const PLCF_GLSY_FIB_INDEX: usize = 10;
const STTB_GLSY_STYLE_FIB_INDEX: usize = 83;

static NEXT_FILE: AtomicUsize = AtomicUsize::new(0);

fn temporary_doc_path() -> std::path::PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-glossary-{}-{}.doc",
        std::process::id(),
        NEXT_FILE.fetch_add(1, Ordering::Relaxed)
    ))
}

fn metadata() -> GlossaryMetadata {
    GlossaryMetadata::try_new(
        vec![
            GlossaryItem::try_new("greeting", GlossaryItemKind::NamedAutoText, Some(0), 0, 9)
                .unwrap(),
            GlossaryItem::try_new("teh", GlossaryItemKind::FormattedAutoCorrect, None, 9, 15)
                .unwrap(),
        ],
        vec![GlossaryStyle::try_new("Normal", 1).unwrap()],
        15,
        16,
        17,
    )
    .unwrap()
}

fn writer() -> DocWriter {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Greeting").unwrap();
    writer.add_paragraph("World").unwrap();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();
    writer.set_glossary_metadata(metadata());
    writer
}

fn assert_glossary(package: &mut Package<Cursor<Vec<u8>>>) {
    let document = package.document().unwrap();
    assert!(document.fib().is_glossary_document());
    let glossary = document.glossary_metadata().unwrap().unwrap();
    assert_eq!(glossary, &metadata());
    assert_eq!(document.glossary_item_text(0).unwrap(), Some("Greeting"));
    assert_eq!(document.glossary_item_text(1).unwrap(), Some("World"));
    assert_eq!(document.glossary_item_text(2).unwrap(), None);
    for index in [
        STTBF_GLSY_FIB_INDEX,
        PLCF_GLSY_FIB_INDEX,
        STTB_GLSY_STYLE_FIB_INDEX,
    ] {
        assert!(
            document
                .fib()
                .get_table_pointer(index)
                .is_some_and(|(_, length)| length > 0)
        );
    }
}

#[test]
fn glossary_only_document_round_trips_through_write_to() {
    let mut writer = writer();
    assert_eq!(writer.glossary_metadata(), Some(&metadata()));

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_glossary(&mut package);
}

#[test]
fn glossary_only_document_round_trips_through_file_save_and_can_be_cleared() {
    let mut writer = writer();
    let removed = writer.clear_glossary_metadata().unwrap();
    assert_eq!(removed, metadata());
    assert!(writer.glossary_metadata().is_none());
    writer.set_glossary_metadata(removed);

    let path = temporary_doc_path();
    writer.save(&path).unwrap();
    let bytes = std::fs::read(&path).unwrap();
    std::fs::remove_file(path).unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    assert_glossary(&mut package);
}

#[test]
fn mismatched_main_story_length_is_rejected_before_output_changes() {
    let mut writer = DocWriter::new();
    writer.add_paragraph("short").unwrap();
    writer.set_glossary_metadata(metadata());

    let original = vec![0xA5; 16];
    let mut output = Cursor::new(original.clone());
    let error = writer.write_to(&mut output).unwrap_err();
    assert!(error.to_string().contains("glossary ccpText"));
    assert_eq!(output.into_inner(), original);
}
