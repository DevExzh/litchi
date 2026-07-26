use litchi_cfb::{OleFile, OleWriter};
use litchi_ole::doc::writer::DocEncryptionProfile;
use litchi_ole::doc::{
    DocOpenOptions, DocWriter, GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle,
    Package,
};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

const STTBF_GLSY_FIB_INDEX: usize = 9;
const PLCF_GLSY_FIB_INDEX: usize = 10;
const STTB_GLSY_STYLE_FIB_INDEX: usize = 83;
const FIB_BYTES: usize = 1248;
const FIB_POINTERS_OFFSET: usize = 154;
const FIB_POINTER_BYTES: usize = 8;
const FIB_PN_NEXT_OFFSET: usize = 8;
const FIB_FLAGS_OFFSET: usize = 10;
const FIB_CB_MAC_OFFSET: usize = 64;

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

fn attached_template_bytes(mutate: impl FnOnce(&mut [u8], usize)) -> Vec<u8> {
    let mut source_writer = writer();
    let mut source = Cursor::new(Vec::new());
    source_writer.write_to(&mut source).unwrap();

    let mut source_ole = OleFile::open(Cursor::new(source.into_inner())).unwrap();
    let mut word_document = source_ole.open_stream(&["WordDocument"]).unwrap();
    let table_stream = source_ole.open_stream(&["1Table"]).unwrap();
    let data_stream = source_ole.open_stream(&["Data"]).unwrap();

    let secondary_offset = word_document.len().next_multiple_of(512);
    assert_eq!(secondary_offset % 512, 0);
    let secondary_fib = word_document[..FIB_BYTES].to_vec();
    word_document.resize(secondary_offset, 0);
    word_document.extend_from_slice(&secondary_fib);

    let page = u16::try_from(secondary_offset / 512).unwrap();
    word_document[FIB_PN_NEXT_OFFSET..FIB_PN_NEXT_OFFSET + 2].copy_from_slice(&page.to_le_bytes());
    let mut main_flags = u16::from_le_bytes(
        word_document[FIB_FLAGS_OFFSET..FIB_FLAGS_OFFSET + 2]
            .try_into()
            .unwrap(),
    );
    main_flags = (main_flags & !0x0002) | 0x0001;
    word_document[FIB_FLAGS_OFFSET..FIB_FLAGS_OFFSET + 2]
        .copy_from_slice(&main_flags.to_le_bytes());
    for index in [
        STTBF_GLSY_FIB_INDEX,
        PLCF_GLSY_FIB_INDEX,
        STTB_GLSY_STYLE_FIB_INDEX,
    ] {
        let length_offset = FIB_POINTERS_OFFSET + index * FIB_POINTER_BYTES + 4;
        word_document[length_offset..length_offset + 4].fill(0);
    }

    let logical_size = u32::try_from(word_document.len()).unwrap();
    word_document[FIB_CB_MAC_OFFSET..FIB_CB_MAC_OFFSET + 4]
        .copy_from_slice(&logical_size.to_le_bytes());
    let secondary_cb_mac = secondary_offset + FIB_CB_MAC_OFFSET;
    word_document[secondary_cb_mac..secondary_cb_mac + 4]
        .copy_from_slice(&logical_size.to_le_bytes());
    mutate(&mut word_document, secondary_offset);

    let mut ole = OleWriter::new();
    ole.create_stream(&["WordDocument"], &word_document)
        .unwrap();
    ole.create_stream(&["1Table"], &table_stream).unwrap();
    ole.create_stream(&["Data"], &data_stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    ole.write_to(&mut output).unwrap();
    output.into_inner()
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

#[test]
fn glossary_utf16_ranges_round_trip_non_bmp_text_and_reject_split_pairs() {
    let metadata = |end_cp| {
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new("emoji", GlossaryItemKind::NamedAutoText, None, 0, end_cp)
                    .unwrap(),
            ],
            vec![],
            end_cp,
            4,
            5,
        )
        .unwrap()
    };
    let mut writer = DocWriter::new();
    writer.add_paragraph("😀").unwrap();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();
    writer.set_glossary_metadata(metadata(3));
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        package.document().unwrap().glossary_item_text(0).unwrap(),
        Some("😀")
    );

    writer.set_glossary_metadata(metadata(1));
    let original = vec![0x5A; 8];
    let mut output = Cursor::new(original.clone());
    let error = writer.write_to(&mut output).unwrap_err();
    assert!(error.to_string().contains("surrogate pair"));
    assert_eq!(output.into_inner(), original);
}

#[test]
fn template_secondary_fib_exposes_passive_attached_glossary() {
    let bytes = attached_template_bytes(|_, _| {});
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.fib().is_template());
    assert!(!document.fib().is_glossary_document());
    assert!(document.glossary_metadata().unwrap().is_none());

    let attached = document.attached_glossary().unwrap().unwrap();
    assert!(attached.fib().is_glossary_document());
    assert_eq!(attached.metadata(), &metadata());
    assert_eq!(attached.item_text(0), Some("Greeting"));
    assert_eq!(attached.item_text(1), Some("World"));
    assert_eq!(attached.item_text(2), None);
}

#[test]
fn distinct_attached_glossary_round_trips_through_public_writer() {
    let mut template = DocWriter::new();
    template.add_paragraph("Template body only").unwrap();
    assert!(template.set_attached_glossary(writer()).unwrap().is_none());
    assert!(template.attached_glossary().is_some());

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();

    assert!(document.fib().is_template());
    assert!(!document.fib().is_glossary_document());
    assert_eq!(document.text().unwrap(), "Template body only\r");
    let attached = document.attached_glossary().unwrap().unwrap();
    assert_eq!(attached.metadata(), &metadata());
    assert_eq!(attached.item_text(0), Some("Greeting"));
    assert_eq!(attached.item_text(1), Some("World"));
    assert!(!attached.text().contains("Template body only"));

    let removed = template.clear_attached_glossary().unwrap();
    assert_eq!(removed.glossary_metadata(), Some(&metadata()));
    assert!(template.attached_glossary().is_none());
}

#[test]
fn invalid_attached_glossary_configuration_is_atomic() {
    let mut template = DocWriter::new();
    template.add_paragraph("Template").unwrap();
    let mut output = Cursor::new(vec![0xA5; 16]);

    let error = match template.set_attached_glossary(DocWriter::new()) {
        Ok(_) => panic!("glossary metadata must be required"),
        Err(error) => error,
    };
    assert!(error.to_string().contains("requires glossary metadata"));
    assert!(template.attached_glossary().is_none());

    template.set_glossary_metadata(metadata());
    template.set_attached_glossary(writer()).unwrap();
    let error = template.write_to(&mut output).unwrap_err();
    assert!(error.to_string().contains("both glossary-only"));
    assert_eq!(output.into_inner(), vec![0xA5; 16]);
}

#[test]
fn attached_glossary_round_trips_inside_encrypted_template() {
    let mut template = DocWriter::new();
    template.add_paragraph("Protected template").unwrap();
    template.set_attached_glossary(writer()).unwrap();
    template
        .set_password("secret", DocEncryptionProfile::OfficeBinaryRc4)
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package
        .document_with_options(DocOpenOptions {
            password: Some("secret"),
        })
        .unwrap();

    assert_eq!(document.text().unwrap(), "Protected template\r");
    let attached = document.attached_glossary().unwrap().unwrap();
    assert_eq!(attached.item_text(0), Some("Greeting"));
    assert_eq!(attached.item_text(1), Some("World"));
}

#[test]
fn malformed_attached_fib_topologies_are_deferred_and_rejected() {
    let cases = [
        attached_template_bytes(|word, _| {
            word[FIB_FLAGS_OFFSET..FIB_FLAGS_OFFSET + 2].copy_from_slice(&0x12F4u16.to_le_bytes());
        }),
        attached_template_bytes(|word, secondary| {
            let flags = secondary + FIB_FLAGS_OFFSET;
            let value = u16::from_le_bytes(word[flags..flags + 2].try_into().unwrap()) & !0x0002;
            word[flags..flags + 2].copy_from_slice(&value.to_le_bytes());
        }),
        attached_template_bytes(|word, secondary| {
            let chpx_length = secondary + FIB_POINTERS_OFFSET + 12 * FIB_POINTER_BYTES + 4;
            let value =
                u32::from_le_bytes(word[chpx_length..chpx_length + 4].try_into().unwrap()) + 1;
            word[chpx_length..chpx_length + 4].copy_from_slice(&value.to_le_bytes());
        }),
        attached_template_bytes(|word, secondary| {
            let cb_mac = secondary + FIB_CB_MAC_OFFSET;
            word[cb_mac..cb_mac + 4].copy_from_slice(&1u32.to_le_bytes());
        }),
        attached_template_bytes(|word, _| {
            word[FIB_PN_NEXT_OFFSET..FIB_PN_NEXT_OFFSET + 2].copy_from_slice(&1u16.to_le_bytes());
        }),
        attached_template_bytes(|word, _| {
            word[FIB_PN_NEXT_OFFSET..FIB_PN_NEXT_OFFSET + 2]
                .copy_from_slice(&u16::MAX.to_le_bytes());
        }),
    ];

    for bytes in cases {
        let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
        let document = package.document().unwrap();
        assert!(document.attached_glossary().is_err());
        assert!(!document.text().unwrap().is_empty());
    }
}
