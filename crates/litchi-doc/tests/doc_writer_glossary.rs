use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::writer::{
    CharacterFormatting, DocPicture, EncryptionProfile, FloatingPosition, Kind as DrawingKind,
    ParagraphFormatting, Shape as DrawingShape,
};
use litchi_doc::{
    DocWriter, GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle, OpenOptions,
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

fn image_fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/images")
            .join(relative),
    )
    .unwrap()
}

fn glossary_with_drawings() -> DocWriter {
    let mut writer = DocWriter::new();
    writer.add_paragraph("Greeting").unwrap();
    writer
        .insert_picture(DocPicture::new(image_fixture("png/lena.png")).unwrap())
        .unwrap();
    writer
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 1440, 720)
                .unwrap()
                .with_fill(0x80, 0x40, 0x20),
            FloatingPosition::new(720, 360),
            "AutoText box",
        )
        .unwrap();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();
    writer.set_glossary_metadata(
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new("greeting", GlossaryItemKind::NamedAutoText, Some(0), 0, 9)
                    .unwrap(),
                GlossaryItem::try_new("logo", GlossaryItemKind::NamedAutoText, None, 9, 11)
                    .unwrap(),
                GlossaryItem::try_new("shape", GlossaryItemKind::NamedAutoText, None, 11, 13)
                    .unwrap(),
            ],
            vec![GlossaryStyle::try_new("Normal", 1).unwrap()],
            13,
            14,
            15,
        )
        .unwrap(),
    );
    writer
}

fn glossary_with_hyperlink() -> DocWriter {
    const DISPLAY_TEXT: &str = "OpenAI";
    const TARGET: &str = "https://example.com";

    let mut writer = DocWriter::new();
    writer
        .add_hyperlink(DISPLAY_TEXT, TARGET, ParagraphFormatting::default())
        .unwrap();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();

    let instruction = format!("HYPERLINK \"{TARGET}\"");
    let item_end = u32::try_from(
        1 + instruction.encode_utf16().count() + 1 + DISPLAY_TEXT.encode_utf16().count() + 1 + 1,
    )
    .unwrap();
    writer.set_glossary_metadata(
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new(
                    "website",
                    GlossaryItemKind::NamedAutoText,
                    None,
                    0,
                    item_end,
                )
                .unwrap(),
            ],
            vec![],
            item_end,
            item_end + 1,
            item_end + 2,
        )
        .unwrap(),
    );
    writer
}

fn glossary_with_non_plcf_fields() -> DocWriter {
    const INSTRUCTIONS: [&str; 5] = [
        r#"TC "Illustration 1" \f i \l 4 \n"#,
        r#"TA \l "Baldwin v. Alberti" \c 1 \s Baldwin"#,
        r#"XE "Office Open XML:Syntax" \b \f Intro"#,
        r#"RD "chapters/Chapter 1.doc" \f"#,
        r#"PRIVATE "converter payload""#,
    ];

    let mut writer = DocWriter::new();
    let marker_format = CharacterFormatting {
        special: Some(true),
        ..CharacterFormatting::default()
    };
    let instruction_format = CharacterFormatting {
        field_vanish: Some(true),
        ..CharacterFormatting::default()
    };
    let mut item_end = 0u32;
    for instruction in INSTRUCTIONS {
        writer
            .add_paragraph_runs(
                vec![
                    ("\u{0013}".to_string(), marker_format.clone()),
                    (instruction.to_string(), instruction_format.clone()),
                    ("\u{0015}".to_string(), marker_format.clone()),
                ],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let paragraph_units = u32::try_from(instruction.encode_utf16().count() + 3).unwrap();
        item_end = item_end.checked_add(paragraph_units).unwrap();
    }
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();
    writer.set_glossary_metadata(
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new(
                    "field-indexes",
                    GlossaryItemKind::NamedAutoText,
                    None,
                    0,
                    item_end,
                )
                .unwrap(),
            ],
            vec![],
            item_end,
            item_end + 1,
            item_end + 2,
        )
        .unwrap(),
    );
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
        .set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package
        .document_with_options(OpenOptions {
            password: Some("secret"),
            ..Default::default()
        })
        .unwrap();

    assert_eq!(document.text().unwrap(), "Protected template\r");
    let attached = document.attached_glossary().unwrap().unwrap();
    assert_eq!(attached.item_text(0), Some("Greeting"));
    assert_eq!(attached.item_text(1), Some("World"));
}

#[test]
fn attached_glossary_preserves_shared_data_and_drawing_graphs() {
    const DGG_INFO_INDEX: usize = 50;
    const PLCF_SPA_MOM_INDEX: usize = 40;
    const PIC_LOCATION_OPCODE: [u8; 2] = 0x6A03u16.to_le_bytes();

    let parent_image = image_fixture("jpg/abstract1.jpg");
    let child_image = image_fixture("png/lena.png");
    let mut template = DocWriter::new();
    template.add_paragraph("Template").unwrap();
    template
        .insert_picture(DocPicture::new(parent_image.clone()).unwrap())
        .unwrap();
    template
        .set_attached_glossary(glossary_with_drawings())
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let mut ole = OleFile::open(Cursor::new(&bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let data = ole.open_stream(&["Data"]).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.text().unwrap().starts_with("Template\r"));
    let attached = document.attached_glossary().unwrap().unwrap();
    assert_eq!(attached.item_text(0), Some("Greeting"));
    assert_eq!(attached.item_text(1), Some("\u{0001}"));
    assert_eq!(attached.item_text(2), Some("\u{0008}"));
    assert_eq!(attached.images().len(), 1);
    assert!(attached.images()[0].pic_offset() as usize >= 4096);
    assert_eq!(attached.shape_positions().len(), 1);
    assert!(
        attached
            .shapes()
            .iter()
            .any(|shape| shape.shape_id == attached.shape_positions()[0].spa.shape_id)
    );
    assert_eq!(attached.text_boxes().len(), 1);
    assert_eq!(attached.text_boxes()[0].text, "AutoText box\r");
    let paragraphs = attached.paragraphs().unwrap();
    assert!(
        paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.runs().unwrap())
            .any(|run| run.has_image())
    );
    assert_eq!(
        document
            .image_data(&attached.images()[0])
            .unwrap()
            .data()
            .unwrap(),
        child_image.as_slice()
    );
    assert!(
        attached
            .fib()
            .get_table_pointer(DGG_INFO_INDEX)
            .is_some_and(|(_, length)| length > 0)
    );
    assert!(
        attached
            .fib()
            .get_table_pointer(PLCF_SPA_MOM_INDEX)
            .is_some_and(|(_, length)| length > 0)
    );

    assert!(
        data.windows(parent_image.len())
            .any(|window| window == parent_image)
    );
    assert!(
        data.windows(child_image.len())
            .any(|window| window == child_image)
    );
    let picture_locations: Vec<u32> = word_document
        .windows(6)
        .filter(|window| window[..2] == PIC_LOCATION_OPCODE)
        .map(|window| u32::from_le_bytes(window[2..6].try_into().unwrap()))
        .filter(|offset| (*offset as usize) < data.len())
        .collect();
    assert!(picture_locations.contains(&0));
    assert!(
        picture_locations
            .iter()
            .any(|offset| *offset as usize >= 4096)
    );
}

#[test]
fn attached_glossary_exposes_typed_inert_fields() {
    let mut template = DocWriter::new();
    template.add_paragraph("Template").unwrap();
    template
        .set_attached_glossary(glossary_with_hyperlink())
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();
    let attached = document.attached_glossary().unwrap().unwrap();

    let main_fields = attached.fields_table().main_document_fields();
    assert_eq!(main_fields.len(), 1);
    let field = attached.field_text(&main_fields[0]).unwrap();
    assert_eq!(field.instruction, r#"HYPERLINK "https://example.com""#);
    assert_eq!(field.result.as_deref(), Some("OpenAI"));
    assert_eq!(attached.fields().unwrap(), vec![field]);

    let links = attached.hyperlink_fields().unwrap();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].external_target(), Some("https://example.com"));
    assert_eq!(links[0].cached_result(), Some("OpenAI"));
}

#[test]
fn attached_glossary_reconstructs_all_fields_excluded_from_plcfld() {
    let mut template = DocWriter::new();
    template.add_paragraph("Template").unwrap();
    template
        .set_attached_glossary(glossary_with_non_plcf_fields())
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    template.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();
    let attached = document.attached_glossary().unwrap().unwrap();

    assert!(attached.fields_table().main_document_fields().is_empty());
    let fields = attached.non_plcf_fields();
    assert_eq!(fields.len(), 5);
    assert!(!fields.is_empty());
    assert_eq!(
        fields.table_of_contents_entries()[0].entry(),
        "Illustration 1"
    );
    assert_eq!(fields.table_of_authorities_entries().len(), 1);
    assert_eq!(fields.index_entries()[0].entry(), "Office Open XML:Syntax");
    assert_eq!(
        fields.referenced_documents()[0].source(),
        "chapters/Chapter 1.doc"
    );
    assert!(fields.referenced_documents()[0].uses_relative_path());
    assert_eq!(
        fields.private_fields()[0].opaque_instructions(),
        r#""converter payload""#
    );
}

#[test]
fn malformed_attached_fib_topologies_are_deferred_and_rejected() {
    const DGG_INFO_INDEX: usize = 50;
    const PLCF_FLD_MOM_INDEX: usize = 16;
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
        attached_template_bytes(|word, secondary| {
            let pointer = secondary + FIB_POINTERS_OFFSET + DGG_INFO_INDEX * FIB_POINTER_BYTES;
            word[pointer..pointer + 4].copy_from_slice(&u32::MAX.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&1u32.to_le_bytes());
        }),
        attached_template_bytes(|word, secondary| {
            let pointer = secondary + FIB_POINTERS_OFFSET + PLCF_FLD_MOM_INDEX * FIB_POINTER_BYTES;
            word[pointer..pointer + 4].copy_from_slice(&u32::MAX.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&4u32.to_le_bytes());
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
