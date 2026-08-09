#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

use litchi_ppt::odraw::ShapeExt as _;
use litchi_ppt::writer::{FontEntity, NotesPage, Paragraph, SmartTagDefinition, TextRun, Writer};
use litchi_ppt::{
    EscherTextboxWrapper, Font, FontEmbeddingFlags, FontFacet, FontScope, Package,
    ProgBinaryTagVersion, ProgTag, ProgTagLimits, ProgTagScope, ProgTags, Record, RecordType,
};

fn configured_writer() -> Writer {
    let mut writer = Writer::new();
    let base = writer
        .add_font_model(FontScope::Base, Font::new("Writer Embedded"))
        .unwrap();
    writer
        .set_embedded_font(FontScope::Base, base, FontFacet::Plain, minimal_eot())
        .unwrap();
    writer
        .add_font_model(FontScope::International, Font::new("Writer Intl"))
        .unwrap();
    writer
        .set_font_embedding_flags(Some(FontEmbeddingFlags::new(true, true)))
        .unwrap();
    writer
}

fn minimal_eot() -> Vec<u8> {
    let mut eot = vec![0; 96];
    eot[0..4].copy_from_slice(&96u32.to_le_bytes());
    eot[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
    eot[34..36].copy_from_slice(&0x504cu16.to_le_bytes());
    eot
}

fn write_to_bytes(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn assert_custom_catalog(bytes: Vec<u8>) {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    assert!(
        presentation
            .document_atom()
            .unwrap()
            .unwrap()
            .save_with_fonts
    );
    let fonts = presentation.fonts().unwrap();
    assert_eq!(fonts.base.as_ref().unwrap().fonts.len(), 2);
    assert_eq!(fonts.get_base(1).unwrap().name, "Writer Embedded");
    assert_eq!(fonts.get_base(1).unwrap().embedded_fonts.len(), 1);
    assert_eq!(
        fonts.get_base(1).unwrap().embedded_fonts[0].bytes(),
        minimal_eot().as_slice()
    );
    assert_eq!(fonts.get_international(0).unwrap().name, "Writer Intl");
    assert_eq!(
        fonts.embedding_flags,
        Some(FontEmbeddingFlags::new(true, true))
    );
}

fn temporary_path(label: &str) -> PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-ppt-font-{label}-{}-{}.ppt",
        std::process::id(),
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ))
}

#[test]
fn default_writer_retains_arial_without_embedding_flags() {
    let mut writer = Writer::new();
    let mut package = Package::from_reader(Cursor::new(write_to_bytes(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    assert!(
        !presentation
            .document_atom()
            .unwrap()
            .unwrap()
            .save_with_fonts
    );
    let fonts = presentation.fonts().unwrap();
    assert_eq!(fonts.base.as_ref().unwrap().fonts.len(), 1);
    assert_eq!(fonts.get_base(0).unwrap().name, "Arial");
    assert!(fonts.international.is_none());
    assert!(fonts.embedding_flags.is_none());
}

#[test]
fn custom_catalog_round_trips_through_write_to() {
    assert_custom_catalog(write_to_bytes(&mut configured_writer()));
}

#[test]
fn custom_catalog_round_trips_through_save() {
    let path = temporary_path("save");
    configured_writer().save(&path).unwrap();
    let bytes = std::fs::read(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    assert_custom_catalog(bytes);
}

#[test]
fn checked_font_add_enforces_name_and_collection_limits_atomically() {
    let exact = "12345678901234567890123456789012";
    let mut writer = Writer::new();
    assert_eq!(writer.add_font_checked(FontEntity::new(exact)).unwrap(), 1);
    assert_eq!(FontEntity::new(exact).try_build().unwrap().len(), 68);

    let before = writer.font_count();
    assert!(
        writer
            .add_font_checked(FontEntity::new(format!("{exact}x")))
            .is_err()
    );
    assert_eq!(writer.font_count(), before);

    for index in writer.font_count()..129 {
        writer
            .add_font_checked(FontEntity::new(format!("Font {index}")))
            .unwrap();
    }
    assert_eq!(writer.font_count(), 129);
    assert!(
        writer
            .add_font_checked(FontEntity::new("Overflow"))
            .is_err()
    );
    assert_eq!(writer.font_count(), 129);
}

#[test]
fn every_text_font_reference_fails_before_either_destination_is_modified() {
    let cases = [
        (
            "primary",
            Paragraph::with_runs(vec![TextRun::new("bad").font(9)]),
        ),
        (
            "east-asian",
            Paragraph::with_runs(vec![TextRun::new("bad").asian_font(9)]),
        ),
        (
            "ansi",
            Paragraph::with_runs(vec![TextRun::new("bad").ansi_font(9)]),
        ),
        (
            "symbol",
            Paragraph::with_runs(vec![TextRun::new("bad").symbol_font(9)]),
        ),
        (
            "international-east-asian",
            Paragraph::with_runs(vec![TextRun::new("bad").international_east_asian_font(9)]),
        ),
        (
            "complex-script",
            Paragraph::with_runs(vec![TextRun::new("bad").complex_script_font(9)]),
        ),
        ("bullet", Paragraph::new("bad").bullet_font(9)),
    ];

    for (label, paragraph) in cases {
        let mut writer = Writer::new();
        let slide = writer.add_slide().unwrap();
        writer
            .add_rich_textbox(slide, 10, 10, 100, 40, vec![paragraph])
            .unwrap();

        let mut output = Cursor::new(Vec::new());
        assert!(writer.write_to(&mut output).is_err(), "{label}");
        assert!(output.get_ref().is_empty(), "{label}");

        let path = temporary_path(label);
        assert!(writer.save(&path).is_err(), "{label}");
        assert!(!path.exists(), "{label}");
    }
}

#[test]
fn out_of_range_international_font_reference_is_atomic() {
    let mut writer = Writer::new();
    writer
        .add_font_model(
            FontScope::International,
            FontEntity::new("International").into(),
        )
        .unwrap();
    let slide = writer.add_slide().unwrap();
    writer
        .add_rich_textbox(
            slide,
            10,
            10,
            100,
            40,
            vec![Paragraph::with_runs(vec![
                TextRun::new("bad")
                    .international_east_asian_font(1)
                    .complex_script_font(1),
            ])],
        )
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut output).is_err());
    assert!(output.get_ref().is_empty());

    let path = temporary_path("international-out-of-range");
    assert!(writer.save(&path).is_err());
    assert!(!path.exists());
}

#[test]
fn font_ppt10_and_smart_tag_ppt11_share_one_document_prog_tags_owner() {
    let mut writer = configured_writer();
    let tag = writer
        .add_smart_tag(
            SmartTagDefinition::new("urn:litchi:test", "fruit").with_property("name", "lychee"),
        )
        .unwrap();
    let slide = writer.add_slide().unwrap();
    writer
        .add_rich_textbox(
            slide,
            10,
            10,
            160,
            40,
            vec![Paragraph::with_runs(vec![
                TextRun::new("lychee").with_smart_tag(tag),
                TextRun::new(" international")
                    .international_east_asian_font(0)
                    .complex_script_font(0),
            ])],
        )
        .unwrap();
    let bytes = write_to_bytes(&mut writer);

    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&bytes)).unwrap();
    let stream = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let (document, _) = Record::parse(&stream, 0).unwrap();
    let doc_info = document
        .children
        .iter()
        .find(|record| record.record_type_raw == 2000)
        .unwrap();
    let owners = doc_info
        .children
        .iter()
        .filter(|record| record.record_type_raw == 5000)
        .collect::<Vec<_>>();
    assert_eq!(
        owners.len(),
        1,
        "DocInfo must contain one DocProgTags owner"
    );

    let raw_tags =
        ProgTags::parse(owners[0], ProgTagScope::Document, ProgTagLimits::default()).unwrap();
    assert_eq!(
        raw_tags
            .tags
            .iter()
            .filter(|entry| matches!(
                entry,
                ProgTag::Binary(binary)
                    if binary.version == ProgBinaryTagVersion::PowerPoint10
            ))
            .count(),
        1,
        "the shared owner must contain one ___PPT10 tag"
    );
    assert_eq!(
        raw_tags
            .tags
            .iter()
            .filter(|entry| matches!(
                entry,
                ProgTag::Binary(binary)
                    if binary.version == ProgBinaryTagVersion::PowerPoint11
            ))
            .count(),
        1,
        "the shared owner must retain the ___PPT11 smart-tag store"
    );

    let ppt10_tag = raw_tags
        .binary_tag(ProgBinaryTagVersion::PowerPoint10)
        .unwrap();
    let ppt10_records = ppt10_tag.records().unwrap();
    let extensions = raw_tags.document_extensions().unwrap();
    let ppt10 = extensions.powerpoint10.as_ref().unwrap();
    assert_eq!(
        ppt10_records
            .iter()
            .map(|record| record.record_type_raw)
            .collect::<Vec<_>>(),
        vec![
            ppt10.font_collection.as_ref().unwrap().record_type_raw,
            ppt10.grid_spacing.as_ref().unwrap().record_type_raw,
            ppt10.font_embed_flags.as_ref().unwrap().record_type_raw,
        ],
        "PP10 font records must precede optional grid spacing and embedding flags"
    );
    assert!(
        extensions
            .powerpoint11
            .as_ref()
            .unwrap()
            .smart_tag_store
            .is_some()
    );

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(
        presentation.fonts().unwrap().get_base(1).unwrap().name,
        "Writer Embedded"
    );
    let store = presentation.smart_tags().unwrap().unwrap();
    assert_eq!(store.tags[0].properties[0].value, "lychee");
    let shape_tags = presentation.shape_programmable_tags().unwrap();
    assert_eq!(
        shape_tags[0].programmable_tags.powerpoint11().unwrap().runs[0].smart_tag_indices,
        vec![tag.as_u32()]
    );
    let powerpoint10 = shape_tags[0].programmable_tags.powerpoint10().unwrap();
    assert_eq!(powerpoint10.runs[1].new_east_asian_font_ref, Some(0));
    assert_eq!(powerpoint10.runs[1].complex_script_font_ref, Some(0));
}

#[test]
fn notes_font_references_round_trip_through_the_rich_text_path() {
    let path = temporary_path("notes-rich-fonts");
    let mut writer = configured_writer();
    let slide = writer.add_slide().unwrap();
    let paragraph = Paragraph::with_runs(vec![
        TextRun::new("notes fonts")
            .font(1)
            .asian_font(1)
            .ansi_font(1)
            .symbol_font(1)
            .international_east_asian_font(0)
            .complex_script_font(0),
    ])
    .bullet_font(1);
    writer
        .set_notes_page(slide, NotesPage::new(0).with_paragraphs(vec![paragraph]))
        .unwrap();
    writer.save(&path).unwrap();

    let bytes = std::fs::read(&path).unwrap();
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let mut offset = 0usize;
    let notes = loop {
        let (record, consumed) = Record::parse(&stream, offset).unwrap();
        offset += consumed;
        if record.record_type == RecordType::Notes {
            break record;
        }
    };
    let drawing = notes
        .children
        .iter()
        .find(|record| record.record_type == RecordType::PPDrawing)
        .unwrap();
    let shapes = litchi_ppt::odraw::parse(&drawing.data).unwrap();
    let mut pending = shapes.iter().collect::<Vec<_>>();
    let (textbox, tags) = loop {
        let shape = pending.pop().unwrap();
        pending.extend(shape.children());
        let Some(textbox) = shape.textbox() else {
            continue;
        };
        let wrapper = EscherTextboxWrapper::new(textbox.data().to_vec()).unwrap();
        if wrapper.text().is_empty() {
            continue;
        }
        break (wrapper, shape.programmable_tags().unwrap().unwrap());
    };

    assert_eq!(textbox.text(), "notes fonts");
    assert_eq!(textbox.runs().len(), 1);
    let formatting = &textbox.runs()[0].formatting;
    assert_eq!(formatting.font_index, Some(1));
    assert_eq!(formatting.asian_font_index, Some(1));
    assert_eq!(formatting.ansi_font_index, Some(1));
    assert_eq!(formatting.symbol_font_index, Some(1));
    assert_eq!(
        textbox.paragraph_runs()[0].formatting.bullet_font_index,
        Some(1)
    );
    let powerpoint10 = tags.powerpoint10().unwrap();
    assert_eq!(powerpoint10.runs[0].new_east_asian_font_ref, Some(0));
    assert_eq!(powerpoint10.runs[0].complex_script_font_ref, Some(0));

    std::fs::remove_file(path).unwrap();
}

#[test]
fn invalid_notes_font_references_leave_both_destinations_untouched() {
    let cases = [
        (
            "notes-primary",
            Paragraph::with_runs(vec![TextRun::new("bad").font(9)]),
        ),
        (
            "notes-east-asian",
            Paragraph::with_runs(vec![TextRun::new("bad").asian_font(9)]),
        ),
        (
            "notes-ansi",
            Paragraph::with_runs(vec![TextRun::new("bad").ansi_font(9)]),
        ),
        (
            "notes-symbol",
            Paragraph::with_runs(vec![TextRun::new("bad").symbol_font(9)]),
        ),
        (
            "notes-international-east-asian",
            Paragraph::with_runs(vec![TextRun::new("bad").international_east_asian_font(9)]),
        ),
        (
            "notes-complex-script",
            Paragraph::with_runs(vec![TextRun::new("bad").complex_script_font(9)]),
        ),
        ("notes-bullet", Paragraph::new("bad").bullet_font(9)),
    ];

    for (label, paragraph) in cases {
        let mut writer = Writer::new();
        let slide = writer.add_slide().unwrap();
        writer
            .set_notes_page(slide, NotesPage::new(0).with_paragraphs(vec![paragraph]))
            .unwrap();

        let mut output = Cursor::new(Vec::new());
        assert!(writer.write_to(&mut output).is_err(), "{label}");
        assert!(output.get_ref().is_empty(), "{label}");

        let path = temporary_path(label);
        assert!(writer.save(&path).is_err(), "{label}");
        assert!(!path.exists(), "{label}");
    }
}
