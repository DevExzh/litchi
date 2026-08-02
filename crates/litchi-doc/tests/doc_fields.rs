use litchi_doc::writer::DocWriter;
use litchi_doc::{FieldStory, FieldType, IndexEntryOption, Package, TableOfAuthoritiesEntryOption};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name)
}

fn reference_plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
    assert_eq!(cps.len(), descriptors.len() + 1);
    let mut bytes = Vec::with_capacity(cps.len() * 4 + descriptors.len() * 2);
    for cp in cps {
        bytes.extend_from_slice(&cp.to_le_bytes());
    }
    for descriptor in descriptors {
        bytes.extend_from_slice(descriptor);
    }
    bytes
}

type StoryCase = (
    FieldStory,
    &'static [u32],
    &'static [[u8; 2]],
    &'static [FieldType],
);

#[test]
fn apache_poi_reference_plcfs_cover_all_seven_story_tables() {
    // Ported from Apache POI `TestFieldsTables.EXPECTED`, whose source fixture
    // has a malformed CFB root name that strict litchi-cfb intentionally rejects.
    let cases: &[StoryCase] = &[
        (
            FieldStory::Comment,
            &[19, 43, 54, 59],
            &[[0x13, 0x1F], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::Date],
        ),
        (
            FieldStory::Endnote,
            &[31, 59, 61, 66],
            &[[0x13, 0x45], [0x14, 0xFF], [0x15, 0x80]],
            &[FieldType::FileSize],
        ),
        (
            FieldStory::Footnote,
            &[23, 49, 64, 69],
            &[[0x13, 0x11], [0x14, 0xFF], [0x15, 0x80]],
            &[FieldType::Author],
        ),
        (
            FieldStory::Header,
            &[18, 42, 44, 47, 75, 85, 91],
            &[
                [0x13, 0x21],
                [0x14, 0xFF],
                [0x15, 0x81],
                [0x13, 0x1D],
                [0x14, 0xFF],
                [0x15, 0x81],
            ],
            &[FieldType::Page, FieldType::FileName],
        ),
        (
            FieldStory::HeaderTextbox,
            &[30, 54, 62, 68],
            &[[0x13, 0x20], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::Time],
        ),
        (
            FieldStory::Main,
            &[1, 31, 51, 541],
            &[[0x13, 0x15], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::CreateDate],
        ),
        (
            FieldStory::Textbox,
            &[19, 47, 49, 55],
            &[[0x13, 0x19], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::EditTime],
        ),
    ];

    for &(story, cps, descriptors, expected_types) in cases {
        let reference = reference_plcf(cps, descriptors);
        let table =
            litchi_doc::FieldStoryTable::parse_plcf(story, *cps.last().unwrap(), &reference)
                .unwrap();
        assert_eq!(table.terminal_cp(), *cps.last().unwrap());
        assert_eq!(
            table
                .fields()
                .iter()
                .map(|field| field.field_type)
                .collect::<Vec<_>>(),
            expected_types
        );
        assert!(
            table
                .fields()
                .iter()
                .all(|field| field.end_flags.has_separator)
        );
        assert_eq!(table.to_plcf_bytes().unwrap(), reference);
    }
}

#[test]
fn poi_hyperlink_fixture_retains_inert_hyperlink_field() {
    let mut package = Package::open(fixture("hyperlink.doc")).unwrap();
    let document = package.document().unwrap();
    let fields = document.fields_table().expect("fields table");
    assert!(
        fields
            .main_document_fields()
            .iter()
            .any(|field| field.field_type == FieldType::Hyperlink)
    );
}

#[test]
fn generated_tc_document_discovers_table_of_contents_entries() {
    let mut writer = DocWriter::new();
    writer
        .add_paragraph(concat!(
            "\u{0013} TC \"Illustration 1\" \\f i \\l 4 ",
            "\\n \u{0014}cached entry\u{0015}",
        ))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let document = package.document().unwrap();
    let entries = document.table_of_contents_entries().unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].story(), FieldStory::Main);
    assert_eq!(entries[0].entry(), "Illustration 1");
    assert_eq!(entries[0].cached_result(), Some("cached entry"));
    assert_eq!(
        document.table_of_contents_entry_count().unwrap(),
        entries.len()
    );
}

#[test]
fn generated_ta_document_discovers_table_of_authorities_entries() {
    let mut writer = DocWriter::new();
    writer
        .add_paragraph(concat!(
            "\u{0013} TA \\l \"Baldwin v. Alberti\" \\c 1 \\s Baldwin ",
            "\\b \\i \\r PageRange \u{0014}cached authority\u{0015}",
        ))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let document = package.document().unwrap();
    let entries = document.table_of_authorities_entries().unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].story(), FieldStory::Main);
    assert_eq!(entries[0].cached_result(), Some("cached authority"));
    assert_eq!(
        entries[0].options(),
        &[
            TableOfAuthoritiesEntryOption::LongCitation("Baldwin v. Alberti".to_string()),
            TableOfAuthoritiesEntryOption::Category("1".to_string()),
            TableOfAuthoritiesEntryOption::ShortCitation("Baldwin".to_string()),
            TableOfAuthoritiesEntryOption::BoldPageNumber,
            TableOfAuthoritiesEntryOption::ItalicPageNumber,
            TableOfAuthoritiesEntryOption::PageRangeBookmark("PageRange".to_string()),
        ]
    );
    assert_eq!(
        document.table_of_authorities_entry_count().unwrap(),
        entries.len()
    );
}

#[test]
fn generated_xe_document_discovers_index_entries() {
    let mut writer = DocWriter::new();
    writer
        .add_paragraph(concat!(
            "\u{0013} XE \"Office Open XML:Syntax\" \\b \\f Intro \\i ",
            "\\r PageRange \\t \"See syntax\" \\y Office \u{0014}cached entry\u{0015}",
        ))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let document = package.document().unwrap();
    let entries = document.index_entries().unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].story(), FieldStory::Main);
    assert_eq!(entries[0].entry(), "Office Open XML:Syntax");
    assert_eq!(entries[0].cached_result(), Some("cached entry"));
    assert_eq!(
        entries[0].options(),
        &[
            IndexEntryOption::BoldPageNumber,
            IndexEntryOption::EntryType("Intro".to_string()),
            IndexEntryOption::ItalicPageNumber,
            IndexEntryOption::PageRangeBookmark("PageRange".to_string()),
            IndexEntryOption::CrossReference("See syntax".to_string()),
            IndexEntryOption::Yomi("Office".to_string()),
        ]
    );
    assert_eq!(document.index_entry_count().unwrap(), entries.len());
}

#[test]
fn generated_rd_document_discovers_referenced_documents() {
    let mut writer = DocWriter::new();
    writer
        .add_paragraph(concat!(
            "\u{0013} RD \"chapters/Chapter 1.doc\" \\f \\* MERGEFORMAT ",
            "\u{0014}cached reference\u{0015}",
        ))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let document = package.document().unwrap();
    let references = document.referenced_documents().unwrap();
    assert_eq!(references.len(), 1);
    assert_eq!(references[0].story(), FieldStory::Main);
    assert_eq!(references[0].source(), "chapters/Chapter 1.doc");
    assert!(references[0].uses_relative_path());
    assert_eq!(references[0].cached_result(), Some("cached reference"));
    assert_eq!(references[0].switches()[0].name(), 'f');
    assert_eq!(references[0].switches()[1].name(), '*');
    assert_eq!(
        document.referenced_document_count().unwrap(),
        references.len()
    );
}

#[test]
fn generated_private_document_discovers_private_fields() {
    let mut writer = DocWriter::new();
    writer
        .add_paragraph(concat!(
            "\u{0013} PRIVATE \"converter payload\" \\* MERGEFORMAT ",
            "\u{0014}cached private payload\u{0015}",
        ))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let document = package.document().unwrap();
    let private_fields = document.private_fields().unwrap();
    assert_eq!(private_fields.len(), 1);
    assert_eq!(private_fields[0].story(), FieldStory::Main);
    assert_eq!(
        private_fields[0].opaque_instructions(),
        "\"converter payload\" \\* MERGEFORMAT"
    );
    assert_eq!(
        private_fields[0].cached_result(),
        Some("cached private payload")
    );
    assert_eq!(
        document.private_field_count().unwrap(),
        private_fields.len()
    );
}
