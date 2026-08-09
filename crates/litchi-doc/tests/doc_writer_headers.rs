#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::headers::HeaderFooterType;
use litchi_doc::writer::FootnoteEntry;
use litchi_doc::{
    CharacterFormatting, HeaderFooterParagraph, Package, ParagraphFormatting, Writer,
};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

static NEXT_FILE: AtomicUsize = AtomicUsize::new(0);

fn temporary_doc_path() -> std::path::PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-rich-header-{}-{}.doc",
        std::process::id(),
        NEXT_FILE.fetch_add(1, Ordering::Relaxed)
    ))
}

#[test]
fn formatted_multi_paragraph_header_round_trips_through_plcfhdd() {
    let mut writer = Writer::new();
    writer.add_paragraph("Body").unwrap();

    let bold = CharacterFormatting {
        bold: Some(true),
        font_name: Some("Courier New".to_string()),
        ..CharacterFormatting::default()
    };
    writer
        .set_odd_header_paragraphs(vec![
            HeaderFooterParagraph::plain("First"),
            HeaderFooterParagraph::from_runs(
                vec![
                    ("Bold".to_string(), bold),
                    (" tail".to_string(), CharacterFormatting::default()),
                ],
                ParagraphFormatting::default(),
            ),
        ])
        .unwrap();

    let path = temporary_doc_path();
    writer.save(&path).unwrap();
    let mut package = Package::open(&path).unwrap();
    let document = package.document().unwrap();
    let headers = document.headers().unwrap();
    let header = headers
        .iter()
        .find(|header| header.header_footer_type == HeaderFooterType::OddPageHeader)
        .unwrap();
    assert_eq!(header.text(), "First\rBold tail\r\r");
    assert_eq!(header.paragraphs().unwrap().len(), 3);
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rich_header_api_rejects_ambiguous_or_unbounded_structure() {
    let mut writer = Writer::new();
    assert!(writer.set_odd_header_paragraphs(Vec::new()).is_err());
    assert!(
        writer
            .set_odd_header_paragraphs(vec![HeaderFooterParagraph::plain("one\rtwo")])
            .is_err()
    );
    assert!(
        writer
            .set_odd_header_paragraphs(vec![HeaderFooterParagraph::plain("\u{0013}")])
            .is_err()
    );

    let special = CharacterFormatting {
        special: Some(true),
        ..CharacterFormatting::default()
    };
    assert!(
        writer
            .set_odd_header_paragraphs(vec![HeaderFooterParagraph::from_runs(
                vec![("\u{0013}".to_string(), special.clone())],
                ParagraphFormatting::default(),
            )])
            .is_err()
    );
    assert!(
        HeaderFooterParagraph::field("PA\u{0013}GE", "1", CharacterFormatting::default(),).is_err()
    );

    let excessive_nesting = "\u{0013}".repeat(129);
    assert!(
        writer
            .set_odd_header_paragraphs(vec![HeaderFooterParagraph::from_runs(
                vec![(excessive_nesting, special)],
                ParagraphFormatting::default(),
            )])
            .is_err()
    );
}

#[test]
fn inert_header_field_round_trips_with_structural_markers() {
    let mut writer = Writer::new();
    writer.add_paragraph("Body").unwrap();
    let field = HeaderFooterParagraph::field("PAGE", "1", CharacterFormatting::default()).unwrap();
    writer.set_odd_footer_paragraphs(vec![field]).unwrap();

    let path = temporary_doc_path();
    writer.save(&path).unwrap();
    let mut package = Package::open(&path).unwrap();
    let document = package.document().unwrap();
    let footer = document
        .footers()
        .unwrap()
        .into_iter()
        .find(|footer| footer.header_footer_type == HeaderFooterType::OddPageFooter)
        .unwrap();
    assert_eq!(footer.text(), "\u{0013}PAGE\u{0014}1\u{0015}\r\r");
    std::fs::remove_file(path).unwrap();
}

#[test]
fn writes_header_and_note_plcfs_through_fib_pointers() {
    fn table_field<'a>(
        fib: &FileInformationBlock,
        table_stream: &'a [u8],
        index: usize,
    ) -> &'a [u8] {
        let (offset, length) = fib
            .get_table_pointer(index)
            .unwrap_or_else(|| panic!("missing FIB table pointer {index}"));
        let start = usize::try_from(offset).unwrap();
        let end = start + usize::try_from(length).unwrap();
        &table_stream[start..end]
    }

    fn cps(plcf: &[u8]) -> Vec<u32> {
        plcf.chunks_exact(4)
            .map(|chunk| u32::from_le_bytes(chunk.try_into().unwrap()))
            .collect()
    }

    let mut writer = Writer::new();
    writer.add_paragraph("Main").unwrap();
    writer.add_paragraph("End").unwrap();
    writer.set_odd_header("Header");
    writer.set_odd_footer("Footer");
    // These positions are the ends of the two main-story paragraphs, where
    // the writer injects the automatic reference characters.
    writer.add_footnote(FootnoteEntry::new(4, "Ftn", 2));
    writer.add_endnote(FootnoteEntry::new(9, "Edn", 3));

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let document_bytes = output.into_inner();
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(&document_bytes)).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let table_name = if FileInformationBlock::parse(&word_document)
        .unwrap()
        .which_table_stream()
    {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();

    let main_length = fib.get_main_doc_range().1;
    assert_eq!(main_length, 11);
    assert_eq!(fib.get_footnote_range(), Some((11, 17)));
    assert_eq!(fib.get_header_range(), Some((17, 34)));
    assert_eq!(fib.get_endnote_range(), Some((34, 40)));

    let header_plcf = table_field(&fib, &table_stream, 11);
    assert_eq!(header_plcf.len(), 14 * 4);
    let header_cps = cps(header_plcf);
    let (header_start, header_end) = fib.get_header_range().unwrap();
    assert_eq!(header_cps[7], 0);
    assert_eq!(header_cps[8], 8);
    assert_eq!(header_cps[9], 8);
    assert_eq!(header_cps[10], 16);
    assert_eq!(header_cps[12], 16);
    assert_eq!(header_cps[13], 17);
    assert_eq!(header_cps[12], header_end - header_start - 1);

    let footnote_ref = table_field(&fib, &table_stream, 2);
    assert_eq!(footnote_ref.len(), 10);
    assert_eq!(cps(&footnote_ref[..8]), vec![4, main_length]);
    assert_eq!(
        u16::from_le_bytes(footnote_ref[8..10].try_into().unwrap()),
        2
    );

    let footnote_txt = table_field(&fib, &table_stream, 3);
    assert_eq!(cps(footnote_txt), vec![0, 5, 6]);

    let endnote_ref = table_field(&fib, &table_stream, 46);
    assert_eq!(endnote_ref.len(), 10);
    assert_eq!(cps(&endnote_ref[..8]), vec![9, main_length]);
    assert_eq!(
        u16::from_le_bytes(endnote_ref[8..10].try_into().unwrap()),
        3
    );

    let endnote_txt = table_field(&fib, &table_stream, 47);
    assert_eq!(cps(endnote_txt), vec![0, 5, 6]);

    let mut package = Package::from_reader(Cursor::new(document_bytes)).unwrap();
    let document = package.document().unwrap();
    assert_eq!(document.footnotes().unwrap()[0].reference_position, 4);
    assert_eq!(document.endnotes().unwrap()[0].reference_position, 9);
    assert_eq!(document.headers().unwrap().len(), 1);
    assert_eq!(document.footers().unwrap().len(), 1);
}
