use litchi_doc::parts::headers::HeaderFooterType;
use litchi_doc::{
    CharacterFormatting, DocWriter, HeaderFooterParagraph, Package, ParagraphFormatting,
};
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
    let mut writer = DocWriter::new();
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
    let mut writer = DocWriter::new();
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
    let mut writer = DocWriter::new();
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
