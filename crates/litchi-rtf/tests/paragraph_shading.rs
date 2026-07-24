use litchi_rtf::{Paragraph, RtfDocument, RtfWriter, Shading, StyleBlock};

fn block<'a>(document: &'a RtfDocument<'a>, text: &str) -> &'a StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(text))
        .unwrap_or_else(|| panic!("missing block containing {text}"))
}

fn write(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn parses_inherits_restores_and_pard_resets_exact_values() {
    let document = RtfDocument::parse(concat!(
        r"{\rtf1\pard\cbpat3\shading1234\cfpat2 Outer\par ",
        r"{\shading0\cfpat0\cbpat0 Inner\par }Tail\par ",
        r"{\pard Reset\par }{\*\unknown\shading9999\cfpat8\cbpat9 Ignored}",
        r"Visible\par}"
    ))
    .unwrap();

    let outer = block(&document, "Outer").paragraph.shading;
    assert_eq!(outer.amount, Some(1234));
    assert_eq!(outer.foreground_color, Some(2));
    assert_eq!(outer.background_color, Some(3));

    let inner = block(&document, "Inner").paragraph.shading;
    assert_eq!(inner.amount, Some(0));
    assert_eq!(inner.foreground_color, Some(0));
    assert_eq!(inner.background_color, Some(0));
    assert_eq!(block(&document, "Tail").paragraph.shading, outer);
    assert_eq!(block(&document, "Reset").paragraph.shading, Shading::default());
    assert_eq!(block(&document, "Visible").paragraph.shading, outer);
}

#[test]
fn stylesheet_and_body_round_trip_in_canonical_order() {
    let source = concat!(
        r"{\rtf1{\stylesheet{\s9\cbpat3\shading6250\cfpat2 Shade;}}",
        r"\pard\s9\cbpat3\shading6250\cfpat2 Body\par}"
    );
    let document = RtfDocument::parse(source).unwrap();
    let style = document.stylesheet().get(9).unwrap().paragraph.unwrap();
    assert_eq!(style.shading.amount, Some(6250));
    assert_eq!(style.shading.foreground_color, Some(2));
    assert_eq!(style.shading.background_color, Some(3));

    let once = write(&document);
    assert!(once.contains(r"\shading6250\cfpat2\cbpat3"));
    let reparsed = RtfDocument::parse(&once).unwrap();
    assert_eq!(block(&reparsed, "Body").paragraph.shading, style.shading);
    assert_eq!(write(&reparsed), once);
}

#[test]
fn mutation_preserves_omission_and_explicit_zero() {
    let mut shading = Shading::default();
    assert!(!shading.is_present());
    shading.set_amount(Some(0)).unwrap();
    shading.set_foreground_color(Some(0));
    shading.set_background_color(Some(0));
    assert!(shading.is_present());
    assert!(!shading.is_visible());
    assert!(shading.set_amount(Some(10_001)).is_err());

    let mut paragraph = Paragraph::default();
    paragraph.set_shading(shading).unwrap();
    assert_eq!(paragraph.shading, shading);
    paragraph.clear_shading();
    assert_eq!(paragraph.shading, Shading::default());
}

#[test]
fn rejects_missing_negative_and_out_of_range_values_but_keeps_destinations_inert() {
    for source in [
        r"{\rtf1\shading X}",
        r"{\rtf1\cfpat X}",
        r"{\rtf1\cbpat X}",
        r"{\rtf1\shading-1 X}",
        r"{\rtf1\shading10001 X}",
        r"{\rtf1\cfpat-1 X}",
        r"{\rtf1\cfpat65536 X}",
        r"{\rtf1\cbpat-1 X}",
        r"{\rtf1\cbpat65536 X}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }

    let inert = RtfDocument::parse(
        r"{\rtf1{\*\unknown\shading10001\cfpat65536\cbpat-1 hidden}Visible}",
    )
    .unwrap();
    assert_eq!(block(&inert, "Visible").paragraph.shading, Shading::default());
}

#[test]
fn parses_clean_libreoffice_explicit_zero_body_and_style_producers() {
    let explicit_zero = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo47764.rtf"
    )))
    .unwrap();
    assert!(
        explicit_zero
            .blocks()
            .iter()
            .any(|block| block.paragraph.shading.background_color == Some(0))
    );

    let style_and_body = RtfDocument::parse(include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf108955.rtf"
    )))
    .unwrap();
    assert_eq!(
        style_and_body
            .stylesheet()
            .get(2)
            .unwrap()
            .paragraph
            .unwrap()
            .shading
            .background_color,
        Some(17)
    );
    assert!(
        style_and_body
            .blocks()
            .iter()
            .any(|block| block.paragraph.shading.background_color == Some(17))
    );
}
