use litchi_rtf::{Formatting, RtfDocument, RtfWriter, StyleBlock};

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
fn parses_group_inheritance_restoration_plain_and_last_wins() {
    let document = RtfDocument::parse(concat!(
        r"{\rtf1\cb0 Zero {\cb2 Inner} Tail ",
        r"\cb1\cb3 Last{\plain Reset} {\*\unknown\cb4 Hidden}Visible}"
    ))
    .unwrap();

    assert_eq!(block(&document, "Zero").formatting.background_color, Some(0));
    assert_eq!(block(&document, "Inner").formatting.background_color, Some(2));
    assert_eq!(block(&document, "Tail").formatting.background_color, Some(0));
    assert_eq!(block(&document, "Last").formatting.background_color, Some(3));
    assert_eq!(block(&document, "Reset").formatting.background_color, None);
    assert_eq!(block(&document, "Visible").formatting.background_color, Some(3));
}

#[test]
fn character_style_and_body_round_trip_in_canonical_color_order() {
    let source = concat!(
        r"{\rtf1{\stylesheet{\*\cs9\additive\cb2 Back;}}",
        r"\highlight3\cb2\cf1 Body}"
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(
        document
            .stylesheet()
            .get(9)
            .unwrap()
            .formatting
            .background_color,
        Some(2)
    );

    let once = write(&document);
    assert!(once.contains(r"\cf1\cb2\highlight3"));
    let reparsed = RtfDocument::parse(&once).unwrap();
    assert_eq!(block(&reparsed, "Body").formatting.background_color, Some(2));
    assert_eq!(write(&reparsed), once);
}

#[test]
fn mutation_preserves_explicit_zero_and_omission() {
    let mut formatting = Formatting::default();
    assert_eq!(formatting.background_color, None);
    formatting.set_background_color(Some(0));
    assert_eq!(formatting.background_color, Some(0));
    formatting.clear_background_color();
    assert_eq!(formatting.background_color, None);
}

#[test]
fn rejects_missing_negative_overflow_and_unrepresentable_shape_runs() {
    for source in [
        r"{\rtf1\cb Missing}",
        r"{\rtf1\cb-1 Negative}",
        r"{\rtf1\cb65536 Overflow}",
        concat!(
            r"{\rtf1{\shp{\*\shpinst{\shptxt\cb1 one",
            r"\cb2 two}}}}"
        ),
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }

    let inert = RtfDocument::parse(concat!(
        r"{\rtf1{\*\unknown\cb65536 hidden}",
        r"{\field{\*\fldinst TEST \cb65536}{\fldrslt cached}}Visible}"
    ))
    .unwrap();
    assert_eq!(block(&inert, "Visible").formatting.background_color, None);
}

#[test]
fn parses_libreoffice_producer_fragment_and_round_trips_uniform_shape_text() {
    let fixture = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/fdo53556.rtf"
    ));
    let start = fixture.find(r"{\pard\cb1 ").unwrap();
    let producer = &fixture[start..];
    let end = producer.find('}').unwrap() + 1;
    let document = RtfDocument::parse(&format!(r"{{\rtf1{}}}", &producer[..end])).unwrap();
    assert_eq!(
        block(&document, "ARL STATISTICS")
            .formatting
            .background_color,
        Some(1)
    );

    let shape = RtfDocument::parse(r"{\rtf1{\shp{\*\shpinst{\shptxt\cb1 uniform}}}}")
        .unwrap();
    assert_eq!(
        shape.shapes()[0]
            .text_formatting
            .unwrap()
            .background_color,
        Some(1)
    );
    let serialized = write(&shape);
    assert!(serialized.contains(r"{\shptxt \cb1"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert!(reparsed.shapes().iter().any(|shape| {
        shape
            .text_formatting
            .is_some_and(|formatting| formatting.background_color == Some(1))
    }));
}
