use litchi_rtf::{DocumentLineSpacingCompatibility, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

fn all_enabled() -> DocumentLineSpacingCompatibility {
    DocumentLineSpacingCompatibility {
        suppress_extra_spacing_for_raised_lowered_text: true,
        suppress_extra_spacing_at_top_of_page: true,
        suppress_space_before_after_hard_break: true,
        suppress_wordperfect_extra_line_spacing: true,
        suppress_extra_spacing_at_bottom_of_page: true,
    }
}

#[test]
fn parses_all_flags_as_passive_metadata_without_changing_body() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\noextrasprl\sprstsp\sprsspbf"#,
        r#"\sprslnsp\sprsbsp Body}"#,
    ))
    .unwrap();
    assert_eq!(*document.line_spacing_compatibility(), all_enabled());
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_is_false_and_serializes_nothing() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(document.line_spacing_compatibility().is_empty());
    let serialized = String::from_utf8(write(&document)).unwrap();
    for name in ["noextrasprl", "sprstsp", "sprsspbf", "sprslnsp", "sprsbsp"] {
        assert!(!serialized.contains(name));
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_passively() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_line_spacing_compatibility(all_enabled());
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\noextrasprl",
        "\\sprstsp",
        "\\sprsspbf",
        "\\sprslnsp",
        "\\sprsbsp",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.line_spacing_compatibility(),
        document.line_spacing_compatibility()
    );
    assert_eq!(reparsed.text(), "Body");

    document.clear_line_spacing_compatibility();
    assert!(document.line_spacing_compatibility().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_bundled_producer_flags_without_applying_compatibility_layout() {
    let fixture =
        include_bytes!("../../../3rdparty/libreoffice-core/sw/qa/extras/layout/data/A020-min.rtf");
    let document = RtfDocument::parse_bytes(fixture).unwrap();
    assert!(
        document
            .line_spacing_compatibility()
            .suppress_extra_spacing_for_raised_lowered_text
    );
    assert!(
        document
            .line_spacing_compatibility()
            .suppress_space_before_after_hard_break
    );
    assert!(!document.text().is_empty());
}

#[test]
fn accepts_each_independent_compatibility_request() {
    for name in ["noextrasprl", "sprstsp", "sprsspbf", "sprslnsp", "sprsbsp"] {
        let source = format!(r#"{{\rtf1\{name} Body}}"#);
        let document = RtfDocument::parse(&source).unwrap();
        assert!(!document.line_spacing_compatibility().is_empty());
        assert_eq!(document.text(), "Body");
    }
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in ["noextrasprl", "sprstsp", "sprsspbf", "sprslnsp", "sprsbsp"] {
        for suffix in ["0", "1", "-1", "2147483647", "99999999999"] {
            let source = format!(r#"{{\rtf1\{name}{suffix} Body}}"#);
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
        for source in [
            format!(r#"{{\rtf1\{name}\{name} Body}}"#),
            format!(r#"{{\rtf1{{\*\{name}}}Body}}"#),
            format!(r#"{{\rtf1{{\{name}}}Body}}"#),
            format!(r#"{{\rtf1 Body\{name}}}"#),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
