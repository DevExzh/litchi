use litchi_rtf::{DocumentStyleRestrictions, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_four_flags_as_passive_metadata() {
    let document = RtfDocument::parse(
        r#"{\rtf1\stylelock\stylelockenforced\stylelockbackcomp\autofmtoverride Body}"#,
    )
    .unwrap();
    assert_eq!(
        *document.style_restrictions(),
        DocumentStyleRestrictions {
            restrictions_present: true,
            enforced: true,
            backward_compatibility: true,
            allow_auto_format_override: true,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn accepts_independent_producer_combinations_without_enforcement() {
    for (source, expected) in [
        (
            r#"{\rtf1\stylelock Body}"#,
            DocumentStyleRestrictions {
                restrictions_present: true,
                ..DocumentStyleRestrictions::default()
            },
        ),
        (
            r#"{\rtf1\stylelockenforced Body}"#,
            DocumentStyleRestrictions {
                enforced: true,
                ..DocumentStyleRestrictions::default()
            },
        ),
        (
            r#"{\rtf1\stylelockbackcomp Body}"#,
            DocumentStyleRestrictions {
                backward_compatibility: true,
                ..DocumentStyleRestrictions::default()
            },
        ),
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert_eq!(*document.style_restrictions(), expected);
        assert_eq!(document.text(), "Body");
    }
}

#[test]
fn omission_remains_empty_and_is_not_serialized() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(document.style_restrictions().is_empty());
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("stylelock"));
    assert!(!serialized.contains("autofmtoverride"));
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_passively() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_style_restrictions(DocumentStyleRestrictions {
        restrictions_present: true,
        enforced: true,
        backward_compatibility: true,
        allow_auto_format_override: true,
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\stylelock\\",
        "\\stylelockenforced",
        "\\stylelockbackcomp",
        "\\autofmtoverride",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_restrictions(), document.style_restrictions());
    assert_eq!(reparsed.text(), "Body");

    document.clear_style_restrictions();
    assert!(document.style_restrictions().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_modern_style_and_revision_policies() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\stylelock\stylelocktheme\trackmoves1"#,
        r#"\stylelockenforced\usenormstyforlist\stylelockbackcomp Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_restrictions(), document.style_restrictions());
    assert_eq!(reparsed.style_policies(), document.style_policies());
    assert_eq!(reparsed.revision_policies(), document.revision_policies());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn coexists_with_word_style_protection_sequence_without_enforcing_it() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\stylelock\stylelockenforced\stylelockbackcomp\autofmtoverride"#,
        r#"\readprot\enforceprot1\protlevel3 Body}"#,
    ))
    .unwrap();
    assert!(document.style_restrictions().allow_auto_format_override);
    assert_eq!(document.text(), "Body");

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_restrictions(), document.style_restrictions());
    assert_eq!(reparsed.protection(), document.protection());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in [
        "stylelock",
        "stylelockenforced",
        "stylelockbackcomp",
        "autofmtoverride",
    ] {
        for suffix in ["0", "1", "2147483647", "99999999999"] {
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
