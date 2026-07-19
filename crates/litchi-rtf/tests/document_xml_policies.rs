use litchi_rtf::{DocumentXmlPolicies, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_common_producer_policy_sequence() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\relyonvml0\validatexml1\showplaceholdtext0"#,
        r#"\ignoremixedcontent0\saveinvalidxml0\showxmlerrors1 Body}"#,
    ))
    .unwrap();
    assert_eq!(
        *document.xml_policies(),
        DocumentXmlPolicies {
            rely_on_vml: Some(false),
            validate_custom_xml: Some(true),
            show_placeholder_text: Some(false),
            ignore_mixed_content: Some(false),
            save_invalid_xml: Some(false),
            show_xml_errors: Some(true),
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_is_preserved_with_each_normative_effective_default() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let policies = document.xml_policies();
    assert!(policies.is_empty());
    assert!(!policies.effective_rely_on_vml());
    assert!(policies.effective_validate_custom_xml());
    assert!(!policies.effective_show_placeholder_text());
    assert!(!policies.effective_ignore_mixed_content());
    assert!(!policies.effective_save_invalid_xml());
    assert!(!policies.effective_show_xml_errors());
    let serialized = String::from_utf8(write(&document)).unwrap();
    for control in [
        "relyonvml",
        "validatexml",
        "showplaceholdtext",
        "ignoremixedcontent",
        "saveinvalidxml",
        "showxmlerrors",
    ] {
        assert!(!serialized.contains(control));
    }
}

#[test]
fn typed_api_round_trips_explicit_values_in_stable_order_and_clears() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_xml_policies(DocumentXmlPolicies {
        rely_on_vml: Some(true),
        validate_custom_xml: Some(false),
        show_placeholder_text: Some(true),
        ignore_mixed_content: Some(true),
        save_invalid_xml: Some(true),
        show_xml_errors: Some(false),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\relyonvml1",
        "\\validatexml0",
        "\\showplaceholdtext1",
        "\\ignoremixedcontent1",
        "\\saveinvalidxml1",
        "\\showxmlerrors0",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.xml_policies(), document.xml_policies());
    assert_eq!(reparsed.text(), "Body");

    document.clear_xml_policies();
    assert!(document.xml_policies().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_theme_languages_and_inert_transform_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\showxmlerrors1\themelang1033"#,
        r#"{\*\xform transforms\\safe.xsl}\validatexml0"#,
        r#"\usexform\saveinvalidxml1\relyonvml0 Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.xml_policies(), document.xml_policies());
    assert_eq!(reparsed.theme_languages(), document.theme_languages());
    assert_eq!(reparsed.xsl_transform(), document.xsl_transform());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_missing_non_boolean_overflow_and_duplicate_values() {
    for name in [
        "relyonvml",
        "validatexml",
        "showplaceholdtext",
        "ignoremixedcontent",
        "saveinvalidxml",
        "showxmlerrors",
    ] {
        for suffix in ["", "-1", "2", "32767", "99999999999"] {
            let source = format!(r#"{{\rtf1\{name}{suffix} Body}}"#);
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
        let source = format!(r#"{{\rtf1\{name}0\{name}1 Body}}"#);
        assert!(
            RtfDocument::parse(&source).is_err(),
            "accepted duplicate {source}"
        );
    }
}

#[test]
fn rejects_every_starred_grouped_and_late_policy() {
    for control in [
        r#"\relyonvml0"#,
        r#"\validatexml1"#,
        r#"\showplaceholdtext0"#,
        r#"\ignoremixedcontent0"#,
        r#"\saveinvalidxml0"#,
        r#"\showxmlerrors1"#,
    ] {
        for source in [
            format!(r#"{{\rtf1{{\*{control}}}Body}}"#),
            format!(r#"{{\rtf1{{{control}}}Body}}"#),
            format!(r#"{{\rtf1 Body{control}}}"#),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
