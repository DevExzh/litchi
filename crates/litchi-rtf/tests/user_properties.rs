#![allow(
    clippy::expect_used,
    clippy::float_cmp,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design, rebind fixture names across steps, and compare exact fixture fractions"
)]

use litchi_rtf::{RtfDocument, RtfWriter, UserPropertyType, UserPropertyValue};

// The fixture value 3.1400 exercises lexical round-tripping; it is not the constant PI.
#[allow(
    clippy::approx_constant,
    reason = "the fixture value 3.1400 exercises lexical round-tripping and is not the constant PI"
)]
#[test]
fn parses_all_typed_values_in_normative_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\userprops"#,
        r#"{\propname Text}\proptype30{\staticval hello}{\linkval https://invalid.example/no-fetch}"#,
        r#"{\propname Integer}\proptype3{\staticval -42}"#,
        r#"{\propname Real}\proptype5{\staticval 3.1400}"#,
        r#"{\propname Boolean}\proptype11{\staticval 1}"#,
        r#"{\propname Date}\proptype64{\staticval 2016. 01. 30.}"#,
        r#"{\propname Future}\proptype999{\staticval opaque}"#,
        r#"{\propname Caf\'e9}\proptype30{\staticval na\'efve}"#,
        r#"}Body}"#,
    ))
    .unwrap();
    let properties = document.user_properties();
    assert_eq!(properties.len(), 7);
    assert_eq!(properties[0].name, "Text");
    assert_eq!(
        properties[0].link_value.as_deref(),
        Some("https://invalid.example/no-fetch")
    );
    assert!(matches!(
        &properties[1].value,
        UserPropertyValue::Integer { value: -42, lexical } if lexical == "-42"
    ));
    assert!(matches!(
        &properties[2].value,
        UserPropertyValue::Real { value, lexical } if *value == 3.14 && lexical == "3.1400"
    ));
    assert!(matches!(
        &properties[3].value,
        UserPropertyValue::Boolean { value: true, lexical } if lexical == "1"
    ));
    assert!(matches!(
        &properties[4].value,
        UserPropertyValue::Date { value, lexical }
            if value.year == 2016 && value.month == 1 && value.day == 30
                && value.hour.is_none() && lexical == "2016. 01. 30."
    ));
    assert_eq!(
        properties[4].value.property_type(),
        UserPropertyType::DateTime
    );
    assert!(matches!(
        &properties[5].value,
        UserPropertyValue::Unknown { type_code: 999, lexical } if lexical == "opaque"
    ));
    assert_eq!(
        properties[5].value.property_type(),
        UserPropertyType::Unknown(999)
    );
    assert_eq!(properties[6].name, "Caf\u{e9}");
    assert!(matches!(
        &properties[6].value,
        UserPropertyValue::Text { value } if value == "na\u{ef}ve"
    ));
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_libreoffice_user_property_fixtures_when_available() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/extras");
    let fixtures = [
        "rtfexport/data/custom-doc-props.rtf",
        "rtfexport/data/classification-confidential.rtf",
        "rtfexport/data/classification-yes.rtf",
        "rtfexport/data/tdf158762.rtf",
        "rtfimport/data/tdf163003.rtf",
        "rtfimport/data/fdo57708.rtf",
    ];
    if !root.exists() {
        return;
    }
    for fixture in fixtures {
        let source = std::fs::read_to_string(root.join(fixture)).unwrap();
        let document = RtfDocument::parse(&source)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(!document.user_properties().is_empty(), "{fixture}");
    }
    let source = std::fs::read_to_string(root.join("rtfexport/data/tdf158762.rtf")).unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    assert!(!document.user_properties().is_empty());
    assert!(!document.document_variables().is_empty());

    let source = std::fs::read_to_string(root.join("rtfexport/data/custom-doc-props.rtf")).unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    let date = document
        .user_properties()
        .iter()
        .find(|property| property.name == "d")
        .unwrap();
    assert!(matches!(
        &date.value,
        UserPropertyValue::Date { value, .. }
            if (value.year, value.month, value.day) == (2016, 1, 30)
    ));
    let owned = date.clone().into_owned();
    assert_eq!(owned, *date);
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.user_properties(), document.user_properties());
}

#[test]
fn rejects_noncanonical_malformed_and_active_properties() {
    for source in [
        r"{\rtf1{\userprops{\propname A}\proptype30{\staticval x}}}",
        r"{\rtf1{\*\userprops{\propname A}{\staticval x}\proptype30}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype{\staticval x}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype11{\staticval true}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype64{\staticval 2023. 02. 29.}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype30{\staticval x}{\propname A}\proptype30{\staticval y}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype30{\staticval {nested}}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype30{\staticval \bin4 abcd}}}",
        r"{\rtf1{\*\userprops{\propname A}\proptype30{\staticval \field danger}}}",
        r"{\rtf1{\*\userprops}Body{\*\userprops}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }
}

#[test]
fn writer_round_trips_lexical_forms_unicode_links_and_coexisting_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\userprops"#,
        r#"{\propname A\{B\}\\C}\proptype5{\staticval 3.1400}{\linkval file:///never-opened}"#,
        r#"{\propname Emoji}\proptype30{\staticval \u-10179?\u-8704?}"#,
        r#"}{\*\docvar {Name}{Value}}Body}"#,
    ))
    .unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let output = String::from_utf8(output).unwrap();
    assert!(output.contains("{\\*\\userprops"));
    assert!(output.contains("\\proptype5{\\staticval 3.1400}"));
    let reparsed = RtfDocument::parse(&output).unwrap();
    assert_eq!(reparsed.user_properties(), document.user_properties());
    assert_eq!(reparsed.document_variables(), document.document_variables());
    assert_eq!(reparsed.text(), "Body");
}
