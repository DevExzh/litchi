use litchi_rtf::{DocumentOutputSettings, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_both_passive_parameterless_flags() {
    let document = RtfDocument::parse(r#"{\rtf1\muser\psover Body}"#).unwrap();
    assert_eq!(
        *document.output_settings(),
        DocumentOutputSettings {
            word97_compatibility_marker: true,
            postscript_over_text: true,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_preserves_no_requested_output_behavior() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(document.output_settings().is_empty());
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_side_effects() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_output_settings(DocumentOutputSettings {
        word97_compatibility_marker: true,
        postscript_over_text: true,
    });

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\muser").unwrap() < serialized.find("\\psover").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.output_settings(), document.output_settings());
    assert_eq!(reparsed.text(), "Body");

    document.clear_output_settings();
    assert!(document.output_settings().is_empty());
    let cleared = write(&document);
    let cleared_text = String::from_utf8(cleared.clone()).unwrap();
    assert!(!cleared_text.contains("\\muser"));
    assert!(!cleared_text.contains("\\psover"));
    assert_eq!(RtfDocument::parse_bytes(&cleared).unwrap().text(), "Body");
}

#[test]
fn coexists_with_adjacent_document_properties_in_canonical_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\fromhtml1\deff0{\*\nextfile next.rtf}"#,
        r#"\psover\makebackup\muser\defformat\doctemp\doctype2 Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();

    let makebackup = serialized.find("\\makebackup").unwrap();
    let defformat = serialized.find("\\defformat").unwrap();
    let doctemp = serialized.find("\\doctemp").unwrap();
    let muser = serialized.find("\\muser").unwrap();
    let psover = serialized.find("\\psover").unwrap();
    assert!(makebackup < defformat);
    assert!(defformat < doctemp);
    assert!(doctemp < muser);
    assert!(muser < psover);

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.output_settings(), document.output_settings());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_late_and_overlong_numeric_forms() {
    for source in [
        r#"{\rtf1\muser0 Body}"#,
        r#"{\rtf1\muser1 Body}"#,
        r#"{\rtf1\muser2147483647 Body}"#,
        r#"{\rtf1\psover0 Body}"#,
        r#"{\rtf1\psover-2147483648 Body}"#,
        r#"{\rtf1\psover99999999999 Body}"#,
        r#"{\rtf1\muser\muser Body}"#,
        r#"{\rtf1\psover\psover Body}"#,
        r#"{\rtf1{\*\muser}Body}"#,
        r#"{\rtf1{\*\psover}Body}"#,
        r#"{\rtf1{\muser}Body}"#,
        r#"{\rtf1{\psover}Body}"#,
        r#"{\rtf1 Body\muser}"#,
        r#"{\rtf1 Body\psover}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
