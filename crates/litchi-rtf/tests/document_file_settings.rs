use litchi_rtf::{DocumentFileSettings, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_three_passive_parameterless_flags() {
    let document = RtfDocument::parse(r#"{\rtf1\makebackup\defformat\doctemp Body}"#).unwrap();
    assert_eq!(
        *document.file_settings(),
        DocumentFileSettings {
            automatic_backup: true,
            default_save_format_rtf: true,
            template_or_stationery: true,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_is_preserved_as_no_requested_file_behavior() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(document.file_settings().is_empty());
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_side_effects() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_file_settings(DocumentFileSettings {
        automatic_backup: true,
        default_save_format_rtf: true,
        template_or_stationery: true,
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\makebackup").unwrap() < serialized.find("\\defformat").unwrap());
    assert!(serialized.find("\\defformat").unwrap() < serialized.find("\\doctemp").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.file_settings(), document.file_settings());
    assert_eq!(reparsed.text(), "Body");
    document.clear_file_settings();
    assert!(document.file_settings().is_empty());
}

#[test]
fn coexists_with_origin_external_reference_and_classification_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\fromhtml1\deff0{\*\nextfile next.rtf}"#,
        r#"\doctemp\defformat\makebackup\doctype2 Body}"#,
    ))
    .unwrap();
    assert!(document.file_settings().automatic_backup);
    assert!(document.file_settings().default_save_format_rtf);
    assert!(document.file_settings().template_or_stationery);
    assert_eq!(document.text(), "Body");
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_late_and_overlong_numeric_forms() {
    for source in [
        r#"{\rtf1\makebackup0 Body}"#,
        r#"{\rtf1\makebackup2147483647 Body}"#,
        r#"{\rtf1\defformat1 Body}"#,
        r#"{\rtf1\defformat-2147483648 Body}"#,
        r#"{\rtf1\doctemp0 Body}"#,
        r#"{\rtf1\doctemp99999999999 Body}"#,
        r#"{\rtf1\makebackup\makebackup Body}"#,
        r#"{\rtf1\defformat\defformat Body}"#,
        r#"{\rtf1\doctemp\doctemp Body}"#,
        r#"{\rtf1{\*\makebackup}Body}"#,
        r#"{\rtf1{\*\defformat}Body}"#,
        r#"{\rtf1{\*\doctemp}Body}"#,
        r#"{\rtf1{\makebackup}Body}"#,
        r#"{\rtf1{\defformat}Body}"#,
        r#"{\rtf1{\doctemp}Body}"#,
        r#"{\rtf1 Body\makebackup}"#,
        r#"{\rtf1 Body\defformat}"#,
        r#"{\rtf1 Body\doctemp}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
