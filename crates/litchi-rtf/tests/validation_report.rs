#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "validation fixtures are fixed and assertions intentionally fail loudly"
)]

use litchi_rtf::{
    Document, ValidationDependency, ValidationStatus,
    validation::{ValidationReport, validate_with_limits},
};

const CLEAN: &str = r#"{\rtf1\ansi Clean text\par}"#;

#[test]
fn clean_report_is_content_free_and_explicit_about_scope() {
    let report = ValidationReport::from_str(CLEAN).unwrap();

    assert_eq!(report.syntax().status(), ValidationStatus::Valid);
    assert_eq!(report.root().status(), ValidationStatus::Valid);
    assert_eq!(report.document().status(), ValidationStatus::Valid);
    assert_eq!(
        report.compressed_transport().status(),
        ValidationStatus::NotApplicable
    );
    assert_eq!(
        report.compressed_transport().dependency(),
        ValidationDependency::CompressedTransport
    );
    assert_eq!(report.fields().status(), ValidationStatus::Absent);
    assert_eq!(report.external_links().status(), ValidationStatus::Absent);
    assert_eq!(report.objects().status(), ValidationStatus::Absent);
    assert_eq!(report.pictures().status(), ValidationStatus::Absent);
    assert_eq!(report.active_content().status(), ValidationStatus::Absent);
    assert_eq!(
        report.unsupported_syntax().status(),
        ValidationStatus::Absent
    );
    assert_eq!(
        report.external_resolution().status(),
        ValidationStatus::NotApplicable
    );
    assert_eq!(report.execution().status(), ValidationStatus::NotApplicable);
    assert_eq!(report.repair().status(), ValidationStatus::NotApplicable);
    assert_eq!(report.security().status(), ValidationStatus::Absent);
    assert!(report.is_conservatively_clean());
    assert_eq!(report.counts().fields(), 0);
    assert_eq!(report.counts().source_bytes(), CLEAN.len());
    assert_eq!(report.limits().max_binary_bytes(), 256 * 1_048_576);
    assert_eq!(report.limits().max_opaque_node_bytes(), 8 * 1_048_576);
    assert_eq!(report.limits().max_object_payload_bytes(), 64 * 1_048_576);
    assert_eq!(report.limits().max_picture_payload_bytes(), 64 * 1_048_576);
    assert_eq!(report.counts().unknown_syntax_markers(), 0);
}

#[test]
fn external_field_is_present_but_never_resolved_or_exposed_in_report() {
    let source = concat!(
        r#"{\rtf1\ansi before"#,
        r#"{\field{\*\fldinst HYPERLINK "https://example.invalid/a"}"#,
        r#"{\fldrslt link}} after}"#,
    );
    let document = Document::parse(source).unwrap();
    let before = document.to_bytes().unwrap();
    let report = document.validation_report();

    assert_eq!(report.fields().status(), ValidationStatus::Present);
    assert_eq!(report.external_links().status(), ValidationStatus::Present);
    assert_eq!(
        report.external_resolution().status(),
        ValidationStatus::Unsupported
    );
    assert_eq!(
        report.external_resolution().dependency(),
        ValidationDependency::ExternalProvider
    );
    assert_eq!(report.execution().status(), ValidationStatus::NotApplicable);
    assert_eq!(report.security().status(), ValidationStatus::Present);
    assert_eq!(report.counts().fields(), 1);
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn malformed_external_field_keyword_is_unknown_not_proven_present() {
    let source = concat!(
        r#"{\rtf1\ansi before"#,
        r#"{\field{\*\fldinst DDE}{\fldrslt cached}}"#,
        r#" after}"#,
    );
    let report = ValidationReport::from_str(source).unwrap();

    assert_eq!(report.fields().status(), ValidationStatus::Present);
    assert_eq!(report.external_links().status(), ValidationStatus::Unknown);
    assert_eq!(
        report.external_resolution().status(),
        ValidationStatus::Unknown
    );
    assert_eq!(report.security().status(), ValidationStatus::Unknown);
}

#[test]
fn dde_and_ddeauto_are_both_external_and_active() {
    for instruction in [
        r#"DDE Excel "source.xlsx" "Sheet1!A1""#,
        r#"DDEAUTO Excel "source.xlsx" "Sheet1!A1""#,
    ] {
        let source = format!(
            r#"{{\rtf1\ansi{{\field{{\*\fldinst {instruction}}}{{\fldrslt cached}}}}Body}}"#
        );
        let report = ValidationReport::from_str(&source).unwrap();
        assert_eq!(report.external_links().status(), ValidationStatus::Present);
        assert_eq!(report.active_content().status(), ValidationStatus::Present);
        assert_eq!(report.execution().status(), ValidationStatus::Unsupported);
        assert_eq!(
            report.external_resolution().status(),
            ValidationStatus::Unsupported
        );
    }
}

#[test]
fn xsl_and_mail_merge_metadata_are_external_and_active_surfaces() {
    let xsl =
        ValidationReport::from_str(r#"{\rtf1{\*\xform transform.xsl}\usexform Body}"#).unwrap();
    assert_eq!(xsl.external_links().status(), ValidationStatus::Present);
    assert_eq!(xsl.active_content().status(), ValidationStatus::Present);
    assert_eq!(xsl.execution().status(), ValidationStatus::Unsupported);

    let mail_merge = concat!(
        r#"{\rtf1\ansi{\*\mailmerge\mmlinktoquery"#,
        r#"{\*\mmconnectstr Provider=SQLOLEDB;Server=invalid.example}"#,
        r#"{\*\mmdatasource file:///definitely/not-opened.csv}"#,
        r#"{\*\mmodso\mmodsoactive7}"#,
        r#"}Body}"#,
    );
    let merge = ValidationReport::from_str(mail_merge).unwrap();
    assert_eq!(merge.external_links().status(), ValidationStatus::Present);
    assert_eq!(merge.active_content().status(), ValidationStatus::Present);
    assert_eq!(merge.execution().status(), ValidationStatus::Unsupported);
}

#[test]
fn mail_merge_fields_are_external_and_active_even_with_cached_results() {
    for instruction in [
        r#"MERGEFIELD Customer"#,
        r#"DATABASE \\d "NeverOpened.csv" \\c "DSN=NeverConnect" \\s "SELECT 1" \\h"#,
        r#"DATA "NeverOpened.csv""#,
        "MERGEREC",
        "MERGESEQ",
        "NEXT",
        "NEXTIF Customer = Ada",
        "SKIPIF Customer = Ada",
        r#"ADDRESSBLOCK \\f "<<_FIRST0_>>""#,
        r#"GREETINGLINE \\f "Hello""#,
    ] {
        let source = format!(
            r#"{{\rtf1\ansi{{\field{{\*\fldinst {instruction}}}{{\fldrslt cached}}}}Body}}"#
        );
        let report = ValidationReport::from_str(&source).unwrap();
        assert_eq!(
            report.external_links().status(),
            ValidationStatus::Present,
            "{instruction}"
        );
        assert_eq!(
            report.active_content().status(),
            ValidationStatus::Present,
            "{instruction}"
        );
        assert_eq!(
            report.security().status(),
            ValidationStatus::Present,
            "{instruction}"
        );
    }
}

#[test]
fn destination_binary_limits_reject_one_over_before_arena_copy() {
    let limits = litchi_rtf::read::Limits::default()
        .with_max_binary_bytes(3)
        .with_max_total_binary_bytes(3);
    let picture = br#"{\rtf1{\pict\pngblip\bin3 abc}}"#;
    let picture_report = ValidationReport::from_bytes_with_limits(picture, limits).unwrap();
    assert_eq!(picture_report.counts().pictures(), 1);

    let over_picture = br#"{\rtf1{\pict\pngblip\bin4 abcd}}"#;
    assert!(ValidationReport::from_bytes_with_limits(over_picture, limits).is_err());

    let object = br#"{\rtf1{\object\objemb{\*\objdata\bin3 abc}}}"#;
    let object_report = ValidationReport::from_bytes_with_limits(object, limits).unwrap();
    assert_eq!(object_report.counts().objects(), 1);

    let over_object = br#"{\rtf1{\object\objemb{\*\objdata\bin4 abcd}}}"#;
    assert!(ValidationReport::from_bytes_with_limits(over_object, limits).is_err());
}

#[test]
fn malformed_active_field_keyword_is_unknown_not_proven_executable() {
    let source = concat!(
        r#"{\rtf1\ansi before"#,
        r#"{\field{\*\fldinst MACROBUTTON}{\fldrslt cached}}"#,
        r#" after}"#,
    );
    let report = ValidationReport::from_str(source).unwrap();

    assert_eq!(report.fields().status(), ValidationStatus::Present);
    assert_eq!(report.active_content().status(), ValidationStatus::Unknown);
    assert_eq!(report.execution().status(), ValidationStatus::Unknown);
    assert_eq!(report.security().status(), ValidationStatus::Unknown);
}

#[test]
fn known_objects_and_pictures_are_counted_without_inspecting_payload_content() {
    let source = concat!(
        r#"{\rtf1\ansi"#,
        r#"{\*\shppict{\pict\pngblip 89504e470d0a1a0a}}"#,
        r#"{\object\objemb{\*\objdata 00}}"#,
        r#"Body}"#,
    );
    let report = ValidationReport::from_str(source).unwrap();

    assert_eq!(report.counts().pictures(), 1);
    assert_eq!(report.counts().objects(), 1);
    assert_eq!(report.pictures().status(), ValidationStatus::Present);
    assert_eq!(report.objects().status(), ValidationStatus::Present);
    assert_eq!(report.security().status(), ValidationStatus::Present);
}

#[test]
fn linked_objects_are_external_and_unknown_object_kinds_fail_closed() {
    let linked = concat!(
        r#"{\rtf1\ansi"#,
        r#"{\object\objlink{\*\objclass Package}{\*\objdata 00}}"#,
        r#"Body}"#,
    );
    let linked_report = ValidationReport::from_str(linked).unwrap();
    assert_eq!(
        linked_report.external_links().status(),
        ValidationStatus::Present
    );
    assert_eq!(
        linked_report.external_resolution().status(),
        ValidationStatus::Unsupported
    );

    // No object-kind control is a parser-accepted retained object with an
    // unknown storage mode. Its payload remains opaque to this report.
    let unknown = concat!(r#"{\rtf1\ansi"#, r#"{\object{\*\objdata 00}}"#, r#"Body}"#,);
    let unknown_report = ValidationReport::from_str(unknown).unwrap();
    assert_eq!(unknown_report.objects().status(), ValidationStatus::Unknown);
    assert_eq!(
        unknown_report.external_links().status(),
        ValidationStatus::Unknown
    );
    assert_eq!(
        unknown_report.security().status(),
        ValidationStatus::Unknown
    );

    let unknown_picture = r#"{\rtf1\ansi{\pict 00}Body}"#;
    let unknown_picture_report = ValidationReport::from_str(unknown_picture).unwrap();
    assert_eq!(
        unknown_picture_report.pictures().status(),
        ValidationStatus::Unknown
    );
    assert_eq!(
        unknown_picture_report.security().status(),
        ValidationStatus::Unknown
    );
}

#[test]
fn opaque_and_unknown_fields_fail_closed_without_leaking_content() {
    let source = concat!(
        r#"{\rtf1\ansi"#,
        r#"{\*\futureDangerDestination opaque}"#,
        r#"{\field{\*\fldinst FUTUREACTIVE "opaque"}{\fldrslt cached}}"#,
        r#"Body}"#,
    );
    let report = ValidationReport::from_str(source).unwrap();

    assert_eq!(
        report.unsupported_syntax().status(),
        ValidationStatus::Present
    );
    assert_eq!(report.fields().status(), ValidationStatus::Unknown);
    assert_eq!(report.external_links().status(), ValidationStatus::Unknown);
    assert_eq!(report.active_content().status(), ValidationStatus::Unknown);
    assert_eq!(
        report.external_resolution().status(),
        ValidationStatus::Unknown
    );
    assert_eq!(report.execution().status(), ValidationStatus::Unknown);
    assert_eq!(report.security().status(), ValidationStatus::Unknown);
    assert!(!report.is_conservatively_clean());
}

#[test]
fn compressed_and_raw_byte_ingress_use_one_bounded_parse() {
    let raw = CLEAN.as_bytes();
    let compressed = litchi_rtf::transport::compress(raw, true).unwrap();
    let raw_report = ValidationReport::from_bytes(raw).unwrap();
    let compressed_report = ValidationReport::from_bytes(&compressed).unwrap();

    assert_eq!(
        raw_report.compressed_transport().status(),
        ValidationStatus::NotApplicable
    );
    assert_eq!(
        compressed_report.compressed_transport().status(),
        ValidationStatus::Present
    );
    assert_eq!(compressed_report.counts().source_bytes(), compressed.len());
    assert_eq!(
        compressed_report.security().status(),
        ValidationStatus::Absent
    );
}

#[test]
fn malformed_root_and_finite_limits_are_errors_not_safe_reports() {
    assert!(ValidationReport::from_str("plain text").is_err());
    assert!(ValidationReport::from_str(r#"{\ansi text}"#).is_err());
    assert!(ValidationReport::from_str(r#"{\rtf text}"#).is_err());
    assert!(ValidationReport::from_str(r#"{\rtf2 text}"#).is_err());
    assert!(ValidationReport::from_str(r#"{\rtf1\ansi unterminated"#).is_err());
    assert!(ValidationReport::from_str(r#"{\rtf1 text} trailing"#).is_err());
    assert!(ValidationReport::from_str(r#"{\rtf1 text}{\rtf1 second}"#).is_err());
    assert!(ValidationReport::from_str("{\\rtf1 text} \t\r\n").is_ok());

    let limits = litchi_rtf::read::Limits::default().with_max_source_bytes(CLEAN.len() - 1);
    assert!(validate_with_limits(raw_bytes(CLEAN), limits).is_err());

    let report = ValidationReport::from_str_with_limits(
        CLEAN,
        litchi_rtf::read::Limits::default().with_max_source_bytes(CLEAN.len()),
    )
    .unwrap();
    assert_eq!(report.limits().max_source_bytes(), CLEAN.len());
    assert_eq!(report.limits().max_group_depth(), 32);
}

#[test]
fn empty_picture_and_dropped_nested_destination_leave_unknown_markers() {
    let empty = ValidationReport::from_str(r#"{\rtf1{\pict}Body}"#).unwrap();
    assert_eq!(empty.pictures().status(), ValidationStatus::Unknown);
    assert!(empty.counts().unknown_syntax_markers() > 0);
    assert_eq!(empty.security().status(), ValidationStatus::Unknown);

    let nested = r#"{\rtf1{\pict\pngblip{\*\futurePictureNested ignored}89504e}}"#;
    let nested_report = ValidationReport::from_str(nested).unwrap();
    assert!(nested_report.counts().unknown_syntax_markers() > 0);
    assert_eq!(nested_report.security().status(), ValidationStatus::Unknown);

    let object = r#"{\rtf1{\object\objemb{\*\objdata 00{\*\futureObjectNested ignored}}}}"#;
    let object_report = ValidationReport::from_str(object).unwrap();
    assert!(object_report.counts().unknown_syntax_markers() > 0);
    assert_eq!(object_report.objects().status(), ValidationStatus::Unknown);
    assert_eq!(object_report.security().status(), ValidationStatus::Unknown);
}

#[test]
fn deep_iterative_object_data_is_rejected_with_a_finite_depth_bound() {
    let mut source = String::from(r#"{\rtf1{\object\objemb{\*\objdata "#);
    for _ in 0..40 {
        source.push('{');
    }
    source.push_str("00");
    for _ in 0..40 {
        source.push('}');
    }
    source.push_str("}Body}");
    assert!(ValidationReport::from_str(&source).is_err());
}

#[test]
fn deeply_nested_generic_fields_are_rejected_before_stack_growth() {
    let mut source = String::from(r#"{\rtf1{\field{\*\fldinst PAGE}{\fldrslt "#);
    for _ in 0..40 {
        source.push_str(r#"{\field{\*\fldinst PAGE}{\fldrslt "#);
    }
    source.push_str("leaf");
    for _ in 0..41 {
        source.push_str("}}");
    }
    source.push('}');
    assert!(ValidationReport::from_str(&source).is_err());
}

fn raw_bytes(source: &str) -> &[u8] {
    source.as_bytes()
}
