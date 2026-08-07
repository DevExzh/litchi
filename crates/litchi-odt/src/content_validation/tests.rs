//! Regression tests for bounded and lossless content-validation handling.

use super::*;
const PREFIX: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet>"#;
const SUFFIX: &str = "</office:spreadsheet></office:body></office:document>";

#[test]
fn parses_messages_macro_metadata_and_round_trips() {
    let body = r#"<table:content-validations><table:content-validation table:name="number" table:condition="of:cell-content-is-whole-number()" table:base-cell-address="Sheet1.A1" table:allow-empty-cell="0" table:display-list="sort-ascending"><table:help-message table:title="Input" table:display="true"><text:p>Use <text:span>whole</text:span> numbers.</text:p></table:help-message><table:error-message table:title="Invalid" table:display="1" table:message-type="warning"><text:p>Try again.</text:p></table:error-message></table:content-validation><table:content-validation table:name="macro"><table:error-macro table:execute="false"/><office:event-listeners><script:event-listener script:event-name="dom:invalid" xlink:href="macro://ignored"/></office:event-listeners></table:content-validation></table:content-validations>"#;
    let parsed = parse_content_validations(&format!("{PREFIX}{body}{SUFFIX}")).unwrap();
    assert_eq!(parsed.validations.len(), 2);
    assert_eq!(
        parsed.validations[0]
            .help_message
            .as_ref()
            .unwrap()
            .paragraphs[0]
            .text(),
        "Use whole numbers."
    );
    assert!(matches!(
        parsed.validations[1].failure,
        Some(ValidationFailure::Macro {
            execute: Some(false),
            event_listeners: Some(_),
            ..
        })
    ));
    let fragment = parsed.to_xml_fragment().unwrap();
    let reparsed = parse_content_validations(&format!("{PREFIX}{fragment}{SUFFIX}")).unwrap();
    assert_eq!(reparsed, parsed);
}

#[test]
fn rejects_invalid_validation_grammar() {
    for body in [
        "<table:content-validations/>",
        r"<table:content-validations><table:content-validation/></table:content-validations>",
        r#"<table:content-validations><table:content-validation table:name="x" table:display-list="sorted"/></table:content-validations>"#,
        r#"<table:content-validations><table:content-validation table:name="x"><table:error-message/><table:help-message/></table:content-validation></table:content-validations>"#,
        r#"<table:content-validations><table:content-validation table:name="x"><office:event-listeners/></table:content-validation></table:content-validations>"#,
        r#"<table:content-validations><table:content-validation table:name="x"><table:error-macro><text:p>x</text:p></table:error-macro></table:content-validation></table:content-validations>"#,
        r#"<table:content-validations><table:content-validation table:name="x"/><table:content-validation table:name="x"/></table:content-validations>"#,
    ] {
        assert!(
            parse_content_validations(&format!("{PREFIX}{body}{SUFFIX}")).is_err(),
            "accepted {body}"
        );
    }
    assert!(parse_content_validations("<!DOCTYPE x><x/>").is_err());
}

#[test]
fn parses_libreoffice_content_validation_when_available() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/libreoffice-core/sc/qa/unit/data/functions/mathematical/fods/aggregate.fods");
    let Ok(xml) = std::fs::read_to_string(path) else {
        return;
    };
    let parsed = parse_content_validations(&xml).unwrap();
    let value = parsed.get("val1").unwrap();
    assert_eq!(
        value.effective_display_list(),
        ValidationDisplayList::Unsorted
    );
    assert!(
        value
            .condition
            .as_ref()
            .unwrap()
            .as_str()
            .contains("cell-content-is-in-list")
    );
}
