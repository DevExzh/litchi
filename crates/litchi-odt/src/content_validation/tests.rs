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

/// Every sequence of at most `max_len` indices below `alphabet`, shortest
/// first.
fn sequences(alphabet: usize, max_len: usize) -> Vec<Vec<usize>> {
    let mut all = vec![Vec::new()];
    let mut previous = vec![Vec::new()];
    for _ in 0..max_len {
        let mut next = Vec::new();
        for prefix in &previous {
            for index in 0..alphabet {
                let mut sequence: Vec<usize> = prefix.clone();
                sequence.push(index);
                next.push(sequence);
            }
        }
        all.extend(next.iter().cloned());
        previous = next;
    }
    all
}

fn assert_lines(actual: &[String], expected: &[&str]) {
    let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
    assert!(
        actual
            .iter()
            .map(String::as_str)
            .eq(expected.iter().copied()),
        "observed lines:\n{}",
        listing.join("\n")
    );
}

fn validation_child_outcome(case: &[(&str, &str)]) -> String {
    let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
    let body: String = case.iter().map(|(_, xml)| *xml).collect();
    let xml = format!(
        "{PREFIX}<table:content-validations><table:content-validation table:name=\"x\">{body}</table:content-validation></table:content-validations>{SUFFIX}"
    );
    let outcome = match parse_content_validations(&xml) {
        Ok(parsed) => {
            let value = &parsed.validations[0];
            let failure = match &value.failure {
                None => "none",
                Some(ValidationFailure::Message(_)) => "message",
                Some(ValidationFailure::Macro {
                    event_listeners: None,
                    ..
                }) => "macro",
                Some(ValidationFailure::Macro {
                    event_listeners: Some(_),
                    ..
                }) => "macro+listeners",
            };
            format!("ok help={} failure={failure}", value.help_message.is_some())
        },
        Err(error) => format!("{error:?}"),
    };
    format!("{} => {outcome}", label.join(","))
}

const VALIDATION_CHILDREN: [(&str, &str); 8] = [
    ("help", "<table:help-message/>"),
    ("error", "<table:error-message/>"),
    ("macro", "<table:error-macro/>"),
    ("listeners", "<office:event-listeners/>"),
    ("help-open", "<table:help-message></table:help-message>"),
    ("error-open", "<table:error-message></table:error-message>"),
    ("macro-open", "<table:error-macro></table:error-macro>"),
    (
        "listeners-open",
        "<office:event-listeners></office:event-listeners>",
    ),
];

#[test]
fn validation_child_order_pair_matrix_is_stable() {
    let mut cases: Vec<Vec<(&str, &str)>> = sequences(VALIDATION_CHILDREN.len(), 2)
        .into_iter()
        .map(|sequence| {
            sequence
                .into_iter()
                .map(|index| VALIDATION_CHILDREN[index])
                .collect()
        })
        .collect();
    let unknown = ("unknown", "<table:unknown/>");
    let unknown_open = ("unknown-open", "<table:unknown></table:unknown>");
    let foreign = ("foreign help", "<x:help-message xmlns:x=\"urn:other\"/>");
    cases.push(vec![unknown]);
    cases.push(vec![unknown_open]);
    cases.push(vec![VALIDATION_CHILDREN[0], unknown]);
    cases.push(vec![VALIDATION_CHILDREN[2], unknown_open]);
    cases.push(vec![foreign]);
    let actual: Vec<String> = cases
        .iter()
        .map(|case| validation_child_outcome(case))
        .collect();
    assert_lines(
        &actual,
        &[
            " => ok help=false failure=none",
            "help => ok help=true failure=none",
            "error => ok help=false failure=message",
            "macro => ok help=false failure=macro",
            "listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help-open => ok help=true failure=none",
            "error-open => ok help=false failure=message",
            "macro-open => ok help=false failure=macro",
            "listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error => ok help=true failure=message",
            "help,macro => ok help=true failure=macro",
            "help,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error-open => ok help=true failure=message",
            "help,macro-open => ok help=true failure=macro",
            "help,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "macro,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners => ok help=false failure=macro+listeners",
            "macro,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners-open => ok help=false failure=macro+listeners",
            "listeners,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,help-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help-open,error => ok help=true failure=message",
            "help-open,macro => ok help=true failure=macro",
            "help-open,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help-open,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help-open,error-open => ok help=true failure=message",
            "help-open,macro-open => ok help=true failure=macro",
            "help-open,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error-open,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,error-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,macro-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error-open,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "macro-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,listeners => ok help=false failure=macro+listeners",
            "macro-open,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,error-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,macro-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro-open,listeners-open => ok help=false failure=macro+listeners",
            "listeners-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,help-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "unknown => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "unknown-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,unknown => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,unknown-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "foreign help => InvalidFormat(\"spoofed table:help-message element namespace\")",
        ],
    );
}

#[test]
fn validation_child_order_triple_matrix_is_stable() {
    // Triples over the empty-element forms plus the start-tag form of
    // `office:event-listeners`, which takes a different path from the
    // empty-element form.
    let forms = [
        VALIDATION_CHILDREN[0],
        VALIDATION_CHILDREN[1],
        VALIDATION_CHILDREN[2],
        VALIDATION_CHILDREN[3],
        VALIDATION_CHILDREN[7],
    ];
    let actual: Vec<String> = sequences(forms.len(), 3)
        .into_iter()
        .filter(|sequence| sequence.len() == 3)
        .map(|sequence| {
            let case: Vec<(&str, &str)> = sequence.into_iter().map(|index| forms[index]).collect();
            validation_child_outcome(&case)
        })
        .collect();
    assert_lines(
        &actual,
        &[
            "help,help,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,help,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,help,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,help,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,help,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,error,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,error,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,macro,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,macro,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,macro,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,macro,listeners => ok help=true failure=macro+listeners",
            "help,macro,listeners-open => ok help=true failure=macro+listeners",
            "help,listeners,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,listeners,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,listeners,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,listeners,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,listeners,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "help,listeners-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,listeners-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,listeners-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,listeners-open,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "help,listeners-open,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,help,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,error,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,macro,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "error,listeners,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners-open,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners-open,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners-open,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners-open,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "error,listeners-open,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "macro,help,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,help,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,help,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,help,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,help,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,error,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,macro,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "macro,listeners,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "macro,listeners-open,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "macro,listeners-open,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners,help,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,help,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,help,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,help,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,help,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,error,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,macro,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open,help => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open,error => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open,macro => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open,listeners => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners,listeners-open,listeners-open => InvalidFormat(\"office:event-listeners must follow exactly one table:error-macro\")",
            "listeners-open,help,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,help,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,help,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,help,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,help,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,error,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,macro,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open,help => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open,error => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open,macro => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open,listeners => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
            "listeners-open,listeners-open,listeners-open => InvalidFormat(\"invalid, duplicate, or out-of-order content-validation child\")",
        ],
    );
}
