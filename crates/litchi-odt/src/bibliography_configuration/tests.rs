//! Regression tests for the bibliography model, codec, and mutation adapter.

use super::{
    Configuration, Field, SortKey, parse_bibliography_configuration,
    parse_bibliography_configuration_parts, remove_bibliography_configuration_xml,
    set_bibliography_configuration_xml,
};
use crate::FlatDocument;
use crate::variable_declaration::Part;

const PREFIX: &str = r#"<o:document-styles
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:styles>"#;
const SUFFIX: &str = "</o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>";

#[test]
fn parses_complete_bibliography_policy_and_ordered_keys() {
    let xml = format!(
        r#"{PREFIX}<t:bibliography-configuration t:prefix="[" t:suffix="]"
                t:numbered-entries="true" t:sort-by-position="false"
                t:sort-algorithm="unicode" f:language="en" f:country="US"
                f:script="Latn" s:rfc-language-tag="en-US">
                <t:sort-key t:key="author" t:sort-ascending="true"/>
                <t:sort-key t:key="year" t:sort-ascending="false"/>
                <t:sort-key t:key="isbn"/>
            </t:bibliography-configuration>{SUFFIX}"#
    );
    let configuration = parse_bibliography_configuration(&xml).unwrap().unwrap();
    assert_eq!(configuration.prefix.as_deref(), Some("["));
    assert!(configuration.effective_numbered_entries());
    assert!(!configuration.effective_sort_by_position());
    assert_eq!(configuration.sort_keys.len(), 3);
    assert_eq!(configuration.sort_keys[0].field, Field::Author);
    assert!(!configuration.sort_keys[1].effective_ascending());
    assert_eq!(configuration.sort_keys[2].field, Field::Isbn);
}

#[test]
fn applies_effective_defaults_and_accepts_empty_configuration() {
    let xml = format!(r#"{PREFIX}<t:bibliography-configuration/>{SUFFIX}"#);
    let configuration = parse_bibliography_configuration(&xml).unwrap().unwrap();
    assert!(!configuration.effective_numbered_entries());
    assert!(configuration.effective_sort_by_position());
    assert!(configuration.sort_keys.is_empty());
}

#[test]
fn rejects_invalid_structure_values_and_duplicates() {
    let bodies = [
        r#"<t:p><t:bibliography-configuration/></t:p>"#,
        r#"<t:bibliography-configuration t:numbered-entries="yes"/>"#,
        r#"<t:bibliography-configuration><t:sort-key/></t:bibliography-configuration>"#,
        r#"<t:bibliography-configuration><t:sort-key t:key="unknown"/></t:bibliography-configuration>"#,
        r#"<t:bibliography-configuration><t:sort-key t:key="author">x</t:sort-key></t:bibliography-configuration>"#,
        r#"<t:bibliography-configuration/><t:bibliography-configuration/>"#,
    ];
    for body in bodies {
        let xml = format!("{PREFIX}{body}{SUFFIX}");
        assert!(
            parse_bibliography_configuration(&xml).is_err(),
            "accepted {body}"
        );
    }
}

#[test]
fn rejects_configuration_outside_styles_metadata() {
    let xml = r#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <o:body><o:text><t:bibliography-configuration/></o:text></o:body>
        </o:document-content>"#;
    assert!(parse_bibliography_configuration_parts(&[(xml, Part::Content)]).is_err());
}

#[test]
fn flat_document_reads_configuration_from_styles_metadata() {
    let xml = r#"<o:document
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
            o:mimetype="application/vnd.oasis.opendocument.text">
            <o:styles><t:bibliography-configuration t:prefix="["/></o:styles>
            <o:body><o:text/></o:body>
        </o:document>"#;
    let document = FlatDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
    assert_eq!(
        document.bibliography_configuration().unwrap(),
        Some(Configuration {
            prefix: Some("[".to_string()),
            ..Default::default()
        })
    );
}

#[test]
fn replaces_and_removes_only_bibliography_metadata() {
    let original = format!(r#"{PREFIX}<s:style s:name="Keep"/>{SUFFIX}"#);
    let configuration = Configuration {
        prefix: Some("[".to_string()),
        suffix: Some("]".to_string()),
        numbered_entries: Some(true),
        sort_keys: vec![SortKey {
            field: Field::Author,
            ascending: Some(false),
        }],
        ..Default::default()
    };
    let inserted = set_bibliography_configuration_xml(&original, &configuration).unwrap();
    assert!(inserted.contains(r#"<s:style s:name="Keep"/>"#));
    assert_eq!(
        parse_bibliography_configuration(&inserted).unwrap(),
        Some(configuration.clone())
    );

    let removed = remove_bibliography_configuration_xml(&inserted).unwrap();
    assert!(removed.contains(r#"<s:style s:name="Keep"/>"#));
    assert_eq!(parse_bibliography_configuration(&removed).unwrap(), None);
}

#[test]
fn inserts_into_an_empty_styles_element() {
    let original = r#"<o:document-styles
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
            <o:styles/>
        </o:document-styles>"#;
    let configuration = Configuration {
        prefix: Some("[".to_string()),
        ..Default::default()
    };
    let inserted = set_bibliography_configuration_xml(original, &configuration).unwrap();
    assert_eq!(
        parse_bibliography_configuration(&inserted).unwrap(),
        Some(configuration)
    );
}
