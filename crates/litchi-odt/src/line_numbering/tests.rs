use super::*;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

fn styles(body: &str) -> String {
    format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:styles>{body}</o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>"#
    )
}

#[test]
fn parses_and_round_trips_complete_line_numbering() {
    // ODF 1.2/1.3 text:linenumbering-configuration grammar. LibreOffice
    // writes the same configuration in office:styles.
    let xml = styles(
        r#"<t:linenumbering-configuration t:number-lines="1" s:num-format="A" s:num-letter-sync="false" t:style-name="Line &amp; Number" t:increment="5" t:number-position="outer" t:offset="0.25in" t:count-empty-lines="true" t:count-in-text-boxes="0" t:restart-on-page="false"><t:linenumbering-separator t:increment="10"> / &amp; </t:linenumbering-separator></t:linenumbering-configuration>"#,
    );
    let configuration = parse(&xml).unwrap().unwrap();
    assert_eq!(configuration.number_lines, Some(true));
    assert_eq!(configuration.number_format, Some(Format::UpperAlpha));
    assert_eq!(configuration.number_position, Some(Position::Outer));
    assert_eq!(configuration.offset.as_ref().unwrap().as_str(), "0.25in");
    assert_eq!(configuration.separator.as_ref().unwrap().text, " / & ");

    let serialized = configuration.to_xml().unwrap();
    let reparsed = parse(&styles(&serialized)).unwrap().unwrap();
    assert_eq!(reparsed, configuration);
}

#[test]
fn preserves_empty_and_custom_number_formats() {
    for format in [Format::Empty, Format::Custom("一, 二, 三".to_string())] {
        let configuration = Configuration {
            number_format: Some(format.clone()),
            ..Configuration::default()
        };
        let parsed = parse(&styles(&configuration.to_xml().unwrap()))
            .unwrap()
            .unwrap();
        assert_eq!(parsed.number_format, Some(format));
    }
}

#[test]
fn rejects_malformed_or_misplaced_configurations() {
    for body in [
        r#"<t:linenumbering-configuration t:number-lines="yes"/>"#,
        r#"<t:linenumbering-configuration s:num-format="1" s:num-letter-sync="true"/>"#,
        r#"<t:linenumbering-configuration t:number-position="center"/>"#,
        r#"<t:linenumbering-configuration t:offset="-1cm"/>"#,
        r#"<t:linenumbering-configuration t:increment="-1"/>"#,
        r#"<t:linenumbering-configuration><t:linenumbering-separator/><t:linenumbering-separator/></t:linenumbering-configuration>"#,
        r#"<t:linenumbering-configuration><t:linenumbering-separator><t:span/></t:linenumbering-separator></t:linenumbering-configuration>"#,
    ] {
        assert!(parse(&styles(body)).is_err(), "{body}");
    }
    let misplaced = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body><t:linenumbering-configuration/></o:body></o:document-content>"#
    );
    assert!(parse(&misplaced).is_err());
    assert!(
        parse(&styles(
            r#"<t:linenumbering-configuration/><t:linenumbering-configuration/>"#
        ))
        .is_err()
    );
}

#[test]
fn replaces_inserts_and_removes_configuration_without_rewriting_other_styles() {
    let original = styles(
        r#"<s:style s:name="Preserved"/><t:linenumbering-configuration t:number-lines="false" s:num-format="1"/>"#,
    );
    let configuration = Configuration {
        number_lines: Some(true),
        number_format: Some(Format::LowerAlpha),
        letter_sync: Some(true),
        style_name: Some("LineNumbers".to_string()),
        increment: Some(3),
        number_position: Some(Position::Outer),
        offset: Some(NonNegativeLength::new("0.25in").unwrap()),
        count_empty_lines: Some(true),
        count_in_text_boxes: Some(false),
        restart_on_page: Some(true),
        separator: Some(Separator {
            increment: Some(6),
            text: " · ".to_string(),
        }),
    };

    let replaced = set_xml(&original, &configuration).unwrap();
    assert!(replaced.contains(r#"<s:style s:name="Preserved"/>"#));
    assert_eq!(parse(&replaced).unwrap(), Some(configuration.clone()));

    let removed = remove_xml(&replaced).unwrap();
    assert!(removed.contains(r#"<s:style s:name="Preserved"/>"#));
    assert_eq!(parse(&removed).unwrap(), None);

    let empty_styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:styles/><o:automatic-styles/><o:master-styles/></o:document-styles>"#
    );
    let inserted = set_xml(&empty_styles, &configuration).unwrap();
    assert!(inserted.contains("<o:styles>"));
    assert_eq!(parse(&inserted).unwrap(), Some(configuration.clone()));

    let flat_xml = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}" o:mimetype="application/vnd.oasis.opendocument.text" o:version="1.3"><o:styles>{}</o:styles><o:body><o:text/></o:body></o:document>"#,
        configuration.to_xml().unwrap()
    );
    let flat = crate::FlatDocument::from_bytes(flat_xml.into_bytes()).unwrap();
    assert_eq!(
        flat.line_numbering_configuration().unwrap(),
        Some(configuration)
    );
}
