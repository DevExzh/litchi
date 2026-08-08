use litchi_odt::{
    Document,
    core::PackageWriter,
    generic::{FlatDocument, Package},
    variable_declaration::{Declaration, Kind, Value},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

#[test]
fn parses_ordered_typed_inert_declarations_and_sequence_defaults() {
    let xml = flat(concat!(
        r#"<t:variable-decls><t:variable-decl t:name="simple" o:value-type="float"/></t:variable-decls>"#,
        r#"<t:user-field-decls>"#,
        r#"<t:user-field-decl t:name="float" o:value-type="float" o:value="3.1400"/>"#,
        r#"<t:user-field-decl t:name="percentage" o:value-type="percentage" o:value="0.25"/>"#,
        r#"<t:user-field-decl t:name="currency" o:value-type="currency" o:value="42" o:currency="EUR"/>"#,
        r#"<t:user-field-decl t:name="date" o:value-type="date" o:date-value="2026-07-16"/>"#,
        r#"<t:user-field-decl t:name="datetime" o:value-type="date" o:date-value="2026-07-16T12:30:00Z"/>"#,
        r#"<t:user-field-decl t:name="time" o:value-type="time" o:time-value="PT1H2M3.40S"/>"#,
        r#"<t:user-field-decl t:name="bool" o:value-type="boolean" o:boolean-value="true"/>"#,
        r#"<t:user-field-decl t:name="string" o:value-type="string" o:string-value="hello" t:formula="of:=WEBSERVICE(&quot;https://never.invalid&quot;)"/>"#,
        r#"<t:user-field-decl t:name="void" o:value-type="void"/>"#,
        r#"</t:user-field-decls>"#,
        r##"<t:sequence-decls><t:sequence-decl t:name="Figure" t:display-outline-level="3" t:separation-character="#"/><t:sequence-decl t:name="Table" t:display-outline-level="2"/></t:sequence-decls>"##,
        r#"<t:p><t:variable-get t:name="simple">1</t:variable-get><t:user-field-get t:name="string">hello</t:user-field-get><t:sequence t:name="Figure" t:formula="ooow:Figure+1">1</t:sequence></t:p>"#,
    ));
    let declarations = FlatDocument::from_bytes(xml)
        .unwrap()
        .variable_declarations()
        .unwrap();
    assert_eq!(declarations.groups.len(), 3);
    assert_eq!(declarations.declarations().count(), 12);
    let float = declarations.find(Kind::User, "float").unwrap();
    assert!(matches!(
        float,
        Declaration::User {
            value: Some(Value::Float { value, lexical }),
            ..
        } if *value == lexical.parse::<f64>().unwrap() && lexical == "3.1400"
    ));
    let string = declarations.find(Kind::User, "string").unwrap();
    assert!(matches!(
        string,
        Declaration::User { formula: Some(formula), .. }
            if formula.contains("WEBSERVICE")
    ));
    assert_eq!(
        declarations
            .find(Kind::Sequence, "Figure")
            .unwrap()
            .effective_separation_character(),
        Some('#')
    );
    assert_eq!(
        declarations
            .find(Kind::Sequence, "Table")
            .unwrap()
            .effective_separation_character(),
        Some('.')
    );
}

#[test]
fn rejects_malformed_spoofed_active_and_ambiguous_declarations() {
    for inner in [
        r#"<t:variable-decls><t:variable-decl t:name="x"/></t:variable-decls>"#,
        r#"<t:sequence-decls><t:sequence-decl t:name="x" t:display-outline-level="11"/></t:sequence-decls>"#,
        r#"<t:sequence-decls><t:sequence-decl t:name="x" t:display-outline-level="0" t:separation-character="."/></t:sequence-decls>"#,
        r#"<t:sequence-decls><t:sequence-decl t:name="x" t:display-outline-level="1" t:separation-character=".."/></t:sequence-decls>"#,
        r#"<t:user-field-decls><t:user-field-decl t:name="x" o:value-type="boolean" o:boolean-value="1"/></t:user-field-decls>"#,
        r#"<t:user-field-decls><t:user-field-decl t:name="x" o:value-type="float" o:value="not-number"/></t:user-field-decls>"#,
        r#"<t:user-field-decls><t:user-field-decl t:name="x" o:value-type="boolean" o:boolean-value="true" o:value="1"/></t:user-field-decls>"#,
        r#"<t:user-field-decls><t:user-field-decl t:name="x" o:string-value="hidden"/></t:user-field-decls>"#,
        r#"<t:user-field-decls><t:user-field-decl t:name="x" o:value-type="string" o:string-value="x">content</t:user-field-decl></t:user-field-decls>"#,
        r#"<t:variable-decls><t:variable-decl t:name="x" o:value-type="float"/></t:variable-decls><t:variable-decls/>"#,
        r#"<t:variable-get t:name="late"/><t:variable-decls><t:variable-decl t:name="late" o:value-type="float"/></t:variable-decls>"#,
        r#"<t:user-field-get t:name="missing"/>"#,
    ] {
        let document = FlatDocument::from_bytes(flat(inner)).unwrap();
        assert!(document.variable_declarations().is_err(), "{inner}");
    }
    let spoofed = format!(r#"<o:document xmlns:o="{OFFICE}" xmlns:t="urn:not-text" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><t:variable-decls/></o:text></o:body></o:document>"#).into_bytes();
    assert!(
        FlatDocument::from_bytes(spoofed)
            .unwrap()
            .variable_declarations()
            .is_err()
    );
}

#[test]
fn packaged_content_and_styles_are_exposed_by_current_odt_facades() {
    let declaration = r#"<t:variable-decls><t:variable-decl t:name="body" o:value-type="float"/></t:variable-decls>"#;
    let styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" o:version="1.3"><o:styles/><o:automatic-styles/><o:master-styles><s:master-page s:name="Standard"><s:header><t:user-field-decls><t:user-field-decl t:name="header" o:value-type="string" o:string-value="safe"/></t:user-field-decls><t:p><t:user-field-get t:name="header">safe</t:user-field-get></t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#
    );
    let text_content = content(&format!(
        r#"<o:text>{declaration}<t:p><t:variable-get t:name="body">1</t:variable-get></t:p></o:text>"#
    ));
    let bytes = package(&text_content, Some(&styles));
    let generic = Package::from_bytes(bytes.clone()).unwrap();
    let declarations = generic.variable_declarations().unwrap();
    assert_eq!(declarations.groups.len(), 2);
    assert!(declarations.find(Kind::User, "header").is_some());
    assert_eq!(
        Document::from_bytes(bytes)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        2
    );
}

#[test]
fn parses_compact_flat_text_oracles() {
    for (inner, kind, name) in [
        (
            r#"<t:variable-decls><t:variable-decl t:name="Counter" o:value-type="float"/></t:variable-decls><t:p><t:variable-get t:name="Counter">1</t:variable-get></t:p>"#,
            Kind::Simple,
            "Counter",
        ),
        (
            r#"<t:user-field-decls><t:user-field-decl t:name="Language" o:value-type="string" o:string-value="en-US"/></t:user-field-decls><t:p><t:user-field-get t:name="Language">en-US</t:user-field-get></t:p>"#,
            Kind::User,
            "Language",
        ),
        (
            r#"<t:sequence-decls><t:sequence-decl t:name="Illustration" t:display-outline-level="1"/></t:sequence-decls><t:p><t:sequence t:name="Illustration" t:formula="ooow:Illustration+1">1</t:sequence></t:p>"#,
            Kind::Sequence,
            "Illustration",
        ),
    ] {
        let declarations = FlatDocument::from_bytes(flat(inner))
            .unwrap()
            .variable_declarations()
            .unwrap();
        assert!(declarations.find(kind, name).is_some());
    }
}

#[test]
fn parses_compact_odt_package_oracles() {
    for (body, kind, name) in [
        (
            r#"<o:text><t:user-field-decls><t:user-field-decl t:name="Status" o:value-type="string" o:string-value="ready"/></t:user-field-decls><t:p><t:user-field-get t:name="Status">ready</t:user-field-get></t:p></o:text>"#,
            Kind::User,
            "Status",
        ),
        (
            r#"<o:text><t:sequence-decls><t:sequence-decl t:name="Illustration" t:display-outline-level="2"/></t:sequence-decls><t:p><t:sequence t:name="Illustration" t:formula="ooow:Illustration+1">1</t:sequence></t:p></o:text>"#,
            Kind::Sequence,
            "Illustration",
        ),
    ] {
        let xml = content(body);
        let bytes = package(&xml, None);
        let generic = Package::from_bytes(bytes.clone()).unwrap();
        assert!(
            generic
                .variable_declarations()
                .unwrap()
                .find(kind, name)
                .is_some()
        );
        let document = Document::from_bytes(bytes).unwrap();
        assert!(
            document
                .variable_declarations()
                .unwrap()
                .find(kind, name)
                .is_some()
        );
    }
}

#[test]
fn enforces_name_count_depth_and_aggregate_limits() {
    let oversized_name = "n".repeat(65_537);
    let document = FlatDocument::from_bytes(flat(&format!(r#"<t:variable-decls><t:variable-decl t:name="{oversized_name}" o:value-type="float"/></t:variable-decls>"#))).unwrap();
    assert!(document.variable_declarations().is_err());

    let mut many = String::from("<t:variable-decls>");
    for index in 0..=65_536 {
        many.push_str(&format!(
            r#"<t:variable-decl t:name="v{index}" o:value-type="float"/>"#
        ));
    }
    many.push_str("</t:variable-decls>");
    let document = FlatDocument::from_bytes(flat(&many)).unwrap();
    assert!(document.variable_declarations().is_err());

    let deep = format!("{}{}", "<t:span>".repeat(260), "</t:span>".repeat(260));
    let document = FlatDocument::from_bytes(flat(&deep)).unwrap();
    assert!(document.variable_declarations().is_err());

    let chunk = "x".repeat(1_000_000);
    let mut aggregate = String::from("<t:user-field-decls>");
    for index in 0..17 {
        aggregate.push_str(&format!(r#"<t:user-field-decl t:name="u{index}" o:value-type="string" o:string-value="{chunk}"/>"#));
    }
    aggregate.push_str("</t:user-field-decls>");
    let document = FlatDocument::from_bytes(flat(&aggregate)).unwrap();
    assert!(document.variable_declarations().is_err());
}

fn content(body: &str) -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" o:version="1.3"><o:body>{body}</o:body></o:document-content>"#
    )
}

fn flat(inner: &str) -> Vec<u8> {
    format!(r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text>{inner}</o:text></o:body></o:document>"#).into_bytes()
}

fn package(content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}
