use litchi_odt::{
    ChartDocument, Document, DrawingDocument, FlatOpenDocument, MasterDocument,
    OdfVariableDeclaration, OdfVariableKind, OdfVariableValue, OpenDocumentPackage, Presentation,
    Spreadsheet,
};
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn flat(inner: &str) -> Vec<u8> {
    format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text>{inner}</o:text></o:body></o:document>"#
    )
    .into_bytes()
}

fn content(body: &str) -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" o:version="1.3"><o:body>{body}</o:body></o:document-content>"#
    )
}

fn package(mimetype: &str, content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    if let Some(styles) = styles {
        zip.start_file("styles.xml", deflated).unwrap();
        zip.write_all(styles.as_bytes()).unwrap();
    }
    let manifest = format!(
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{mimetype}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/>{}</m:manifest>"#,
        styles.map_or(
            "",
            |_| r#"<m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/>"#
        )
    );
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
// The fixture value 3.1400 exercises lexical round-tripping; it is not the constant PI.
#[allow(clippy::approx_constant)]
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
    let declarations = FlatOpenDocument::from_bytes(xml)
        .unwrap()
        .variable_declarations()
        .unwrap();
    assert_eq!(declarations.groups.len(), 3);
    assert_eq!(declarations.declarations().count(), 12);
    let float = declarations.find(OdfVariableKind::User, "float").unwrap();
    assert!(matches!(
        float,
        OdfVariableDeclaration::User {
            value: Some(OdfVariableValue::Float { value, lexical }), ..
        } if *value == 3.14 && lexical == "3.1400"
    ));
    let string = declarations.find(OdfVariableKind::User, "string").unwrap();
    assert!(matches!(
        string,
        OdfVariableDeclaration::User { formula: Some(formula), .. }
            if formula.contains("WEBSERVICE")
    ));
    assert_eq!(
        declarations
            .find(OdfVariableKind::Sequence, "Figure")
            .unwrap()
            .effective_separation_character(),
        Some('#')
    );
    assert_eq!(
        declarations
            .find(OdfVariableKind::Sequence, "Table")
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
        let document = FlatOpenDocument::from_bytes(flat(inner)).unwrap();
        assert!(document.variable_declarations().is_err(), "{inner}");
    }
    let spoofed = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="urn:not-text" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><t:variable-decls/></o:text></o:body></o:document>"#
    )
    .into_bytes();
    assert!(
        FlatOpenDocument::from_bytes(spoofed)
            .unwrap()
            .variable_declarations()
            .is_err()
    );
}

#[test]
fn packaged_content_styles_and_every_specialized_facade_expose_declarations() {
    let declaration = r#"<t:variable-decls><t:variable-decl t:name="body" o:value-type="float"/></t:variable-decls>"#;
    let styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" o:version="1.3"><o:styles/><o:automatic-styles/><o:master-styles><s:master-page s:name="Standard"><s:header><t:user-field-decls><t:user-field-decl t:name="header" o:value-type="string" o:string-value="safe"/></t:user-field-decls><t:p><t:user-field-get t:name="header">safe</t:user-field-get></t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#
    );
    let text_content = content(&format!(
        r#"<o:text>{declaration}<t:p><t:variable-get t:name="body">1</t:variable-get></t:p></o:text>"#
    ));
    let bytes = package(
        "application/vnd.oasis.opendocument.text",
        &text_content,
        Some(&styles),
    );
    let generic = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    let declarations = generic.variable_declarations().unwrap();
    assert_eq!(declarations.groups.len(), 2);
    assert!(declarations.find(OdfVariableKind::User, "header").is_some());
    assert_eq!(
        Document::from_bytes(bytes)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        2
    );

    let families = [
        (
            "application/vnd.oasis.opendocument.spreadsheet",
            format!(r#"<o:spreadsheet>{declaration}<table:table table:name="S"/></o:spreadsheet>"#),
        ),
        (
            "application/vnd.oasis.opendocument.presentation",
            format!(r#"<o:presentation>{declaration}<d:page d:name="P"/></o:presentation>"#),
        ),
        (
            "application/vnd.oasis.opendocument.graphics",
            format!(r#"<o:drawing>{declaration}<d:page d:name="P"/></o:drawing>"#),
        ),
        (
            "application/vnd.oasis.opendocument.chart",
            format!(r#"<o:chart>{declaration}<c:chart><c:plot-area/></c:chart></o:chart>"#),
        ),
    ];
    let spreadsheet = package(families[0].0, &content(&families[0].1), None);
    assert_eq!(
        Spreadsheet::from_bytes(spreadsheet)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        1
    );
    let presentation = package(families[1].0, &content(&families[1].1), None);
    assert_eq!(
        Presentation::from_bytes(presentation)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        1
    );
    let drawing = package(families[2].0, &content(&families[2].1), None);
    assert_eq!(
        DrawingDocument::from_bytes(drawing)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        1
    );
    let chart = package(families[3].0, &content(&families[3].1), None);
    assert_eq!(
        ChartDocument::from_bytes(chart)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        1
    );
    let master = package(
        "application/vnd.oasis.opendocument.text-master",
        &text_content,
        None,
    );
    assert_eq!(
        MasterDocument::from_bytes(master)
            .unwrap()
            .variable_declarations()
            .unwrap()
            .declarations()
            .count(),
        1
    );
}

#[test]
fn parses_bundled_flat_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let fixtures = [
        "test-data/libreoffice-core/sw/qa/extras/layout/data/field_hide_section.fodt",
        "test-data/libreoffice-core/sw/qa/extras/layout/data/user-field-type-language.fodt",
        "test-data/libreoffice-core/sw/qa/extras/uiwriter/data/tdf91801.fodt",
        "test-data/libreoffice-core/xmloff/qa/unit/data/scale-width-redline.fodt",
        "test-data/odfdo/tests/samples/images.fodt",
    ];
    for fixture in fixtures {
        let path = root.join(fixture);
        if !path.exists() {
            continue;
        }
        let document = FlatOpenDocument::from_bytes(std::fs::read(path).unwrap()).unwrap();
        let declarations = document
            .variable_declarations()
            .unwrap_or_else(|error| panic!("failed declaration scan for {fixture}: {error}"));
        assert!(!declarations.groups.is_empty(), "{fixture}");
    }
}

#[test]
fn parses_bundled_odfpy_and_odfdo_package_oracles() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for fixture in [
        "test-data/odfpy/tests/examples/simpletable.odt",
        "test-data/odfdo/tests/samples/variable.odt",
    ] {
        let path = root.join(fixture);
        if !path.exists() {
            continue;
        }
        let declarations = Document::open(path)
            .unwrap()
            .variable_declarations()
            .unwrap_or_else(|error| panic!("failed declaration scan for {fixture}: {error}"));
        assert!(!declarations.groups.is_empty(), "{fixture}");
    }

    let example = root.join("test-data/odfdo/tests/samples/example.xml");
    if example.exists() {
        let xml = std::fs::read_to_string(example).unwrap();
        let bytes = package("application/vnd.oasis.opendocument.text", &xml, None);
        let declarations = OpenDocumentPackage::from_bytes(bytes)
            .unwrap()
            .variable_declarations()
            .unwrap();
        assert!(
            declarations
                .find(OdfVariableKind::Sequence, "Illustration")
                .is_some()
        );
    }
}

#[test]
fn enforces_name_count_depth_and_aggregate_limits() {
    let oversized_name = "n".repeat(65_537);
    let document = FlatOpenDocument::from_bytes(flat(&format!(
        r#"<t:variable-decls><t:variable-decl t:name="{oversized_name}" o:value-type="float"/></t:variable-decls>"#
    )))
    .unwrap();
    assert!(document.variable_declarations().is_err());

    let mut many = String::from("<t:variable-decls>");
    for index in 0..=65_536 {
        many.push_str(&format!(
            r#"<t:variable-decl t:name="v{index}" o:value-type="float"/>"#
        ));
    }
    many.push_str("</t:variable-decls>");
    let document = FlatOpenDocument::from_bytes(flat(&many)).unwrap();
    assert!(document.variable_declarations().is_err());

    let mut deep = String::new();
    for _ in 0..260 {
        deep.push_str("<t:span>");
    }
    for _ in 0..260 {
        deep.push_str("</t:span>");
    }
    let document = FlatOpenDocument::from_bytes(flat(&deep)).unwrap();
    assert!(document.variable_declarations().is_err());

    let chunk = "x".repeat(1_000_000);
    let mut aggregate = String::from("<t:user-field-decls>");
    for index in 0..17 {
        aggregate.push_str(&format!(
            r#"<t:user-field-decl t:name="u{index}" o:value-type="string" o:string-value="{chunk}"/>"#
        ));
    }
    aggregate.push_str("</t:user-field-decls>");
    let document = FlatOpenDocument::from_bytes(flat(&aggregate)).unwrap();
    assert!(document.variable_declarations().is_err());
}
