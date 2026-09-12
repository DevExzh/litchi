use litchi_core::{Error, Result};
use litchi_ods::{
    Spreadsheet,
    document::{NumberStyleNode, Snapshot, StyleGraph},
};

// Tracked LibreOffice source; content.xml owns scientific number style N81.
const FORMATS: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sc/qa/unit/data/ods/formats.ods");

#[test]
fn native_scientific_style_cannot_be_discarded_by_legacy_decimal_replacement() -> Result<()> {
    let original = Spreadsheet::from_bytes(FORMATS.to_vec())?;
    assert!(original.content_xml().contains("number:scientific-number"));
    let snapshot = Snapshot::from_bytes(FORMATS.to_vec())?;
    let mut edit = snapshot.edit();
    let graph = StyleGraph {
        number_styles: vec![NumberStyleNode {
            name: "N81".to_owned(),
            decimal_places: 3,
            min_integer_digits: 1,
            prefix: None,
            suffix: None,
        }],
        ..StyleGraph::default()
    };
    let error = edit
        .replace_style_graph(&graph)
        .expect_err("a scientific body is outside the legacy decimal replacement envelope");
    assert!(matches!(error, Error::Unsupported(_)), "{error}");
    assert_eq!(edit.as_bytes(), FORMATS);
    let commit = edit.commit()?;
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().as_bytes(), FORMATS);
    Ok(())
}

#[test]
fn unresolved_entity_in_style_text_never_becomes_a_rewrite_target() -> Result<()> {
    let xml = br#"<office:document-content
        xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
        xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
        office:version="1.3">
        <office:automatic-styles>
          <number:number-style style:name="Unresolved">
            <number:text>&future;</number:text>
            <number:number number:decimal-places="2" number:min-integer-digits="1"/>
          </number:number-style>
        </office:automatic-styles>
        <office:body><office:spreadsheet/></office:body>
        </office:document-content>"#;
    let known = std::str::from_utf8(xml)
        .expect("the fixture is UTF-8")
        .replace("&future;", "&amp;");
    Snapshot::from_bytes(support::raw_package(&[(
        "content.xml",
        known.as_bytes(),
        "text/xml",
    )]))?;
    let bytes = support::raw_package(&[("content.xml", xml, "text/xml")]);
    let Ok(snapshot) = Snapshot::from_bytes(bytes.clone()) else {
        return Ok(());
    };
    let mut edit = snapshot.edit();
    let graph = StyleGraph {
        number_styles: vec![NumberStyleNode {
            name: "Unresolved".to_owned(),
            decimal_places: 3,
            min_integer_digits: 1,
            prefix: None,
            suffix: None,
        }],
        ..StyleGraph::default()
    };
    assert!(edit.replace_style_graph(&graph).is_err());
    assert_eq!(edit.as_bytes(), bytes);
    Ok(())
}
mod support;
