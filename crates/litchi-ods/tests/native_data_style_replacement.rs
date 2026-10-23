use litchi_core::Error as CoreError;
use litchi_odf_common::Error;
use litchi_ods::{
    Package,
    document::{NumberStyleNode, StyleGraph},
};

// Tracked LibreOffice source; content.xml owns scientific number style N81.
const FORMATS: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sc/qa/unit/data/ods/formats.ods");

#[test]
fn native_scientific_style_cannot_be_discarded_by_legacy_decimal_replacement() -> Result<(), Error>
{
    let original = Package::from_bytes(FORMATS)?.workbook()?;
    assert!(original.content_xml().contains("number:scientific-number"));
    let snapshot = Package::from_bytes(FORMATS)?;
    let mut edit = snapshot.edit()?;
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
    assert!(
        matches!(
            error,
            Error::Core(CoreError::Unsupported(_)) | Error::Refused(_)
        ),
        "{error}"
    );
    assert_eq!(edit.as_bytes(), FORMATS);
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().as_bytes(), FORMATS);
    Ok(())
}

#[test]
fn unresolved_entity_in_style_text_never_becomes_a_rewrite_target() -> Result<(), Error> {
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
    Package::from_bytes(&support::raw_package(&[(
        "content.xml",
        known.as_bytes(),
        "text/xml",
    )]))?;
    let bytes = support::raw_package(&[("content.xml", xml, "text/xml")]);
    let Ok(snapshot) = Package::from_bytes(&bytes) else {
        return Ok(());
    };
    let mut edit = snapshot.edit()?;
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
