use litchi_ooxml::docx::{Package, WebSettings, WebSettingsConformance};
use std::io::Cursor;
use std::path::Path;

#[test]
fn writes_both_namespace_families_deterministically() {
    let settings = WebSettings::default();
    let transitional = settings.to_xml().unwrap();
    let strict = settings
        .to_xml_with_conformance(WebSettingsConformance::Strict)
        .unwrap();
    assert!(transitional.contains("http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
    assert!(strict.contains("http://purl.oclc.org/ooxml/wordprocessingml/main"));
    assert!(strict.contains("http://purl.oclc.org/ooxml/officeDocument/relationships"));
    assert_eq!(
        settings
            .to_xml_with_conformance(WebSettingsConformance::Strict)
            .unwrap(),
        strict
    );
}

#[test]
fn opens_real_poi_div_metadata_and_round_trips_a_package() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..//test-data/poi/test-data/document/heading123.docx");
    let package = Package::open(fixture).unwrap();
    let web_settings = package.document().unwrap().web_settings().unwrap().unwrap();
    let xml = web_settings.to_xml().unwrap();
    assert!(xml.contains("<w:divs>"));
    assert!(xml.contains("<w:allowPNG/>"));
    assert!(xml.contains("<w:doNotSaveAsSingleFile/>"));

    let mut synthetic = Package::new().unwrap();
    synthetic.web_settings_mut().unwrap().set_encoding("utf-8");
    let mut output = Cursor::new(Vec::new());
    synthetic.to_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let serialized = reopened
        .document()
        .unwrap()
        .web_settings()
        .unwrap()
        .unwrap()
        .to_xml()
        .unwrap();
    assert!(serialized.contains(r#"<w:encoding w:val="utf-8"/>"#));
}
