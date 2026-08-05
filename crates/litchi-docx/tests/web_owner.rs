use litchi_docx::{
    Package,
    web::{Conformance, Settings},
};
use std::io::Cursor;
use std::path::Path;

fn contains(xml: &[u8], value: &[u8]) -> bool {
    xml.windows(value.len()).any(|window| window == value)
}

#[test]
fn writes_both_namespace_families_deterministically() {
    let settings = Settings::default();
    let transitional = settings.xml(Conformance::Transitional).unwrap();
    let strict = settings.xml(Conformance::Strict).unwrap();
    assert!(contains(
        &transitional,
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ));
    assert!(contains(
        &strict,
        b"http://purl.oclc.org/ooxml/wordprocessingml/main"
    ));
    assert!(contains(
        &strict,
        b"http://purl.oclc.org/ooxml/officeDocument/relationships"
    ));
    assert_eq!(settings.xml(Conformance::Strict).unwrap(), strict);
}

#[test]
fn opens_real_poi_div_metadata_and_round_trips_a_package() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/heading123.docx");
    let package = Package::open(fixture).unwrap();
    let (settings, conformance) = package.document().unwrap().web().unwrap().unwrap();
    let xml = settings.xml(conformance).unwrap();
    assert!(contains(&xml, b"<w:divs>"));
    assert!(contains(&xml, b"<w:allowPNG/>"));
    assert!(contains(&xml, b"<w:doNotSaveAsSingleFile/>"));

    let mut synthetic = Package::new().unwrap();
    let (mut settings, conformance) = synthetic.web().unwrap().unwrap();
    settings.set_encoding("utf-8").unwrap();
    assert!(synthetic.put_web(settings, conformance).unwrap());
    let mut output = Cursor::new(Vec::new());
    synthetic.to_stream(&mut output).unwrap();
    let reopened = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let (settings, conformance) = reopened.document().unwrap().web().unwrap().unwrap();
    let serialized = settings.xml(conformance).unwrap();
    assert!(contains(&serialized, br#"<w:encoding w:val="utf-8"/>"#));
}
