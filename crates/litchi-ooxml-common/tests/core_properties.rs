use litchi_ooxml_common::{Props, properties};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
const STRICT_NS: &str = "http://purl.oclc.org/ooxml/package/metadata/core-properties";
const CORE_PATH: &str = "/custom/MetaData.XML";

fn strict_package() -> (OpcPackage, Vec<u8>) {
    let mut package = OpcPackage::new();
    let xml = format!(
        "<?xml version=\"1.0\"?>\n<cp:coreProperties xmlns:cp=\"{STRICT_NS}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n  <dc:title>Original</dc:title>\n</cp:coreProperties>"
    )
    .into_bytes();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new(CORE_PATH).expect("test URI"),
        ct::OPC_CORE_PROPERTIES.to_owned(),
        xml.clone(),
    )));
    package.relate_to(CORE_PATH.trim_start_matches('/'), STRICT_REL);
    (package, xml)
}

#[test]
fn strict_core_properties_preserve_the_selected_graph() {
    let (mut package, _) = strict_package();
    assert_eq!(
        properties::read(&package)
            .unwrap()
            .unwrap()
            .title
            .as_deref(),
        Some("Original")
    );

    properties::write(&mut package, Props::new().title("Changed")).unwrap();
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == STRICT_REL)
        .expect("strict core relationship");
    assert_eq!(relationship.target_ref(), CORE_PATH.trim_start_matches('/'));
    let target = relationship.target_partname().expect("internal target");
    let part = package.get_part(&target).unwrap();
    assert_eq!(part.partname().as_str(), CORE_PATH);
    let xml = std::str::from_utf8(part.blob()).unwrap();
    assert!(xml.contains(STRICT_NS));
    assert!(xml.contains("<dc:title>Changed</dc:title>"));
    assert!(
        !package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::CORE_PROPERTIES)
    );

    assert!(properties::clear(&mut package).unwrap());
    assert!(properties::read(&package).unwrap().is_none());
    assert!(
        !package
            .iter_parts()
            .any(|part| part.content_type() == ct::OPC_CORE_PROPERTIES)
    );
}

#[test]
fn absent_core_properties_are_idempotent() {
    let mut package = OpcPackage::new();
    assert!(properties::read(&package).unwrap().is_none());
    assert!(!properties::clear(&mut package).unwrap());
    assert!(properties::write(&mut package, Props::new().title("Created")).unwrap());
    assert_eq!(
        properties::read(&package)
            .unwrap()
            .unwrap()
            .title
            .as_deref(),
        Some("Created")
    );
}
