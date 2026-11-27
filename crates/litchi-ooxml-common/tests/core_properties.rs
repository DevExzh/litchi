use litchi_core::Metadata;
use litchi_ooxml_common::custom::Host;
use litchi_ooxml_common::properties;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

/// Read core properties, asserting the conformant fixtures need no repair.
fn read_clean(package: &OpcPackage) -> litchi_ooxml_common::Result<Option<Metadata>> {
    let mut report = litchi_ooxml_common::leniency::ToleranceReport::default();
    let props = properties::read_with_report(package, &mut report)?;
    assert!(report.is_clean(), "{report:?}");
    Ok(props)
}

fn titled(title: &str) -> Metadata {
    Metadata {
        title: Some(title.to_owned()),
        ..Metadata::default()
    }
}

fn publish_core(package: &mut OpcPackage, value: Option<Metadata>) -> bool {
    let mut edit = properties::Edit::new(package, Host::Word).expect("metadata edit");
    match value {
        Some(value) => {
            edit.set_metadata(value);
        },
        None => {
            edit.clear_metadata();
        },
    }
    let commit = edit.commit().expect("metadata commit");
    let changed = commit.diagnostics().changed();
    *package = commit.into_package();
    changed
}

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
        read_clean(&package).unwrap().unwrap().title.as_deref(),
        Some("Original")
    );

    assert!(publish_core(&mut package, Some(titled("Changed"))));
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

    assert!(publish_core(&mut package, None));
    assert!(read_clean(&package).unwrap().is_none());
    assert!(
        !package
            .iter_parts()
            .any(|part| part.content_type() == ct::OPC_CORE_PROPERTIES)
    );
}

#[test]
fn absent_core_properties_are_idempotent() {
    let mut package = OpcPackage::new();
    assert!(read_clean(&package).unwrap().is_none());
    assert!(!publish_core(&mut package, None));
    assert!(publish_core(&mut package, Some(titled("Created"))));
    assert_eq!(
        read_clean(&package).unwrap().unwrap().title.as_deref(),
        Some("Created")
    );
}

#[test]
fn retained_metadata_graph_matches_eager_repairs_and_refusals() {
    use litchi_opc::{PackageWriter, SourceBackedPackage, SourceTopologyPlan};
    for (relationship, content_type, external) in [
        (STRICT_REL, ct::OPC_CORE_PROPERTIES, false),
        (rt::CORE_PROPERTIES, ct::OPC_CORE_PROPERTIES, false),
        (STRICT_REL, "application/xml", false),
        (STRICT_REL, ct::OPC_CORE_PROPERTIES, true),
    ] {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(PackURI::new(CORE_PATH).unwrap(), content_type.to_owned(), format!("<cp:coreProperties xmlns:cp=\"{STRICT_NS}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>first</dc:title><dc:title>duplicate</dc:title></cp:coreProperties>").into_bytes())));
        if external {
            package.relate_to_external("https://example.invalid/core.xml", relationship);
        } else {
            package.relate_to(CORE_PATH.trim_start_matches('/'), relationship);
        }
        let physical = SourceBackedPackage::from_vec(
            PackageWriter::to_bytes(&package).unwrap(),
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            None,
        )
        .unwrap();
        let prepared = physical
            .prepare_topology(SourceTopologyPlan::new())
            .unwrap();
        let mut expected_report = litchi_ooxml_common::leniency::ToleranceReport::default();
        let expected = properties::read_with_report(&package, &mut expected_report);
        let mut actual_report = litchi_ooxml_common::leniency::ToleranceReport::default();
        let actual = properties::read_prepared_with_report(&prepared, &mut actual_report);
        match (expected, actual) {
            (Ok(expected), Ok(actual)) => {
                assert_eq!(actual, expected);
                assert_eq!(actual_report, expected_report);
            },
            (Err(expected), Err(actual)) => assert_eq!(actual.to_string(), expected.to_string()),
            (expected, actual) => panic!("graph parity differs: {expected:?} / {actual:?}"),
        }
    }
}
