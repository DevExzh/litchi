#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::OpcPackage;
use litchi_pptx::{
    Error, Package,
    presentation_properties::metadata::{changes, revision},
};

const LOCAL_REVISION_INFORMATION: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/revision-information/basic_revision.xml");
const LOCAL_CHANGES_INFORMATION: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/changes-information/basic_changes.xml");

#[test]
fn package_and_presentation_manage_local_revision_and_changes_information() {
    let mut opc = base_opc();
    assert_eq!(revision::load(&opc).unwrap(), None);
    assert_eq!(changes::load(&opc).unwrap(), None);

    let revision_part = revision::Part {
        relationship_id: "rIdRevisionInformation".to_string(),
        part_name: "/ppt/revisionInfo.xml".to_string(),
        revision_information: revision::Info::parse(LOCAL_REVISION_INFORMATION).unwrap(),
    };
    let changes_part = changes::Part {
        relationship_id: "rIdChangesInformation".to_string(),
        part_name: "/ppt/changesInfo.xml".to_string(),
        changes_information: changes::Info::parse(LOCAL_CHANGES_INFORMATION).unwrap(),
    };

    revision::store(&mut opc, &revision_part).unwrap();
    changes::store(&mut opc, &changes_part).unwrap();
    assert!(matches!(
        revision::store(&mut opc, &revision_part),
        Err(Error::Invalid(_))
    ));
    assert!(matches!(
        changes::store(&mut opc, &changes_part),
        Err(Error::Invalid(_))
    ));

    let package = Package::from_opc_package(opc).unwrap();
    let opc = package.opc().unwrap();
    assert_eq!(revision::load(opc).unwrap(), Some(revision_part.clone()));
    assert_eq!(changes::load(opc).unwrap(), Some(changes_part.clone()));

    // The package and presentation facades both expose the same borrowed OPC
    // owner; direct metadata loaders are the canonical standalone surface.
    assert_eq!(
        revision::load(package.presentation().unwrap().package()).unwrap(),
        Some(revision_part)
    );
    assert_eq!(
        changes::load(package.presentation().unwrap().package()).unwrap(),
        Some(changes_part)
    );
}

fn base_opc() -> OpcPackage {
    let mut package = Package::new().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    OpcPackage::from_bytes(&package_bytes).unwrap()
}
