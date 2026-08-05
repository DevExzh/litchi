use litchi_ooxml_common::Error as CommonError;
use litchi_ooxml_common::web::{Conformance, Store, load, put, remove};
use litchi_ooxml_common::web::raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};
use litchi_opc::part::{Part, XmlPart};
use litchi_opc::{OpcPackage, PackURI};

const TASK_PANES_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
const WEB_EXTENSION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");

#[test]
fn common_web_owner_loads_inert_task_panes() {
    let package = package_with_task_panes(false);

    let panes = load(&package).unwrap().unwrap();
    assert_eq!(panes.len(), 1);
    let pane = panes.iter().next().unwrap();
    assert!(pane.visible());
    assert_eq!(pane.add_in().reference().store(), Store::Omex);
}

#[test]
fn common_web_owner_moves_and_removes_a_validated_task_pane_graph() {
    let source = package_with_task_panes(false);
    let panes = load(&source).unwrap().unwrap();
    let mut package = OpcPackage::new();

    put(&mut package, panes, Conformance::Transitional).unwrap();
    let stored = load(&package).unwrap().unwrap();
    assert_eq!(stored.len(), 1);
    assert_eq!(
        stored.iter().next().unwrap().add_in().reference().store(),
        Store::Omex
    );

    assert!(remove(&mut package).unwrap());
    assert!(load(&package).unwrap().is_none());
    assert!(!remove(&mut package).unwrap());
}

#[test]
fn common_web_owner_rejects_external_web_extension_relationships() {
    let package = package_with_task_panes(true);

    assert!(matches!(
        load(&package),
        Err(CommonError::Relationship(message)) if message.contains("internal")
    ));
}

fn package_with_task_panes(external_extension: bool) -> OpcPackage {
    let task_panes_name = PackURI::new("/ppt/webextensions/taskpanes.xml").unwrap();
    let extension_name = PackURI::new("/ppt/webextensions/webextension1.xml").unwrap();
    let mut task_panes = XmlPart::new(
        task_panes_name,
        TASK_PANES_CONTENT_TYPE.to_owned(),
        TASK_PANES_XML.to_vec(),
    );
    task_panes.rels_mut().add_relationship(
        ADD_IN_RELATIONSHIP.to_owned(),
        if external_extension {
            "https://example.invalid/add-in"
        } else {
            "webextension1.xml"
        }
        .to_owned(),
        "rId1".to_owned(),
        external_extension,
    );

    let mut package = OpcPackage::new();
    package.relationships_mut().add_relationship(
        TASK_PANES_RELATIONSHIP.to_owned(),
        "ppt/webextensions/taskpanes.xml".to_owned(),
        "rIdTaskPanes".to_owned(),
        false,
    );
    package.add_part(Box::new(task_panes));
    if !external_extension {
        package.add_part(Box::new(XmlPart::new(
            extension_name,
            ADD_IN_CONTENT_TYPE.to_owned(),
            WEB_EXTENSION_XML.to_vec(),
        )));
    }
    package
}
