use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_ooxml_common::Error as CommonError;
use litchi_ooxml_common::web::raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};
use litchi_ooxml_common::web::{Conformance, Store};
use litchi_opc::part::{Part, XmlPart};

const TASK_PANES_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
const WEB_EXTENSION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");

#[test]
fn presentation_loads_inert_task_panes_through_semantic_views() {
    let package = package_with_task_panes(false);

    let panes = package
        .presentation()
        .unwrap()
        .task_panes()
        .unwrap()
        .unwrap();
    assert_eq!(panes.len(), 1);
    let pane = panes.iter().next().unwrap();
    assert!(pane.visible());
    assert_eq!(pane.add_in().reference().store(), Store::Omex);
}

#[test]
fn package_moves_and_removes_a_validated_task_pane_graph() {
    let source = package_with_task_panes(false);
    let panes = source.task_panes().unwrap().unwrap();
    let mut package = Package::new().unwrap();

    package
        .put_task_panes(panes, Conformance::Transitional)
        .unwrap();
    let stored = package.task_panes().unwrap().unwrap();
    assert_eq!(stored.len(), 1);
    assert_eq!(
        stored.iter().next().unwrap().add_in().reference().store(),
        Store::Omex
    );

    assert!(package.remove_task_panes().unwrap());
    assert!(package.task_panes().unwrap().is_none());
    assert!(!package.remove_task_panes().unwrap());
}

#[test]
fn presentation_rejects_external_web_extension_relationships() {
    let package = package_with_task_panes(true);

    assert!(matches!(
        package.presentation().unwrap().task_panes(),
        Err(OoxmlError::Common(CommonError::Relationship(message)))
            if message.contains("internal")
    ));
}

fn package_with_task_panes(external_extension: bool) -> Package {
    let mut package = Package::new().unwrap();
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

    package
        .edit_opc(|opc| {
            opc.relationships_mut().add_relationship(
                TASK_PANES_RELATIONSHIP.to_owned(),
                "ppt/webextensions/taskpanes.xml".to_owned(),
                "rIdTaskPanes".to_owned(),
                false,
            );
            opc.add_part(Box::new(task_panes));
            if !external_extension {
                opc.add_part(Box::new(XmlPart::new(
                    extension_name,
                    ADD_IN_CONTENT_TYPE.to_owned(),
                    WEB_EXTENSION_XML.to_vec(),
                )));
            }
            Ok(())
        })
        .unwrap();
    package
}
