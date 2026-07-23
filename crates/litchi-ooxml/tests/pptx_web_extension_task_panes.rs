use litchi_ooxml::pptx::Package;
use litchi_ooxml::web_extensions::{
    TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP_TYPE, WEB_EXTENSION_CONTENT_TYPE,
    WEB_EXTENSION_RELATIONSHIP_TYPE, WebExtensionStoreType,
};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::part::{Part, XmlPart};

const TASK_PANES_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
const WEB_EXTENSION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");

#[test]
fn presentation_loads_inert_web_extension_task_panes() {
    let package = package_with_task_panes(false);

    let panes = package
        .presentation()
        .unwrap()
        .web_extension_task_panes()
        .unwrap()
        .unwrap();
    assert_eq!(panes.panes.len(), 1);
    assert!(panes.panes[0].visible);
    assert_eq!(panes.panes[0].relationship_id, "rId1");
    assert_eq!(
        panes.panes[0].web_extension.reference.store_type,
        WebExtensionStoreType::Omex
    );
}

#[test]
fn presentation_rejects_external_web_extension_relationships() {
    let mut package = package_with_task_panes(true);
    let task_panes_name = PackURI::new("/ppt/webextensions/taskpanes.xml").unwrap();
    let task_panes = package
        .opc_package_mut()
        .get_part_mut(&task_panes_name)
        .unwrap();
    task_panes.rels_mut().remove("rId1");
    task_panes.rels_mut().add_relationship(
        WEB_EXTENSION_RELATIONSHIP_TYPE.to_string(),
        "https://example.invalid/add-in".to_string(),
        "rId1".to_string(),
        true,
    );

    assert!(matches!(
        package
            .presentation()
            .unwrap()
            .web_extension_task_panes(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("must be internal")
    ));
}

fn package_with_task_panes(external_extension: bool) -> Package {
    let mut package = Package::new().unwrap();
    let task_panes_name = PackURI::new("/ppt/webextensions/taskpanes.xml").unwrap();
    let extension_name = PackURI::new("/ppt/webextensions/webextension1.xml").unwrap();
    let mut task_panes = XmlPart::new(
        task_panes_name,
        TASK_PANES_CONTENT_TYPE.to_string(),
        TASK_PANES_XML.to_vec(),
    );
    task_panes.rels_mut().add_relationship(
        WEB_EXTENSION_RELATIONSHIP_TYPE.to_string(),
        if external_extension {
            "https://example.invalid/add-in"
        } else {
            "webextension1.xml"
        }
        .to_string(),
        "rId1".to_string(),
        external_extension,
    );

    let opc = package.opc_package_mut();
    opc.relationships_mut().add_relationship(
        TASK_PANES_RELATIONSHIP_TYPE.to_string(),
        "ppt/webextensions/taskpanes.xml".to_string(),
        "rIdTaskPanes".to_string(),
        false,
    );
    opc.add_part(Box::new(task_panes));
    if !external_extension {
        opc.add_part(Box::new(XmlPart::new(
            extension_name,
            WEB_EXTENSION_CONTENT_TYPE.to_string(),
            WEB_EXTENSION_XML.to_vec(),
        )));
    }
    package
}
