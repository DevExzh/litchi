use litchi_ooxml::web_extensions::{
    TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP_TYPE, WEB_EXTENSION_CONTENT_TYPE,
    WEB_EXTENSION_RELATIONSHIP_TYPE, WebExtensionStoreType,
};
use litchi_ooxml::{OpcPackage, PackURI};
use litchi_opc::{Part, XmlPart};

const LOCAL_EXTENSION: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");
const LOCAL_TASK_PANES: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");

#[test]
fn package_wrappers_discover_local_task_panes_without_activation() {
    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    install_task_panes(docx.opc_package_mut());
    assert_task_pane(docx.web_extension_task_panes().unwrap().unwrap());

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    install_task_panes(xlsx.opc_package_mut());
    assert_task_pane(xlsx.web_extension_task_panes().unwrap().unwrap());

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    install_task_panes(pptx.opc_package_mut());
    assert_task_pane(pptx.web_extension_task_panes().unwrap().unwrap());
}

fn install_task_panes(package: &mut OpcPackage) {
    package.rels_mut().add_relationship(
        TASK_PANES_RELATIONSHIP_TYPE.into(),
        "webextensions/taskpanes.xml".into(),
        "rIdTaskPanes".into(),
        false,
    );
    let mut task_panes = XmlPart::new(
        PackURI::new("/webextensions/taskpanes.xml").unwrap(),
        TASK_PANES_CONTENT_TYPE.into(),
        LOCAL_TASK_PANES.to_vec(),
    );
    task_panes.rels_mut().add_relationship(
        WEB_EXTENSION_RELATIONSHIP_TYPE.into(),
        "webextension1.xml".into(),
        "rId1".into(),
        false,
    );
    package.add_part(Box::new(task_panes));
    package.add_part(Box::new(XmlPart::new(
        PackURI::new("/webextensions/webextension1.xml").unwrap(),
        WEB_EXTENSION_CONTENT_TYPE.into(),
        LOCAL_EXTENSION.to_vec(),
    )));
}

fn assert_task_pane(task_panes: litchi_ooxml::web_extensions::WebExtensionTaskPanes) {
    assert_eq!(task_panes.panes.len(), 1);
    let pane = &task_panes.panes[0];
    assert_eq!(pane.dock_state, "right");
    assert!(pane.visible);
    assert_eq!(
        pane.web_extension.reference.store_type,
        WebExtensionStoreType::Omex
    );
    assert_eq!(pane.web_extension.reference.id, "local-omex");
}
