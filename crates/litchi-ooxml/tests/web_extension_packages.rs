use litchi_ooxml::web_extensions::{
    OoxmlConformance, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP_TYPE,
    WEB_EXTENSION_CONTENT_TYPE, WEB_EXTENSION_RELATIONSHIP_TYPE, WebExtension,
    WebExtensionStoreReference, WebExtensionStoreType, WebExtensionTaskPane, WebExtensionTaskPanes,
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

#[test]
fn package_wrappers_author_and_remove_inert_task_panes() {
    let authored = authored_task_panes();
    let directory = tempfile::tempdir().unwrap();

    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    docx.set_web_extension_task_panes(&authored, OoxmlConformance::Transitional)
        .unwrap();
    assert_eq!(
        docx.web_extension_task_panes().unwrap(),
        Some(authored.clone())
    );
    let docx_path = directory.path().join("task-panes.docx");
    docx.save(&docx_path).unwrap();
    assert_eq!(
        litchi_ooxml::docx::Package::open(&docx_path)
            .unwrap()
            .web_extension_task_panes()
            .unwrap(),
        Some(authored.clone())
    );
    assert!(docx.remove_web_extension_task_panes().unwrap());

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    xlsx.set_web_extension_task_panes(&authored, OoxmlConformance::Strict)
        .unwrap();
    assert_eq!(
        xlsx.web_extension_task_panes().unwrap(),
        Some(authored.clone())
    );
    let xlsx_path = directory.path().join("task-panes.xlsx");
    xlsx.save(&xlsx_path).unwrap();
    assert_eq!(
        litchi_ooxml::xlsx::Workbook::open(&xlsx_path)
            .unwrap()
            .web_extension_task_panes()
            .unwrap(),
        Some(authored.clone())
    );
    assert!(xlsx.remove_web_extension_task_panes().unwrap());

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    pptx.set_web_extension_task_panes(&authored, OoxmlConformance::Transitional)
        .unwrap();
    assert_eq!(
        pptx.web_extension_task_panes().unwrap(),
        Some(authored.clone())
    );
    let pptx_path = directory.path().join("task-panes.pptx");
    pptx.save(&pptx_path).unwrap();
    assert_eq!(
        litchi_ooxml::pptx::Package::open(&pptx_path)
            .unwrap()
            .web_extension_task_panes()
            .unwrap(),
        Some(authored)
    );
    assert!(pptx.remove_web_extension_task_panes().unwrap());
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

fn authored_task_panes() -> WebExtensionTaskPanes {
    WebExtensionTaskPanes {
        panes: vec![WebExtensionTaskPane {
            dock_state: "right".into(),
            visible: true,
            width: 360.0,
            row: 0,
            locked: false,
            relationship_id: "rId1".into(),
            web_extension: WebExtension {
                id: "{10000000-0000-0000-0000-000000000001}".into(),
                frozen: true,
                reference: WebExtensionStoreReference {
                    id: "inert-test-add-in".into(),
                    version: "1.0".into(),
                    store: Some("en-US".into()),
                    store_type: WebExtensionStoreType::Omex,
                    extension_list: None,
                },
                alternate_references: vec![],
                properties: vec![],
                bindings: vec![],
                snapshot: None,
                extension_list: None,
            },
            snapshot_resources: vec![],
            extension_list: None,
        }],
    }
}
