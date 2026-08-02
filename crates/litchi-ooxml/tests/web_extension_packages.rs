use litchi_ooxml::web::{
    AddIn, Binding, Compression, Conformance, Dock, Effect, EffectKind, ExtKind, ExtList, Pane,
    Panes, Property, Reference, Store,
    raw::{
        ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
    },
};
use litchi_ooxml::{OpcPackage, PackURI};
use litchi_opc::{Part, XmlPart};
use std::sync::Arc;

const LOCAL_EXTENSION: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/omex_webextension.xml");
const LOCAL_TASK_PANES: &[u8] =
    include_bytes!("../../../test-data/ooxml/web_extensions/visible_taskpanes.xml");
const SNAPSHOT_IMAGE: &[u8] = include_bytes!("../../../test-data/images/jpg/abstract1.jpg");

#[test]
fn package_wrappers_discover_local_task_panes_without_activation() {
    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    install_task_panes(docx.opc_package_mut());
    assert_task_pane(docx.task_panes().unwrap().unwrap());

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    install_task_panes(xlsx.opc_package_mut());
    assert_task_pane(xlsx.task_panes().unwrap().unwrap());

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    install_task_panes(pptx.opc_package_mut());
    assert_task_pane(pptx.task_panes().unwrap().unwrap());
}

#[test]
fn package_wrappers_author_and_remove_inert_task_panes() {
    let directory = tempfile::tempdir().unwrap();

    let mut docx = litchi_ooxml::docx::Package::new().unwrap();
    docx.put_task_panes(authored_task_panes(), Conformance::Transitional)
        .unwrap();
    assert_task_pane(docx.task_panes().unwrap().unwrap());
    let docx_path = directory.path().join("task-panes.docx");
    docx.save(&docx_path).unwrap();
    assert_task_pane(
        litchi_ooxml::docx::Package::open(&docx_path)
            .unwrap()
            .task_panes()
            .unwrap()
            .unwrap(),
    );
    assert!(docx.remove_task_panes().unwrap());

    let mut xlsx = litchi_ooxml::xlsx::Workbook::create().unwrap();
    xlsx.put_task_panes(authored_task_panes(), Conformance::Strict)
        .unwrap();
    assert_task_pane(xlsx.task_panes().unwrap().unwrap());
    let xlsx_path = directory.path().join("task-panes.xlsx");
    xlsx.save(&xlsx_path).unwrap();
    assert_task_pane(
        litchi_ooxml::xlsx::Workbook::open(&xlsx_path)
            .unwrap()
            .task_panes()
            .unwrap()
            .unwrap(),
    );
    assert!(xlsx.remove_task_panes().unwrap());

    let mut pptx = litchi_ooxml::pptx::Package::new().unwrap();
    pptx.put_task_panes(authored_task_panes(), Conformance::Transitional)
        .unwrap();
    assert_task_pane(pptx.task_panes().unwrap().unwrap());
    let pptx_path = directory.path().join("task-panes.pptx");
    pptx.save(&pptx_path).unwrap();
    assert_task_pane(
        litchi_ooxml::pptx::Package::open(&pptx_path)
            .unwrap()
            .task_panes()
            .unwrap()
            .unwrap(),
    );
    assert!(pptx.remove_task_panes().unwrap());
}

#[test]
fn public_web_facade_updates_and_round_trips_crud() {
    const ADD_IN_ID: &str = "{20000000-0000-0000-0000-000000000002}";

    let primary = Reference::new("primary", "1", Store::Omex)
        .unwrap()
        .with_ext(test_ext(ExtKind::AddIn, "primary-old"))
        .unwrap();
    let alternate = Reference::new("alternate", "1", Store::FileSystem)
        .unwrap()
        .with_ext(test_ext(ExtKind::AddIn, "alternate-old"))
        .unwrap();
    let binding = Binding::new("selection", "matrix", "Sheet1!A1")
        .unwrap()
        .with_ext(test_ext(ExtKind::AddIn, "binding-old"))
        .unwrap();
    let mut add_in = AddIn::new(ADD_IN_ID, primary).unwrap();
    add_in
        .push_reference(alternate)
        .unwrap()
        .push_property(Property::new("theme", "light").unwrap())
        .unwrap()
        .push_binding(binding)
        .unwrap()
        .set_ext(test_ext(ExtKind::AddIn, "add-in-old"))
        .unwrap();

    let mut pane = Pane::new(add_in);
    pane.set_compression(Some(Compression::Screen));
    pane.push_effect(test_effect("grayscl")).unwrap();
    pane.snapshot_mut()
        .set_ext(test_ext(ExtKind::DrawingMl, "snapshot-old"))
        .unwrap();
    pane.set_ext(test_ext(ExtKind::TaskPane, "pane-old"))
        .unwrap();

    let mut panes = Panes::new();
    panes.push(pane).unwrap();
    assert!(
        panes
            .edit(ADD_IN_ID, |pane| {
                pane.set_visible(false).set_row(7).set_locked(true);
                pane.set_width(512.5).unwrap().set_dock(Dock::Left).unwrap();

                assert_ext(pane.ext(), ExtKind::TaskPane, "pane-old");
                assert!(pane.clear_ext().is_some());
                pane.set_ext(test_ext(ExtKind::TaskPane, "pane-final"))
                    .unwrap();

                {
                    let add_in = pane.add_in_mut();
                    assert_ext(add_in.reference().ext(), ExtKind::AddIn, "primary-old");
                    assert!(add_in.reference_mut().clear_ext().is_some());
                    add_in
                        .set_reference(
                            Reference::new("primary-final", "2", Store::Registry)
                                .unwrap()
                                .with_ext(test_ext(ExtKind::AddIn, "primary-final"))
                                .unwrap(),
                        )
                        .unwrap();

                    let alternate = add_in.alternate_reference_mut("alternate").unwrap();
                    assert_ext(alternate.ext(), ExtKind::AddIn, "alternate-old");
                    assert!(alternate.clear_ext().is_some());
                    add_in
                        .upsert_reference(
                            Reference::new("alternate", "2", Store::Registry)
                                .unwrap()
                                .with_ext(test_ext(ExtKind::AddIn, "alternate-final"))
                                .unwrap(),
                        )
                        .unwrap();

                    assert_eq!(add_in.property("theme").unwrap().value(), "light");
                    add_in
                        .upsert_property(Property::new("theme", "dark").unwrap())
                        .unwrap();

                    let binding = add_in.binding_mut("selection").unwrap();
                    assert_ext(binding.ext(), ExtKind::AddIn, "binding-old");
                    assert!(binding.clear_ext().is_some());
                    add_in
                        .upsert_binding(
                            Binding::new("selection", "table", "Sheet1!A1:B2")
                                .unwrap()
                                .with_ext(test_ext(ExtKind::AddIn, "binding-final"))
                                .unwrap(),
                        )
                        .unwrap();

                    assert_ext(add_in.ext(), ExtKind::AddIn, "add-in-old");
                    assert!(add_in.clear_ext().is_some());
                    add_in
                        .set_ext(test_ext(ExtKind::AddIn, "add-in-final"))
                        .unwrap();
                }

                assert_ext(
                    pane.add_in().snapshot().and_then(|snapshot| snapshot.ext()),
                    ExtKind::DrawingMl,
                    "snapshot-old",
                );
                let snapshot = pane.snapshot_mut();
                assert!(snapshot.clear_ext().is_some());
                snapshot
                    .set_ext(test_ext(ExtKind::DrawingMl, "snapshot-final"))
                    .unwrap();

                pane.set_compression(Some(Compression::Print));
                assert!(pane.clear_compression());
                pane.set_compression(Some(Compression::HighQualityPrint));
                let blur = test_effect("blur");
                assert_eq!(
                    pane.replace_effect(0, blur.clone())
                        .unwrap()
                        .unwrap()
                        .kind(),
                    EffectKind::Grayscale
                );
                assert!(pane.clear_effects());
                pane.push_effect(blur).unwrap();

                let image = Arc::new(SNAPSHOT_IMAGE.to_vec());
                let shared = Arc::clone(&image);
                pane.set_image(
                    "/webextensions/media/crud-snapshot.jpg",
                    "image/jpeg",
                    image,
                )
                .unwrap();
                assert!(Arc::ptr_eq(&shared, &pane.image().unwrap().shared()));
                pane.set_external_link("https://example.invalid/crud-snapshot.jpg")
                    .unwrap();
                Ok(())
            })
            .unwrap()
    );

    assert_eq!(panes.get(0usize).unwrap().add_in().id(), ADD_IN_ID);
    assert!(panes.get(1usize).is_none());

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("public-web-crud.xlsx");
    let mut workbook = litchi_ooxml::xlsx::Workbook::create().unwrap();
    workbook
        .put_task_panes(panes, Conformance::Transitional)
        .unwrap();
    workbook.save(&path).unwrap();

    let workbook = litchi_ooxml::xlsx::Workbook::open(&path).unwrap();
    let panes = workbook.task_panes().unwrap().unwrap();
    assert_eq!(panes.get(0usize).unwrap().add_in().id(), ADD_IN_ID);
    let pane = panes.get(ADD_IN_ID).unwrap();
    assert!(!pane.visible());
    assert_eq!(pane.pane_width(), 512.5);
    assert_eq!(pane.dock_kind(), &Dock::Left);
    assert_eq!(pane.row(), 7);
    assert!(pane.locked());

    let add_in = pane.add_in();
    assert_eq!(add_in.reference().id(), "primary-final");
    assert_eq!(add_in.reference().version(), "2");
    assert_eq!(add_in.reference().store(), Store::Registry);
    assert_ext(add_in.reference().ext(), ExtKind::AddIn, "primary-final");
    let alternate = add_in.alternate_reference("alternate").unwrap();
    assert_eq!(alternate.version(), "2");
    assert_eq!(alternate.store(), Store::Registry);
    assert_ext(alternate.ext(), ExtKind::AddIn, "alternate-final");
    assert_eq!(add_in.property("theme").unwrap().value(), "dark");
    let binding = add_in.binding("selection").unwrap();
    assert_eq!(binding.kind_name(), "table");
    assert_eq!(binding.app_ref(), "Sheet1!A1:B2");
    assert_ext(binding.ext(), ExtKind::AddIn, "binding-final");
    assert_ext(add_in.ext(), ExtKind::AddIn, "add-in-final");

    let snapshot = add_in.snapshot().unwrap();
    assert_eq!(snapshot.compression(), Some(Compression::HighQualityPrint));
    assert_eq!(snapshot.effects().len(), 1);
    assert_eq!(snapshot.effects()[0].kind(), EffectKind::Blur);
    assert_ext(snapshot.ext(), ExtKind::DrawingMl, "snapshot-final");
    assert_ext(pane.ext(), ExtKind::TaskPane, "pane-final");
    let image = pane.image().unwrap();
    assert_eq!(
        image.name().as_str(),
        "/webextensions/media/crud-snapshot.jpg"
    );
    assert_eq!(image.content_type(), "image/jpeg");
    assert_eq!(image.bytes(), SNAPSHOT_IMAGE);
    assert_eq!(
        pane.link().unwrap().external(),
        Some("https://example.invalid/crud-snapshot.jpg")
    );
}

fn install_task_panes(package: &mut OpcPackage) {
    package.rels_mut().add_relationship(
        TASK_PANES_RELATIONSHIP.into(),
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
        ADD_IN_RELATIONSHIP.into(),
        "webextension1.xml".into(),
        "rId1".into(),
        false,
    );
    package.add_part(Box::new(task_panes));
    package.add_part(Box::new(XmlPart::new(
        PackURI::new("/webextensions/webextension1.xml").unwrap(),
        ADD_IN_CONTENT_TYPE.into(),
        LOCAL_EXTENSION.to_vec(),
    )));
}

fn assert_task_pane(task_panes: Panes) {
    assert_eq!(task_panes.len(), 1);
    let pane = task_panes.iter().next().unwrap();
    assert_eq!(pane.dock_state(), "right");
    assert!(pane.visible());
    assert_eq!(pane.add_in().reference().store(), Store::Omex);
    assert!(matches!(
        pane.add_in().reference().id(),
        "local-omex" | "inert-test-add-in"
    ));
}

fn authored_task_panes() -> Panes {
    let reference = Reference::new("inert-test-add-in", "1.0", Store::Omex)
        .unwrap()
        .catalog("en-US")
        .unwrap();
    let add_in = AddIn::new("{10000000-0000-0000-0000-000000000001}", reference)
        .unwrap()
        .frozen(true);
    let pane = Pane::new(add_in).width(360.0).unwrap();
    let mut panes = Panes::new();
    panes.push(pane).unwrap();
    panes
}

fn test_effect(name: &str) -> Effect {
    Effect::from_xml(
        format!(
            r#"<a:{name} xmlns:a="{}"/>"#,
            ExtKind::DrawingMl.namespace()
        )
        .as_bytes(),
    )
    .unwrap()
}

fn test_ext(kind: ExtKind, marker: &str) -> ExtList {
    let prefix = match kind {
        ExtKind::AddIn => "we",
        ExtKind::TaskPane => "wetp",
        ExtKind::DrawingMl | ExtKind::StrictDrawingMl => "a",
    };
    let drawing_namespace = if kind == ExtKind::StrictDrawingMl {
        ExtKind::StrictDrawingMl.namespace()
    } else {
        ExtKind::DrawingMl.namespace()
    };
    let drawing_declaration = if prefix != "a" {
        format!(r#" xmlns:a="{drawing_namespace}""#)
    } else {
        String::new()
    };
    ExtList::from_xml(
        format!(
            r#"<{prefix}:extLst xmlns:{prefix}="{}"{drawing_declaration}><a:ext uri="urn:litchi:{marker}"><v:marker xmlns:v="urn:litchi:test" name="{marker}"/></a:ext></{prefix}:extLst>"#,
            kind.namespace()
        )
        .as_bytes(),
    )
    .unwrap()
}

fn assert_ext(extension: Option<&ExtList>, kind: ExtKind, marker: &str) {
    let extension = extension.unwrap();
    assert_eq!(extension.kind(), kind);
    assert!(extension.xml().contains(marker));
}
