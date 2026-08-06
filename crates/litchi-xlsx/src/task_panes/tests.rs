use super::*;
use crate::Package;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const ADD_IN_NS: &str = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const TASK_PANE_NS: &str = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn pane(id: &str) -> Pane {
    let reference = Reference::new(format!("{id}-reference"), "1.0", Store::Omex).unwrap();
    Pane::new(AddIn::new(id, reference).unwrap())
}

fn opaque_add_in_extension() -> ExtList {
    ExtList::from_xml(
        format!(
            r#"<we:extLst xmlns:we="{ADD_IN_NS}" xmlns:a="{DRAWING_NS}" xmlns:vendor="urn:vendor"><a:ext uri="urn:vendor"><vendor:payload data="kept"/></a:ext></we:extLst>"#
        )
        .as_bytes(),
    )
    .unwrap()
}

fn opaque_task_pane_extension() -> ExtList {
    ExtList::from_xml(
        format!(
            r#"<wetp:extLst xmlns:wetp="{TASK_PANE_NS}" xmlns:a="{DRAWING_NS}" xmlns:vendor="urn:vendor"><a:ext uri="urn:vendor"><vendor:payload><![CDATA[kept]]></vendor:payload></a:ext></wetp:extLst>"#
        )
        .as_bytes(),
    )
    .unwrap()
}

fn root_task_pane_relationship(package: &Package) -> (String, String) {
    let parsed = OpcPackage::from_vec(package.to_bytes().unwrap()).unwrap();
    let relationship = parsed
        .rels()
        .iter()
        .find(|relationship| {
            relationship.reltype() == litchi_ooxml_common::web::raw::TASK_PANES_RELATIONSHIP
        })
        .unwrap();
    (
        relationship.r_id().to_owned(),
        relationship.target_ref().to_owned(),
    )
}

#[test]
fn crud_is_atomic_and_preserves_relationships_and_opaque_extensions() {
    let mut package = Package::create().unwrap();
    let mut value = pane("addin-one");
    value
        .add_in_mut()
        .set_ext(opaque_add_in_extension())
        .unwrap();
    value.set_ext(opaque_task_pane_extension()).unwrap();

    let mut transaction = package.edit_task_panes().unwrap();
    transaction.add(value).unwrap();
    transaction.commit().unwrap();
    let relationship_before = root_task_pane_relationship(&package);

    let mut transaction = package.edit_task_panes().unwrap();
    assert!(
        transaction
            .edit("addin-one", |pane| {
                pane.set_visible(false);
                Ok(())
            })
            .unwrap()
    );
    let failed = transaction.edit("addin-one", |_pane| {
        Err(litchi_ooxml_common::Error::Invalid(
            "deliberate edit failure".into(),
        ))
    });
    assert!(failed.is_err());
    transaction.commit().unwrap();

    assert_eq!(root_task_pane_relationship(&package), relationship_before);
    let workbook = package.workbook().unwrap();
    let panes = workbook.task_panes().unwrap().unwrap();
    let stored = panes.get("addin-one").unwrap();
    assert!(!stored.visible());
    assert!(
        stored
            .add_in()
            .ext()
            .unwrap()
            .xml()
            .contains("data=\"kept\"")
    );
    assert!(stored.ext().unwrap().xml().contains("<![CDATA[kept]]>"));

    let mut transaction = package.edit_task_panes().unwrap();
    assert!(transaction.remove("addin-one").unwrap().is_some());
    transaction.commit().unwrap();
    assert!(package.task_panes().unwrap().is_none());
}

#[test]
fn strict_task_panes_round_trip_through_facade() {
    let mut package = Package::create().unwrap();
    let mut transaction = package.edit_task_panes_with(Conformance::Strict).unwrap();
    transaction.add(pane("strict-addin")).unwrap();
    transaction.commit().unwrap();

    let bytes = package.to_bytes().unwrap();
    let xml_package = OpcPackage::from_vec(bytes).unwrap();
    let relationship = xml_package
        .rels()
        .iter()
        .find(|relationship| {
            relationship.reltype() == litchi_ooxml_common::web::raw::TASK_PANES_RELATIONSHIP
        })
        .unwrap();
    let part = xml_package
        .get_part(&relationship.target_partname().unwrap())
        .unwrap();
    assert!(
        part.blob()
            .windows(b"http://purl.oclc.org/ooxml/officeDocument/relationships".len())
            .any(|window| window == b"http://purl.oclc.org/ooxml/officeDocument/relationships")
    );

    let mut transaction = package.edit_task_panes().unwrap();
    assert_eq!(transaction.panes().unwrap().len(), 1);
    transaction.clear().unwrap();
    transaction.commit().unwrap();
    assert!(package.task_panes().unwrap().is_none());
}

#[test]
fn malformed_task_pane_graph_is_rejected_without_mutation() {
    let mut raw: OpcPackage = Package::create().unwrap().into();
    let part_name = PackURI::new("/xl/webextensions/taskpanes.xml").unwrap();
    raw.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        litchi_ooxml_common::web::raw::TASK_PANES_CONTENT_TYPE.into(),
        b"<wetp:taskpanes xmlns:wetp=\"urn:wrong\"/>".to_vec(),
    )))
    .unwrap();
    raw.rels_mut().add_relationship(
        litchi_ooxml_common::web::raw::TASK_PANES_RELATIONSHIP.into(),
        part_name.relative_ref("/"),
        "rIdPanesMalformed".into(),
        false,
    );
    let mut package = Package::from_opc(raw).unwrap();
    let before = package.to_bytes().unwrap();
    assert!(package.edit_task_panes().is_err());
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn unknown_root_parts_survive_task_pane_edits() {
    let mut package: Package = {
        let mut raw: OpcPackage = Package::create().unwrap().into();
        let opaque_name = PackURI::new("/xl/opaque.bin").unwrap();
        raw.try_add_part(Box::new(BlobPart::new(
            opaque_name.clone(),
            "application/octet-stream".into(),
            b"opaque payload".to_vec(),
        )))
        .unwrap();
        raw.rels_mut().add_relationship(
            "urn:vendor:opaque".into(),
            opaque_name.relative_ref("/"),
            "rIdOpaque".into(),
            false,
        );
        Package::from_opc(raw).unwrap()
    };
    let mut transaction = package.edit_task_panes().unwrap();
    transaction.add(pane("opaque-survivor")).unwrap();
    transaction.commit().unwrap();

    let parsed = OpcPackage::from_vec(package.to_bytes().unwrap()).unwrap();
    let opaque_name = PackURI::new("/xl/opaque.bin").unwrap();
    assert_eq!(
        parsed.get_part(&opaque_name).unwrap().blob(),
        b"opaque payload"
    );
    assert!(parsed.rels().get("rIdOpaque").is_some());
}
