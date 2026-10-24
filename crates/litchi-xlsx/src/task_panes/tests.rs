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
    let parsed = OpcPackage::from_vec(
        package.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
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

fn same_reopened_content(left: &Package, right: &Package) -> bool {
    let left = Package::from_bytes(&left.to_bytes().unwrap()).unwrap();
    let right = Package::from_bytes(&right.to_bytes().unwrap()).unwrap();
    left.same_content(&right).unwrap()
}

#[test]
fn crud_is_atomic_and_preserves_relationships_and_opaque_extensions() {
    let mut package = Package::new().unwrap();
    let mut value = pane("addin-one");
    value
        .add_in_mut()
        .set_ext(opaque_add_in_extension())
        .unwrap();
    value.set_ext(opaque_task_pane_extension()).unwrap();

    let mut transaction = package.edit_task_panes().unwrap();
    transaction.add(value).unwrap();
    package = transaction.commit().unwrap().into_package();
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
    package = transaction.commit().unwrap().into_package();

    assert_eq!(root_task_pane_relationship(&package), relationship_before);
    let workbook = package.workbook();
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
    package = transaction.commit().unwrap().into_package();
    assert!(package.task_panes().unwrap().is_none());
}

#[test]
fn strict_task_panes_round_trip_through_facade() {
    let mut package = Package::new().unwrap();
    let mut transaction = package
        .edit_task_panes_with_options(EditOptions::default().conformance(Conformance::Strict))
        .unwrap();
    transaction.add(pane("strict-addin")).unwrap();
    package = transaction.commit().unwrap().into_package();

    let bytes = package.to_bytes().unwrap();
    let xml_package = OpcPackage::from_vec(bytes, litchi_opc::SourceLimits::default()).unwrap();
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
    package = transaction.commit().unwrap().into_package();
    assert!(package.task_panes().unwrap().is_none());
}

#[test]
fn malformed_task_pane_graph_is_rejected_without_mutation() {
    let mut raw = Package::new().unwrap().into_plain_opc();
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
    let package = Package::from_opc(raw).unwrap();
    let before = package.to_bytes().unwrap();
    let edit = package.edit_task_panes().unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn unknown_root_parts_survive_task_pane_edits() {
    let mut package: Package = {
        let mut raw = Package::new().unwrap().into_plain_opc();
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
    package = transaction.commit().unwrap().into_package();

    let parsed = OpcPackage::from_vec(
        package.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    let opaque_name = PackURI::new("/xl/opaque.bin").unwrap();
    assert_eq!(
        parsed.get_part(&opaque_name).unwrap().blob(),
        b"opaque payload"
    );
    assert!(parsed.rels().get("rIdOpaque").is_some());
}

#[test]
fn commits_publish_new_packages_and_patches_replay_root_relationships_exactly() {
    let source = Package::new().unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_task_panes().unwrap();
    edit.add(pane("addin-one")).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (after, patch) = commit.into_parts();
    // The source is unchanged; the new package relates the task panes from
    // the package root.
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    assert!(source.task_panes().unwrap().is_none());
    assert!(!root_task_pane_relationship(&after).0.is_empty());

    let replayed = source.apply_task_panes(&patch).unwrap().into_package();
    assert!(replayed.same_content(&after).unwrap());
    assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
    assert_eq!(replayed.task_panes().unwrap(), after.task_panes().unwrap());
    let reopened_source = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
    let reopened_after = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
    let reopened_forward = reopened_source
        .apply_task_panes(&patch)
        .unwrap()
        .into_package();
    assert!(same_reopened_content(&reopened_forward, &after));
    let reopened_restored = reopened_after
        .apply_task_panes(&patch.inverse())
        .unwrap()
        .into_package();
    assert!(same_reopened_content(&reopened_restored, &source));
    let restored = after
        .apply_task_panes(&patch.inverse())
        .unwrap()
        .into_package();
    assert!(restored.same_content(&source).unwrap());
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);
    assert!(restored.task_panes().unwrap().is_none());
    assert!(matches!(
        after.apply_task_panes(&patch),
        Err(Error::PatchConflict { .. })
    ));

    // A dropped or reset edit stages nothing.
    let mut edit = source.edit_task_panes().unwrap();
    edit.add(pane("addin-two")).unwrap();
    edit.reset();
    let noop = edit.commit().unwrap();
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().to_bytes().unwrap(), source_bytes);
}
