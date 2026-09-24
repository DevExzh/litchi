#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the focused transaction tests use panic-on-failure assertions"
)]

use super::*;
use crate::web::raw::{
    ADD_IN_CONTENT_TYPE, ADD_IN_RELATIONSHIP, TASK_PANES_CONTENT_TYPE, TASK_PANES_RELATIONSHIP,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, PackageWriter, TargetMode};

const TASK_XML: &[u8] = br#"<we:taskpanes xmlns:we="http://schemas.microsoft.com/office/webextensions/webextensiontaskpanes/2010/11"><we:taskpane/></we:taskpanes>"#;
const TASK_XML_AFTER: &[u8] = br#"<we:taskpanes xmlns:we="http://schemas.microsoft.com/office/webextensions/webextensiontaskpanes/2010/11"><we:taskpane visible="0"/></we:taskpanes>"#;
const EXT_XML: &[u8] = br#"<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"/>"#;

fn package() -> (OpcPackage, PackURI, PackURI, PackURI, PackURI) {
    let task = PackURI::new("/web/taskpanes.xml").unwrap();
    let extension = PackURI::new("/web/extension.xml").unwrap();
    let old_image = PackURI::new("/media/old.png").unwrap();
    let new_image = PackURI::new("/media/new.png").unwrap();
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        task.clone(),
        TASK_PANES_CONTENT_TYPE.into(),
        TASK_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        extension.clone(),
        ADD_IN_CONTENT_TYPE.into(),
        EXT_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        old_image.clone(),
        "image/png".into(),
        vec![1, 2, 3, 4],
    )));
    package.relate_to("web/taskpanes.xml", TASK_PANES_RELATIONSHIP);
    package
        .get_part_mut(&task)
        .unwrap()
        .rels_mut()
        .add_relationship(
            ADD_IN_RELATIONSHIP.into(),
            "extension.xml".into(),
            "rIdExtension".into(),
            false,
        );
    (package, task, extension, old_image, new_image)
}

fn task_noop_patch(package: &OpcPackage, task: &PackURI) -> Patch {
    let root = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .map(RelationshipState::capture);
    let task_state = PartState::capture(package.get_part(task).unwrap());
    let planned_task = PlannedPart {
        name: task.clone(),
        content_type: task_state.content_type.clone(),
        data: Arc::clone(&task_state.data),
        relationships: task_state
            .relationships
            .iter()
            .map(|relationship| PlannedRelationship {
                id: relationship.id.clone(),
                relationship_type: relationship.relationship_type.clone(),
                target: relationship.target.clone(),
                external: relationship.external,
            })
            .collect(),
    };
    Patch::planned(
        package,
        PatchPlan {
            before: PlannedGraph {
                root: root.clone(),
                owned_parts: vec![task.clone()],
            },
            after: PlannedGraph {
                root,
                owned_parts: vec![task.clone()],
            },
            parts: vec![planned_task],
            deletions: Vec::new(),
            limits: Limits::standard(),
        },
    )
    .unwrap()
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn stored_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut archive = Vec::new();
    let mut central = Vec::new();
    for (name, bytes) in entries {
        let name = name.as_bytes();
        let offset = u32::try_from(archive.len()).unwrap();
        let length = u32::try_from(bytes.len()).unwrap();
        let name_length = u16::try_from(name.len()).unwrap();
        push_u32(&mut archive, 0x0403_4b50);
        push_u16(&mut archive, 20);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u32(&mut archive, crc32(bytes));
        push_u32(&mut archive, length);
        push_u32(&mut archive, length);
        push_u16(&mut archive, name_length);
        push_u16(&mut archive, 0);
        archive.extend_from_slice(name);
        archive.extend_from_slice(bytes);

        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 20);
        push_u16(&mut central, 20);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc32(bytes));
        push_u32(&mut central, length);
        push_u32(&mut central, length);
        push_u16(&mut central, name_length);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, offset);
        central.extend_from_slice(name);
    }
    let central_offset = u32::try_from(archive.len()).unwrap();
    let central_size = u32::try_from(central.len()).unwrap();
    archive.extend_from_slice(&central);
    push_u32(&mut archive, 0x0605_4b50);
    push_u16(&mut archive, 0);
    push_u16(&mut archive, 0);
    let entry_count = u16::try_from(entries.len()).unwrap();
    push_u16(&mut archive, entry_count);
    push_u16(&mut archive, entry_count);
    push_u32(&mut archive, central_size);
    push_u32(&mut archive, central_offset);
    push_u16(&mut archive, 0);
    archive
}

fn lexical_source_package() -> (OpcPackage, PackURI) {
    const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/web/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/><Override PartName="/web/extension.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>"#;
    const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPanes" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="web/taskpanes.xml"/></Relationships>"#;
    const TASK_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0"?>
<!-- before -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExtension" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="extension.xml"/>
  <!-- between -->
</Relationships>
<!-- after -->"#;
    let task = PackURI::new("/web/taskpanes.xml").unwrap();
    let archive = stored_zip(&[
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELATIONSHIPS),
        ("web/taskpanes.xml", TASK_XML),
        ("web/_rels/taskpanes.xml.rels", TASK_RELATIONSHIPS),
        ("web/extension.xml", EXT_XML),
    ]);
    (OpcPackage::from_bytes(&archive).unwrap(), task)
}

#[test]
fn inverse_restores_content_types_relationships_and_xml_after_reopen() {
    let (mut package, task, extension, old_image, new_image) = package();
    let root = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP)
        .map(RelationshipState::capture);
    let before_content_types = package.source_content_types().unwrap();
    let before_relationships = package.source_relationships(&task).unwrap();
    let before_xml = package.source_xml_part(&task).unwrap();
    let old_image_bytes = package.get_part(&old_image).unwrap().blob_arc();
    let extension_state = PartState::capture(package.get_part(&extension).unwrap());
    let task_relationships = [
        RelationshipState {
            id: "rIdExtension".into(),
            relationship_type: ADD_IN_RELATIONSHIP.into(),
            target: "extension.xml".into(),
            external: false,
        },
        RelationshipState {
            id: "rIdNewImage".into(),
            relationship_type: rt::IMAGE.into(),
            target: "../media/new.png".into(),
            external: false,
        },
    ];
    let root_after = root.clone();
    let patch = Patch::planned(
        &package,
        PatchPlan {
            before: PlannedGraph {
                root: root.clone(),
                owned_parts: vec![task.clone(), extension.clone(), old_image.clone()],
            },
            after: PlannedGraph {
                root: root_after,
                owned_parts: vec![task.clone(), extension.clone(), new_image.clone()],
            },
            parts: vec![
                PlannedPart {
                    name: task.clone(),
                    content_type: TASK_PANES_CONTENT_TYPE.into(),
                    data: Arc::new(TASK_XML_AFTER.to_vec()),
                    relationships: task_relationships
                        .iter()
                        .map(|relationship| PlannedRelationship {
                            id: relationship.id.clone(),
                            relationship_type: relationship.relationship_type.clone(),
                            target: relationship.target.clone(),
                            external: relationship.external,
                        })
                        .collect(),
                },
                PlannedPart {
                    name: new_image.clone(),
                    content_type: "image/png".into(),
                    data: Arc::new(vec![9, 8, 7, 6]),
                    relationships: Vec::new(),
                },
            ],
            deletions: vec![old_image.clone()],
            limits: Limits::standard(),
        },
    )
    .unwrap();
    assert!(!patch.is_empty());
    let mut signed = package.clone();
    signed.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(patch.apply(&mut signed).is_err());
    assert!(signed.is_signed());
    assert!(patch.apply(&mut package).unwrap());
    assert_eq!(package.get_part(&task).unwrap().blob(), TASK_XML_AFTER);
    assert!(package.get_part(&old_image).is_err());
    assert!(package.get_part(&new_image).is_ok());

    let inverse = patch.inverse();
    assert!(inverse.apply(&mut package).unwrap());
    assert_eq!(
        package.source_content_types().unwrap(),
        before_content_types
    );
    assert_eq!(
        package.source_relationships(&task).unwrap(),
        before_relationships
    );
    assert_eq!(package.source_xml_part(&task).unwrap(), before_xml);
    assert!(Arc::ptr_eq(
        &package.get_part(&old_image).unwrap().blob_arc(),
        &old_image_bytes
    ));
    assert!(package.get_part(&new_image).is_err());
    assert!(PartState::capture(package.get_part(&extension).unwrap()) == extension_state);

    let mut bytes = Vec::new();
    PackageWriter::write_to_stream(&mut bytes, &package).unwrap();
    let reopened = OpcPackage::from_vec(bytes).unwrap();
    assert_eq!(
        reopened.source_content_types().unwrap(),
        before_content_types
    );
    assert_eq!(
        reopened.source_relationships(&task).unwrap(),
        before_relationships
    );
    assert_eq!(reopened.source_xml_part(&task).unwrap(), before_xml);
}

#[test]
fn source_bound_noop_checks_source_and_inverse_before_returning_false() {
    let (mut source_package, task, _extension, _old_image, _new_image) = package();
    let patch = task_noop_patch(&source_package, &task);
    assert!(patch.is_empty());
    assert!(!patch.apply(&mut source_package).unwrap());

    let (mut lexical_package, lexical_task) = lexical_source_package();
    let lexical_patch = task_noop_patch(&lexical_package, &lexical_task);
    let before_relationships = lexical_package.source_relationships(&lexical_task).unwrap();
    let removed = before_relationships
        .without_relationship("rIdExtension", Limits::standard().xml_bytes)
        .unwrap();
    let rewritten = removed
        .with_relationship(
            ADD_IN_RELATIONSHIP,
            "extension.xml",
            "rIdExtension",
            TargetMode::Internal,
            Limits::standard().xml_bytes,
        )
        .unwrap();
    assert_ne!(before_relationships, rewritten);
    lexical_package
        .try_replace_relationships(&before_relationships, &removed)
        .unwrap();
    lexical_package
        .try_replace_relationships(&removed, &rewritten)
        .unwrap();
    assert!(lexical_patch.apply(&mut lexical_package).is_err());

    source_package
        .get_part_mut(&task)
        .unwrap()
        .set_blob(TASK_XML_AFTER.to_vec());
    let stale_bytes = source_package.get_part(&task).unwrap().blob().to_vec();
    assert!(patch.apply(&mut source_package).is_err());
    assert_eq!(source_package.get_part(&task).unwrap().blob(), stale_bytes);

    let (mut inverse_package, inverse_task, _extension, _old_image, _new_image) = package();
    let inverse = patch.inverse();
    assert!(!inverse.apply(&mut inverse_package).unwrap());
    inverse_package
        .get_part_mut(&inverse_task)
        .unwrap()
        .set_blob(TASK_XML_AFTER.to_vec());
    assert!(inverse.apply(&mut inverse_package).is_err());
}
