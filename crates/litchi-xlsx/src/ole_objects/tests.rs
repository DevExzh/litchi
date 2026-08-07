use super::model::{MAX_XML_BYTES, REL, SML, STRICT_SML};
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

const POI: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/spreadsheet/bug64512_embed.xlsx");

fn marker(row: u32) -> OleObjectMarker {
    OleObjectMarker {
        column: 1,
        column_offset: 0,
        row,
        row_offset: 0,
    }
}
fn value() -> OleObjects {
    OleObjects {
        objects: vec![OleObject {
            program_id: Some("Package.2".into()),
            data_or_view_aspect: Some(Aspect::Icon),
            link: None,
            update: Some(OleObjectUpdate::OnCall),
            auto_load: Some(false),
            shape_id: 1025,
            relationship_id: "rIdOle".into(),
            relationship_kind: OleObjectRelationshipKind::OleObject,
            target: Some(OleObjectTarget::Internal(OleObjectResource {
                part_name: "/xl/embeddings/oleObject1.bin".into(),
                content_type: ct::OFC_OLE_OBJECT.into(),
                data: vec![0xd0, 0xcf, 0x11, 0xe0],
            })),
            properties: Some(OleObjectProperties {
                preview_relationship_id: "rIdPreview".into(),
                preview: Some(OleObjectResource {
                    part_name: "/xl/media/image1.emf".into(),
                    content_type: "image/x-emf".into(),
                    data: vec![1, 2, 3],
                }),
                default_size: Some(false),
                print: Some(true),
                disabled: None,
                ui_object: None,
                auto_fill: Some(false),
                auto_line: Some(false),
                auto_pict: None,
                dde: None,
                macro_name: None,
                alt_text: Some("Object preview".into()),
                anchor: OleObjectAnchor {
                    move_with_cells: Some(true),
                    size_with_cells: Some(false),
                    from: marker(1),
                    to: marker(3),
                },
            }),
        }],
    }
}
fn package(conformance: OleObjectConformance) -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::SML_WORKSHEET.into(),
        format!(
            "<x:worksheet xmlns:x=\"{}\"><x:sheetData/><x:tableParts/></x:worksheet>",
            conformance.sml()
        )
        .into_bytes(),
    )));
    (package, uri)
}

fn package_with_unknown_markup() -> (OpcPackage, PackURI) {
    let (mut package, uri) = package(OleObjectConformance::Strict);
    store_ole_objects(&mut package, &uri, &value(), OleObjectConformance::Strict).unwrap();
    let xml = String::from_utf8(package.get_part(&uri).unwrap().blob().to_vec())
        .unwrap()
        .replace(
            "xmlns:xdr=\"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing\"",
            "xmlns:xdr=\"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing\" xmlns:future=\"urn:example:future\"",
        )
        .replace(
            "<x:oleObject ",
            "<x:oleObject futureAttr=\"preserve-me\"",
        )
        .replace(
            "</x:oleObject>",
            "<future:opaque futureValue=\"preserve-me\"/></x:oleObject>",
        );
    package
        .get_part_mut(&uri)
        .unwrap()
        .set_blob(xml.into_bytes());
    (package, uri)
}

#[test]
fn strict_round_trip_covers_complete_typed_properties() {
    let expected = value();
    let fragment = write_ole_objects(&expected, OleObjectConformance::Strict).unwrap();
    let xml = [
        format!("<x:worksheet xmlns:x=\"{STRICT_SML}\">").as_bytes(),
        fragment.as_slice(),
        b"</x:worksheet>",
    ]
    .concat();
    let parsed = parse_ole_objects(&xml).unwrap().unwrap();
    assert_eq!(parsed.objects[0].program_id.as_deref(), Some("Package.2"));
    assert_eq!(
        parsed.objects[0].properties.as_ref().unwrap().anchor.to.row,
        3
    );
    assert!(parsed.objects[0].target.is_none());
}

#[test]
fn loads_real_poi_mce_objects_without_opening_payloads() {
    let package = OpcPackage::from_bytes(POI).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let objects = load_ole_objects(&package, &uri).unwrap().unwrap();
    assert_eq!(objects.objects.len(), 2);
    assert_eq!(objects.objects[0].program_id.as_deref(), Some("Package"));
    assert_eq!(objects.objects[1].program_id.as_deref(), Some("Package2"));
    assert!(
        objects
            .objects
            .iter()
            .all(|object| matches!(object.target, Some(OleObjectTarget::Internal(_))))
    );
    assert!(objects.objects.iter().all(|object| {
        object
            .properties
            .as_ref()
            .unwrap()
            .preview
            .as_ref()
            .unwrap()
            .data
            .starts_with(b"\x01\x00\x00\x00")
    }));
}

#[test]
fn strict_package_writer_round_trips_and_inserts_in_schema_order() {
    let (mut package, uri) = package(OleObjectConformance::Strict);
    let expected = value();
    store_ole_objects(&mut package, &uri, &expected, OleObjectConformance::Strict).unwrap();
    assert_eq!(load_ole_objects(&package, &uri).unwrap().unwrap(), expected);
    let xml = package.get_part(&uri).unwrap().blob();
    assert!(
        memchr::memmem::find(xml, b"<x:oleObjects").unwrap()
            < memchr::memmem::find(xml, b"<x:tableParts").unwrap()
    );
}

#[test]
fn accepts_inert_external_package_target() {
    let (mut package, uri) = package(OleObjectConformance::Transitional);
    let mut expected = value();
    let object = &mut expected.objects[0];
    object.relationship_kind = OleObjectRelationshipKind::Package;
    object.target = Some(OleObjectTarget::External(
        "https://example.invalid/object.xlsx".into(),
    ));
    object.link = Some("'https://example.invalid/object.xlsx'!A1".into());
    object.properties = None;
    store_ole_objects(
        &mut package,
        &uri,
        &expected,
        OleObjectConformance::Transitional,
    )
    .unwrap();
    assert_eq!(load_ole_objects(&package, &uri).unwrap().unwrap(), expected);
}

#[test]
fn rejects_malformed_markup_caps_and_graphs() {
    for xml in [
        format!(
            r#"<worksheet xmlns="{SML}"><oleObjects><oleObject shapeId="0"/></oleObjects></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><oleObjects><oleObject shapeId="1" r:id="rId1"><objectPr r:id="rId2"><anchor><to/></anchor></objectPr></oleObject></oleObjects></worksheet>"#
        ),
        format!(r#"<!DOCTYPE x><worksheet xmlns="{SML}"/>"#),
    ] {
        assert!(parse_ole_objects(xml.as_bytes()).is_err(), "{xml}");
    }
    assert!(parse_ole_objects(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    let (mut missing, uri) = package(OleObjectConformance::Transitional);
    let fragment = write_ole_objects(&value(), OleObjectConformance::Transitional).unwrap();
    missing.get_part_mut(&uri).unwrap().set_blob(
        [
            format!("<x:worksheet xmlns:x=\"{SML}\">").as_bytes(),
            fragment.as_slice(),
            b"</x:worksheet>",
        ]
        .concat(),
    );
    assert!(load_ole_objects(&missing, &uri).is_err());
    let (mut unreferenced, uri) = package(OleObjectConformance::Transitional);
    unreferenced
        .get_part_mut(&uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::OLE_OBJECT.into(),
            "../embeddings/x.bin".into(),
            "rIdX".into(),
            false,
        );
    assert!(load_ole_objects(&unreferenced, &uri).is_err());

    let (mut orphan, uri) = package(OleObjectConformance::Transitional);
    orphan.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/embeddings/orphan.bin").unwrap(),
        ct::OFC_OLE_OBJECT.into(),
        vec![0xD0, 0xCF],
    )));
    assert!(Snapshot::load(&orphan, &uri).is_err());
}

#[test]
fn source_bound_anchor_edits_preserve_unknown_markup_and_payload_bytes() {
    let (mut package, uri) = package_with_unknown_markup();
    let source_before = package.get_part(&uri).unwrap().blob().to_vec();
    let payload_before = package
        .get_part(&PackURI::new("/xl/embeddings/oleObject1.bin").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = Transaction::new(&mut package, &uri).unwrap();
    let mut anchor = transaction.objects().unwrap().objects[0]
        .properties
        .as_ref()
        .unwrap()
        .anchor
        .clone();
    anchor.to.row = 9;
    assert!(transaction.set_anchor(1025, anchor).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let source_after = package.get_part(&uri).unwrap().blob();
    assert_ne!(source_after, source_before.as_slice());
    assert!(
        source_after
            .windows(b"futureAttr=\"preserve-me\"".len())
            .any(|window| { window == b"futureAttr=\"preserve-me\"" })
    );
    assert!(
        source_after
            .windows(b"<future:opaque futureValue=\"preserve-me\"/>".len())
            .any(|window| window == b"<future:opaque futureValue=\"preserve-me\"/>")
    );
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/embeddings/oleObject1.bin").unwrap())
            .unwrap()
            .blob(),
        payload_before.as_slice()
    );
    assert_eq!(
        load_ole_objects(&package, &uri).unwrap().unwrap().objects[0]
            .properties
            .as_ref()
            .unwrap()
            .anchor
            .to
            .row,
        9
    );
}

#[test]
fn source_bound_metadata_edits_insert_and_remove_known_attributes_without_loss() {
    let (mut package, uri) = package_with_unknown_markup();
    let mut transaction = Transaction::new(&mut package, &uri).unwrap();
    assert!(
        transaction
            .edit_object(1025, |object| {
                object.program_id = None;
                object.link = Some("https://example.invalid/linked".into());
                object.properties.as_mut().unwrap().alt_text = None;
                Ok(())
            })
            .unwrap()
    );
    transaction.commit().unwrap();
    let source = package.get_part(&uri).unwrap().blob();
    assert!(
        source
            .windows(b"link=\"https://example.invalid/linked\"".len())
            .any(|window| window == b"link=\"https://example.invalid/linked\"")
    );
    assert!(
        !source
            .windows(b"progId=\"Package.2\"".len())
            .any(|window| { window == b"progId=\"Package.2\"" })
    );
    assert!(
        !source
            .windows(b"altText=\"Object preview\"".len())
            .any(|window| { window == b"altText=\"Object preview\"" })
    );
    assert!(
        source
            .windows(b"futureValue=\"preserve-me\"".len())
            .any(|window| window == b"futureValue=\"preserve-me\"")
    );
    let object = &load_ole_objects(&package, &uri).unwrap().unwrap().objects[0];
    assert_eq!(object.program_id, None);
    assert_eq!(
        object.link.as_deref(),
        Some("https://example.invalid/linked")
    );
    assert_eq!(object.properties.as_ref().unwrap().alt_text, None);
}

#[test]
fn no_op_is_byte_exact_and_invalid_edits_are_staged_atomically() {
    let (mut package, uri) = package_with_unknown_markup();
    let source = package.get_part(&uri).unwrap().blob().to_vec();
    let mut transaction = Transaction::new(&mut package, &uri).unwrap();
    assert!(!transaction.edit_object(1025, |_object| Ok(())).unwrap());
    let staged = transaction.objects().cloned();
    assert!(
        transaction
            .edit_object(1025, |object| {
                object.shape_id = 0;
                Ok(())
            })
            .is_err()
    );
    assert_eq!(transaction.objects().cloned(), staged);
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(package.get_part(&uri).unwrap().blob(), source.as_slice());
}

#[test]
fn patches_replay_exact_sources_invert_and_reject_stale_targets_atomically() {
    let (mut source, uri) = package_with_unknown_markup();
    let original = source.get_part(&uri).unwrap().blob().to_vec();
    let mut transaction = Transaction::new(&mut source, &uri).unwrap();
    assert!(
        transaction
            .edit_object(1025, |object| {
                object.update = None;
                Ok(())
            })
            .unwrap()
    );
    let patch = transaction.commit().unwrap().patch().clone();
    let changed = source.get_part(&uri).unwrap().blob().to_vec();

    let (mut replay, replay_uri) = package_with_unknown_markup();
    assert_eq!(uri, replay_uri);
    patch.apply(&mut replay).unwrap();
    assert_eq!(replay.get_part(&uri).unwrap().blob(), changed.as_slice());
    patch.inverse().apply(&mut replay).unwrap();
    assert_eq!(replay.get_part(&uri).unwrap().blob(), original.as_slice());

    let (mut stale, stale_uri) = package_with_unknown_markup();
    let mut stale_xml = stale.get_part(&stale_uri).unwrap().blob().to_vec();
    stale_xml.extend_from_slice(b"\n");
    stale.get_part_mut(&stale_uri).unwrap().set_blob(stale_xml);
    let stale_before = stale.get_part(&stale_uri).unwrap().blob().to_vec();
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(
        stale.get_part(&stale_uri).unwrap().blob(),
        stale_before.as_slice()
    );
}

#[test]
fn mce_choice_and_fallback_sources_remain_editable_without_activation() {
    let mut package = OpcPackage::from_bytes(POI).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let source_before = package.get_part(&uri).unwrap().blob().to_vec();
    let mut transaction = Transaction::new(&mut package, &uri).unwrap();
    let mut anchor = transaction.objects().unwrap().objects[1]
        .properties
        .as_ref()
        .unwrap()
        .anchor
        .clone();
    anchor.from.row = 8;
    assert!(transaction.set_anchor(1026, anchor).unwrap());
    transaction.commit().unwrap();
    let source_after = package.get_part(&uri).unwrap().blob();
    assert_ne!(source_after, source_before.as_slice());
    assert!(
        source_after
            .windows(b"<mc:Fallback><oleObject progId=\"Package2\"".len())
            .any(|window| window == b"<mc:Fallback><oleObject progId=\"Package2\"")
    );
    assert_eq!(
        load_ole_objects(&package, &uri).unwrap().unwrap().objects[1]
            .properties
            .as_ref()
            .unwrap()
            .anchor
            .from
            .row,
        8
    );
}
