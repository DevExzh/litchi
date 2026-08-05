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
            data_or_view_aspect: Some(OleObjectAspect::Icon),
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
}
