//! Regression tests for the typed Custom XML Maps model and package service.

use super::model::{CONTENT_TYPE, MAX_OPAQUE_BYTES, NS, REL, STRICT_NS, STRICT_REL};
use super::*;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI, PackageWriter};

fn package(bytes: &[u8]) -> XmlMapInfo {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    load_from_package(&package).unwrap().unwrap()
}

fn fixture_info() -> XmlMapInfo {
    XmlMapInfo {
        selection_namespaces: "xmlns:xs='http://www.w3.org/2001/XMLSchema'".into(),
        schemas: vec![XmlMapSchema {
            id: "schema-1".into(),
            schema_reference: Some("urn:litchi:example".into()),
            namespace: Some("urn:litchi:example".into()),
            payload_xml: Some(
                br#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
            ),
        }],
        maps: vec![XmlMap {
            id: 1,
            name: "Example map".into(),
            root_element: "example".into(),
            schema_id: "schema-1".into(),
            show_import_export_validation_errors: true,
            auto_fit: true,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: Some(XmlMapDataBinding {
                data_binding_name: Some("inert binding".into()),
                file_binding: Some(true),
                connection_id: Some(7),
                file_binding_name: None,
                load_mode: 1,
                payload_xml: Some(br#"<binding xmlns="urn:litchi:binding"/>"#.to_vec()),
            }),
        }],
    }
}

fn workbook_package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        format!(
            r#"<workbook xmlns="{}"><sheets/></workbook>"#,
            std::str::from_utf8(NS).unwrap()
        )
        .into_bytes(),
    );
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    package
}

fn synthetic_package(
    relationship_type: &str,
    external: bool,
    content_type: &str,
    outbound: bool,
) -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let mut workbook = BlobPart::new(
        workbook_uri.clone(),
        ct::SML_SHEET_MAIN.into(),
        format!(
            r#"<workbook xmlns="{}"><sheets/></workbook>"#,
            std::str::from_utf8(NS).unwrap()
        )
        .into_bytes(),
    );
    if external {
        workbook.relate_to_ext("https://example.invalid/xmlMaps.xml", relationship_type);
    } else {
        workbook.relate_to("xmlMaps.xml", relationship_type);
    }
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    if !external {
        let mut maps = BlobPart::new(
            PackURI::new("/xl/xmlMaps.xml").unwrap(),
            content_type.into(),
            fixture_info().to_xml(false).unwrap(),
        );
        if outbound {
            maps.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        }
        package.add_part(Box::new(maps));
    }
    package
}

#[test]
fn reads_poi_real_fixture_and_round_trips_strict() {
    let maps = package(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
    ));
    assert_eq!(maps.schemas.len(), 1);
    assert_eq!(maps.maps[0].name, "CORSO_mapping");
    let strict = maps.to_xml(true).unwrap();
    assert_eq!(XmlMapInfo::parse(&strict).unwrap(), maps);
}

#[test]
fn keeps_poi_xxe_schema_inert() {
    let maps = package(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/xxe_in_schema.xlsx"
    ));
    let payload = maps.schemas[0].payload_xml.as_deref().unwrap();
    let text = std::str::from_utf8(payload).unwrap();
    assert!(text.contains("schemaLocation=\"http://localhost\""));
    assert!(text.contains("redefine"));
}

#[test]
fn reads_libreoffice_unqualified_children_and_binding() {
    let maps = package(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx"
    ));
    let binding = maps.maps[0].data_binding.as_ref().unwrap();
    assert_eq!(binding.file_binding, Some(true));
    assert_eq!(binding.connection_id, Some(1));
    assert_eq!(binding.load_mode, 1);
}

#[test]
fn handles_strict_and_mce_fallback() {
    let strict = std::str::from_utf8(STRICT_NS).unwrap();
    let xml = format!(
        r#"<MapInfo xmlns="{strict}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" SelectionNamespaces=""><Schema ID="s"><x:schema xmlns:x="urn:x"/></Schema><mc:AlternateContent><mc:Choice Requires="u"><u:Map/></mc:Choice><mc:Fallback><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="false" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"/></mc:Fallback></mc:AlternateContent></MapInfo>"#
    );
    let parsed = XmlMapInfo::parse(xml.as_bytes()).unwrap();
    assert_eq!(parsed.maps.len(), 1);
    assert_eq!(
        XmlMapInfo::parse(&parsed.to_xml(false).unwrap()).unwrap(),
        parsed
    );
}

#[test]
fn rejects_malformed_unsafe_and_invalid_models() {
    let ns = std::str::from_utf8(NS).unwrap();
    for xml in [
        format!(
            r#"<MapInfo xmlns="{ns}" SelectionNamespaces=""><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="0" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
        ),
        format!(
            r#"<MapInfo xmlns="{ns}" SelectionNamespaces=""><Schema ID="s"/><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="maybe" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
        ),
        format!(
            r#"<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><MapInfo xmlns="{ns}" SelectionNamespaces=""><Schema ID="s"><x:schema xmlns:x="urn:x">&e;</x:schema></Schema><Map ID="1" Name="m" RootElement="r" SchemaID="s" ShowImportExportValidationErrors="0" AutoFit="1" Append="0" PreserveSortAFLayout="1" PreserveFormat="1"/></MapInfo>"#
        ),
    ] {
        assert!(XmlMapInfo::parse(xml.as_bytes()).is_err(), "accepted {xml}");
    }
    let mut valid = package(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
    ));
    valid.schemas[0].payload_xml = Some(b"<?unsafe?><x/>".to_vec());
    assert!(valid.to_xml(false).is_err());
}

#[test]
fn serializer_rejects_oversized_output_before_final_append() {
    fn large_payload() -> Vec<u8> {
        let mut payload = Vec::with_capacity(MAX_OPAQUE_BYTES);
        payload.extend_from_slice(b"<x>");
        payload.resize(MAX_OPAQUE_BYTES - 4, b'x');
        payload.extend_from_slice(b"</x>");
        payload
    }

    let mut value = fixture_info();
    value.schemas[0].payload_xml = Some(large_payload());
    value.maps[0]
        .data_binding
        .as_mut()
        .expect("fixture has a data binding")
        .payload_xml = Some(large_payload());

    let error = value.to_xml(false).unwrap_err();
    assert_eq!(
        error.to_string(),
        "serialized custom XML maps part exceeds 32 MiB"
    );
}

#[test]
fn borrowed_codec_view_does_not_clone_large_opaque_payloads() {
    let mut value = fixture_info();
    value.schemas[0].payload_xml = Some(vec![b'x'; 4 * 1024 * 1024]);
    value.maps[0].data_binding.as_mut().unwrap().payload_xml = Some(vec![b'y'; 4 * 1024 * 1024]);
    let schema_payload = value.schemas[0].payload_xml.as_deref().unwrap();
    let binding_payload = value.maps[0]
        .data_binding
        .as_ref()
        .unwrap()
        .payload_xml
        .as_deref()
        .unwrap();

    let view = value.to_common_ref().unwrap();
    assert_eq!(
        view.schemas[0].payload_xml.unwrap().as_ptr(),
        schema_payload.as_ptr()
    );
    assert_eq!(
        view.maps[0]
            .data_binding
            .as_ref()
            .unwrap()
            .payload_xml
            .unwrap()
            .as_ptr(),
        binding_payload.as_ptr()
    );
}

#[test]
fn stores_rewrites_and_removes_inert_xml_maps_parts() {
    let mut package = workbook_package();
    let value = fixture_info();

    store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();
    assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));
    assert_eq!(
        load_from_package_with_conformance(&package).unwrap(),
        Some((value.clone(), XmlMapConformance::Transitional))
    );

    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == REL)
        .unwrap();
    let relationship_id = relationship.r_id().to_string();
    let part_name = relationship.target_partname().unwrap();
    assert_eq!(part_name, PackURI::new("/xl/xmlMaps.xml").unwrap());
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(std::str::from_utf8(NS).unwrap())
    );

    let mut replacement = value.clone();
    replacement.maps[0].name = "Strict replacement".into();
    store_in_package(&mut package, &replacement, XmlMapConformance::Strict).unwrap();
    let workbook = package.main_document_part().unwrap();
    let relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == relationship_id)
        .unwrap();
    assert_eq!(relationship.reltype(), STRICT_REL);
    assert_eq!(relationship.target_partname().unwrap(), part_name);
    assert!(
        std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
            .unwrap()
            .contains(std::str::from_utf8(STRICT_NS).unwrap())
    );
    assert_eq!(
        load_from_package_with_conformance(&package).unwrap(),
        Some((replacement, XmlMapConformance::Strict))
    );

    assert!(remove_from_package(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_err());
    assert_eq!(load_from_package(&package).unwrap(), None);
    assert!(!remove_from_package(&mut package).unwrap());
}

#[test]
fn preserves_unrelated_references_when_removing_xml_maps() {
    let mut package = workbook_package();
    let value = fixture_info();
    store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();

    let part_name = PackURI::new("/xl/xmlMaps.xml").unwrap();
    let mut referring_part = BlobPart::new(
        PackURI::new("/xl/retained-reference.xml").unwrap(),
        ct::XML.into(),
        b"<reference/>".to_vec(),
    );
    referring_part.relate_to("xmlMaps.xml", "urn:litchi:test:xml-maps-reference");
    package.add_part(Box::new(referring_part));

    assert!(remove_from_package(&mut package).unwrap());
    assert!(package.get_part(&part_name).is_ok());
    assert_eq!(load_from_package(&package).unwrap(), None);

    store_in_package(&mut package, &value, XmlMapConformance::Transitional).unwrap();
    let relationship = package
        .main_document_part()
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == REL)
        .unwrap();
    assert_eq!(
        relationship.target_partname().unwrap(),
        PackURI::new("/xl/xmlMaps1.xml").unwrap()
    );
}

#[test]
fn writes_real_poi_xml_maps_package_without_resolving_schema_payloads() {
    let mut package = OpcPackage::from_bytes(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/CustomXMLMappings.xlsx"
    ))
    .unwrap();
    let (value, conformance) = load_from_package_with_conformance(&package)
        .unwrap()
        .unwrap();
    store_in_package(&mut package, &value, conformance).unwrap();

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("xml-maps.xlsx");
    package.save(&path).unwrap();
    let reopened = OpcPackage::open(&path).unwrap();
    assert_eq!(
        load_from_package_with_conformance(&reopened).unwrap(),
        Some((value, conformance))
    );
}

#[test]
fn package_xml_maps_mutators_reject_invalid_existing_graphs_before_replacement() {
    let value = fixture_info();
    let mut wrong_content_type = synthetic_package(REL, false, ct::SML_STYLES, false);
    let part_name = PackURI::new("/xl/xmlMaps.xml").unwrap();
    let original = wrong_content_type
        .get_part(&part_name)
        .unwrap()
        .blob()
        .to_vec();
    assert!(
        store_in_package(
            &mut wrong_content_type,
            &value,
            XmlMapConformance::Transitional,
        )
        .is_err()
    );
    assert_eq!(
        wrong_content_type.get_part(&part_name).unwrap().blob(),
        original
    );
    assert!(remove_from_package(&mut wrong_content_type).is_err());

    let mut duplicate = synthetic_package(REL, false, CONTENT_TYPE, false);
    duplicate
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            REL.into(),
            "xmlMaps.xml".into(),
            "rIdDuplicateXmlMaps".into(),
            false,
        );
    assert!(store_in_package(&mut duplicate, &value, XmlMapConformance::Transitional).is_err());
    assert!(remove_from_package(&mut duplicate).is_err());

    let mut external = synthetic_package(REL, true, CONTENT_TYPE, false);
    assert!(store_in_package(&mut external, &value, XmlMapConformance::Transitional).is_err());
    assert!(remove_from_package(&mut external).is_err());

    let mut outbound = synthetic_package(REL, false, CONTENT_TYPE, true);
    assert!(store_in_package(&mut outbound, &value, XmlMapConformance::Transitional).is_err());
    assert!(remove_from_package(&mut outbound).is_err());

    let mut root_relationship = workbook_package();
    root_relationship.relate_to("xl/xmlMaps.xml", REL);
    assert!(
        store_in_package(
            &mut root_relationship,
            &value,
            XmlMapConformance::Transitional,
        )
        .is_err()
    );
}

fn package_with_source(conformance: XmlMapConformance) -> (OpcPackage, XmlMapInfo) {
    let mut package = workbook_package();
    let namespace = if conformance.is_strict() {
        std::str::from_utf8(STRICT_NS).unwrap()
    } else {
        std::str::from_utf8(NS).unwrap()
    };
    let raw = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><MapInfo xmlns="{namespace}" xmlns:future="urn:litchi:future" SelectionNamespaces="xmlns:xs='http://www.w3.org/2001/XMLSchema'"><Schema ID="schema-1" SchemaRef="urn:litchi:example" Namespace="urn:litchi:example"><x:payload xmlns:x="urn:litchi:payload"><x:future marker="keep"/></x:payload></Schema><Map ID="1" Name="Example map" RootElement="example" SchemaID="schema-1" ShowImportExportValidationErrors="true" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"><DataBinding DataBindingName="inert binding" FileBinding="true" ConnectionID="7" DataBindingLoadMode="1"><x:binding xmlns:x="urn:litchi:binding"/></DataBinding></Map></MapInfo>"#
    );
    let parsed = XmlMapInfo::parse(raw.as_bytes()).unwrap();
    store_in_package(&mut package, &parsed, conformance).unwrap();
    let part_name = package
        .main_document_part()
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == REL || relationship.reltype() == STRICT_REL)
        .unwrap()
        .target_partname()
        .unwrap();
    package
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(raw.into_bytes());
    assert_eq!(
        load_from_package_with_conformance(&package).unwrap(),
        Some((parsed.clone(), conformance))
    );
    (package, parsed)
}

#[test]
fn transaction_crud_preserves_source_payload_and_round_trips_both_conformances() {
    for conformance in [XmlMapConformance::Transitional, XmlMapConformance::Strict] {
        let mut package = workbook_package();
        let value = fixture_info();
        let mut create = Transaction::new_with_conformance(&mut package, conformance).unwrap();
        assert!(create.info().is_none());
        assert!(create.set(value.clone()).unwrap());
        let commit = create.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(commit.snapshot().conformance(), conformance);
        assert_eq!(
            load_from_package_with_conformance(&package).unwrap(),
            Some((value.clone(), conformance))
        );

        let mut edit = Transaction::new(&mut package).unwrap();
        assert!(
            edit.edit_map(1, |map| {
                map.name = "edited map".into();
                Ok(())
            })
            .unwrap()
        );
        let commit = edit.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(commit.snapshot().info().unwrap().maps[0].name, "edited map");

        let mut remove_map = Transaction::new(&mut package).unwrap();
        assert!(remove_map.remove_map(1).is_err());
        assert_eq!(remove_map.info().unwrap().maps.len(), 1);

        let mut remove = Transaction::new(&mut package).unwrap();
        assert!(remove.remove().unwrap().is_some());
        assert!(remove.commit().unwrap().changed());
        assert!(Snapshot::load(&package).unwrap().is_empty());
    }
}

#[test]
fn transaction_noop_inverse_and_stale_patch_are_atomic() {
    let (mut package, value) = package_with_source(XmlMapConformance::Transitional);
    let original_bytes = PackageWriter::to_bytes(&package).unwrap();
    let before = Snapshot::load(&package).unwrap();
    let original_relationships = package.main_document_part().unwrap().rels().to_xml();

    let mut no_op = Transaction::new(&mut package).unwrap();
    assert!(!no_op.edit_map(1, |_map| Ok(())).unwrap());
    let commit = no_op.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), original_bytes);

    let mut changed = Transaction::new(&mut package).unwrap();
    changed
        .edit_map(1, |map| {
            map.name = "changed".into();
            Ok(())
        })
        .unwrap();
    let patch = changed.commit().unwrap().patch().clone();
    assert!(
        patch
            .after()
            .source_xml()
            .unwrap()
            .windows(b"xmlns:future=\"urn:litchi:future\"".len())
            .any(|window| window == b"xmlns:future=\"urn:litchi:future\"")
    );
    assert!(
        patch
            .after()
            .source_xml()
            .unwrap()
            .windows(b"<x:future marker=\"keep\"/>".len())
            .any(|window| window == b"<x:future marker=\"keep\"/>")
    );
    assert_eq!(
        package.main_document_part().unwrap().rels().to_xml(),
        original_relationships
    );
    patch.inverse().apply(&mut package).unwrap();
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), original_bytes);
    assert_eq!(Snapshot::load(&package).unwrap(), before);
    assert_eq!(Snapshot::load(&package).unwrap().info(), Some(&value));

    let mut stale = package.clone();
    let part_name = Snapshot::load(&stale).unwrap().part_name().unwrap().clone();
    stale
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(b"<stale/>".to_vec());
    let stale_bytes = PackageWriter::to_bytes(&stale).unwrap();
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(PackageWriter::to_bytes(&stale).unwrap(), stale_bytes);
}
