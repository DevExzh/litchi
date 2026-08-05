//! Compatibility and package-boundary coverage for SpreadsheetML connections.

use super::codec::BoundedXml;
use super::model::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP, CORE_NAMESPACE, MAX_STRING_BYTES,
    MAX_XML_BYTES, STRICT_NAMESPACE,
};
use super::*;
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
fn f(b: &[u8]) -> Connections {
    let p = OpcPackage::from_bytes(b).unwrap();
    load_from_package(&p).unwrap().unwrap()
}
fn f_without_broken_thumbnail(b: &[u8]) -> Connections {
    let reader = PhysPkgReader::new(b).unwrap();
    let mut writer = PhysPkgWriter::new();
    for name in reader.member_names().unwrap() {
        if name == "docProps/thumbnail.jpeg" {
            continue;
        }
        let uri = PackURI::new(format!("/{name}")).unwrap();
        let mut data = reader.blob_for(&uri).unwrap();
        if name == "_rels/.rels" {
            let xml = String::from_utf8(data).unwrap();
            data = xml.replace("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\" Target=\"docProps/thumbnail.jpeg\"/>", "").into_bytes();
        }
        writer.write(&uri, &data).unwrap();
    }
    f(&writer.finish().unwrap())
}
#[test]
fn poi_web_paths_are_inert() {
    let v = f(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/56169.xlsx"
    ));
    assert_eq!(v.connections.len(), 3);
    assert!(
        v.connections[0]
            .web
            .as_ref()
            .unwrap()
            .url
            .as_ref()
            .unwrap()
            .starts_with("\\\\snb.ch")
    );
}
#[test]
fn poi_database_mce_and_strict_roundtrip() {
    let v = f(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/ExcelPivotTableSample.xlsx"
    ));
    let db = v.connections[0].database.as_ref().unwrap();
    assert!(db.connection.contains("Microsoft.ACE.OLEDB"));
    assert_eq!(db.command.as_deref(), Some("Office Address List"));
    let x = v.to_xml(true).unwrap();
    assert_eq!(
        Connections::parse(&x).unwrap().connections[0]
            .database
            .as_ref()
            .unwrap()
            .command_type,
        Some(3)
    );
}
#[test]
fn libreoffice_text_import_fields() {
    let v = f_without_broken_thumbnail(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/queryTableExport.xlsx"
    ));
    assert_eq!(v.connections.len(), 2);
    assert_eq!(
        v.connections[0]
            .text
            .as_ref()
            .unwrap()
            .fields
            .as_ref()
            .unwrap()
            .len(),
        2
    );
    assert_eq!(v.connections[1].text.as_ref().unwrap().comma, Some(true));
}
#[test]
fn libreoffice_olap_and_extensions() {
    let v = f(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf66377.xlsx"
    ));
    assert_eq!(
        v.connections[0].olap.as_ref().unwrap().row_drill_count,
        Some(1000)
    );
    assert!(
        std::str::from_utf8(v.connections[1].extension_xml.as_deref().unwrap())
            .unwrap()
            .contains("x15:rangePr")
    );
}
#[test]
fn libreoffice_prefixed_core_namespace() {
    let v = f(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx"
    ));
    assert_eq!(
        v.connections[0].web.as_ref().unwrap().xml_source,
        Some(true)
    );
}
#[test]
fn standards_parameters_tables_strict_and_mce() {
    let xml = format!(
        r#"<connections xmlns="{STRICT_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><connection id="9" refreshedVersion="8" credentials="stored"><webPr htmlFormat="rtf"><tables count="3"><m/><s v="A"/><x v="2"/></tables></webPr><parameters count="1"><parameter name="p" sqlType="4" parameterType="value" double="1.5"/></parameters></connection></mc:Fallback></mc:AlternateContent></connections>"#
    );
    let v = Connections::parse(xml.as_bytes()).unwrap();
    assert_eq!(
        v.connections[0].parameters.as_ref().unwrap()[0].double,
        Some(1.5)
    );
    assert_eq!(Connections::parse(&v.to_xml(false).unwrap()).unwrap(), v);
}
#[test]
fn rejects_malformed_and_unsafe() {
    for xml in [
        format!(r#"<connections xmlns="{CORE_NAMESPACE}"/>"#),
        format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"><parameters count="2"><parameter/></parameters></connection></connections>"#
        ),
        format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"><parameters><parameter double="NaN"/></parameters></connection></connections>"#
        ),
        format!(
            r#"<!DOCTYPE x><connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"/></connections>"#
        ),
    ] {
        assert!(
            Connections::parse(xml.as_bytes()).is_err(),
            "accepted {xml}"
        );
    }
}

#[test]
fn failed_add_preserves_the_existing_connection_set() {
    let xml = format!(
        r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"/></connections>"#
    );
    let mut value = Connections::parse(xml.as_bytes()).unwrap();
    let before = value.clone();

    let mut invalid = value.connections[0].clone();
    invalid.id = 2;
    invalid.description = Some("x".repeat(MAX_STRING_BYTES + 1));
    assert!(value.add(invalid).is_err());
    assert_eq!(value, before);

    let duplicate = value.connections[0].clone();
    assert!(value.add(duplicate).is_err());
    assert_eq!(value, before);
}

#[test]
fn failed_reorder_preserves_the_existing_connection_order() {
    let xml = format!(
        r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"/><connection id="2" refreshedVersion="0"/></connections>"#
    );
    let mut value = Connections::parse(xml.as_bytes()).unwrap();
    let before = value.clone();

    assert!(value.reorder(&[1, 1]).is_err());
    assert_eq!(value, before);

    value.reorder(&[2, 1]).unwrap();
    assert_eq!(
        value
            .connections
            .iter()
            .map(|connection| connection.id)
            .collect::<Vec<_>>(),
        vec![2, 1]
    );
}

#[test]
fn bounded_serializer_rejects_oversized_output_before_appending() {
    let mut output = BoundedXml::new();
    let error = output
        .push_bytes(&vec![b'x'; MAX_XML_BYTES + 1])
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "serialized connections part exceeds 16 MiB"
    );
    assert!(output.bytes.is_empty());
}

fn package(content_type: &str, external: bool, outbound: bool) -> OpcPackage {
    let mut p = OpcPackage::new();
    let wb = PackURI::new("/xl/workbook.xml").unwrap();
    let mut w = BlobPart::new(
        wb,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        Vec::new(),
    );
    if external {
        w.rels_mut().add_relationship(
            CONNECTIONS_RELATIONSHIP.into(),
            "https://example.invalid/c.xml".into(),
            "rId1".into(),
            true,
        );
    } else {
        w.relate_to("connections.xml", CONNECTIONS_RELATIONSHIP);
    }
    p.relate_to(
        "xl/workbook.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    );
    p.add_part(Box::new(w));
    let mut c=BlobPart::new(PackURI::new("/xl/connections.xml").unwrap(),content_type.into(),format!(r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="0"/></connections>"#).into_bytes());
    if outbound {
        c.relate_to("other.xml", "urn:forbidden");
    }
    p.add_part(Box::new(c));
    p
}
#[test]
fn rejects_external_wrong_content_and_outbound_package_edges() {
    assert!(load_from_package(&package(CONNECTIONS_CONTENT_TYPE, true, false)).is_err());
    assert!(load_from_package(&package("application/xml", false, false)).is_err());
    assert!(load_from_package(&package(CONNECTIONS_CONTENT_TYPE, false, true)).is_err());
}
