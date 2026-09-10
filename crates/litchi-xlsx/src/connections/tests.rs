//! Compatibility and package-boundary coverage for SpreadsheetML connections.

use super::codec::BoundedXml;
use super::model::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP, CORE_NAMESPACE, Conformance,
    MAX_STRING_BYTES, MAX_XML_BYTES, STRICT_NAMESPACE,
};
use super::*;
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

const TEST_QUERY_TABLE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
const TEST_WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

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
fn accepts_processing_instructions_while_still_validating_connections() {
    let xml = format!(
        r#"<?before?><connections xmlns="{CORE_NAMESPACE}"><?inside?><connection id="1" refreshedVersion="0"/><?between?></connections><?after?>"#
    );
    assert_eq!(
        Connections::parse(xml.as_bytes())
            .unwrap()
            .connections
            .len(),
        1
    );

    let malformed = format!(
        r#"<?keep?><connections xmlns="{CORE_NAMESPACE}"><connection id="1"/></connections>"#
    );
    assert!(Connections::parse(malformed.as_bytes()).is_err());
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

fn transaction_connection(id: u32, name: &str) -> Connection {
    Connection {
        id,
        source_file: None,
        odc_file: None,
        keep_alive: None,
        interval: None,
        name: Some(name.into()),
        description: None,
        connection_type: Some(1),
        reconnection_method: None,
        refreshed_version: 7,
        min_refreshable_version: None,
        save_password: None,
        new_connection: None,
        deleted: None,
        only_use_connection_file: None,
        background: None,
        refresh_on_load: None,
        save_data: None,
        credentials: None,
        single_sign_on_id: None,
        database: None,
        olap: None,
        web: None,
        text: None,
        parameters: None,
        extension_xml: None,
    }
}

fn transaction_package(with_connection: bool) -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let mut workbook_part = BlobPart::new(
        workbook.clone(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    );
    package.relate_to(
        "xl/workbook.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    );
    if with_connection {
        workbook_part.rels_mut().add_relationship(
            CONNECTIONS_RELATIONSHIP.into(),
            "connections.xml".into(),
            "rIdConnections".into(),
            false,
        );
    }
    package.add_part(Box::new(workbook_part));
    if with_connection {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/connections.xml").unwrap(),
            CONNECTIONS_CONTENT_TYPE.into(),
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7" name="before"><x:future xmlns:x="urn:future" marker="keep"/></connection></connections>"#
            )
            .into_bytes(),
        )));
    }
    package
}

fn many_connections_source(count: u32) -> Vec<u8> {
    let mut source = format!(r#"<connections xmlns="{CORE_NAMESPACE}">"#);
    for id in 1..=count {
        source.push_str(&format!(
            r#"<connection id="{id}" refreshedVersion="7" name="n{id}"/>"#
        ));
    }
    source.push_str("</connections>");
    source.into_bytes()
}

#[test]
fn typed_transaction_preserves_opaque_connection_xml_and_inverse() {
    let mut package = transaction_package(true);
    let before = Snapshot::load(&package).unwrap();
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(
        transaction
            .edit(1, |connection| {
                connection.name = Some("after".into());
                Ok(())
            })
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let source = package
        .get_part(&PackURI::new("/xl/connections.xml").unwrap())
        .unwrap()
        .blob();
    assert!(
        source
            .windows(b"marker=\"keep\"".len())
            .any(|window| window == b"marker=\"keep\"")
    );
    assert!(
        source
            .windows(b"name=\"after\"".len())
            .any(|window| window == b"name=\"after\"")
    );
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(Snapshot::load(&package).unwrap(), before);
}

#[test]
fn extension_owner_transaction_projection_handles_compact_pi_mce_and_foreign_contexts() {
    let mce = litchi_ooxml_common::mce::NAMESPACE;
    let cases = [
        (
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><extLst/></connection></connections>"#
            ),
            format!(r#"<extLst xmlns="{CORE_NAMESPACE}"><ext uri="u"/></extLst>"#),
            None,
        ),
        (
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><extLst><?old?><ext uri="old"/></extLst><?outside?></connection></connections>"#
            ),
            format!(r#"<extLst xmlns="{CORE_NAMESPACE}"><?new?><ext uri="pi"/></extLst>"#),
            Some("pi"),
        ),
        (
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{mce}" xmlns:u="urn:unsupported"><connection id="1" refreshedVersion="7"><extLst><mc:AlternateContent><mc:Choice Requires="u"><u:old/></mc:Choice><mc:Fallback><ext uri="old"/></mc:Fallback></mc:AlternateContent></extLst></connection></connections>"#
            ),
            format!(
                r#"<extLst xmlns="{CORE_NAMESPACE}" xmlns:mc="{mce}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><u:inactive/></mc:Choice><mc:Fallback><ext uri="mce"/></mc:Fallback></mc:AlternateContent></extLst>"#
            ),
            Some("mce"),
        ),
        (
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{mce}" xmlns:u="urn:ignored" mc:Ignorable="u"><connection id="1" refreshedVersion="7"><extLst/></connection></connections>"#
            ),
            format!(
                r#"<extLst xmlns="{CORE_NAMESPACE}" xmlns:u="urn:ignored"><u:ignored/></extLst>"#
            ),
            Some("ignored"),
        ),
        (
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{mce}" xmlns:u="urn:process" mc:Ignorable="u" mc:ProcessContent="u:wrap"><connection id="1" refreshedVersion="7"><extLst/></connection></connections>"#
            ),
            format!(
                r#"<extLst xmlns="{CORE_NAMESPACE}" xmlns:u="urn:process"><u:wrap><ext uri="process"/></u:wrap></extLst>"#
            ),
            Some("process"),
        ),
        (
            format!(
                r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns="urn:foreign" xmlns:f="urn:foreign"><p:connection id="1" refreshedVersion="7"><p:extLst/></p:connection></p:connections>"#
            ),
            format!(r#"<extLst xmlns="{CORE_NAMESPACE}"><ext uri="foreign"/></extLst>"#),
            Some("foreign"),
        ),
        (
            format!(
                r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns="urn:foreign"><p:connection id="1" refreshedVersion="7"><p:extLst/></p:connection></p:connections>"#
            ),
            format!(r#"<p:extLst xmlns:p="{CORE_NAMESPACE}"><p:ext uri="prefixed"/></p:extLst>"#),
            Some("prefixed"),
        ),
    ];

    for (source, replacement, marker) in cases {
        let mut package = transaction_package(true);
        package
            .get_part_mut(&PackURI::new("/xl/connections.xml").unwrap())
            .unwrap()
            .set_blob(source.into_bytes());
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(
            transaction
                .edit(1, |connection| {
                    connection.extension_xml = Some(replacement.into_bytes());
                    Ok(())
                })
                .unwrap()
        );
        let commit = transaction
            .commit()
            .unwrap_or_else(|error| panic!("extension owner case {marker:?}: {error}"));
        assert!(commit.changed());

        let source = package
            .get_part(&PackURI::new("/xl/connections.xml").unwrap())
            .unwrap()
            .blob();
        let reopened = Snapshot::load(&package).unwrap();
        let extension = reopened.connections().unwrap().connections[0]
            .extension_xml
            .as_deref()
            .unwrap();
        if marker != Some("ignored") {
            assert!(std::str::from_utf8(extension).unwrap().contains("uri="));
        }
        if marker == Some("pi") {
            assert!(
                source
                    .windows(b"<?new?>".len())
                    .any(|window| window == b"<?new?>")
            );
            assert!(
                !extension
                    .windows(b"<?new?>".len())
                    .any(|window| window == b"<?new?>")
            );
            assert!(
                source
                    .windows(b"<?outside?>".len())
                    .any(|window| window == b"<?outside?>")
            );
        }
        if marker == Some("mce") {
            assert!(
                source
                    .windows(b"mc:Choice".len())
                    .any(|window| window == b"mc:Choice")
            );
            assert!(
                std::str::from_utf8(extension)
                    .unwrap()
                    .contains(r#"uri="mce""#)
            );
        }
        if marker == Some("ignored") {
            let source = std::str::from_utf8(source).unwrap();
            let extension = std::str::from_utf8(extension).unwrap();
            assert!(source.contains("u:ignored"));
            assert!(!extension.contains("u:ignored"));
        }
        if marker == Some("process") {
            let source = std::str::from_utf8(source).unwrap();
            let extension = std::str::from_utf8(extension).unwrap();
            assert!(source.contains("u:wrap"));
            assert!(!extension.contains("u:wrap"));
            assert!(extension.contains(r#"uri="process""#));
        }
        if marker == Some("foreign") {
            assert!(
                std::str::from_utf8(source)
                    .unwrap()
                    .contains(r#"xmlns:f="urn:foreign""#)
            );
            assert!(
                std::str::from_utf8(extension)
                    .unwrap()
                    .contains(r#"xmlns:f="urn:foreign""#)
            );
        }
        if marker == Some("prefixed") {
            let source = std::str::from_utf8(source).unwrap();
            let extension = std::str::from_utf8(extension).unwrap();
            assert!(source.contains("p:extLst"));
            assert!(
                source.contains(
                    r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#
                )
            );
            assert!(extension.contains("p:extLst"));
            assert!(extension.contains(r#"uri="prefixed""#));
        }
    }
}

#[test]
fn extension_projection_uses_a_bounded_id_index_for_many_connections() {
    const COUNT: usize = 8_192;
    let mut actual = Connections {
        connections: Vec::new(),
    };
    actual.connections.try_reserve_exact(COUNT).unwrap();
    for id in 0..COUNT as u32 {
        let mut connection = transaction_connection(id, "actual");
        connection.extension_xml = Some(
            format!(r#"<extLst xmlns="{CORE_NAMESPACE}"><ext uri="{id}"/></extLst>"#).into_bytes(),
        );
        actual.connections.push(connection);
    }
    let mut staged = actual.clone();
    for connection in &mut staged.connections {
        connection.extension_xml = Some(
            format!(
                r#"<extLst xmlns="{CORE_NAMESPACE}"><ext uri="staged-{}"/></extLst>"#,
                connection.id
            )
            .into_bytes(),
        );
    }
    codec::normalize_connections_source_projection(&actual, &mut staged).unwrap();
    assert_eq!(staged, actual);
}

#[test]
fn large_transaction_scalar_commit_reopens_with_stable_ids() {
    const COUNT: u32 = 8_192;
    let mut package = transaction_package(true);
    package
        .get_part_mut(&PackURI::new("/xl/connections.xml").unwrap())
        .unwrap()
        .set_blob(many_connections_source(COUNT));
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(
        transaction
            .edit(COUNT, |connection| {
                connection.name = Some("updated".into());
                Ok(())
            })
            .unwrap()
    );
    assert!(transaction.commit().unwrap().changed());
    let reopened = Snapshot::load(&package).unwrap();
    assert_eq!(
        reopened.connections().unwrap().connections.len(),
        COUNT as usize
    );
    assert_eq!(
        reopened
            .connections()
            .unwrap()
            .connections
            .last()
            .unwrap()
            .name
            .as_deref(),
        Some("updated")
    );
}

#[test]
fn large_transaction_reorder_publishes_the_requested_order() {
    const COUNT: u32 = 8_192;
    let mut package = transaction_package(true);
    package
        .get_part_mut(&PackURI::new("/xl/connections.xml").unwrap())
        .unwrap()
        .set_blob(many_connections_source(COUNT));
    let mut transaction = Transaction::new(&mut package).unwrap();
    let mut reordered = transaction.connections().unwrap().clone();
    reordered.connections.reverse();
    assert!(transaction.replace(Some(reordered)).unwrap());
    assert!(transaction.commit().unwrap().changed());
    let reopened = Snapshot::load(&package).unwrap();
    let connections = &reopened.connections().unwrap().connections;
    assert_eq!(connections.first().unwrap().id, COUNT);
    assert_eq!(connections.last().unwrap().id, 1);
}

#[test]
fn transaction_noop_and_stale_failure_are_source_checked() {
    let mut package = transaction_package(true);
    let before = Snapshot::load(&package).unwrap();
    let mut transaction = Transaction::new(&mut package).unwrap();
    assert!(!transaction.edit(1, |_connection| Ok(())).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(Snapshot::load(&package).unwrap(), before);

    let patch = {
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction
            .edit(1, |connection| {
                connection.name = Some("changed".into());
                Ok(())
            })
            .unwrap();
        transaction.commit().unwrap().patch().clone()
    };
    package
        .get_part_mut(&PackURI::new("/xl/connections.xml").unwrap())
        .unwrap()
        .set_blob(b"<connections/>".to_vec());
    assert!(patch.apply(&mut package).is_err());
}

#[test]
fn transaction_creates_and_removes_the_connections_owner() {
    let mut package = transaction_package(false);
    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction
        .set(transaction_connection(9, "created"))
        .unwrap();
    transaction.commit().unwrap();
    assert_eq!(
        Snapshot::load(&package)
            .unwrap()
            .connections()
            .unwrap()
            .connections[0]
            .id,
        9
    );

    let mut transaction = Transaction::new(&mut package).unwrap();
    transaction.remove(9).unwrap();
    transaction.commit().unwrap();
    assert!(Snapshot::load(&package).unwrap().connections().is_none());
    assert!(
        package
            .get_part(&PackURI::new("/xl/connections.xml").unwrap())
            .is_err()
    );
}

fn query_table_package(
    source_content_type: &str,
    target: &str,
    target_external: bool,
    query_content_type: &str,
) -> OpcPackage {
    let mut package = transaction_package(true);
    let worksheet_name = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let mut worksheet = BlobPart::new(
        worksheet_name.clone(),
        source_content_type.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
            .to_vec(),
    );
    worksheet.rels_mut().add_relationship(
        TEST_QUERY_TABLE_RELATIONSHIP.into(),
        target.into(),
        "rIdQueryTable".into(),
        target_external,
    );
    package.add_part(Box::new(worksheet));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/queryTables/queryTable1.xml").unwrap(),
        query_content_type.into(),
        format!(r#"<?query?><queryTable xmlns="{CORE_NAMESPACE}" name="query" connectionId="1"/>"#)
            .into_bytes(),
    )));
    package
}

#[test]
fn query_table_validation_accepts_processing_instructions() {
    let package = query_table_package(
        TEST_WORKSHEET_CONTENT_TYPE,
        "../queryTables/queryTable1.xml",
        false,
        QUERY_TABLE_CONTENT_TYPE,
    );
    let snapshot = Snapshot::load(&package).unwrap();
    assert_eq!(snapshot.query_table_parts().count(), 1);
}

#[test]
fn graph_closure_rejects_foreign_connections_and_invalid_query_table_edges() {
    let mut foreign = transaction_package(true);
    let mut other = BlobPart::new(
        PackURI::new("/xl/other.xml").unwrap(),
        "application/xml".into(),
        b"<other/>".to_vec(),
    );
    other.relate_to("connections.xml", CONNECTIONS_RELATIONSHIP);
    foreign.add_part(Box::new(other));
    assert!(validate_graph(&foreign).is_err());

    let wrong_target = query_table_package(
        TEST_WORKSHEET_CONTENT_TYPE,
        "../other.xml",
        false,
        QUERY_TABLE_CONTENT_TYPE,
    );
    assert!(validate_graph(&wrong_target).is_err());

    let wrong_source = query_table_package(
        "application/xml",
        "../queryTables/queryTable1.xml",
        false,
        QUERY_TABLE_CONTENT_TYPE,
    );
    assert!(validate_graph(&wrong_source).is_err());

    let mut wrong_workbook = transaction_package(true);
    wrong_workbook
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap()
        .set_content_type("application/xml".into())
        .unwrap();
    assert!(validate_graph(&wrong_workbook).is_err());
}

#[test]
fn recognized_owner_targets_reject_query_and_fragment_components() {
    let mut package = transaction_package(true);
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    package
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .remove("rIdConnections");
    package
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .add_relationship(
            CONNECTIONS_RELATIONSHIP.into(),
            "connections.xml#part".into(),
            "rIdConnections".into(),
            false,
        );
    assert!(validate_graph(&package).is_err());
}

#[test]
fn conformance_uses_the_root_namespace_only() {
    let mut package = transaction_package(true);
    let connection = PackURI::new("/xl/connections.xml").unwrap();
    package
        .get_part_mut(&connection)
        .unwrap()
        .set_blob(
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><x:future xmlns:x="urn:future" marker="{STRICT_NAMESPACE}"/></connection></connections>"#
            )
            .into_bytes(),
        );
    assert_eq!(
        Snapshot::load(&package).unwrap().conformance(),
        Conformance::Transitional
    );

    package
        .get_part_mut(&connection)
        .unwrap()
        .set_blob(
            format!(
                r#"<connections xmlns="{STRICT_NAMESPACE}"><connection id="1" refreshedVersion="7"/></connections>"#
            )
            .into_bytes(),
        );
    assert_eq!(
        Snapshot::load(&package).unwrap().conformance(),
        Conformance::Strict
    );
}
