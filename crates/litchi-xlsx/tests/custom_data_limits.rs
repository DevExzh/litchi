#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "The bounded OPC fixtures are authored finite probes whose construction failures are test failures."
)]

//! Public host-limit coverage for the source-bound XLSX Custom Data owner.
//!
//! The XML and relationship members are written into a physical ZIP and then
//! reopened through the public XLSX package. This keeps the tests on public
//! package, transaction, and patch APIs while proving that limits are enforced
//! at the host boundary after source provenance is captured.

use std::collections::BTreeMap;

use litchi_core::Resource;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_xlsx::{CustomData, CustomDataLimits, Error, Package, custom_data::RemovalDisposition};
use quick_xml::{events::Event, reader::NsReader};
use soapberry_zip::office::ArchiveReader;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const STYLES: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CONNECTIONS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const PROPERTIES_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
const DATA_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
const MARKUP_COMPATIBILITY: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
const DATA_CONTENT_TYPE: &str = "application/binary";
const WORKBOOK_URI: &str = "/xl/workbook.xml";
const CONNECTIONS_URI: &str = "/xl/connections.xml";
const QUERY_TABLE_URI: &str = "/xl/queryTables/queryTable1.xml";
const QUERY_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
const QUERY_TABLE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";

const EMBEDDED_DATA_EXTENSION: &str = "{D79990A0-CA42-45E3-83F4-45C500A0EAA5}";

#[derive(Debug, Clone)]
struct Fixture {
    package: Package,
    properties: BTreeMap<String, Vec<u8>>,
    connections: Option<Vec<u8>>,
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("fixture URI is valid")
}

fn source_bytes(value: String) -> Vec<u8> {
    value.replace("\\r\n", "\r\n").into_bytes()
}

fn properties_xml(uid: &str, formatted: bool, extension: bool) -> Vec<u8> {
    if !extension {
        return format!(r#"<alias:datastoreItem xmlns:alias="{X14}" id="{uid}"/>"#).into_bytes();
    }
    if formatted {
        return source_bytes(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>\r
<!-- retained before root -->\r
<alias:datastoreItem xmlns:alias="{X14}" xmlns:s="{SML}" xmlns:v="urn:synthetic-vendor" id="{uid}">\r
  <?root-pi?>\r
  <alias:extLst>\r
    <?extension-pi?>\r
    <s:ext uri="urn:synthetic">\r
      <v:opaque v:raw="&amp;raw">\r
        <v:leaf><![CDATA[<? retained-looking text ]]></v:leaf>\r
        <!-- opaque comment -->\r
      </v:opaque>\r
    </s:ext>\r
  </alias:extLst>\r
  <!-- retained after extension -->\r
</alias:datastoreItem>\r
"#
        ));
    }
    source_bytes(format!(
        r#"<alias:datastoreItem xmlns:alias="{X14}" xmlns:s="{SML}" xmlns:v="urn:synthetic-vendor" id="{uid}"><?root-pi?><alias:extLst><?extension-pi?><s:ext uri="urn:synthetic"><v:opaque v:raw="&amp;raw"><v:leaf><![CDATA[<? retained-looking text ]]></v:leaf><!-- opaque comment --></v:opaque></s:ext></alias:extLst><!-- retained after extension --></alias:datastoreItem>"#
    ))
}

fn connections_xml(references: &[&str], formatted: bool) -> Vec<u8> {
    let mut output = if formatted {
        format!(
            r#"<s:connections xmlns:s="{SML}" xmlns:a="{X14}" xmlns:q="urn:synthetic-opaque">\r
"#
        )
    } else {
        format!(r#"<s:connections xmlns:s="{SML}" xmlns:a="{X14}" xmlns:q="urn:synthetic-opaque">"#)
    };
    for (ordinal, reference) in references.iter().enumerate() {
        let id = ordinal + 7;
        if formatted {
            output.push_str(&format!(
                r#"  <s:connection id="{id}" type="5" refreshedVersion="3">\r
    <s:extLst>\r
      <s:ext uri="{EMBEDDED_DATA_EXTENSION}">\r
        <a:connection embeddedDataId="{reference}" culture="en-US"><q:opaque marker="keep"/></a:connection>\r
      </s:ext>\r
    </s:extLst>\r
  </s:connection>\r
"#
            ));
        } else {
            output.push_str(&format!(
                r#"<s:connection id="{id}" type="5"><s:extLst><s:ext uri="{EMBEDDED_DATA_EXTENSION}"><a:connection embeddedDataId="{reference}"/></s:ext></s:extLst></s:connection>"#
            ));
        }
    }
    if formatted {
        output.push_str("</s:connections>\r\n");
    } else {
        output.push_str("</s:connections>");
    }
    source_bytes(output)
}

fn shallow_connections_xml(reference: &str) -> Vec<u8> {
    source_bytes(format!(
        r#"<s:connections xmlns:s="{SML}" xmlns:a="{X14}"><s:connection id="7" type="5" refreshedVersion="3"><s:extLst><s:ext uri="{EMBEDDED_DATA_EXTENSION}"><a:connection embeddedDataId="{reference}"/></s:ext></s:extLst></s:connection></s:connections>"#
    ))
}

fn connections_with_opaque_cdata(reference: &str, cdata: &str) -> Vec<u8> {
    let mut output =
        String::from_utf8(connections_xml(&[reference], true)).expect("connection XML is UTF-8");
    let marker = r#"<q:opaque marker="keep"/>"#;
    let replacement = format!(r#"<q:opaque marker="keep"><![CDATA[{cdata}]]></q:opaque>"#);
    assert!(
        output.contains(marker),
        "connection XML has its opaque marker"
    );
    output = output.replace(marker, &replacement);
    output.into_bytes()
}

fn connections_with_escaped_x14_namespace(reference: &str) -> Vec<u8> {
    let mut output =
        String::from_utf8(connections_xml(&[reference], true)).expect("connection XML is UTF-8");
    let marker = format!(r#"xmlns:a="{X14}""#);
    let escaped = X14.replace("2009", "20&#48;09");
    let replacement = format!(r#"xmlns:a="{escaped}""#);
    assert!(
        output.contains(&marker),
        "connection XML has the X14 namespace"
    );
    output = output.replace(&marker, &replacement);
    output.into_bytes()
}

fn connections_with_root_namespaces(references: &[&str], declarations: &[(&str, &str)]) -> Vec<u8> {
    let mut output =
        String::from_utf8(connections_xml(references, true)).expect("connection XML is UTF-8");
    let marker = r#" xmlns:q="urn:synthetic-opaque""#;
    let insertion = declarations
        .iter()
        .map(|(prefix, namespace)| format!(r#" xmlns:{prefix}="{namespace}""#))
        .collect::<String>();
    let position = output
        .find(marker)
        .map(|position| position + marker.len())
        .expect("connection root has the opaque namespace");
    output.insert_str(position, &insertion);
    output.into_bytes()
}

fn add_relationship(package: &mut OpcPackage, source: &str, reltype: &str, target: &str, id: &str) {
    package
        .get_part_mut(&uri(source))
        .expect("fixture relationship source exists")
        .rels_mut()
        .add_relationship(reltype.to_owned(), target.to_owned(), id.to_owned(), false);
}

fn workbook_relationships(storage_count: usize, has_connections: bool) -> Vec<u8> {
    let mut output = format!(
        r#"<?xml version='1.0'?>\r
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>\r
 <!-- retained workbook relationship source -->\r
 <r:Relationship Target='worksheets/sheet1.xml' Type='{WORKSHEET}' Id='rId1'></r:Relationship>\r
 <r:Relationship Target='styles.xml' Type='{STYLES}' Id='rId2'></r:Relationship>\r
"#
    );
    for index in 1..=storage_count {
        output.push_str(&format!(
            " <r:Relationship Target='customData/props{index}.xml' Type='{PROPERTIES_RELATIONSHIP}' Id='rIdCustomDataProps{index}'></r:Relationship>\r\n"
        ));
    }
    if has_connections {
        output.push_str(&format!(
            " <r:Relationship Target='connections.xml' Type='{CONNECTIONS_RELATIONSHIP}' Id='rIdConnections'></r:Relationship>\r\n"
        ));
    }
    output.push_str("</r:Relationships>\r\n");
    source_bytes(output)
}

fn properties_relationships(index: usize) -> Vec<u8> {
    source_bytes(format!(
        r#"<?xml version='1.0'?>\r
<p:Relationships xmlns:p='http://schemas.openxmlformats.org/package/2006/relationships'>\r
 <?keep-properties?>\r
 <p:Relationship Target='data{index}.bin' Type='{DATA_RELATIONSHIP}' Id='rIdCustomData{index}'></p:Relationship>\r
</p:Relationships>\r
"#
    ))
}

fn empty_relationships() -> Vec<u8> {
    source_bytes(String::from(
        r#"<?xml version='1.0'?>\r
<d:Relationships xmlns:d='http://schemas.openxmlformats.org/package/2006/relationships'>\r
 <!-- explicit empty payload relationship member -->\r
</d:Relationships>\r
"#,
    ))
}

fn physical_source(
    package: &OpcPackage,
    relationship_overrides: &BTreeMap<String, Vec<u8>>,
) -> Vec<u8> {
    physical_source_with_content_types(package, relationship_overrides, None)
}

fn physical_source_with_content_types(
    package: &OpcPackage,
    relationship_overrides: &BTreeMap<String, Vec<u8>>,
    content_types_override: Option<Vec<u8>>,
) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    let mut parts = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));

    let mut generated_content_types = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\r
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\r
"#,
    );
    for part in &parts {
        generated_content_types.push_str(&format!(
            " <Override PartName=\"{}\" ContentType=\"{}\"/>\r\n",
            part.partname(),
            part.content_type()
        ));
        writer
            .write_stored(
                part.partname().as_str().trim_start_matches('/'),
                part.blob(),
            )
            .expect("fixture part writes");

        let rels_name = part
            .partname()
            .rels_uri()
            .expect("fixture part has a relationship URI")
            .as_str()
            .to_owned();
        let owner = part.partname().as_str().to_owned();
        if let Some(bytes) = relationship_overrides.get(&owner) {
            writer
                .write_stored(rels_name.trim_start_matches('/'), bytes)
                .expect("fixture relationship override writes");
        } else if !part.rels().is_empty() {
            writer
                .write_stored(
                    rels_name.trim_start_matches('/'),
                    part.rels().to_xml().as_bytes(),
                )
                .expect("fixture relationships write");
        }
    }
    generated_content_types.push_str("</Types>\r\n");
    let content_types =
        content_types_override.unwrap_or_else(|| source_bytes(generated_content_types));
    writer
        .write_stored("[Content_Types].xml", &content_types)
        .expect("fixture content types write");
    writer
        .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
        .expect("fixture package relationships write");
    writer.finish_to_bytes().expect("fixture ZIP closes")
}

fn generated_content_types(
    package: &OpcPackage,
    expanded_overrides: bool,
    comment_count: usize,
) -> Vec<u8> {
    let mut parts = package.iter_parts().collect::<Vec<_>>();
    parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut output = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\r
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\r
"#,
    );
    for (index, part) in parts.iter().enumerate() {
        for comment in 0..comment_count {
            output.push_str(&format!(
                " <!-- retained content-types comment {index}-{comment} -->\r\n"
            ));
        }
        if expanded_overrides {
            output.push_str(&format!(
                " <Override PartName=\"{}\" ContentType=\"{}\"></Override>\r\n",
                part.partname(),
                part.content_type()
            ));
        } else {
            output.push_str(&format!(
                " <Override PartName=\"{}\" ContentType=\"{}\"/>\r\n",
                part.partname(),
                part.content_type()
            ));
        }
    }
    output.push_str("</Types>\r\n");
    source_bytes(output)
}

fn synthetic_fixture(
    storage_ids: &[&str],
    references: &[&str],
    payloads: &[&[u8]],
    formatted: bool,
    extension: bool,
) -> Fixture {
    assert_eq!(storage_ids.len(), payloads.len());
    let mut package = Package::create()
        .expect("minimal XLSX package creates")
        .into_plain_opc();
    let mut properties = BTreeMap::new();

    for (index, (id, payload)) in storage_ids.iter().zip(payloads).enumerate() {
        let ordinal = index + 1;
        let properties_uri = format!("/xl/customData/props{ordinal}.xml");
        let data_uri = format!("/xl/customData/data{ordinal}.bin");
        let property_bytes = properties_xml(id, formatted, extension);
        package.add_part(Box::new(BlobPart::new(
            uri(&properties_uri),
            PROPERTIES_CONTENT_TYPE.to_owned(),
            property_bytes.clone(),
        )));
        package.add_part(Box::new(BlobPart::new(
            uri(&data_uri),
            DATA_CONTENT_TYPE.to_owned(),
            payload.to_vec(),
        )));
        add_relationship(
            &mut package,
            &properties_uri,
            DATA_RELATIONSHIP,
            &format!("data{ordinal}.bin"),
            &format!("rIdCustomData{ordinal}"),
        );
        properties.insert(properties_uri, property_bytes);
    }

    let connection_bytes = (!references.is_empty()).then(|| connections_xml(references, formatted));
    add_relationships_to_workbook(&mut package, storage_ids.len(), connection_bytes.is_some());
    if let Some(bytes) = &connection_bytes {
        package.add_part(Box::new(BlobPart::new(
            uri(CONNECTIONS_URI),
            CONNECTIONS_CONTENT_TYPE.to_owned(),
            bytes.clone(),
        )));
    }

    let mut relationship_overrides = BTreeMap::new();
    relationship_overrides.insert(
        WORKBOOK_URI.to_owned(),
        workbook_relationships(storage_ids.len(), connection_bytes.is_some()),
    );
    for index in 1..=storage_ids.len() {
        relationship_overrides.insert(
            format!("/xl/customData/props{index}.xml"),
            properties_relationships(index),
        );
        relationship_overrides.insert(
            format!("/xl/customData/data{index}.bin"),
            empty_relationships(),
        );
    }
    let bytes = physical_source(&package, &relationship_overrides);
    let package = Package::from_bytes(bytes).expect("physical XLSX fixture reopens");
    Fixture {
        package,
        properties,
        connections: connection_bytes,
    }
}

fn heavy_workbook_relationships(storage_count: usize) -> Vec<u8> {
    let mut output = String::from_utf8(workbook_relationships(storage_count, false))
        .expect("workbook relationship fixture is UTF-8");
    let comments = (0..48)
        .map(|index| format!(" <!-- retained relationship comment {index} -->\r\n"))
        .collect::<String>();
    let insertion = output
        .find(" <r:Relationship")
        .expect("workbook relationship has a relationship element");
    output.insert_str(insertion, &comments);
    output.into_bytes()
}

fn relationship_heavy_fixture() -> Fixture {
    let mut package = Package::create()
        .expect("minimal XLSX package creates")
        .into_plain_opc();
    let property_uri = "/xl/customData/props1.xml";
    let data_uri = "/xl/customData/data1.bin";
    let property_bytes = properties_xml("uid-A", false, false);
    package.add_part(Box::new(BlobPart::new(
        uri(property_uri),
        PROPERTIES_CONTENT_TYPE.to_owned(),
        property_bytes.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(data_uri),
        DATA_CONTENT_TYPE.to_owned(),
        b"payload".to_vec(),
    )));
    add_relationship(
        &mut package,
        property_uri,
        DATA_RELATIONSHIP,
        "data1.bin",
        "rIdCustomData1",
    );
    add_relationships_to_workbook(&mut package, 1, false);

    let mut relationship_overrides = BTreeMap::new();
    relationship_overrides.insert(WORKBOOK_URI.to_owned(), heavy_workbook_relationships(1));
    let bytes = physical_source(&package, &relationship_overrides);
    Fixture {
        package: Package::from_bytes(bytes).expect("relationship fixture reopens"),
        properties: BTreeMap::from([(property_uri.to_owned(), property_bytes)]),
        connections: None,
    }
}

fn content_types_heavy_fixture() -> Fixture {
    content_types_heavy_fixture_with_overrides(true)
}

fn content_types_self_closed_heavy_fixture() -> Fixture {
    content_types_heavy_fixture_with_overrides(false)
}

fn content_types_heavy_fixture_with_overrides(expanded_overrides: bool) -> Fixture {
    let mut package = Package::create()
        .expect("minimal XLSX package creates")
        .into_plain_opc();
    let property_uri = "/xl/customData/props1.xml";
    let data_uri = "/xl/customData/data1.bin";
    let property_bytes = properties_xml("uid-A", false, false);
    package.add_part(Box::new(BlobPart::new(
        uri(property_uri),
        PROPERTIES_CONTENT_TYPE.to_owned(),
        property_bytes.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(data_uri),
        DATA_CONTENT_TYPE.to_owned(),
        b"payload".to_vec(),
    )));
    add_relationship(
        &mut package,
        property_uri,
        DATA_RELATIONSHIP,
        "data1.bin",
        "rIdCustomData1",
    );
    add_relationships_to_workbook(&mut package, 1, false);
    let content_types = generated_content_types(&package, expanded_overrides, 8);
    let bytes = physical_source_with_content_types(&package, &BTreeMap::new(), Some(content_types));
    Fixture {
        package: Package::from_bytes(bytes).expect("content-types fixture reopens"),
        properties: BTreeMap::from([(property_uri.to_owned(), property_bytes)]),
        connections: None,
    }
}

fn query_table_xml(nested_depth: usize, with_mce: bool) -> Vec<u8> {
    let mut output = format!(r#"<queryTable xmlns="{SML}" name="Query1" connectionId="7""#);
    if with_mce {
        output.push_str(&format!(
            r#" xmlns:mc="{MARKUP_COMPATIBILITY}" xmlns:x="urn:unsupported" mc:Ignorable="x""#
        ));
    }
    output.push('>');
    if with_mce {
        output.push_str(
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback>"#,
        );
    }
    for _ in 0..nested_depth {
        output.push_str("<opaque>");
    }
    for _ in 0..nested_depth {
        output.push_str("</opaque>");
    }
    if with_mce {
        output.push_str("</mc:Fallback></mc:AlternateContent>");
    }
    output.push_str("</queryTable>");
    output.into_bytes()
}

fn query_table_xml_with_empty_child(nested_depth: usize) -> Vec<u8> {
    let mut output = format!(r#"<queryTable xmlns="{SML}" name="Query1" connectionId="7">"#);
    for _ in 0..nested_depth {
        output.push_str("<opaque>");
    }
    output.push_str("<opaque/>");
    for _ in 0..nested_depth {
        output.push_str("</opaque>");
    }
    output.push_str("</queryTable>");
    output.into_bytes()
}

fn query_table_fixture(query_table: Vec<u8>) -> Fixture {
    query_table_fixture_with_connections(query_table, connections_xml(&["uid-A"], true))
}

fn query_table_fixture_with_connections(
    query_table: Vec<u8>,
    connection_bytes: Vec<u8>,
) -> Fixture {
    let mut package = Package::create()
        .expect("minimal XLSX package creates")
        .into_plain_opc();
    let property_uri = "/xl/customData/props1.xml";
    let data_uri = "/xl/customData/data1.bin";
    let property_bytes = properties_xml("uid-A", false, false);
    package.add_part(Box::new(BlobPart::new(
        uri(property_uri),
        PROPERTIES_CONTENT_TYPE.to_owned(),
        property_bytes.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(data_uri),
        DATA_CONTENT_TYPE.to_owned(),
        b"payload".to_vec(),
    )));
    add_relationship(
        &mut package,
        property_uri,
        DATA_RELATIONSHIP,
        "data1.bin",
        "rIdCustomData1",
    );
    add_relationships_to_workbook(&mut package, 1, true);
    package.add_part(Box::new(BlobPart::new(
        uri(CONNECTIONS_URI),
        CONNECTIONS_CONTENT_TYPE.to_owned(),
        connection_bytes.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(QUERY_TABLE_URI),
        QUERY_TABLE_CONTENT_TYPE.to_owned(),
        query_table,
    )));
    add_relationship(
        &mut package,
        "/xl/worksheets/sheet1.xml",
        QUERY_TABLE_RELATIONSHIP,
        "../queryTables/queryTable1.xml",
        "rIdQueryTable",
    );
    let bytes = physical_source(&package, &BTreeMap::new());
    Fixture {
        package: Package::from_bytes(bytes).expect("query-table fixture reopens"),
        properties: BTreeMap::from([(property_uri.to_owned(), property_bytes)]),
        connections: Some(connection_bytes),
    }
}

fn signed_fixture() -> Fixture {
    let base = single_fixture();
    let mut package = base.package.clone().into_plain_opc();
    let origin = uri("/_xmlsignatures/origin.sigs");
    package.add_part(Box::new(BlobPart::new(
        origin,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        b"<origin/>".to_vec(),
    )));
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rIdSignature".to_owned(),
        false,
    );
    Fixture {
        package: Package::from_bytes(physical_source(&package, &BTreeMap::new()))
            .expect("signed fixture reopens"),
        properties: base.properties,
        connections: base.connections,
    }
}

fn foreign_inbound_fixture() -> Fixture {
    let base = single_fixture();
    let mut package = base.package.clone().into_plain_opc();
    let foreign = "/xl/customData/foreign.xml";
    package.add_part(Box::new(BlobPart::new(
        uri(foreign),
        "application/xml".to_owned(),
        b"<foreign/>".to_vec(),
    )));
    add_relationship(
        &mut package,
        foreign,
        "urn:foreign-custom-data-edge",
        "props1.xml",
        "rIdForeign",
    );
    Fixture {
        package: Package::from_bytes(physical_source(&package, &BTreeMap::new()))
            .expect("foreign inbound fixture reopens"),
        properties: base.properties,
        connections: base.connections,
    }
}

fn add_relationships_to_workbook(
    package: &mut OpcPackage,
    storage_count: usize,
    has_connections: bool,
) {
    for index in 1..=storage_count {
        add_relationship(
            package,
            WORKBOOK_URI,
            PROPERTIES_RELATIONSHIP,
            &format!("customData/props{index}.xml"),
            &format!("rIdCustomDataProps{index}"),
        );
    }
    if has_connections {
        add_relationship(
            package,
            WORKBOOK_URI,
            CONNECTIONS_RELATIONSHIP,
            "connections.xml",
            "rIdConnections",
        );
    }
}

fn single_fixture() -> Fixture {
    synthetic_fixture(
        &["uid-A"],
        &["uid-A"],
        &[b"\0opaque\xffpayload\n"],
        true,
        true,
    )
}

fn single_fixture_with_connection_relationship_member(explicit: bool) -> Fixture {
    let base = single_fixture();
    let package = base.package.clone().into_plain_opc();
    let mut overrides = BTreeMap::new();
    if explicit {
        overrides.insert(CONNECTIONS_URI.to_owned(), empty_relationships());
    }
    Fixture {
        package: Package::from_bytes(physical_source(&package, &overrides))
            .expect("connection relationship fixture reopens"),
        properties: base.properties,
        connections: base.connections,
    }
}

fn repeated_reference_fixture(count: usize) -> Fixture {
    let references = vec!["uid-A"; count];
    synthetic_fixture(
        &["uid-A"],
        &references,
        &[b"\0opaque\xffpayload\n"],
        true,
        true,
    )
}

fn fixture_with_connections_xml(xml: Vec<u8>) -> Fixture {
    let base = single_fixture();
    let mut package = base.package.clone().into_plain_opc();
    package
        .get_part_mut(&uri(CONNECTIONS_URI))
        .expect("connections part exists")
        .set_blob(xml.clone());
    Fixture {
        package: Package::from_bytes(physical_source(&package, &BTreeMap::new()))
            .expect("custom namespace fixture reopens"),
        properties: base.properties,
        connections: Some(xml),
    }
}

fn minimal_fixture(id: &str) -> Fixture {
    synthetic_fixture(&[id], &[], &[b"payload"], false, false)
}

fn two_storage_fixture() -> Fixture {
    synthetic_fixture(
        &["uid-A", "uid-B"],
        &[],
        &[b"0123456789abcdefghijklmnopqrstuv", b"short"],
        true,
        false,
    )
}

fn shrinking_storage_fixture() -> Fixture {
    synthetic_fixture(
        &["uid-A", "uid-B"],
        &[],
        &[
            b"0123456789abcdefghijklmnopqrstuv0123456789abcdefghijklmnopqrstuv",
            b"short",
        ],
        true,
        false,
    )
}

fn part_bytes(package: &Package, part_name: &str) -> Vec<u8> {
    package
        .clone()
        .into_plain_opc()
        .get_part(&uri(part_name))
        .expect("fixture part exists")
        .blob()
        .to_vec()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    package
        .to_plain_bytes()
        .expect("fixture package serializes")
}

fn archive_members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).expect("package bytes are a valid ZIP");
    let names = archive.file_names().map(str::to_owned).collect::<Vec<_>>();
    names
        .into_iter()
        .map(|name| {
            let member = archive.read(&name).expect("ZIP member reads");
            (name, member)
        })
        .collect()
}

fn logical_package_size(package: &Package) -> usize {
    let opc = package.clone().into_plain_opc();
    let root = uri("/");
    let mut total = opc
        .source_content_types()
        .expect("source content types exist")
        .bytes()
        .len();
    total += opc
        .source_relationships(&root)
        .expect("source package relationships exist")
        .bytes()
        .len();
    for part in opc
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        total += part.blob().len();
        total += opc
            .source_relationships(part.partname())
            .expect("source relationship member exists")
            .bytes()
            .len();
    }
    total
}

fn relationship_xml_size(package: &Package, owner: &str) -> usize {
    package
        .clone()
        .into_plain_opc()
        .source_relationships(&uri(owner))
        .expect("source relationship member exists")
        .bytes()
        .len()
}

fn max_relationship_xml_size(package: &Package) -> usize {
    let opc = package.clone().into_plain_opc();
    let root = uri("/");
    let mut maximum = opc
        .source_relationships(&root)
        .expect("source package relationships exist")
        .bytes()
        .len();
    for part in opc.iter_parts() {
        maximum = maximum.max(
            opc.source_relationships(part.partname())
                .expect("source relationship member exists")
                .bytes()
                .len(),
        );
    }
    maximum
}

fn relationship_count(package: &Package) -> usize {
    let opc = package.clone().into_plain_opc();
    opc.rels().iter().count()
        + opc
            .iter_parts()
            .map(|part| part.rels().iter().count())
            .sum::<usize>()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XmlMetrics {
    events: usize,
    depth: usize,
}

fn xml_metrics(bytes: &[u8]) -> XmlMetrics {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut maximum_depth = 0usize;
    loop {
        events += 1;
        match reader.read_event().expect("fixture XML parses") {
            Event::Start(_) => {
                depth += 1;
                maximum_depth = maximum_depth.max(depth);
            },
            Event::Empty(_) => {
                maximum_depth = maximum_depth.max(depth + 1);
            },
            Event::End(_) => depth = depth.checked_sub(1).expect("fixture XML depth balances"),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    assert_eq!(depth, 0, "fixture XML has a closed element stack");
    XmlMetrics {
        events,
        depth: maximum_depth,
    }
}

fn owned_xml_metrics(package: &Package) -> BTreeMap<String, XmlMetrics> {
    let opc = package.clone().into_plain_opc();
    let mut metrics = BTreeMap::new();
    metrics.insert(
        "[Content_Types].xml".to_owned(),
        xml_metrics(
            opc.source_content_types()
                .expect("source content types exist")
                .bytes(),
        ),
    );
    let root = uri("/");
    metrics.insert(
        "relationships:/".to_owned(),
        xml_metrics(
            opc.source_relationships(&root)
                .expect("source package relationships exist")
                .bytes(),
        ),
    );
    for part in opc
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        let owner = part.partname().as_str().to_owned();
        metrics.insert(
            format!("relationships:{owner}"),
            xml_metrics(
                opc.source_relationships(part.partname())
                    .expect("source relationship member exists")
                    .bytes(),
            ),
        );
        if part.partname().is_equivalent_to(&uri(WORKBOOK_URI))
            || matches!(
                part.content_type(),
                PROPERTIES_CONTENT_TYPE | CONNECTIONS_CONTENT_TYPE | QUERY_TABLE_CONTENT_TYPE
            )
        {
            metrics.insert(format!("part:{owner}"), xml_metrics(part.blob()));
        }
        if part.content_type() == QUERY_TABLE_CONTENT_TYPE {
            let processed = litchi_ooxml_common::mce::process_ooxml(part.blob())
                .expect("query-table MCE fixture processes");
            metrics.insert(format!("mce:{owner}"), xml_metrics(processed.as_ref()));
        }
    }
    metrics
}

fn maximum_xml_events(metrics: &BTreeMap<String, XmlMetrics>) -> usize {
    metrics
        .values()
        .map(|metric| metric.events)
        .max()
        .expect("owned XML metrics are non-empty")
}

fn maximum_xml_depth(metrics: &BTreeMap<String, XmlMetrics>) -> usize {
    metrics
        .values()
        .map(|metric| metric.depth)
        .max()
        .expect("owned XML metrics are non-empty")
}

#[test]
fn host_xml_limits_accept_exact_boundaries_and_reject_one_under() {
    let minimal = minimal_fixture("event");
    let source_properties = part_bytes(&minimal.package, "/xl/customData/props1.xml");
    let minimal_before = package_bytes(&minimal.package);

    let exact_events = CustomDataLimits::new()
        .with_max_xml_events(maximum_xml_events(&owned_xml_metrics(&minimal.package)));
    assert!(
        minimal
            .package
            .custom_data_with_limits(&exact_events)
            .is_ok()
    );
    let one_under_events = exact_events.with_max_xml_events(1);
    assert!(
        minimal
            .package
            .custom_data_with_limits(&one_under_events)
            .is_err()
    );
    assert_eq!(package_bytes(&minimal.package), minimal_before);

    let exact_properties =
        CustomDataLimits::new().with_max_properties_xml_bytes(source_properties.len());
    assert!(
        minimal
            .package
            .custom_data_with_limits(&exact_properties)
            .is_ok()
    );
    let one_under_properties =
        exact_properties.with_max_properties_xml_bytes(source_properties.len() - 1);
    assert!(
        minimal
            .package
            .custom_data_with_limits(&one_under_properties)
            .is_err()
    );
    assert_eq!(package_bytes(&minimal.package), minimal_before);

    let extended = single_fixture();
    let extended_before = package_bytes(&extended.package);
    let snapshot = extended
        .package
        .custom_data()
        .expect("formatted source snapshot loads");
    let extension_bytes = snapshot
        .find("uid-A")
        .expect("uid-A storage exists")
        .value()
        .properties
        .extension_list
        .as_ref()
        .expect("direct extension exists")
        .xml
        .len();
    let exact_extension = CustomDataLimits::new().with_max_extension_xml_bytes(extension_bytes);
    assert!(
        extended
            .package
            .custom_data_with_limits(&exact_extension)
            .is_ok()
    );
    let one_under_extension = exact_extension.with_max_extension_xml_bytes(extension_bytes - 1);
    assert!(
        extended
            .package
            .custom_data_with_limits(&one_under_extension)
            .is_err()
    );
    assert_eq!(package_bytes(&extended.package), extended_before);
}

#[test]
fn relationship_xml_event_and_depth_caps_cover_the_source_member() {
    let fixture = relationship_heavy_fixture();
    let metrics = owned_xml_metrics(&fixture.package);
    let target = metrics
        .get("relationships:/xl/workbook.xml")
        .expect("workbook relationship metrics exist");
    assert_eq!(target.events, maximum_xml_events(&metrics));
    assert!(target.depth < maximum_xml_depth(&metrics));
    assert!(target.events > 1);
    assert!(target.depth > 1);

    let exact = CustomDataLimits::new()
        .with_max_xml_events(maximum_xml_events(&metrics))
        .with_max_xml_depth(target.depth);
    assert!(fixture.package.custom_data_with_limits(&exact).is_ok());

    let before = package_bytes(&fixture.package);
    let one_under_events = exact.with_max_xml_events(target.events - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_events)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
    let one_under_depth = exact.with_max_xml_depth(target.depth - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_depth)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn content_types_xml_event_and_depth_caps_cover_the_source_member() {
    let fixture = content_types_heavy_fixture();
    let metrics = owned_xml_metrics(&fixture.package);
    let target = metrics
        .get("[Content_Types].xml")
        .expect("content-types metrics exist");
    assert_eq!(target.events, maximum_xml_events(&metrics));
    assert!(target.depth < maximum_xml_depth(&metrics));
    assert!(target.events > 1);
    assert!(target.depth > 1);

    let exact = CustomDataLimits::new()
        .with_max_xml_events(maximum_xml_events(&metrics))
        .with_max_xml_depth(target.depth);
    assert!(fixture.package.custom_data_with_limits(&exact).is_ok());

    let before = package_bytes(&fixture.package);
    let one_under_events = exact.with_max_xml_events(target.events - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_events)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
    let one_under_depth = exact.with_max_xml_depth(target.depth - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_depth)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn query_table_mce_event_and_depth_caps_cover_raw_and_processed_xml() {
    let fixture = query_table_fixture(query_table_xml(10, true));
    let metrics = owned_xml_metrics(&fixture.package);
    let raw = metrics
        .get("part:/xl/queryTables/queryTable1.xml")
        .expect("query-table source metrics exist");
    let processed = metrics
        .get("mce:/xl/queryTables/queryTable1.xml")
        .expect("processed query-table metrics exist");
    assert!(raw.events > 1);
    assert!(processed.events > 1);
    let exact_events = raw.events.max(processed.events);
    let exact_depth = raw.depth;
    assert!(exact_events >= processed.events);
    assert_eq!(exact_depth, raw.depth);

    let exact = CustomDataLimits::new()
        .with_max_xml_events(exact_events)
        .with_max_xml_depth(exact_depth);
    assert!(fixture.package.custom_data_with_limits(&exact).is_ok());

    let before = package_bytes(&fixture.package);
    let one_under_events = exact.with_max_xml_events(exact_events - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_events)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
    let one_under_depth = exact.with_max_xml_depth(exact_depth - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&one_under_depth)
            .is_err()
    );
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn query_table_empty_elements_honor_the_xml_depth_cap() {
    let fixture = query_table_fixture_with_connections(
        query_table_xml_with_empty_child(4),
        shallow_connections_xml("uid-A"),
    );
    let exact = CustomDataLimits::new().with_max_xml_depth(6);
    assert!(fixture.package.custom_data_with_limits(&exact).is_ok());

    let before = package_bytes(&fixture.package);
    let one_under = exact.with_max_xml_depth(5);
    let error = fixture
        .package
        .custom_data_with_limits(&one_under)
        .expect_err("empty query-table child must count toward XML depth");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::Depth
    ));
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn zero_xml_event_budget_is_refused_without_widening() {
    let fixture = minimal_fixture("zero-events");
    let before = package_bytes(&fixture.package);
    let error = fixture
        .package
        .custom_data_with_limits(&CustomDataLimits::new().with_max_xml_events(0))
        .expect_err("zero XML events cannot admit the first source event");
    // The bounded OPC reader may reject zero before the owner reaches its XML
    // codec; either path must preserve refusal rather than widen to one event.
    assert!(matches!(error, Error::Package(_) | Error::ResourceLimit(_)));
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn host_uid_limit_counts_utf16_units_exactly() {
    let fixture = minimal_fixture("😀");
    let exact = CustomDataLimits::new().with_max_uid_units(2);
    assert!(fixture.package.custom_data_with_limits(&exact).is_ok());

    let one_under = exact.with_max_uid_units(1);
    let before = package_bytes(&fixture.package);
    assert!(fixture.package.custom_data_with_limits(&one_under).is_err());
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn relationship_xml_cap_includes_inserted_edges_and_removal_inverse_restores_source() {
    let fixture = single_fixture();
    let mut candidate = fixture.package.clone();
    let mut insertion = candidate
        .edit_custom_data()
        .expect("insertion transaction starts");
    insertion
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("relationship insertion stages");
    let insertion_commit = insertion.commit().expect("relationship insertion commits");
    let exact = max_relationship_xml_size(&candidate);
    assert!(
        relationship_xml_size(&candidate, WORKBOOK_URI)
            > relationship_xml_size(&fixture.package, WORKBOOK_URI)
    );
    assert!(exact > 0);

    let mut exact_package = fixture.package.clone();
    let mut exact_transaction = exact_package
        .edit_custom_data_with_limits(
            &CustomDataLimits::new().with_max_relationship_xml_bytes(exact),
        )
        .expect("exact relationship cap admits source");
    exact_transaction
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("exact relationship insertion stages");
    assert!(
        exact_transaction
            .commit()
            .expect("exact relationship cap is inclusive")
            .changed()
    );

    let mut under_package = fixture.package.clone();
    let before = package_bytes(&under_package);
    let mut under_transaction = under_package
        .edit_custom_data_with_limits(
            &CustomDataLimits::new().with_max_relationship_xml_bytes(exact - 1),
        )
        .expect("one-under relationship cap admits source");
    under_transaction
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("one-under relationship insertion stages");
    assert!(under_transaction.commit().is_err());
    assert_eq!(package_bytes(&under_package), before);
    assert!(insertion_commit.changed());

    let two = two_storage_fixture();
    let mut removal_candidate = two.package.clone();
    let mut removal_candidate_transaction = removal_candidate
        .edit_custom_data()
        .expect("removal candidate transaction starts");
    let removal_index = removal_candidate_transaction
        .entries()
        .iter()
        .position(|entry| entry.id() == "uid-B")
        .expect("removal candidate storage exists");
    removal_candidate_transaction
        .remove_with(removal_index, RemovalDisposition::RejectReferenced)
        .expect("removal candidate stages");
    removal_candidate_transaction
        .commit()
        .expect("removal candidate commits");
    let removal_exact = max_relationship_xml_size(&removal_candidate);
    let removal_source = max_relationship_xml_size(&two.package);
    // Relationship XML is bounded while opening the source as well as while
    // publishing the candidate. Removal shrinks this member, so its final
    // maximum cannot admit the larger source; retain the source maximum.
    let removal_limit = removal_source.max(removal_exact);
    assert!(removal_exact > 0);
    assert!(removal_source >= removal_exact);

    let mut exact_removal_package = two.package.clone();
    let mut exact_removal = exact_removal_package
        .edit_custom_data_with_limits(
            &CustomDataLimits::new().with_max_relationship_xml_bytes(removal_limit),
        )
        .expect("source-admitted removal relationship cap admits source");
    let exact_removal_index = exact_removal
        .entries()
        .iter()
        .position(|entry| entry.id() == "uid-B")
        .expect("exact removal storage exists");
    exact_removal
        .remove_with(exact_removal_index, RemovalDisposition::RejectReferenced)
        .expect("exact removal stages");
    assert!(
        exact_removal
            .commit()
            .expect("source-admitted removal relationship cap commits")
            .changed()
    );

    let before_under_removal = package_bytes(&two.package);
    let mut under_removal_package = two.package.clone();
    assert!(
        under_removal_package
            .edit_custom_data_with_limits(
                &CustomDataLimits::new().with_max_relationship_xml_bytes(removal_limit - 1),
            )
            .is_err()
    );
    assert_eq!(package_bytes(&under_removal_package), before_under_removal);
    Package::from_bytes(before_under_removal.clone())
        .expect("one-under removal rejection leaves a reopenable source");

    let before_remove = package_bytes(&two.package);
    let mut removed = two.package.clone();
    let mut remove_transaction = removed
        .edit_custom_data()
        .expect("removal transaction starts");
    let index = remove_transaction
        .entries()
        .iter()
        .position(|entry| entry.id() == "uid-B")
        .expect("unreferenced storage exists");
    remove_transaction
        .remove_with(index, RemovalDisposition::RejectReferenced)
        .expect("unreferenced storage removes");
    let remove_commit = remove_transaction.commit().expect("removal commits");
    removed
        .apply_custom_data_patch(&remove_commit.patch().inverse())
        .expect("removal inverse applies");
    assert_eq!(
        archive_members(&package_bytes(&removed)),
        archive_members(&before_remove)
    );
}

#[test]
fn connection_and_temporary_caps_cover_rename_growth_without_mutating_source() {
    let fixture = repeated_reference_fixture(32);
    let long_id = "renamed-custom-data-uid-with-a-materially-longer-wire-value";
    let source_connections = fixture
        .connections
        .as_ref()
        .expect("connection source exists");
    let source_properties = fixture
        .properties
        .get("/xl/customData/props1.xml")
        .expect("properties source exists");
    let expected_connections = connections_xml(&[long_id; 32], true);
    let expected_properties = properties_xml(long_id, true, true);
    assert_eq!(
        part_bytes(&fixture.package, CONNECTIONS_URI),
        source_connections.as_slice()
    );
    assert_eq!(
        part_bytes(&fixture.package, "/xl/customData/props1.xml"),
        source_properties.as_slice()
    );
    assert!(expected_connections.len() > source_connections.len());
    assert!(expected_properties.len() > source_properties.len());

    let exact_limits = CustomDataLimits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(expected_connections.len());
    let mut exact_package = fixture.package.clone();
    let mut exact = exact_package
        .edit_custom_data_with_limits(&exact_limits)
        .expect("exact rename profile admits source");
    exact.rename(0, long_id).expect("exact rename stages");
    let exact_commit = exact
        .commit()
        .expect("exact connection/property caps are inclusive");
    assert!(exact_commit.changed());
    assert_eq!(
        part_bytes(&exact_package, CONNECTIONS_URI),
        expected_connections
    );
    assert_eq!(
        part_bytes(&exact_package, "/xl/customData/props1.xml"),
        expected_properties
    );

    let before = package_bytes(&fixture.package);
    let mut under_connections_package = fixture.package.clone();
    let under_connections = CustomDataLimits::new()
        .with_max_connections_bytes(expected_connections.len() - 1)
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(expected_connections.len());
    let mut under_connections_transaction = under_connections_package
        .edit_custom_data_with_limits(&under_connections)
        .expect("one-under connection profile admits source");
    under_connections_transaction
        .rename(0, long_id)
        .expect("one-under connection rename stages");
    assert!(under_connections_transaction.commit().is_err());
    assert_eq!(package_bytes(&under_connections_package), before);

    let mut under_properties_package = fixture.package.clone();
    let under_properties = CustomDataLimits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len() - 1)
        .with_max_temporary_bytes(expected_connections.len());
    let mut under_properties_transaction = under_properties_package
        .edit_custom_data_with_limits(&under_properties)
        .expect("one-under property profile admits source");
    under_properties_transaction
        .rename(0, long_id)
        .expect("one-under property rename stages");
    assert!(under_properties_transaction.commit().is_err());
    assert_eq!(package_bytes(&under_properties_package), before);

    let mut under_temporary_package = fixture.package.clone();
    let under_temporary = CustomDataLimits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(source_connections.len());
    let mut under_temporary_transaction = under_temporary_package
        .edit_custom_data_with_limits(&under_temporary)
        .expect("source-sized temporary profile admits source");
    under_temporary_transaction
        .rename(0, long_id)
        .expect("source-sized temporary rename stages");
    assert!(under_temporary_transaction.commit().is_err());
    assert_eq!(package_bytes(&under_temporary_package), before);
}

#[test]
fn retained_xml_can_shrink_below_temporary_cap_and_inverse_restores_source() {
    let fixture = repeated_reference_fixture(32);
    let source_connections = fixture
        .connections
        .as_ref()
        .expect("connection source exists");
    let source_properties = fixture
        .properties
        .get("/xl/customData/props1.xml")
        .expect("properties source exists");
    let expected_connections = connections_xml(&["u"; 32], true);
    let expected_properties = properties_xml("u", true, true);
    let temporary = source_connections.len() - 1;
    assert!(source_connections.len() > temporary);
    assert!(expected_connections.len() < temporary);
    assert!(expected_properties.len() < temporary);

    let profile = CustomDataLimits::new()
        .with_max_connections_bytes(source_connections.len())
        .with_max_properties_xml_bytes(source_properties.len())
        .with_max_temporary_bytes(temporary);
    let before = archive_members(&package_bytes(&fixture.package));
    let mut changed = fixture.package.clone();
    let mut transaction = changed
        .edit_custom_data_with_limits(&profile)
        .expect("source-backed shrinking transaction starts");
    transaction
        .rename(0, "u")
        .expect("shorter UID rename stages");
    let commit = transaction
        .commit()
        .expect("retained source may exceed temporary replacement bytes");
    assert_eq!(part_bytes(&changed, CONNECTIONS_URI), expected_connections);
    assert_eq!(
        part_bytes(&changed, "/xl/customData/props1.xml"),
        expected_properties
    );

    let mut reopened = Package::from_bytes(package_bytes(&changed)).expect("changed reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse restores retained source above temporary cap");
    assert_eq!(archive_members(&package_bytes(&reopened)), before);
}

#[test]
fn retained_properties_xml_can_shrink_below_temporary_cap_and_inverse_restores_source() {
    let long_id = "x".repeat(4096);
    let fixture = minimal_fixture(&long_id);
    let source_properties = part_bytes(&fixture.package, "/xl/customData/props1.xml");
    let expected_properties = properties_xml("u", false, false);
    let temporary = 1024;
    assert!(source_properties.len() > temporary);
    assert!(expected_properties.len() < temporary);

    let profile = CustomDataLimits::new()
        .with_max_properties_xml_bytes(source_properties.len())
        .with_max_temporary_bytes(temporary);
    let before = archive_members(&package_bytes(&fixture.package));
    let mut changed = fixture.package.clone();
    let mut transaction = changed
        .edit_custom_data_with_limits(&profile)
        .expect("source-backed Properties shrink transaction starts");
    transaction
        .rename(0, "u")
        .expect("shorter Properties UID rename stages");
    let commit = transaction
        .commit()
        .expect("retained Properties source may exceed temporary bytes");
    assert_eq!(
        part_bytes(&changed, "/xl/customData/props1.xml"),
        expected_properties
    );

    let mut reopened = Package::from_bytes(package_bytes(&changed)).expect("changed reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse restores retained Properties source above temporary cap");
    assert_eq!(archive_members(&package_bytes(&reopened)), before);
}

#[test]
fn retained_content_types_manifest_can_shrink_below_temporary_cap_and_inverse_restores_source() {
    let fixture = content_types_self_closed_heavy_fixture();
    let source_content_types = fixture
        .package
        .clone()
        .into_plain_opc()
        .source_content_types()
        .expect("source content-types manifest exists")
        .bytes()
        .len();
    let temporary = source_content_types - 1;
    assert!(source_content_types > temporary);

    let profile = CustomDataLimits::new().with_max_temporary_bytes(temporary);
    let before = archive_members(&package_bytes(&fixture.package));
    let mut changed = fixture.package.clone();
    let mut transaction = changed
        .edit_custom_data_with_limits(&profile)
        .expect("source-backed manifest removal transaction starts");
    transaction
        .remove_with(0, RemovalDisposition::RejectReferenced)
        .expect("manifest removal stages");
    let commit = transaction
        .commit()
        .expect("retained content-types source may exceed temporary bytes");
    let candidate_content_types = changed
        .clone()
        .into_plain_opc()
        .source_content_types()
        .expect("candidate content-types manifest exists")
        .bytes()
        .len();
    assert!(candidate_content_types < temporary);

    let mut reopened = Package::from_bytes(package_bytes(&changed)).expect("changed reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse restores retained content-types source above temporary cap");
    assert_eq!(archive_members(&package_bytes(&reopened)), before);
}

#[test]
fn moved_payload_can_exceed_temporary_cap_without_copy() {
    let fixture = synthetic_fixture(&["uid-A"], &[], &[b"source"], false, false);
    let limits = CustomDataLimits::new()
        .with_max_payload_bytes(64 * 1024)
        .with_max_temporary_bytes(8 * 1024);

    let mut noop_package = fixture.package.clone();
    let mut noop_transaction = noop_package
        .edit_custom_data_with_limits(&limits)
        .expect("bounded no-op transaction starts");
    let source_pointer = noop_transaction.entries()[0].value().data().as_ptr();
    assert!(
        !noop_transaction
            .set_data(0, b"source".to_vec())
            .expect("equal payload is a no-op")
    );
    assert_eq!(
        noop_transaction.entries()[0].value().data().as_ptr(),
        source_pointer
    );
    let noop_commit = noop_transaction
        .commit()
        .expect("bounded no-op commit succeeds");
    assert!(!noop_commit.changed());

    let mut package = fixture.package.clone();
    let payload = vec![0x5a; 16 * 1024];
    let payload_pointer = payload.as_ptr();
    let mut transaction = package
        .edit_custom_data_with_limits(&limits)
        .expect("bounded move transaction starts");
    transaction
        .set_data(0, payload)
        .expect("bounded move stages without a temporary copy");
    assert_eq!(
        transaction.entries()[0].value().data().as_ptr(),
        payload_pointer
    );
    let commit = transaction
        .commit()
        .expect("existing payload allocation is not charged as temporary bytes");
    let published = commit.snapshot().find("uid-A").expect("uid-A exists");
    assert_eq!(published.value().data().as_ptr(), payload_pointer);
    assert_eq!(published.value().data(), vec![0x5a; 16 * 1024].as_slice());
}

#[test]
fn final_output_cap_accepts_candidate_shrink_and_rejects_one_under() {
    let fixture = shrinking_storage_fixture();
    let mut candidate = fixture.package.clone();
    let mut baseline = candidate
        .edit_custom_data()
        .expect("shrinking baseline transaction starts");
    baseline
        .set_data(0, b"a".to_vec())
        .expect("first payload shrinks");
    baseline
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("second payload grows");
    baseline.commit().expect("shrinking baseline edit commits");
    let source_size = logical_package_size(&fixture.package);
    let candidate_size = logical_package_size(&candidate);
    // The first payload shrinks by 63 bytes and the second grows by 37; the
    // final candidate is 26 bytes smaller than the source package.
    assert_eq!(candidate_size, source_size - 26);

    let exact_limits = CustomDataLimits::new().with_max_output_bytes(candidate_size);
    let mut exact_package = fixture.package.clone();
    let mut exact = exact_package
        .edit_custom_data_with_limits(&exact_limits)
        .expect("exact shrinking output profile admits source");
    exact
        .set_data(0, b"a".to_vec())
        .expect("exact first shrink stages");
    exact
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("exact second growth stages");
    let exact_commit = exact
        .commit()
        .expect("exact shrinking aggregate cap is inclusive");
    assert!(exact_commit.changed());
    assert_eq!(logical_package_size(&exact_package), candidate_size);

    let mut under_package = fixture.package.clone();
    let before = package_bytes(&under_package);
    let mut under = under_package
        .edit_custom_data_with_limits(
            &CustomDataLimits::new().with_max_output_bytes(candidate_size - 1),
        )
        .expect("one-under shrinking output profile admits source");
    under
        .set_data(0, b"a".to_vec())
        .expect("under first shrink stages");
    under
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("under second growth stages");
    assert!(under.commit().is_err());
    assert_eq!(package_bytes(&under_package), before);
}

#[test]
fn final_output_cap_uses_candidate_aggregate_for_mixed_shrink_and_growth() {
    let fixture = two_storage_fixture();
    let mut candidate = fixture.package.clone();
    let mut baseline = candidate
        .edit_custom_data()
        .expect("baseline transaction starts");
    baseline
        .set_data(0, b"a".to_vec())
        .expect("first payload shrinks");
    baseline
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("second payload grows");
    baseline.commit().expect("baseline mixed edit commits");
    let source_size = logical_package_size(&fixture.package);
    let candidate_size = logical_package_size(&candidate);
    // The first payload shrinks by 31 bytes and the second grows by 37; the
    // final candidate therefore grows by six bytes overall.
    assert_eq!(candidate_size, source_size + 6);

    let exact_limits = CustomDataLimits::new().with_max_output_bytes(candidate_size);
    let mut exact_package = fixture.package.clone();
    let mut exact = exact_package
        .edit_custom_data_with_limits(&exact_limits)
        .expect("exact final output profile admits source");
    exact
        .set_data(0, b"a".to_vec())
        .expect("exact first shrink stages");
    exact
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("exact second growth stages");
    let exact_commit = exact
        .commit()
        .expect("exact final aggregate cap is inclusive");
    assert!(exact_commit.changed());
    assert_eq!(logical_package_size(&exact_package), candidate_size);

    let mut under_package = fixture.package.clone();
    let before = package_bytes(&under_package);
    let mut under = under_package
        .edit_custom_data_with_limits(
            &CustomDataLimits::new().with_max_output_bytes(candidate_size - 1),
        )
        .expect("one-under final output profile admits source");
    under
        .set_data(0, b"a".to_vec())
        .expect("under first shrink stages");
    under
        .set_data(1, b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec())
        .expect("under second growth stages");
    assert!(under.commit().is_err());
    assert_eq!(package_bytes(&under_package), before);
}

#[test]
fn retained_limits_survive_public_patch_inverse_after_reopen() {
    let fixture = single_fixture();
    let before = package_bytes(&fixture.package);
    let profile = CustomDataLimits::new()
        .with_max_uid_units(128)
        .with_max_properties_xml_bytes(64 * 1024)
        .with_max_extension_xml_bytes(32 * 1024)
        .with_max_connections_bytes(64 * 1024)
        .with_max_relationship_xml_bytes(64 * 1024)
        .with_max_temporary_bytes(64 * 1024)
        .with_max_output_bytes(128 * 1024);

    let mut changed = fixture.package.clone();
    let mut transaction = changed
        .edit_custom_data_with_limits(&profile)
        .expect("profiled transaction starts");
    transaction
        .rename(0, "profiled-rename")
        .expect("profiled rename stages");
    let commit = transaction.commit().expect("profiled rename commits");
    assert_eq!(commit.snapshot().limits(), profile);

    let mut reopened =
        Package::from_bytes(package_bytes(&changed)).expect("changed package reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse patch applies with retained profile");
    assert_eq!(
        archive_members(&package_bytes(&reopened)),
        archive_members(&before)
    );
}

#[test]
fn source_bound_noop_and_inverse_preserve_lexical_zip_members() {
    let fixture = single_fixture();
    let before = archive_members(&package_bytes(&fixture.package));

    let mut noop_package = fixture.package.clone();
    let mut noop = noop_package
        .edit_custom_data()
        .expect("no-op transaction starts");
    assert!(!noop.rename(0, "uid-A").expect("same UID is a no-op"));
    let noop_commit = noop.commit().expect("no-op transaction commits");
    assert!(!noop_commit.changed());
    assert!(noop_commit.patch().is_empty());
    assert_eq!(archive_members(&package_bytes(&noop_package)), before);

    let mut changed_package = fixture.package.clone();
    let mut transaction = changed_package
        .edit_custom_data()
        .expect("source-preserving transaction starts");
    transaction.rename(0, "uid-B").expect("UID rename stages");
    let commit = transaction.commit().expect("UID rename commits");
    let changed_bytes = package_bytes(&changed_package);
    let changed_archive = archive_members(&changed_bytes);
    assert_ne!(changed_archive, before);
    assert!(
        changed_archive
            .get("xl/customData/props1.xml")
            .expect("changed properties member exists")
            .windows(b"retained before root".len())
            .any(|window| window == b"retained before root")
    );
    assert!(
        changed_archive
            .get("xl/connections.xml")
            .expect("changed connections member exists")
            .windows(b"urn:synthetic-opaque".len())
            .any(|window| window == b"urn:synthetic-opaque")
    );

    let mut reopened = Package::from_bytes(changed_bytes).expect("changed package reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse applies after reopen");
    assert_eq!(archive_members(&package_bytes(&reopened)), before);
}

#[test]
fn connection_owner_relationship_presence_is_source_bound_and_inverse_exact() {
    let absent = single_fixture_with_connection_relationship_member(false);
    let absent_before = archive_members(&package_bytes(&absent.package));
    let mut changed = absent.package.clone();
    let mut transaction = changed
        .edit_custom_data()
        .expect("absent-owner transaction starts");
    transaction
        .rename(0, "uid-B")
        .expect("absent-owner rename stages");
    let commit = transaction.commit().expect("absent-owner rename commits");

    let explicit = single_fixture_with_connection_relationship_member(true);
    let explicit_before = archive_members(&package_bytes(&explicit.package));
    let mut stale = explicit.package.clone();
    assert!(matches!(
        stale.apply_custom_data_patch(commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(archive_members(&package_bytes(&stale)), explicit_before);

    let mut explicit_changed = explicit.package.clone();
    let mut explicit_transaction = explicit_changed
        .edit_custom_data()
        .expect("explicit-owner transaction starts");
    explicit_transaction
        .rename(0, "uid-B")
        .expect("explicit-owner rename stages");
    let explicit_commit = explicit_transaction
        .commit()
        .expect("explicit-owner rename commits");
    let mut stale_absent = absent.package.clone();
    assert!(matches!(
        stale_absent.apply_custom_data_patch(explicit_commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(
        archive_members(&package_bytes(&stale_absent)),
        absent_before
    );

    let mut reopened = Package::from_bytes(package_bytes(&changed)).expect("changed reopens");
    reopened
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("inverse applies to the changed source");
    assert_eq!(archive_members(&package_bytes(&reopened)), absent_before);

    let mut explicit_reopened =
        Package::from_bytes(package_bytes(&explicit_changed)).expect("explicit change reopens");
    explicit_reopened
        .apply_custom_data_patch(&explicit_commit.patch().inverse())
        .expect("explicit-owner inverse restores the empty relationship member");
    assert_eq!(
        archive_members(&package_bytes(&explicit_reopened)),
        explicit_before
    );
}

#[test]
fn stale_custom_data_patch_is_refused_without_mutating_the_target() {
    let fixture = single_fixture();
    let mut source = fixture.package.clone();
    let mut source_transaction = source
        .edit_custom_data()
        .expect("source transaction starts");
    source_transaction
        .set_data(0, b"source patch".to_vec())
        .expect("source change stages");
    let patch = source_transaction
        .commit()
        .expect("source change commits")
        .patch()
        .clone();

    let mut stale = fixture.package.clone();
    let mut stale_transaction = stale.edit_custom_data().expect("stale transaction starts");
    stale_transaction
        .set_data(0, b"independent target change".to_vec())
        .expect("independent target change stages");
    stale_transaction
        .commit()
        .expect("independent target change commits");
    let before = archive_members(&package_bytes(&stale));

    assert!(matches!(
        stale.apply_custom_data_patch(&patch),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(archive_members(&package_bytes(&stale)), before);
}

#[test]
fn removal_refuses_unexpected_inbound_relationships_atomically() {
    let fixture = foreign_inbound_fixture();
    let before = archive_members(&package_bytes(&fixture.package));
    let mut package = fixture.package.clone();
    let mut transaction = package
        .edit_custom_data()
        .expect("foreign-edge transaction starts");
    transaction
        .remove_with(0, RemovalDisposition::DetachConnections)
        .expect("unreferenced storage removal stages");
    assert!(transaction.commit().is_err());
    assert_eq!(archive_members(&package_bytes(&package)), before);
}

#[test]
fn signed_custom_data_noop_is_exact_but_changed_commit_is_refused() {
    let fixture = signed_fixture();
    let before = archive_members(&package_bytes(&fixture.package));

    let mut noop_package = fixture.package.clone();
    let noop_commit = noop_package
        .edit_custom_data()
        .expect("signed no-op transaction starts")
        .commit()
        .expect("signed no-op commits");
    assert!(!noop_commit.changed());
    assert_eq!(archive_members(&package_bytes(&noop_package)), before);

    let mut changed_package = fixture.package.clone();
    let mut transaction = changed_package
        .edit_custom_data()
        .expect("signed changed transaction starts");
    transaction
        .set_data(0, b"signed change".to_vec())
        .expect("signed change stages");
    assert!(matches!(transaction.commit(), Err(Error::Signed)));
    assert_eq!(archive_members(&package_bytes(&changed_package)), before);
}

#[test]
fn caller_limits_propagate_to_storage_payload_and_relationship_graph_admission() {
    let fixture = two_storage_fixture();
    let before = archive_members(&package_bytes(&fixture.package));
    let relationships = relationship_count(&fixture.package);
    assert!(relationships > 1);

    let exact_relationships = CustomDataLimits::new().with_max_relationships(relationships);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&exact_relationships)
            .is_ok()
    );
    let under_relationships = exact_relationships.with_max_relationships(relationships - 1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&under_relationships)
            .is_err()
    );

    let storage_under = CustomDataLimits::new().with_max_storages(1);
    assert!(
        fixture
            .package
            .custom_data_with_limits(&storage_under)
            .is_err()
    );

    let payload_fixture = minimal_fixture("payload");
    let payload_before = archive_members(&package_bytes(&payload_fixture.package));
    let payload_under = CustomDataLimits::new().with_max_payload_bytes(6);
    assert!(
        payload_fixture
            .package
            .custom_data_with_limits(&payload_under)
            .is_err()
    );
    assert_eq!(
        archive_members(&package_bytes(&payload_fixture.package)),
        payload_before
    );
    assert_eq!(archive_members(&package_bytes(&fixture.package)), before);
}

#[test]
fn namespace_byte_limit_reports_input_bytes_dimension() {
    let fixture = minimal_fixture("namespace");
    let limits = CustomDataLimits::new().with_max_xml_namespace_bytes(1);
    let error = fixture
        .package
        .custom_data_with_limits(&limits)
        .expect_err("tight namespace byte cap rejects the source");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::InputBytes
    ));
}

#[test]
fn connection_namespace_declarations_consume_attribute_budget() {
    let fixture = fixture_with_connections_xml(connections_with_root_namespaces(
        &["uid-A"],
        &[
            ("unused0", "urn:unused0"),
            ("unused1", "urn:unused1"),
            ("unused2", "urn:unused2"),
            ("unused3", "urn:unused3"),
            ("unused4", "urn:unused4"),
            ("unused5", "urn:unused5"),
            ("unused6", "urn:unused6"),
            ("unused7", "urn:unused7"),
        ],
    ));
    let limits = CustomDataLimits::new().with_max_xml_attributes(7);
    let error = fixture
        .package
        .custom_data_with_limits(&limits)
        .expect_err("unused namespace declarations must count as attributes");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
}

#[test]
fn connection_namespace_budget_includes_prefix_bytes() {
    const LONG_PREFIX: &str = "unused_namespace_prefix_that_is_longer_than_the_remaining_budget";
    let fixture = fixture_with_connections_xml(connections_with_root_namespaces(
        &["uid-A"],
        &[(LONG_PREFIX, "u")],
    ));
    let existing_namespace_bytes =
        "s".len() + SML.len() + "a".len() + X14.len() + "q".len() + "urn:synthetic-opaque".len();
    let limits = CustomDataLimits::new()
        .with_max_xml_namespace_bytes(existing_namespace_bytes + LONG_PREFIX.len());
    let error = fixture
        .package
        .custom_data_with_limits(&limits)
        .expect_err("namespace prefixes must consume the byte budget");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::InputBytes
    ));
}

#[test]
fn connection_in_root_cdata_counts_against_xml_string_bytes() {
    let cdata = "opaque connection payload that exceeds the caller string budget";
    let fixture = fixture_with_connections_xml(connections_with_opaque_cdata("uid-A", cdata));
    let limits = CustomDataLimits::new().with_max_xml_string_bytes(cdata.len() - 1);
    let error = fixture
        .package
        .custom_data_with_limits(&limits)
        .expect_err("in-root CDATA must consume the XML string budget");
    assert!(matches!(
        error,
        Error::ResourceLimit(limit) if limit.resource == Resource::InputBytes
    ));
}

#[test]
fn entity_escaped_connection_namespace_preserves_binding_or_refuses() {
    let fixture = fixture_with_connections_xml(connections_with_escaped_x14_namespace("uid-A"));
    let mut package = fixture.package.clone();
    let references = match package.custom_data() {
        Err(_) => return,
        Ok(snapshot) => snapshot.connection_references("uid-A"),
    };
    assert_eq!(
        references, 1,
        "admitted escaped namespace must retain its embedded-data binding"
    );

    let mut transaction = package
        .edit_custom_data()
        .expect("escaped namespace transaction starts");
    transaction
        .rename(0, "uid-B")
        .expect("escaped namespace rename stages");
    let commit = transaction
        .commit()
        .expect("escaped namespace rename commits");
    assert!(commit.changed());
    let xml = part_bytes(&package, CONNECTIONS_URI);
    assert!(
        xml.windows(b"embeddedDataId=\"uid-B\"".len())
            .any(|window| { window == b"embeddedDataId=\"uid-B\"" })
    );
    assert!(
        !xml.windows(b"embeddedDataId=\"uid-A\"".len())
            .any(|window| { window == b"embeddedDataId=\"uid-A\"" })
    );
}

#[test]
fn default_public_custom_data_api_remains_usable() {
    let fixture = single_fixture();
    let snapshot = fixture
        .package
        .custom_data()
        .expect("default snapshot API remains available");
    assert_eq!(snapshot.len(), 1);
    assert_eq!(
        snapshot.find("uid-A").expect("uid-A exists").value().data(),
        b"\0opaque\xffpayload\n"
    );

    let mut package = fixture.package.clone();
    let mut transaction = package
        .edit_custom_data()
        .expect("default transaction API remains available");
    transaction
        .set_data(0, b"default-api-edit".to_vec())
        .expect("default transaction stages");
    transaction.commit().expect("default transaction commits");
    assert_eq!(
        package
            .custom_data()
            .expect("changed default snapshot reads")
            .find("uid-A")
            .expect("uid-A remains")
            .value()
            .data(),
        b"default-api-edit"
    );
}
