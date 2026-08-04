//! Regression coverage for the workbook-metadata owner.

use super::codec::{MAX_STRING, XDA};
use super::model::*;
use super::package::SPREADSHEETML_NAMESPACE as SML;

fn sample() -> WorkbookMetadata {
    WorkbookMetadata {
        types: vec![MetadataType {
            name: "XLDAPR".into(),
            minimum_supported_version: 120000,
            behavior: MetadataBehavior {
                copy: true,
                paste_all: true,
                paste_values: true,
                cell_metadata: true,
                ..Default::default()
            },
        }],
        future: vec![FutureMetadata {
            name: "XLDAPR".into(),
            blocks: vec![MetadataBlock {
                records: Vec::new(),
                extensions: vec![OpaqueMetadataExtension {
                    uri: "u".into(),
                    payload_xml: format!(r#"<p:x xmlns:p="{XDA}" a="1"/>"#).into_bytes(),
                }],
            }],
            extensions: Vec::new(),
        }],
        cell_blocks: vec![MetadataBlock {
            records: vec![MetadataRecord {
                type_index: 1,
                value_index: 0,
            }],
            extensions: Vec::new(),
        }],
        value_blocks: Vec::new(),
        extensions: Vec::new(),
    }
}

#[test]
fn strict_round_trip_preserves_indices_and_extensions() {
    let value = sample();
    let xml = value.to_xml(true).unwrap();
    let parsed = WorkbookMetadata::parse(&xml).unwrap();
    assert_eq!(parsed.cell_block(1).unwrap().records[0].type_index, 1);
    assert!(parsed.cell_block(0).is_none());
    assert_eq!(parsed.to_xml(true).unwrap(), xml);
}

#[test]
fn mce_choice_selects_understood_metadata_branch() {
    let body = String::from_utf8(sample().to_xml(false).unwrap()).unwrap();
    let body = body
        .replace(
            "<metadataTypes",
            r#"<mc:AlternateContent><mc:Choice Requires="xda"><metadataTypes"#,
        )
        .replace(
            "</metadataTypes>",
            "</metadataTypes></mc:Choice><mc:Fallback/></mc:AlternateContent>",
        );
    let xml = body.replace(
            r#"<metadata xmlns=""#,
            &format!(
                r#"<metadata xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xda="{XDA}" xmlns=""#
            ),
        );
    assert_eq!(
        WorkbookMetadata::parse(xml.as_bytes()).unwrap().types.len(),
        1
    );
}

#[test]
fn rejects_malformed_and_out_of_bounds_values() {
    assert!(WorkbookMetadata::parse(br#"<!DOCTYPE x><metadata/>"#).is_err());
    assert!(WorkbookMetadata::parse(
            format!(
                r#"<metadata xmlns="{SML}"><metadataTypes count="2"><metadataType name="x" minSupportedVersion="1"/></metadataTypes></metadata>"#
            )
            .as_bytes(),
        )
        .is_err());

    let mut value = sample();
    value.cell_blocks[0].records[0].type_index = 2;
    assert!(value.to_xml(false).is_err());

    let mut value = sample();
    value.types[0].name = "x".repeat(MAX_STRING + 1);
    assert!(value.to_xml(false).is_err());
}
