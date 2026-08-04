//! OPC relationship boundary for the XLSX workbook-metadata owner.
//!
//! `litchi_xlsx` owns the bounded SpreadsheetML metadata XML grammar,
//! validation, MCE processing, and inert extension payloads. This module
//! `litchi_xlsx` owns the semantic model and XML codec; this module only
//! resolves the workbook relationship and translates package-boundary errors.

use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;
use litchi_xlsx::workbook_metadata as owner;

pub use owner::{
    FutureMetadata, MetadataBehavior, MetadataBlock, MetadataRecord, MetadataType,
    OpaqueMetadataExtension, WorkbookMetadata,
};

#[cfg(test)]
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = owner::SHEET_METADATA_RELATIONSHIP_TYPE;
const REL_STRICT: &str = owner::STRICT_SHEET_METADATA_RELATIONSHIP_TYPE;
const CONTENT_TYPE: &str = owner::SHEET_METADATA_CONTENT_TYPE;
#[cfg(test)]
const XDA: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
#[cfg(test)]
const MAX_STRING: usize = 1024 * 1024;

/// Loads the optional workbook metadata part selected by the workbook relationship.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<WorkbookMetadata>> {
    let workbook = package.main_document_part()?;
    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | REL_STRICT));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(OoxmlError::InvalidRelationship(
            "workbook has multiple sheetMetadata relationships".into(),
        ));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(
            "sheetMetadata relationship must be internal".into(),
        ));
    }
    let uri = relationship.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: CONTENT_TYPE.into(),
            got: part.content_type().into(),
        });
    }
    if part.rels().iter().next().is_some() {
        return Err(OoxmlError::InvalidRelationship(
            "workbook metadata part must not have relationships".into(),
        ));
    }
    Ok(Some(
        WorkbookMetadata::parse(part.blob()).map_err(map_owner_error)?,
    ))
}

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, PackURI, Part};
    fn fixture(bytes: &[u8]) -> WorkbookMetadata {
        let p = OpcPackage::from_bytes(bytes).unwrap();
        load_from_package(&p).unwrap().unwrap()
    }
    #[test]
    fn poi_xlookup_dynamic_array() {
        let m = fixture(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/xlookup.xlsx"
        ));
        assert_eq!(m.types[0].name, "XLDAPR");
        assert_eq!(
            m.cell_blocks[0].records[0],
            MetadataRecord {
                type_index: 1,
                value_index: 0
            }
        );
        assert!(
            std::str::from_utf8(&m.future[0].blocks[0].extensions[0].payload_xml)
                .unwrap()
                .contains("dynamicArrayProperties")
        );
    }
    #[test]
    fn libreoffice_dynamic_and_rich_value_fixtures() {
        let m = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/functions/dynamic_array/xlsx/DynamicArrayFixture.xlsx"
        ));
        assert_eq!(m.future.len(), 1);
        let spill = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/Spill.xlsx"
        ));
        assert_eq!(spill.types.len(), 2);
        assert_eq!(spill.value_blocks.len(), 1);
        assert_eq!(spill.value_blocks[0].records[0].type_index, 2);
        let lambda = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/LambdaAndRelatedFunctions.xlsx"
        ));
        assert_eq!(lambda.future[1].blocks.len(), 2);
        assert_eq!(lambda.value_blocks.len(), 2);
    }
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
                    records: vec![],
                    extensions: vec![OpaqueMetadataExtension {
                        uri: "u".into(),
                        payload_xml: format!("<p:x xmlns:p=\"{XDA}\" a=\"1\"/>").into_bytes(),
                    }],
                }],
                extensions: vec![],
            }],
            cell_blocks: vec![MetadataBlock {
                records: vec![MetadataRecord {
                    type_index: 1,
                    value_index: 0,
                }],
                extensions: vec![],
            }],
            value_blocks: vec![],
            extensions: vec![],
        }
    }
    #[test]
    fn strict_deterministic_roundtrip_and_index_api() {
        let m = sample();
        let x = m.to_xml(true).unwrap();
        let p = WorkbookMetadata::parse(&x).unwrap();
        assert_eq!(p.cell_block(1).unwrap().records[0].type_index, 1);
        assert!(p.cell_block(0).is_none());
        assert_eq!(p.to_xml(true).unwrap(), x);
    }
    #[test]
    fn nested_mce_selects_understood_branch() {
        let body = String::from_utf8(sample().to_xml(false).unwrap()).unwrap();
        let body = body
            .replace(
                "<metadataTypes",
                "<mc:AlternateContent><mc:Choice Requires=\"xda\"><metadataTypes",
            )
            .replace(
                "</metadataTypes>",
                "</metadataTypes></mc:Choice><mc:Fallback/></mc:AlternateContent>",
            );
        let xml=body.replace("<metadata xmlns=\"",&format!("<metadata xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xda=\"{XDA}\" xmlns=\""));
        assert_eq!(
            WorkbookMetadata::parse(xml.as_bytes()).unwrap().types.len(),
            1
        );
    }
    #[test]
    fn malformed_and_bounds() {
        assert!(WorkbookMetadata::parse(br#"<!DOCTYPE x><metadata/>"#).is_err());
        assert!(WorkbookMetadata::parse(format!("<metadata xmlns=\"{SML}\"><metadataTypes count=\"2\"><metadataType name=\"x\" minSupportedVersion=\"1\"/></metadataTypes></metadata>").as_bytes()).is_err());
        let mut m = sample();
        m.cell_blocks[0].records[0].type_index = 2;
        assert!(m.to_xml(false).is_err());
        let mut m = sample();
        m.types[0].name = "x".repeat(MAX_STRING + 1);
        assert!(m.to_xml(false).is_err());
    }
    fn package(external: bool, wrong: bool, outbound: bool) -> OpcPackage {
        let mut wb = BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            b"<workbook/>".to_vec(),
        );
        wb.rels_mut().add_relationship(
            REL.into(),
            if external {
                "https://invalid.example/m".into()
            } else {
                "metadata.xml".into()
            },
            "rId1".into(),
            external,
        );
        let mut md = BlobPart::new(
            PackURI::new("/xl/metadata.xml").unwrap(),
            if wrong {
                "text/xml".into()
            } else {
                CONTENT_TYPE.into()
            },
            sample().to_xml(false).unwrap(),
        );
        if outbound {
            md.rels_mut()
                .add_relationship("x".into(), "other.xml".into(), "rId1".into(), false);
        }
        let mut p = OpcPackage::new();
        p.add_part(Box::new(wb));
        p.add_part(Box::new(md));
        p.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "xl/workbook.xml".into(),
            "rId1".into(),
            false,
        );
        p
    }
    #[test]
    fn package_relationship_matrix() {
        assert!(
            load_from_package(&package(false, false, false))
                .unwrap()
                .is_some()
        );
        assert!(load_from_package(&package(true, false, false)).is_err());
        assert!(load_from_package(&package(false, true, false)).is_err());
        assert!(load_from_package(&package(false, false, true)).is_err());
    }
}
