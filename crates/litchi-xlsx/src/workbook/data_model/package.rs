//! OPC singleton graph lifecycle for an XLSX workbook Data Model.

use std::collections::HashSet;

use litchi_opc::{BlobPart, OpcPackage, PackURI};

use crate::error::Result;
use crate::package::xldm::inspect;

use super::codec::{
    insert_extension, parse_document, validate_definition, workbook_definition, write_data_model,
};
use super::model::{Definition, Model, Payload};
use super::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP_TYPE, DATA_MODEL_CONTENT_TYPE,
    DATA_MODEL_EXTENSION_URI, DATA_MODEL_PART_NAME, MAX_PAYLOAD_BYTES,
    STRICT_CONNECTIONS_RELATIONSHIP_TYPE, invalid, limit,
};

/// OPC content type for an MS-XLDM payload.
/// Load the workbook's optional singleton Data Model and keep its payload inert.
pub fn load_data_model(package: &OpcPackage, workbook_name: &PackURI) -> Result<Option<Model>> {
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = parse_document(workbook.blob())?;
    let (_, definition) = workbook_definition(&workbook_root)?;
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == DATA_MODEL_CONTENT_TYPE);
    let part = parts.next();
    if parts.next().is_some() {
        return Err(invalid("package contains multiple Data Model parts"));
    }
    let (definition, part) = match (definition, part) {
        (Some(definition), Some(part)) => (definition, part),
        (Some(_), None) => {
            return Err(invalid(
                "workbook dataModel extension has no Data Model part",
            ));
        },
        (None, Some(_)) => {
            return Err(invalid(
                "Data Model part has no workbook dataModel extension",
            ));
        },
        (None, None) => return Ok(None),
    };
    if part.partname().as_str() != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part '{}' must be '{DATA_MODEL_PART_NAME}'",
            part.partname()
        )));
    }
    if part.blob().is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if part.blob().len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    inspect(part.blob())?;
    if !part.rels().is_empty() {
        return Err(invalid(
            "Data Model part has forbidden outbound relationships",
        ));
    }
    reject_inbound_relationships(package, part.partname())?;
    validate_connections(package, workbook_name, &definition)?;
    Ok(Some(Model {
        definition,
        payload: Payload {
            part_name: part.partname().to_string(),
            data: part.blob().to_vec(),
        },
    }))
}

/// Store a singleton Data Model after validating the complete mutation plan.
pub fn store_data_model(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Model,
) -> Result<()> {
    validate_definition(&value.definition, false)?;
    validate_payload(&value.payload)?;
    if load_data_model(package, workbook_name)?.is_some() {
        return Err(invalid("workbook already contains a Data Model"));
    }
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == DATA_MODEL_PART_NAME)
    {
        return Err(invalid(format!(
            "part '{DATA_MODEL_PART_NAME}' already exists"
        )));
    }
    validate_connections(package, workbook_name, &value.definition)?;
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    let (core, existing) = workbook_definition(&root)?;
    if existing.is_some() {
        return Err(invalid("workbook already has a dataModel extension"));
    }
    let descriptor = write_data_model(&value.definition)?;
    let mut fragment = Vec::new();
    fragment.extend_from_slice(b"<x:ext xmlns:x=\"");
    escape(&mut fragment, core);
    fragment.extend_from_slice(b"\" uri=\"");
    escape(&mut fragment, DATA_MODEL_EXTENSION_URI);
    fragment.extend_from_slice(b"\">");
    fragment.extend_from_slice(&descriptor);
    fragment.extend_from_slice(b"</x:ext>");
    let updated = insert_extension(workbook.blob(), core, &fragment)?;
    let uri = PackURI::new(&value.payload.part_name).map_err(|error| invalid(error.to_string()))?;
    package.try_add_part(Box::new(BlobPart::new(
        uri,
        DATA_MODEL_CONTENT_TYPE.into(),
        value.payload.data.clone(),
    )))?;
    package.get_part_mut(workbook_name)?.set_blob(updated);
    package.unsign();
    Ok(())
}

fn validate_payload(value: &Payload) -> Result<()> {
    if value.part_name != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part must be '{DATA_MODEL_PART_NAME}'"
        )));
    }
    if value.data.is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if value.data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    inspect(&value.data)?;
    let _ = PackURI::new(&value.part_name).map_err(|error| invalid(error.to_string()))?;
    Ok(())
}

fn validate_connections(
    package: &OpcPackage,
    workbook_name: &PackURI,
    definition: &Definition,
) -> Result<()> {
    if definition.tables.is_empty() {
        return Ok(());
    }
    let workbook = package.get_part(workbook_name)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            CONNECTIONS_RELATIONSHIP_TYPE | STRICT_CONNECTIONS_RELATIONSHIP_TYPE
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("Data Model tables require a workbook Connections part"))?;
    if relationships.next().is_some() {
        return Err(invalid("workbook has multiple Connections relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("Connections relationship cannot be external"));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "Connections part '{target}' has content type '{}', expected '{CONNECTIONS_CONTENT_TYPE}'",
            part.content_type()
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Connections part has forbidden outbound relationships",
        ));
    }
    let connections = crate::connections::Connections::parse(part.blob())
        .map_err(|error| invalid(format!("invalid Connections part: {error}")))?;
    let names: HashSet<String> = connections
        .connections
        .iter()
        .filter_map(|connection| connection.name.as_ref())
        .map(|name| name.to_lowercase())
        .collect();
    for table in &definition.tables {
        if !names.contains(&table.connection.to_lowercase()) {
            return Err(invalid(format!(
                "Data Model table '{}' references unknown workbook connection '{}'",
                table.name, table.connection
            )));
        }
    }
    Ok(())
}

fn reject_inbound_relationships(package: &OpcPackage, target: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname()?.as_str() == target.as_str()
        {
            return Err(invalid(
                "package relationship targets the relationship-free Data Model part",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.as_str() == target.as_str()
            {
                return Err(invalid(format!(
                    "part '{}' has a relationship to the relationship-free Data Model part",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::codec::parse_data_model;
    use super::*;
    use crate::package::xldm::test_xldm_bytes;
    use crate::workbook::data_model::{MAX_XML_BYTES, SML, X15};
    use litchi_opc::Part;

    fn definition() -> Definition {
        Definition {
            min_version_load: 7,
            tables: vec![
                super::super::model::Table {
                    id: "t-sales".into(),
                    name: "Sales".into(),
                    connection: "ModelConnection".into(),
                },
                super::super::model::Table {
                    id: "t-date".into(),
                    name: "Date".into(),
                    connection: "ModelConnection".into(),
                },
            ],
            relationships: vec![super::super::model::Relationship {
                from_table: "Sales".into(),
                from_column: "DateKey".into(),
                to_table: "Date".into(),
                to_column: "DateKey".into(),
            }],
            extension_list: Some(super::super::model::OpaqueXml {
                xml: format!(
                    r#"<x15:extLst xmlns:x15="{X15}"><x15:ext uri="urn:test"><v:opaque xmlns:v="urn:vendor"/></x15:ext></x15:extLst>"#
                )
                .into_bytes(),
            }),
        }
    }

    fn fixture_package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let mut part = BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        );
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        part.rels_mut().add_relationship(
            CONNECTIONS_RELATIONSHIP_TYPE.into(),
            "connections.xml".into(),
            "rIdConnections".into(),
            false,
        );
        package.add_part(Box::new(part));
        package.add_part(Box::new(BlobPart::new(
            connections,
            CONNECTIONS_CONTENT_TYPE.into(),
            format!(r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection" refreshedVersion="7"/></connections>"#).into_bytes(),
        )));
        (package, workbook)
    }

    fn model() -> Model {
        Model {
            definition: definition(),
            payload: Payload {
                part_name: DATA_MODEL_PART_NAME.into(),
                data: test_xldm_bytes(),
            },
        }
    }

    #[test]
    fn typed_descriptor_round_trip() {
        let expected = definition();
        let xml = write_data_model(&expected).unwrap();
        let actual = parse_data_model(&xml).unwrap();
        assert_eq!(actual.min_version_load, expected.min_version_load);
        assert_eq!(actual.tables, expected.tables);
        assert_eq!(actual.relationships, expected.relationships);
        assert!(String::from_utf8_lossy(&actual.extension_list.unwrap().xml).contains("opaque"));
    }

    #[test]
    fn package_round_trip_preserves_inert_payload_and_inline_metadata() {
        let (mut package, workbook) = fixture_package();
        let expected = model();
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let actual = load_data_model(&package, &workbook).unwrap().unwrap();
        assert_eq!(
            actual.definition.min_version_load,
            expected.definition.min_version_load
        );
        assert_eq!(actual.definition.tables, expected.definition.tables);
        assert_eq!(
            actual.definition.relationships,
            expected.definition.relationships
        );
        assert!(
            String::from_utf8_lossy(&actual.definition.extension_list.as_ref().unwrap().xml)
                .contains("opaque")
        );
        assert_eq!(actual.payload, expected.payload);
    }

    #[test]
    fn inserts_into_existing_empty_extension_list() {
        let (mut package, workbook) = fixture_package();
        package.get_part_mut(&workbook).unwrap().set_blob(
            format!(r#"<workbook xmlns="{SML}"><sheets/><extLst /></workbook>"#).into_bytes(),
        );
        store_data_model(&mut package, &workbook, &model()).unwrap();
        assert!(load_data_model(&package, &workbook).unwrap().is_some());
    }

    #[test]
    fn rejects_hostile_xml_schema_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<?bad x?><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}" minVersionLoad="4"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelTables/></x15:dataModel>"#),
            format!(
                r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelRelationships><x15:modelRelationship fromTable="Missing" fromColumn="a" toTable="Missing" toColumn="b"/></x15:modelRelationships></x15:dataModel>"#
            ),
        ] {
            assert!(parse_data_model(xml.as_bytes()).is_err());
        }
        assert!(parse_data_model(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_missing_connection_and_unknown_table_references() {
        let mut value = definition();
        value.tables[0].connection = "Absent".into();
        let (mut package, workbook) = fixture_package();
        assert!(
            store_data_model(
                &mut package,
                &workbook,
                &Model {
                    definition: value,
                    payload: model().payload,
                }
            )
            .is_err()
        );
        let mut value = definition();
        value.relationships[0].to_table = "Absent".into();
        assert!(write_data_model(&value).is_err());
    }

    #[test]
    fn rejects_orphan_duplicate_wrong_path_and_relationship_edges() {
        let (mut package, workbook) = fixture_package();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).unwrap(),
            DATA_MODEL_CONTENT_TYPE.into(),
            vec![1],
        )));
        assert!(load_data_model(&package, &workbook).is_err());
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        package
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "model/item.data".into(),
                "rIdModel".into(),
                false,
            );
        assert!(load_data_model(&package, &workbook).is_err());
        let mut wrong = model();
        wrong.payload.part_name = "/xl/model/other.data".into();
        let (mut package, workbook) = fixture_package();
        assert!(store_data_model(&mut package, &workbook, &wrong).is_err());
    }
}
