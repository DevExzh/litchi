//! OPC graph lifecycle for XLSB External Data Connections.
//!
//! Connection payloads remain inert. This module only validates and mutates
//! the one relationship graph permitted by MS-XLSB 2.1.7.24.

use super::{Connections, parse_connections_part};
use crate::package::connections::write::{
    write_connections_part, write_connections_part_with_unknown,
};
use crate::package::error::{Error, Result};
use litchi_opc::constants::content_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::sync::Arc;

pub const CONNECTIONS_CONTENT_TYPE: &str = "application/vnd.ms-excel.connections";
pub const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
pub const CONNECTIONS_PART_NAME: &str = "/xl/connections.bin";

#[derive(Debug)]
struct ConnectionsGraph {
    relationship_id: String,
    part_name: PackURI,
}

fn invalid(detail: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: "External Data Connections package graph".to_string(),
        val: detail.into(),
    }
}

pub(crate) fn load_from_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<Connections>> {
    let Some(graph) = discover_graph(package, workbook_uri)? else {
        return Ok(None);
    };
    Ok(Some(parse_connections_part(
        package.get_part(&graph.part_name)?.blob(),
    )?))
}

/// Load and validate the workbook-level connections owner.
pub fn load(package: &OpcPackage) -> Result<Option<Connections>> {
    let workbook = package.main_document_part()?;
    load_from_workbook(package, workbook.partname())
}

/// Validate the complete OPC connections graph and return its typed model.
pub fn validate_graph(package: &OpcPackage) -> Result<Option<Connections>> {
    load(package)
}

pub(crate) fn store_on_workbook(
    package: &mut OpcPackage,
    workbook_uri: &PackURI,
    connections: &Connections,
) -> Result<Connections> {
    let existing = discover_graph(package, workbook_uri)?;
    let payload = if let Some(graph) = existing.as_ref() {
        let source = super::codec::parse_source(package.get_part(&graph.part_name)?.blob())?;
        write_connections_part_with_unknown(connections, &source.unknown_records)?
    } else {
        write_connections_part(connections)?
    };
    // Treat the reader as a post-serialization grammar oracle before mutation.
    let canonical_model = parse_connections_part(&payload)?;
    let canonical_part = PackURI::new(CONNECTIONS_PART_NAME).map_err(Error::Encoding)?;
    if existing
        .as_ref()
        .is_none_or(|graph| graph.part_name != canonical_part)
    {
        package.validate_new_part_name(&canonical_part)?;
        ensure_no_inbound_relationship(package, &canonical_part)?;
    }

    package.unsign();
    if let Some(graph) = existing {
        package
            .get_part_mut(workbook_uri)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.part_name);
    }
    package.try_add_part(Box::new(BlobPart::new(
        canonical_part.clone(),
        CONNECTIONS_CONTENT_TYPE.to_string(),
        payload,
    )))?;
    let target = canonical_part.relative_ref(workbook_uri.base_uri());
    package
        .get_part_mut(workbook_uri)?
        .rels_mut()
        .get_or_add(CONNECTIONS_RELATIONSHIP_TYPE, &target);
    load_from_workbook(package, workbook_uri)?
        .ok_or_else(|| invalid("stored connections graph was not discoverable"))?;
    Ok(canonical_model)
}

pub(crate) fn remove_from_workbook(
    package: &mut OpcPackage,
    workbook_uri: &PackURI,
) -> Result<bool> {
    let Some(graph) = discover_graph(package, workbook_uri)? else {
        return Ok(false);
    };
    package.unsign();
    package
        .get_part_mut(workbook_uri)?
        .rels_mut()
        .remove(&graph.relationship_id);
    package.remove_part(&graph.part_name);
    Ok(true)
}

/// Exact source relationship metadata used by source-checked patches.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipImage {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) external: bool,
}

/// Exact connections-part source image used by source-checked patches.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PartImage {
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
}

/// Exact owner graph source image used by source-checked patches.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SourceImage {
    pub(crate) workbook_name: PackURI,
    pub(crate) workbook_bytes: Arc<Vec<u8>>,
    pub(crate) workbook_relationships: Vec<RelationshipImage>,
    pub(crate) root_relationships: Vec<RelationshipImage>,
    pub(crate) connection: Option<PartImage>,
    pub(crate) connection_relationship: Option<RelationshipImage>,
}

pub(crate) fn capture_source(package: &OpcPackage) -> Result<SourceImage> {
    let workbook = package.main_document_part()?;
    let graph = discover_graph(package, workbook.partname())?;
    let connection_relationship = workbook
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE)
        .map(|relationship| RelationshipImage {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            external: relationship.is_external(),
        });
    let connection = graph
        .as_ref()
        .map(|graph| package.get_part(&graph.part_name))
        .transpose()?
        .map(|part| PartImage {
            part_name: part.partname().clone(),
            content_type: part.content_type().to_string(),
            bytes: part.blob_arc(),
        });
    Ok(SourceImage {
        workbook_name: workbook.partname().clone(),
        workbook_bytes: workbook.blob_arc(),
        workbook_relationships: relationship_images(workbook),
        root_relationships: relationship_images_root(package),
        connection,
        connection_relationship,
    })
}

pub(crate) fn restore_source(package: &mut OpcPackage, source: &SourceImage) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    if workbook_name != source.workbook_name {
        return Err(invalid("workbook part identity changed"));
    }
    if let Some(graph) = discover_graph(package, &workbook_name)? {
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.part_name);
    }
    if let (Some(part), Some(relationship)) = (&source.connection, &source.connection_relationship)
    {
        package.try_add_part(Box::new(BlobPart::new(
            part.part_name.clone(),
            part.content_type.clone(),
            part.bytes.as_ref().clone(),
        )))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.external,
            );
    }
    validate_graph(package).map(|_| ())
}

fn relationship_images(part: &dyn Part) -> Vec<RelationshipImage> {
    let mut values = part
        .rels()
        .iter()
        .map(|relationship| RelationshipImage {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            external: relationship.is_external(),
        })
        .collect::<Vec<_>>();
    values.sort_by(|left, right| left.id.cmp(&right.id));
    values
}

fn relationship_images_root(package: &OpcPackage) -> Vec<RelationshipImage> {
    let mut values = package
        .rels()
        .iter()
        .map(|relationship| RelationshipImage {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            external: relationship.is_external(),
        })
        .collect::<Vec<_>>();
    values.sort_by(|left, right| left.id.cmp(&right.id));
    values
}

fn discover_graph(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<ConnectionsGraph>> {
    let workbook = package.get_part(workbook_uri)?;
    if workbook.content_type() != content_type::XLSB_BIN {
        return Err(invalid(
            "relationship source is not the XLSB main workbook part",
        ));
    }
    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE);
    let Some(relationship) = relationships.next() else {
        ensure_no_orphan_parts(package, None)?;
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook declares multiple connections relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("connections relationship cannot be external"));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "connections part '{}' has content type '{}'",
            part_name.as_str(),
            part.content_type()
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid("connections part must not have relationships"));
    }
    ensure_no_orphan_parts(package, Some(&part_name))?;
    ensure_exclusive_inbound_relationship(package, &part_name, workbook_uri, relationship.r_id())?;
    Ok(Some(ConnectionsGraph {
        relationship_id: relationship.r_id().to_string(),
        part_name,
    }))
}

fn ensure_no_orphan_parts(package: &OpcPackage, expected: Option<&PackURI>) -> Result<()> {
    if package.iter_parts().any(|part| {
        part.content_type() == CONNECTIONS_CONTENT_TYPE && expected != Some(part.partname())
    }) {
        return Err(invalid(
            "package contains an orphan or additional connections part",
        ));
    }
    Ok(())
}

fn ensure_no_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Err(invalid(
                "canonical connections part name has a dangling inbound relationship",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Err(invalid(
                    "canonical connections part name has a dangling inbound relationship",
                ));
            }
        }
    }
    Ok(())
}

fn ensure_exclusive_inbound_relationship(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_relationship_id: &str,
) -> Result<()> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Err(invalid(
                "connections part has an unexpected package-level relationship",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external() || relationship.target_partname()? != *target {
                continue;
            }
            if part.partname() != expected_source || relationship.r_id() != expected_relationship_id
            {
                return Err(invalid(
                    "connections part has an unexpected inbound relationship",
                ));
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::Workbook;
    use crate::package::connections::{
        CommandType, Connection, CredentialMethod, DbProperties, Properties, SourceType,
    };
    use crate::package::merged_cells::MergedCell;
    use crate::writer::{MutableWorksheet, WorkbookWriter};
    use std::io::Cursor;

    fn generated_workbook() -> Workbook {
        let mut writer = WorkbookWriter::new();
        writer.add_worksheet(MutableWorksheet::new("Sheet1"));
        let mut bytes = Cursor::new(Vec::new());
        writer.save(&mut bytes).unwrap();
        Workbook::new(Cursor::new(bytes.into_inner())).unwrap()
    }

    fn connections(id: u32, name: &str) -> Connections {
        Connections {
            connections: vec![Connection {
                connection_id: id,
                source_type: SourceType::Odbc,
                name: name.to_string(),
                refresh_interval_minutes: 15,
                credential_method: Some(CredentialMethod::Integrated),
                properties: Properties::Database(DbProperties {
                    command_type: CommandType::Sql,
                    connection_string: "Driver={Generated};Server=example.invalid".to_string(),
                    command: Some("SELECT 1".to_string()),
                    server_command: None,
                }),
                ..Connection::default()
            }],
        }
    }

    fn saved(workbook: &Workbook) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        workbook.save(&mut bytes).unwrap();
        bytes.into_inner()
    }

    #[test]
    fn parsed_workbook_adds_replaces_preserves_and_removes_connections() {
        let mut workbook = generated_workbook();
        let first = connections(7, "Generated warehouse");
        workbook.set_connections(first.clone()).unwrap();
        assert_eq!(workbook.connections(), Some(&first));

        // An unrelated binary worksheet mutation must preserve the package graph.
        let merged = MergedCell::new(0, 0, 0, 1);
        workbook
            .set_merged_cell_ranges(0, std::slice::from_ref(&merged))
            .unwrap();

        let replacement = connections(9, "Replacement");
        workbook.set_connections(replacement.clone()).unwrap();
        let mut reopened = Workbook::new(Cursor::new(saved(&workbook))).unwrap();
        assert_eq!(reopened.connections(), Some(&replacement));
        assert_eq!(reopened.merged_cell_ranges(0).unwrap(), vec![merged]);

        assert!(reopened.remove_connections().unwrap());
        assert!(reopened.connections().is_none());
        assert!(!reopened.remove_connections().unwrap());
        let reopened = Workbook::new(Cursor::new(saved(&reopened))).unwrap();
        assert!(reopened.connections().is_none());
    }

    #[test]
    fn raw_opc_edit_preserves_the_validated_connections_graph() {
        let mut workbook = generated_workbook();
        let original = connections(7, "Raw edit connection");
        workbook.set_connections(original.clone()).unwrap();
        let marker = PackURI::new("/xl/raw-edit-marker.bin").unwrap();

        workbook
            .edit_opc(|package| {
                package.try_add_part(Box::new(BlobPart::new(
                    marker.clone(),
                    "application/octet-stream".to_string(),
                    b"preserve connections".to_vec(),
                )))?;
                Ok::<_, Error>(())
            })
            .unwrap();

        assert_eq!(workbook.connections(), Some(&original));
        assert_eq!(
            workbook.opc_package().get_part(&marker).unwrap().blob(),
            b"preserve connections"
        );
        let reopened = Workbook::new(Cursor::new(saved(&workbook))).unwrap();
        assert_eq!(reopened.connections(), Some(&original));
    }

    #[test]
    fn invalid_replacement_is_rejected_before_package_mutation() {
        let mut workbook = generated_workbook();
        let original = connections(7, "Original");
        workbook.set_connections(original.clone()).unwrap();
        let uri = PackURI::new(CONNECTIONS_PART_NAME).unwrap();
        let original_payload = workbook
            .opc_package()
            .get_part(&uri)
            .unwrap()
            .blob()
            .to_vec();

        let mut invalid = connections(8, "Invalid");
        invalid.connections[0].refresh_interval_minutes = 32_768;
        assert!(workbook.set_connections(invalid).is_err());
        assert_eq!(workbook.connections(), Some(&original));
        assert_eq!(
            workbook.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );

        let mut invalid = connections(8, "Invalid");
        invalid.connections[0].description = Some("x".repeat(256));
        assert!(workbook.set_connections(invalid).is_err());
        assert_eq!(workbook.connections(), Some(&original));
        assert_eq!(
            workbook.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );
    }

    #[test]
    fn malformed_existing_graph_is_not_replaced() {
        let mut workbook = generated_workbook();
        let original = connections(7, "Original");
        workbook.set_connections(original.clone()).unwrap();
        let uri = PackURI::new(CONNECTIONS_PART_NAME).unwrap();
        workbook
            .opc_package_mut()
            .get_part_mut(&uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:generated:forbidden".to_string(),
                "child.bin".to_string(),
                "rIdForbidden".to_string(),
                false,
            );
        let original_payload = workbook
            .opc_package()
            .get_part(&uri)
            .unwrap()
            .blob()
            .to_vec();

        assert!(
            workbook
                .set_connections(connections(9, "Replacement"))
                .is_err()
        );
        assert_eq!(workbook.connections(), Some(&original));
        assert_eq!(
            workbook.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );
    }
}
