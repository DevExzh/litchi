//! OPC graph lifecycle for XLSB External Data Connections.
//!
//! Connection payloads remain inert. This module only validates and mutates
//! the one relationship graph permitted by MS-XLSB 2.1.7.24.

use super::{Connections, parse_connections_part};
use crate::package::connections::write::write_connections_part;
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

fn invalid(detail: &'static str) -> Error {
    Error::Unrecognized {
        typ: "External Data Connections package graph",
        val: detail,
    }
}

pub(crate) fn load_from_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Connections>> {
    let Some(graph) = discover_graph(package, workbook_uri)? else {
        return Ok(None);
    };
    Ok(Some(parse_connections_part(
        package.get_part(&graph.part_name)?.blob(),
        raw_limits,
    )?))
}

/// [`load_from_workbook`], recording Excel-compatible repairs into `report`.
pub(crate) fn load_from_workbook_recording(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    report: &mut crate::leniency::ToleranceReport,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Connections>> {
    let Some(graph) = discover_graph(package, workbook_uri)? else {
        return Ok(None);
    };
    Ok(Some(super::parse::parse_connections_part_recording(
        package.get_part(&graph.part_name)?.blob(),
        report,
        raw_limits,
    )?))
}

/// Admit the inert connections owner under a caller's byte and typed-entry
/// ceilings before the semantic parser allocates its connection vector and
/// strings. The raw record pass borrows the owner payload and performs no
/// connection-specific cloning.
pub(crate) fn preflight(
    package: &OpcPackage,
    max_part_bytes: usize,
    max_connections: usize,
) -> Result<()> {
    let workbook = package.main_document_part()?;
    let Some(graph) = discover_graph(package, workbook.partname())? else {
        return Ok(());
    };
    let part = package.get_part(&graph.part_name)?;
    if part.blob().len() > max_part_bytes {
        return Err(Error::limit(
            crate::package::connections::LimitKind::PartBytes,
            part.blob().len(),
            max_part_bytes,
        ));
    }
    let raw_limit = max_part_bytes.min(crate::raw::MAX_WIRE_PAYLOAD);
    let mut records =
        crate::raw::Records::bounded(part.blob(), crate::raw::Limits::new(raw_limit, 1_048_576))?;
    let mut connection_count = 0_usize;
    for record in &mut records {
        if record?.kind() == crate::raw::kind::BEGIN_EXT_CONNECTION {
            connection_count = connection_count
                .checked_add(1)
                .ok_or(Error::CapacityOverflow {
                    resource: "External Data Connections",
                })?;
            if connection_count > max_connections {
                return Err(Error::limit(
                    crate::package::connections::LimitKind::Connections,
                    connection_count,
                    max_connections,
                ));
            }
        }
    }
    Ok(())
}

/// Load and validate the workbook-level connections owner, including its
/// complete OPC graph, under an explicit raw profile.
pub fn load(package: &OpcPackage, raw_limits: crate::raw::Limits) -> Result<Option<Connections>> {
    let workbook = package.main_document_part()?;
    load_from_workbook(package, workbook.partname(), raw_limits)
}

pub(crate) fn store_on_workbook(
    package: &mut OpcPackage,
    workbook_uri: &PackURI,
    connections: &Connections,
    raw_limits: crate::raw::Limits,
) -> Result<Connections> {
    let existing = discover_graph(package, workbook_uri)?;
    let payload = if let Some(graph) = existing.as_ref() {
        let source =
            super::codec::parse_source(package.get_part(&graph.part_name)?.blob(), raw_limits)?;
        let before =
            parse_connections_part(package.get_part(&graph.part_name)?.blob(), raw_limits)?;
        if has_query_tables(package)
            && before
                .connections
                .iter()
                .any(|old| connections.by_id(old.connection_id).is_none())
        {
            return Err(Error::UnsupportedFeature(
                "removing connections with query-table dependencies",
            ));
        }
        let canonical = write_connections_part(connections, raw_limits)?;
        super::codec::check_opaque_layout(
            package.get_part(&graph.part_name)?.blob(),
            &canonical,
            &source.unknown_records,
            raw_limits,
        )?;
        let (preserved, retained) = super::codec::retain_unchanged_connections(
            package.get_part(&graph.part_name)?.blob(),
            &before,
            connections,
            &canonical,
            raw_limits,
        )?;
        let remaining = source
            .unknown_records
            .into_iter()
            .filter(|r| r.connection_id().is_none_or(|id| !retained.contains(&id)))
            .collect::<Vec<_>>();
        super::codec::restore_opaque_positions(&preserved, &remaining, raw_limits)?
    } else {
        write_connections_part(connections, raw_limits)?
    };
    // Treat the reader as a post-serialization grammar oracle before mutation.
    let canonical_model = parse_connections_part(&payload, raw_limits)?;
    if let Some(graph) = existing.as_ref() {
        crate::signature_policy::ensure_unsigned(package)?;
        package.get_part_mut(&graph.part_name)?.set_blob(payload);
        load_from_workbook(package, workbook_uri, raw_limits)?
            .ok_or_else(|| invalid("stored connections graph was not discoverable"))?;
        return Ok(canonical_model);
    }
    let canonical_part = PackURI::new(CONNECTIONS_PART_NAME)?;
    package.validate_new_part_name(&canonical_part)?;
    ensure_no_inbound_relationship(package, &canonical_part)?;
    crate::signature_policy::ensure_unsigned(package)?;
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
    load_from_workbook(package, workbook_uri, raw_limits)?
        .ok_or_else(|| invalid("stored connections graph was not discoverable"))?;
    Ok(canonical_model)
}

fn has_query_tables(package: &OpcPackage) -> bool {
    package.iter_parts().any(|part| {
        part.content_type()
            .eq_ignore_ascii_case("application/vnd.ms-excel.queryTable")
    })
}

pub(crate) fn remove_from_workbook(
    package: &mut OpcPackage,
    workbook_uri: &PackURI,
) -> Result<bool> {
    let Some(graph) = discover_graph(package, workbook_uri)? else {
        return Ok(false);
    };
    if has_query_tables(package) {
        return Err(Error::UnsupportedFeature(
            "removing connections with query-table dependencies",
        ));
    }
    crate::signature_policy::ensure_unsigned(package)?;
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
    let target_name = relationship.target_partname()?;
    let part = package.get_part(&target_name)?;
    // Relationship resolution is case-insensitive, but mutation APIs remove
    // parts by their stored key. Retain the physical name so a package
    // authored with a case-variant owner can still be replaced, removed, and
    // restored.
    let part_name = part.partname().clone();
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid("connections part has content type"));
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
        part.content_type() == CONNECTIONS_CONTENT_TYPE
            && expected.is_none_or(|expected| !part.partname().is_equivalent_to(expected))
    }) {
        return Err(invalid(
            "package contains an orphan or additional connections part",
        ));
    }
    Ok(())
}

fn ensure_no_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Err(invalid(
                "canonical connections part name has a dangling inbound relationship",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
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
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Err(invalid(
                "connections part has an unexpected package-level relationship",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external()
                || !relationship.target_partname()?.is_equivalent_to(target)
            {
                continue;
            }
            if !part.partname().is_equivalent_to(expected_source)
                || relationship.r_id() != expected_relationship_id
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
    use crate::Package;
    use crate::builder::{self, Builder};
    use crate::merged_cells::MergedCell;
    use crate::package::connections::{
        CommandType, Connection, CredentialMethod, DbProperties, Properties, SourceType,
    };

    fn generated_package() -> Package {
        let mut writer = Builder::new();
        writer.add_worksheet(builder::Worksheet::new("Sheet1"));
        let bytes = writer.build().unwrap().to_bytes().unwrap();
        Package::from_bytes(&bytes).unwrap()
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

    fn set_connections(package: &Package, value: Option<Connections>) -> Result<Package> {
        let mut edit = package.edit_connections()?;
        edit.replace(value)?;
        Ok(edit.commit()?.into_package())
    }

    fn reopened(package: &Package) -> Package {
        Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn parsed_workbook_adds_replaces_preserves_and_removes_connections() {
        let first = connections(7, "Generated warehouse");
        let package = set_connections(&generated_package(), Some(first.clone())).unwrap();
        assert_eq!(package.workbook().unwrap().connections(), Some(&first));

        // An unrelated binary worksheet mutation must preserve the package graph.
        let merged = MergedCell::new(0, 0, 0, 1);
        let mut edit = package.edit_merged_cells().unwrap();
        edit.set_merged_cell_ranges(0, std::slice::from_ref(&merged))
            .unwrap();
        let package = edit.commit().unwrap().into_package();

        let replacement = connections(9, "Replacement");
        let package = set_connections(&package, Some(replacement.clone())).unwrap();
        let reopened_package = reopened(&package);
        let workbook = reopened_package.workbook().unwrap();
        assert_eq!(workbook.connections(), Some(&replacement));
        assert_eq!(workbook.merged_cell_ranges(0).unwrap(), vec![merged]);

        let mut edit = reopened_package.edit_connections().unwrap();
        assert!(edit.replace(None).unwrap());
        let removed = edit.commit().unwrap().into_package();
        assert!(removed.workbook().unwrap().connections().is_none());
        let mut edit = removed.edit_connections().unwrap();
        assert!(!edit.replace(None).unwrap());
        assert!(!edit.commit().unwrap().diagnostics().changed());
        assert!(
            reopened(&removed)
                .workbook()
                .unwrap()
                .connections()
                .is_none()
        );
    }

    #[test]
    fn raw_opc_edit_preserves_the_validated_connections_graph() {
        let original = connections(7, "Raw edit connection");
        let package = set_connections(&generated_package(), Some(original.clone())).unwrap();
        let marker = PackURI::new("/xl/raw-edit-marker.bin").unwrap();

        let mut edit = package.edit_opc().unwrap();
        edit.update(|package| {
            package.try_add_part(Box::new(BlobPart::new(
                marker.clone(),
                "application/octet-stream".to_string(),
                b"preserve connections".to_vec(),
            )))?;
            Ok::<_, Error>(())
        })
        .unwrap();
        let package = edit.commit().unwrap().into_package();

        assert_eq!(package.workbook().unwrap().connections(), Some(&original));
        assert_eq!(
            package.opc_package().get_part(&marker).unwrap().blob(),
            b"preserve connections"
        );
        assert_eq!(
            reopened(&package).workbook().unwrap().connections(),
            Some(&original)
        );
    }

    #[test]
    fn invalid_replacement_is_rejected_before_package_mutation() {
        let original = connections(7, "Original");
        let package = set_connections(&generated_package(), Some(original.clone())).unwrap();
        let uri = PackURI::new(CONNECTIONS_PART_NAME).unwrap();
        let original_payload = package
            .opc_package()
            .get_part(&uri)
            .unwrap()
            .blob()
            .to_vec();

        let mut invalid = connections(8, "Invalid");
        invalid.connections[0].refresh_interval_minutes = 32_768;
        assert!(set_connections(&package, Some(invalid)).is_err());
        assert_eq!(package.workbook().unwrap().connections(), Some(&original));
        assert_eq!(
            package.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );

        let mut invalid = connections(8, "Invalid");
        invalid.connections[0].description = Some("x".repeat(256));
        assert!(set_connections(&package, Some(invalid)).is_err());
        assert_eq!(package.workbook().unwrap().connections(), Some(&original));
        assert_eq!(
            package.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );
    }

    #[test]
    fn malformed_existing_graph_is_not_replaced() {
        let original = connections(7, "Original");
        let package = set_connections(&generated_package(), Some(original.clone())).unwrap();
        let uri = PackURI::new(CONNECTIONS_PART_NAME).unwrap();
        let mut opc = package.opc_package().clone();
        opc.get_part_mut(&uri).unwrap().rels_mut().add_relationship(
            "urn:generated:forbidden".to_string(),
            "child.bin".to_string(),
            "rIdForbidden".to_string(),
            false,
        );
        let malformed = Package::from(opc);
        let original_payload = malformed
            .opc_package()
            .get_part(&uri)
            .unwrap()
            .blob()
            .to_vec();

        assert!(set_connections(&malformed, Some(connections(9, "Replacement"))).is_err());
        assert_eq!(
            malformed.opc_package().get_part(&uri).unwrap().blob(),
            original_payload
        );
    }
}
