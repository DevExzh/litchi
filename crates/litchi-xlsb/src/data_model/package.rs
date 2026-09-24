//! OPC ownership and lazy inspection for the XLSB model part.

use std::collections::HashMap;

use litchi_opc::{OpcPackage, PackURI, TargetMode};

use super::model::{Definition, Model, ModelPart};
use crate::package::connections::Connections;
use crate::package::error::{Error, Result};

/// MS-XLSX 2.1.6 model-part content type reused by MS-XLSB 2.1.7.35.
pub const DATA_MODEL_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.model+data";
/// The fixed OPC part name used by Excel model packages.
pub const DATA_MODEL_PART_NAME: &str = "/xl/model/item.data";
const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;

/// Admit the package metadata that a Data Model snapshot will retain before
/// modeled graph values or the public package facade are cloned. OPC part
/// blobs are Arc-backed; this pass conservatively charges part URI, content-type
/// and relationship fields while borrowing the caller's graph. Source-provenance
/// maps and other OPC-owned state have separate limits and are not a measured
/// part of this estimate.
pub(crate) fn preflight_metadata(package: &OpcPackage, limits: super::ReadLimits) -> Result<()> {
    if package.part_count() > limits.max_graph_parts {
        return Err(limit(
            "Data Model graph parts",
            package.part_count(),
            limits.max_graph_parts,
        ));
    }
    let opc_limits = litchi_opc::ReadLimits::default();
    let mut relationship_count = 0_usize;
    let mut metadata_bytes = 0_usize;
    preflight_relationships(
        package.rels(),
        &opc_limits,
        &mut relationship_count,
        &mut metadata_bytes,
        limits,
    )?;
    for part in package.iter_parts() {
        // Only the model and connections payloads are measured, so only those
        // parts are decoded; every other part stays metadata-only (ADR 0030).
        if part.content_type() == DATA_MODEL_CONTENT_TYPE {
            let actual = package.get_part(part.partname())?.blob().len();
            if actual > limits.max_part_bytes {
                return Err(limit(
                    "Data Model payload bytes",
                    actual,
                    limits.max_part_bytes,
                ));
            }
        }
        if part.content_type() == crate::package::connections::package::CONNECTIONS_CONTENT_TYPE {
            let actual = package.get_part(part.partname())?.blob().len();
            if actual > limits.max_connection_bytes {
                return Err(limit(
                    "Data Model External Data Connections bytes",
                    actual,
                    limits.max_connection_bytes,
                ));
            }
        }
        preflight_xml_attribute(
            "PartName",
            part.partname().as_str(),
            opc_limits.max_xml_attribute_bytes(),
        )?;
        preflight_xml_attribute(
            "ContentType",
            part.content_type(),
            opc_limits.max_xml_attribute_bytes(),
        )?;
        metadata_bytes = checked_metadata_add(
            metadata_bytes,
            metadata_cost([
                part.partname().as_str().len(),
                part.partname().as_str().len(),
                part.content_type().len(),
            ])?,
            limits,
            "Data Model content-type metadata",
        )?;
        preflight_relationships(
            part.rels(),
            &opc_limits,
            &mut relationship_count,
            &mut metadata_bytes,
            limits,
        )?;
    }
    Ok(())
}

pub(crate) fn opc_capture_limits(limits: super::ReadLimits) -> Result<litchi_opc::ReadLimits> {
    if limits.max_part_bytes == 0 || limits.max_metadata_bytes == 0 {
        return Err(limit(
            "Data Model source capture bytes",
            0,
            limits.max_part_bytes.max(limits.max_metadata_bytes),
        ));
    }
    if limits.max_graph_parts == 0 {
        return Err(limit("Data Model graph parts", 1, 0));
    }
    if limits.max_graph_relationships == 0 {
        return Err(limit("Data Model graph relationships", 1, 0));
    }
    // OPC defaults are ordinary safety policy, not hard ceilings. The Data
    // Model limits have already been validated against this feature's finite
    // implementation maxima, so pass them through instead of silently
    // narrowing an explicit caller budget at source-token capture.
    let parts = limits.max_graph_parts;
    let relationships = limits.max_graph_relationships;
    // A package can carry one relationship member alongside each physical
    // part, plus the package content-types and root relationship members.
    // Keep the OPC builder's archive-member invariant finite without treating
    // its default archive-member policy as a Data Model graph ceiling.
    let archive_members = parts
        .checked_mul(2)
        .and_then(|value| value.checked_add(2))
        .ok_or(Error::CapacityOverflow {
            resource: "Data Model OPC archive members",
        })?;
    let content_mappings = parts.max(limits.max_connections);
    let total_relationship_xml = limits.max_metadata_bytes.max(limits.max_part_bytes);
    litchi_opc::ReadLimits::builder()
        .max_archive_members(archive_members)?
        .max_archive_total_entries(archive_members)?
        .max_parts(parts)?
        .max_part_bytes(limits.max_part_bytes as u64)?
        .max_total_part_bytes(limits.max_part_bytes as u64)?
        .max_content_types_bytes(limits.max_part_bytes)?
        .max_content_type_mappings(content_mappings)?
        .max_relationship_parts(parts)?
        .max_relationship_xml_bytes(limits.max_part_bytes)?
        .max_total_relationship_xml_bytes(total_relationship_xml)?
        .max_relationships_per_part(relationships)?
        .max_total_relationships(relationships)?
        .max_relationship_graph_nodes(parts)?
        .build()
        .map_err(Error::Opc)
}

fn preflight_relationships(
    relationships: &litchi_opc::Relationships,
    opc_limits: &litchi_opc::ReadLimits,
    relationship_count: &mut usize,
    metadata_bytes: &mut usize,
    limits: super::ReadLimits,
) -> Result<()> {
    *metadata_bytes = checked_metadata_add(
        *metadata_bytes,
        metadata_cost(
            std::iter::once(relationships.base_uri().len())
                .chain(relationships.source_uri().into_iter().map(str::len)),
        )?,
        limits,
        "Data Model relationship owner metadata",
    )?;
    for relationship in relationships.iter() {
        *relationship_count = relationship_count
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "Data Model graph relationships",
            })?;
        if *relationship_count > limits.max_graph_relationships {
            return Err(limit(
                "Data Model graph relationships",
                *relationship_count,
                limits.max_graph_relationships,
            ));
        }
        for (key, value) in [
            ("Id", relationship.r_id()),
            ("Type", relationship.reltype()),
            ("Target", relationship.target_ref()),
        ] {
            preflight_xml_attribute(key, value, opc_limits.max_xml_attribute_bytes())?;
        }
        if relationship.target_mode() == TargetMode::Internal {
            let resolved = relationship
                .base_uri()
                .len()
                .checked_add(relationship.target_ref().len())
                .ok_or(Error::CapacityOverflow {
                    resource: "Data Model relationship target metadata",
                })?;
            if resolved > opc_limits.max_xml_attribute_bytes() {
                return Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
                    resource: litchi_opc::ReadResource::XmlAttributeBytes,
                    actual: resolved as u64,
                    maximum: opc_limits.max_xml_attribute_bytes() as u64,
                }));
            }
        }
        *metadata_bytes = checked_metadata_add(
            *metadata_bytes,
            metadata_cost(
                std::iter::once(relationship.r_id().len())
                    .chain(std::iter::once(relationship.r_id().len()))
                    .chain(std::iter::once(relationship.reltype().len()))
                    .chain(std::iter::once(relationship.target_ref().len()))
                    .chain(std::iter::once(relationship.base_uri().len()))
                    .chain(relationship.source_uri().into_iter().map(str::len)),
            )?,
            limits,
            "Data Model relationship metadata",
        )?;
    }
    Ok(())
}

fn metadata_cost(lengths: impl IntoIterator<Item = usize>) -> Result<usize> {
    lengths
        .into_iter()
        .try_fold(0_usize, |total, length| total.checked_add(length))
        .and_then(|length| length.checked_mul(6))
        .and_then(|length| length.checked_add(256))
        .ok_or(Error::CapacityOverflow {
            resource: "Data Model package metadata",
        })
}

fn checked_metadata_add(
    current: usize,
    addition: usize,
    limits: super::ReadLimits,
    resource: &'static str,
) -> Result<usize> {
    let total = current
        .checked_add(addition)
        .ok_or(Error::CapacityOverflow { resource })?;
    if total > limits.max_metadata_bytes {
        return Err(limit(resource, total, limits.max_metadata_bytes));
    }
    Ok(total)
}

fn preflight_xml_attribute(key: &str, value: &str, maximum: usize) -> Result<()> {
    let encoded = value.chars().try_fold(0_usize, |length, character| {
        let bytes = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length.checked_add(bytes).ok_or(Error::CapacityOverflow {
            resource: "Data Model XML attribute metadata",
        })
    })?;
    let actual = key
        .len()
        .checked_add(encoded)
        .ok_or(Error::CapacityOverflow {
            resource: "Data Model XML attribute metadata",
        })?;
    if actual > maximum {
        return Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::XmlAttributeBytes,
            actual: actual as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        resource,
        actual,
        maximum,
    }
}

/// Discover the relationship-free model part without interpreting MS-XLDM.
///
/// This function deliberately performs only the bounded OPC envelope checks
/// that are specified locally. DAX, credentials, refresh commands, and other
/// model payload semantics stay opaque and inert.
pub(crate) fn inspect_model_part(package: &OpcPackage) -> Result<Option<ModelPart>> {
    let fixed_uri = model_part_uri()?;
    let mut found = None;
    for part in package.iter_parts() {
        if part.partname().is_equivalent_to(&fixed_uri)
            && part.content_type() != DATA_MODEL_CONTENT_TYPE
        {
            return Err(invalid(format!(
                "fixed Data Model part '{}' has content type '{}' instead of '{DATA_MODEL_CONTENT_TYPE}'",
                part.partname(),
                part.content_type()
            )));
        }
        if part.content_type() != DATA_MODEL_CONTENT_TYPE {
            continue;
        }
        if found.is_some() {
            return Err(invalid("package contains multiple Data Model parts"));
        }
        if !part.partname().is_equivalent_to(&fixed_uri) {
            return Err(invalid(format!(
                "Data Model part '{}' must be '{DATA_MODEL_PART_NAME}'",
                part.partname()
            )));
        }
        // The model payload is the only one this pass reads, so it is the
        // only part decoded here (ADR 0030).
        let payload = package.get_part(part.partname())?;
        if payload.blob().is_empty() {
            return Err(invalid("Data Model payload cannot be empty"));
        }
        if payload.blob().len() > MAX_PAYLOAD_BYTES {
            return Err(invalid("Data Model payload exceeds 512 MiB"));
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "Data Model part has forbidden outbound relationships",
            ));
        }
        reject_inbound_relationships(package, part.partname())?;
        found = Some(ModelPart {
            part_name: part.partname().to_string(),
            content_type: part.content_type().to_string(),
            bytes: payload.blob_arc(),
        });
    }
    Ok(found)
}

/// Validate a caller-provided opaque payload before staging it in a draft.
pub(crate) fn validate_payload(part: &ModelPart) -> Result<()> {
    if !part.part_name.eq_ignore_ascii_case(DATA_MODEL_PART_NAME) {
        return Err(invalid(format!(
            "Data Model part must be '{DATA_MODEL_PART_NAME}'"
        )));
    }
    if part.content_type != DATA_MODEL_CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: DATA_MODEL_CONTENT_TYPE.to_string(),
            got: part.content_type.clone(),
        });
    }
    if part.bytes.is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if part.bytes.len() > MAX_PAYLOAD_BYTES {
        return Err(invalid("Data Model payload exceeds 512 MiB"));
    }
    let _ = model_part_uri()?;
    Ok(())
}

/// Validate the opaque owner before a writer retains a model containing
/// workbook time-grouping records.  The workbook stream and the model part
/// are separate owners; validating only the former lets a caller publish
/// stale groupings after replacing the opaque bytes.
pub(crate) fn validate_model_payload_and_groupings(model: &Model) -> Result<()> {
    super::codec::validate_definition(&model.definition, super::ReadLimits::DEFAULT)?;
    validate_payload(&model.part)?;
    for grouping in &model.definition.time_groupings {
        super::proof::prove_time_grouping(model.part.bytes(), grouping)?;
    }
    Ok(())
}

/// Validate every owner at the final XLSB publication boundary.
pub(crate) fn validate_model_for_write(
    model: &Model,
    connections: Option<&Connections>,
) -> Result<()> {
    validate_model_payload_and_groupings(model)?;
    validate_definition_connections(&model.definition, connections)
}

/// Validate the workbook-owned external connection closure for Data Model
/// table records.
///
/// `BrtModelTable.connection` names the workbook connection that owns the
/// table. The model part is opaque to this crate, so a descriptor mutation is
/// publishable only when the typed workbook owner proves that every table name
/// resolves to exactly one connection. The connections package parser
/// already rejects duplicate names; the explicit count here keeps this
/// boundary safe if that invariant changes later.
pub(crate) fn validate_definition_connections(
    definition: &Definition,
    connections: Option<&Connections>,
) -> Result<()> {
    validate_definition_connection_names(
        definition,
        connections.map(|value| {
            value
                .connections
                .iter()
                .map(|connection| connection.name.as_str())
        }),
    )
}

pub(crate) fn validate_definition_connection_names<'a, I>(
    definition: &Definition,
    connections: Option<I>,
) -> Result<()>
where
    I: IntoIterator<Item = &'a str>,
{
    if definition.tables.is_empty() {
        return Ok(());
    }
    let Some(connections) = connections else {
        return Err(invalid(
            "Data Model tables require the workbook External Data Connections part",
        ));
    };
    let mut names = HashMap::new();
    for name in connections {
        let folded = name.to_lowercase();
        let count = names.entry(folded).or_insert(0_usize);
        *count += 1;
    }
    for table in &definition.tables {
        let wanted = table.connection.to_lowercase();
        match names.get(&wanted).copied().unwrap_or(0) {
            0 => {
                return Err(invalid(format!(
                    "Data Model table '{}' references unknown workbook connection '{}'",
                    table.name, table.connection
                )));
            },
            1 => {},
            _ => {
                return Err(invalid(format!(
                    "Data Model table '{}' references multiple workbook connections named '{}'",
                    table.name, table.connection
                )));
            },
        }
    }
    Ok(())
}

pub(crate) fn model_part_uri() -> Result<PackURI> {
    PackURI::new(DATA_MODEL_PART_NAME).map_err(|error| Error::InvalidUri(error.to_string()))
}

pub(crate) fn reject_inbound_relationships(package: &OpcPackage, target: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship_targets(relationship, target)? {
            return Err(invalid(
                "package relationship targets the relationship-free Data Model part",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship_targets(relationship, target)? {
                return Err(invalid(format!(
                    "part '{}' has a relationship to the relationship-free Data Model part",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

fn relationship_targets(relationship: &litchi_opc::Relationship, target: &PackURI) -> Result<bool> {
    if relationship.target_mode() == TargetMode::External || relationship.is_external() {
        return Ok(false);
    }
    Ok(relationship.target_partname()?.is_equivalent_to(target))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
#[allow(clippy::expect_used)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, PackURI, Part};

    #[test]
    fn model_part_uri_and_content_type_are_fixed() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).expect("URI"),
            DATA_MODEL_CONTENT_TYPE.to_string(),
            vec![1, 2, 3],
        )));
        let part = inspect_model_part(&package)
            .expect("inspect")
            .expect("model");
        assert_eq!(part.part_name(), DATA_MODEL_PART_NAME);
        assert_eq!(part.content_type(), DATA_MODEL_CONTENT_TYPE);
        assert_eq!(part.bytes(), &[1, 2, 3]);
    }

    #[test]
    fn model_part_is_relationship_free() {
        let mut package = OpcPackage::new();
        let mut part = BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).expect("URI"),
            DATA_MODEL_CONTENT_TYPE.to_string(),
            vec![1],
        );
        part.relate_to(
            "/xl/workbook.bin",
            "http://example.invalid/model-dependency",
        );
        package.add_part(Box::new(part));
        assert!(inspect_model_part(&package).is_err());
    }

    #[test]
    fn fixed_model_path_cannot_be_retyped_as_another_part() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).expect("URI"),
            "application/octet-stream".to_string(),
            vec![1],
        )));
        assert!(inspect_model_part(&package).is_err());
    }
}
