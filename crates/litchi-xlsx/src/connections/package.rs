//! Workbook package integration for the SpreadsheetML connections owner.

use super::model::*;
use super::{codec, invalid};
use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use std::collections::HashSet;

pub fn store_in_package(package: &mut OpcPackage, value: &Connections, strict: bool) -> Result<()> {
    store_in_package_with_query_table_validator(package, value, strict, query_table_connection_id)
}

/// Store connections while allowing the migration host to retain its complete
/// query-table parser for cross-part validation.
#[doc(hidden)]
pub fn store_in_package_with_query_table_validator<F>(
    package: &mut OpcPackage,
    value: &Connections,
    strict: bool,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let xml = value.to_xml(strict)?;
    validate_query_table_connection_ids(package, value, query_table_connection_id)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = {
        let workbook = package.get_part(&workbook_name)?;
        let mut found = workbook.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        });
        let first = found
            .next()
            .map(|relationship| {
                if relationship.is_external() {
                    return Err(invalid("connections relationship cannot be external"));
                }
                Ok((
                    relationship.r_id().to_string(),
                    relationship.target_partname()?,
                ))
            })
            .transpose()?;
        if found.next().is_some() {
            return Err(invalid("workbook has multiple connections relationships"));
        }
        first
    };
    if let Some((_, part_name)) = existing {
        let part = package.get_part(&part_name)?;
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(invalid(
                "existing connections part has invalid content type",
            ));
        }
        package.get_part_mut(&part_name)?.set_blob(xml);
    } else {
        let part_name = next_connections_part_name(package)?;
        let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONNECTIONS_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                if strict {
                    STRICT_CONNECTIONS_RELATIONSHIP
                } else {
                    CONNECTIONS_RELATIONSHIP
                }
                .into(),
                part_name.relative_ref(workbook_name.base_uri()),
                relationship_id,
                false,
            );
    }
    package.unsign();
    Ok(())
}

pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    let workbook_name = package.main_document_part()?.partname().clone();
    let relationship = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
            )
        })
        .map(|relationship| {
            relationship
                .target_partname()
                .map(|part_name| (relationship.r_id().to_string(), part_name))
        })
        .transpose()?;
    let Some((relationship_id, part_name)) = relationship else {
        return Ok(false);
    };
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .remove(&relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

fn validate_query_table_connection_ids<F>(
    package: &OpcPackage,
    value: &Connections,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let ids = value
        .connections
        .iter()
        .map(|connection| connection.id)
        .collect::<HashSet<_>>();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        let connection_id = query_table_connection_id(part.blob())?;
        if !ids.contains(&connection_id) {
            return Err(invalid(format!(
                "query-table part '{}' references missing connection ID {}",
                part.partname(),
                connection_id
            )));
        }
    }
    Ok(())
}

fn query_table_connection_id(xml: &[u8]) -> Result<u32> {
    if xml.len() > 8 * 1024 * 1024 {
        return Err(invalid("query-table part exceeds 8 MiB"));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    if processed.len() > 8 * 1024 * 1024 {
        return Err(invalid("processed query-table part exceeds 8 MiB"));
    }
    let root = codec::parse_dom(processed.as_ref())?;
    codec::expect(&root, "queryTable")?;
    let _name = codec::req(&root, "name")?;
    let connection_id = codec::u32req(&root, "connectionId")?;
    codec::only_unqualified(
        &root,
        &[
            "name",
            "headers",
            "rowNumbers",
            "disableRefresh",
            "backgroundRefresh",
            "firstBackgroundRefresh",
            "refreshOnLoad",
            "growShrinkType",
            "fillFormulas",
            "removeDataOnSave",
            "disableEdit",
            "preserveFormatting",
            "adjustColumnWidth",
            "intermediate",
            "connectionId",
            "autoFormatId",
            "applyNumberFormats",
            "applyBorderFormats",
            "applyFontFormats",
            "applyPatternFormats",
            "applyAlignmentFormats",
            "applyWidthHeightFormats",
        ],
    )?;
    codec::kids(&root)?;
    Ok(connection_id)
}

fn next_connections_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/connections.xml".into()
        } else {
            format!("/xl/connections{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections part name"))
}

fn next_connections_relationship_id(package: &OpcPackage, workbook: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdConnections{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}
pub fn load_from_package(package: &OpcPackage) -> Result<Option<Connections>> {
    let workbook = package.main_document_part()?;
    let mut found = workbook.rels().iter().filter(|x| {
        matches!(
            x.reltype(),
            CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
        )
    });
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid("workbook has multiple connections relationships"));
    }
    if rel.is_external() {
        return Err(invalid("connections relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "connections part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("connections part must not have relationships"));
    }
    Ok(Some(Connections::parse(part.blob())?))
}
