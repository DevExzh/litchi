//! Worksheet relationship and package operations for query-table parts.

use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use std::collections::HashSet;

use crate::error::Result;

use super::codec::parse_query_table;
use super::model::{Conformance, Part, Table, invalid};

pub const QUERY_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
pub const QUERY_TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
pub const STRICT_QUERY_TABLE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable";

pub fn is_query_table_relationship_type(value: &str) -> bool {
    matches!(
        value,
        QUERY_TABLE_RELATIONSHIP_TYPE | STRICT_QUERY_TABLE_RELATIONSHIP_TYPE
    )
}

/// Load and validate every inert query-table part owned by a worksheet.
pub fn load_worksheet_query_tables(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Vec<Part>> {
    let worksheet = package.get_part(worksheet_part)?;
    let mut result = Vec::new();
    let mut seen_parts = HashSet::new();
    for relationship in worksheet
        .rels()
        .iter()
        .filter(|relationship| is_query_table_relationship_type(relationship.reltype()))
    {
        if relationship.is_external() {
            return Err(invalid("query-table relationship cannot be external"));
        }
        let part_name = relationship.target_partname()?;
        if !seen_parts.insert(part_name.clone()) {
            return Err(invalid("worksheet has duplicate query-table targets"));
        }
        let part = package.get_part(&part_name)?;
        if part.content_type() != QUERY_TABLE_CONTENT_TYPE {
            return Err(invalid(format!(
                "query-table part '{}' has invalid content type",
                part_name
            )));
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid("query-table parts must not have relationships"));
        }
        result.push(Part::new(
            relationship.r_id().to_string(),
            part_name.to_string(),
            parse_query_table(part.blob())?,
        ));
    }
    result.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(result)
}

pub fn find_worksheet_query_table(
    package: &OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<Option<Part>> {
    Ok(load_worksheet_query_tables(package, worksheet_part)?
        .into_iter()
        .find(|item| item.relationship_id == relationship_id))
}

/// Add an inert query-table part. The referenced connection must already exist.
pub fn add_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    query_table: Table,
    conformance: Conformance,
) -> Result<Part> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let xml = query_table.to_xml(conformance)?;
    parse_query_table(&xml)?;
    let part_name = next_query_table_part_name(package)?;
    let relationship_id = next_query_table_relationship_id(package, worksheet_part)?;
    let target = part_name.relative_ref(worksheet_part.base_uri());
    let relationship_type = match conformance {
        Conformance::Transitional => QUERY_TABLE_RELATIONSHIP_TYPE,
        Conformance::Strict => STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    };
    package.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        QUERY_TABLE_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .add_relationship(
            relationship_type.into(),
            target,
            relationship_id.clone(),
            false,
        );
    package.unsign();
    Ok(Part::new(
        relationship_id,
        part_name.to_string(),
        query_table,
    ))
}

pub fn replace_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
    query_table: Table,
    conformance: Conformance,
) -> Result<()> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let existing = find_worksheet_query_table(package, worksheet_part, relationship_id)?
        .ok_or_else(|| invalid("query-table relationship was not found"))?;
    let xml = query_table.to_xml(conformance)?;
    parse_query_table(&xml)?;
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    package.add_part(Box::new(BlobPart::new(
        part_name,
        QUERY_TABLE_CONTENT_TYPE.into(),
        xml,
    )));
    package.unsign();
    Ok(())
}

pub fn update_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
    query_table: Table,
    conformance: Conformance,
) -> Result<()> {
    replace_worksheet_query_table(
        package,
        worksheet_part,
        relationship_id,
        query_table,
        conformance,
    )
}

pub fn remove_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<bool> {
    let Some(existing) = find_worksheet_query_table(package, worksheet_part, relationship_id)?
    else {
        return Ok(false);
    };
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

/// Reassign deterministic query-table relationship IDs in caller-specified order.
pub fn reorder_worksheet_query_tables(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    ordered_relationship_ids: &[String],
) -> Result<Vec<Part>> {
    let existing = load_worksheet_query_tables(package, worksheet_part)?;
    if existing.len() != ordered_relationship_ids.len() {
        return Err(invalid(
            "query-table reorder must contain every relationship",
        ));
    }
    let existing_ids = existing
        .iter()
        .map(|item| item.relationship_id.clone())
        .collect::<HashSet<_>>();
    let ordered_ids = ordered_relationship_ids
        .iter()
        .cloned()
        .collect::<HashSet<_>>();
    if existing_ids != ordered_ids || ordered_ids.len() != ordered_relationship_ids.len() {
        return Err(invalid("query-table reorder is not a permutation"));
    }
    let mut ordered = Vec::with_capacity(existing.len());
    for id in ordered_relationship_ids {
        ordered.push(
            existing
                .iter()
                .find(|item| &item.relationship_id == id)
                .expect("permutation was validated")
                .clone(),
        );
    }
    let relationship_type = package
        .get_part(worksheet_part)?
        .rels()
        .iter()
        .find(|relationship| is_query_table_relationship_type(relationship.reltype()))
        .map(|relationship| relationship.reltype().to_string())
        .unwrap_or_else(|| QUERY_TABLE_RELATIONSHIP_TYPE.into());
    let worksheet = package.get_part_mut(worksheet_part)?;
    for item in &existing {
        worksheet.rels_mut().remove(&item.relationship_id);
    }
    let mut result = Vec::with_capacity(ordered.len());
    for (offset, item) in ordered.into_iter().enumerate() {
        let id = format!("rIdQueryTable{}", offset + 1);
        let part_name = PackURI::new(item.part_name()).map_err(invalid)?;
        worksheet.rels_mut().add_relationship(
            relationship_type.clone(),
            part_name.relative_ref(worksheet_part.base_uri()),
            id.clone(),
            false,
        );
        result.push(Part::new(id, item.part_name, item.query_table));
    }
    package.unsign();
    Ok(result)
}

fn validate_query_table_connection(package: &OpcPackage, connection_id: u32) -> Result<()> {
    let connections = crate::connections::load_from_package(package)
        .map_err(|_| invalid("connections graph is invalid"))?
        .ok_or_else(|| invalid("query table requires a connections part"))?;
    if connections
        .connections
        .iter()
        .any(|connection| connection.id == connection_id)
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "query table references missing connection ID {connection_id}"
        )))
    }
}

fn next_query_table_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate =
            PackURI::new(format!("/xl/queryTables/queryTable{suffix}.xml")).map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free query-table part name"))
}

fn next_query_table_relationship_id(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(worksheet_part)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdQueryTable{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free query-table relationship ID"))
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
