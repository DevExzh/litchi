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

#[must_use]
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
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    let worksheet = package.get_part(worksheet_part)?;
    let mut owners = vec![worksheet.partname().clone()];
    let mut seen_owners = HashSet::new();
    for relationship in worksheet
        .rels()
        .iter()
        .filter(|r| matches!(r.reltype(), rt::TABLE | rt::STRICT_TABLE))
    {
        if relationship.is_external() || relationship.target_ref().contains(['#', '?']) {
            return Err(invalid(
                "worksheet table relationship must target an internal part",
            ));
        }
        let target = relationship.target_partname()?;
        let table = package.get_part(&target)?;
        if table.content_type() != ct::SML_TABLE {
            return Err(invalid(
                "worksheet table relationship has invalid content type",
            ));
        }
        if !seen_owners.insert(table.partname().as_str().to_ascii_lowercase()) {
            return Err(invalid("worksheet has duplicate table targets"));
        }
        owners.push(table.partname().clone());
    }
    let mut result = Vec::new();
    let mut seen_parts = HashSet::new();
    for owner in owners {
        let part_owner = package.get_part(&owner)?;
        let mut count = 0;
        for relationship in part_owner
            .rels()
            .iter()
            .filter(|r| is_query_table_relationship_type(r.reltype()))
        {
            count += 1;
            if !owner.is_equivalent_to(worksheet_part) && count > 1 {
                return Err(invalid(
                    "table definition has multiple query-table relationships",
                ));
            }
            if relationship.is_external() || relationship.target_ref().contains(['#', '?']) {
                return Err(invalid(
                    "query-table relationship must target an internal part",
                ));
            }
            let part_name = relationship.target_partname()?;
            let part = package.get_part(&part_name)?;
            if !seen_parts.insert(part.partname().as_str().to_ascii_lowercase()) {
                return Err(invalid("worksheet has duplicate query-table targets"));
            }
            if part.content_type() != QUERY_TABLE_CONTENT_TYPE {
                return Err(invalid(format!(
                    "query-table part '{part_name}' has invalid content type"
                )));
            }
            if part.rels().iter().next().is_some() {
                return Err(invalid("query-table parts must not have relationships"));
            }
            result.push(
                Part::new(
                    relationship.r_id().to_owned(),
                    part.partname().to_string(),
                    parse_query_table(part.blob())?,
                )
                .with_owner(&owner),
            );
        }
    }
    result.sort_by(|a, b| {
        a.owner_part_name
            .cmp(&b.owner_part_name)
            .then_with(|| a.relationship_id.cmp(&b.relationship_id))
    });
    Ok(result)
}

pub fn find_worksheet_query_table(
    package: &OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<Option<Part>> {
    let mut matching = load_worksheet_query_tables(package, worksheet_part)?
        .into_iter()
        .filter(|p| p.relationship_id == relationship_id);
    let found = matching.next();
    if matching.next().is_some() {
        return Err(invalid(
            "query-table relationship ID is ambiguous across owner parts",
        ));
    }
    Ok(found)
}

/// Find a query table by its package part identity, including table-owned entries.
pub(crate) fn find_by_part_name(
    package: &OpcPackage,
    worksheet_part: &PackURI,
    part_name: &str,
) -> Result<Option<Part>> {
    let target = PackURI::new(part_name).map_err(invalid)?;
    Ok(load_worksheet_query_tables(package, worksheet_part)?
        .into_iter()
        .find(|p| PackURI::new(p.part_name()).is_ok_and(|name| name.is_equivalent_to(&target))))
}

pub(crate) fn replace_by_part_name(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    part_name: &str,
    query_table: Table,
    conformance: Conformance,
) -> Result<()> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let existing = find_by_part_name(package, worksheet_part, part_name)?
        .ok_or_else(|| invalid("query-table part was not found in the selected worksheet"))?;
    replace_definition(package, &existing, query_table, conformance)
}

fn replace_definition(
    package: &mut OpcPackage,
    existing: &Part,
    query_table: Table,
    conformance: Conformance,
) -> Result<()> {
    if let Some(owner) = existing.owner_part_name() {
        let owner = PackURI::new(owner).map_err(invalid)?;
        if package.get_part(&owner)?.content_type()
            == litchi_opc::constants::content_type::SML_TABLE
        {
            let bindings = |table: &Table| {
                table
                    .refresh()
                    .into_iter()
                    .flat_map(|refresh| refresh.fields())
                    .map(|field| (field.id(), field.table_column_id()))
                    .collect::<Vec<_>>()
            };
            if bindings(existing.query_table()) != bindings(&query_table) {
                return Err(crate::Error::Unsupported {
                    feature: "table-owned query field rebinding requires table-column updates",
                });
            }
        }
    }
    let canonical = query_table.to_xml(conformance)?;
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    let xml = super::codec::patch_name(
        package.get_part(&part_name)?.blob(),
        &query_table,
        conformance,
    )?
    .unwrap_or(canonical);
    parse_query_table(&xml)?;
    crate::signature_policy::ensure_unsigned(package)?;
    package.add_part(Box::new(BlobPart::new(
        part_name,
        QUERY_TABLE_CONTENT_TYPE.into(),
        xml,
    )));
    Ok(())
}

/// Add an inert query-table part. The referenced connection must already exist.
///
/// Adding changes the package, so a signed package is refused with
/// [`crate::Error::Signed`] before anything is changed: remove the signature
/// explicitly first ([`OpcPackage::unsign`]).
pub(crate) fn add_worksheet_query_table(
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
    crate::signature_policy::ensure_unsigned(package)?;
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
    Ok(Part::new(relationship_id, part_name.to_string(), query_table).with_owner(worksheet_part))
}

/// Replace the definition of one existing query-table part.
///
/// Replacing always writes the part, so a signed package is refused with
/// [`crate::Error::Signed`] before anything is changed, even when the
/// serialized definition equals the stored one: remove the signature
/// explicitly first ([`OpcPackage::unsign`]).
pub(crate) fn replace_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
    query_table: Table,
    conformance: Conformance,
) -> Result<()> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let existing = find_worksheet_query_table(package, worksheet_part, relationship_id)?
        .ok_or_else(|| invalid("query-table relationship was not found"))?;
    replace_definition(package, &existing, query_table, conformance)
}

/// Remove one query-table relationship and its unreferenced part.
///
/// An absent relationship leaves the package untouched and returns `false`,
/// signed or not. Removing a query table changes the package, so a signed
/// package is refused with [`crate::Error::Signed`]: remove the signature
/// explicitly first ([`OpcPackage::unsign`]).
pub(crate) fn remove_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<bool> {
    let Some(existing) = find_worksheet_query_table(package, worksheet_part, relationship_id)?
    else {
        return Ok(false);
    };
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    if !existing
        .owner_part_name()
        .is_some_and(|name| name.eq_ignore_ascii_case(worksheet_part.as_str()))
    {
        return Err(crate::Error::Unsupported {
            feature: "table-owned query-table removal requires table conversion",
        });
    }
    crate::signature_policy::ensure_unsigned(package)?;
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    Ok(true)
}

/// Reassign deterministic query-table relationship IDs in caller-specified order.
///
/// When every relationship already carries its deterministic ID, type and
/// target in the requested order (always the case for a worksheet without
/// query tables), nothing is published and a signed package keeps its
/// signature byte for byte. Any other reorder changes the package, so a signed
/// package is refused with [`crate::Error::Signed`]: remove the signature
/// explicitly first ([`OpcPackage::unsign`]).
pub(crate) fn reorder_worksheet_query_tables(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    ordered_relationship_ids: &[String],
) -> Result<Vec<Part>> {
    let existing = load_worksheet_query_tables(package, worksheet_part)?;
    if existing.iter().any(|p| {
        !p.owner_part_name()
            .is_some_and(|name| name.eq_ignore_ascii_case(worksheet_part.as_str()))
    }) {
        return Err(crate::Error::Unsupported {
            feature: "query-table reorder across table-definition owners",
        });
    }
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
                .unwrap_or_else(|| {
                    crate::error::panic_missing_invariant("permutation was validated")
                })
                .clone(),
        );
    }
    let relationship_type = package
        .get_part(worksheet_part)?
        .rels()
        .iter()
        .find(|relationship| is_query_table_relationship_type(relationship.reltype()))
        .map_or_else(
            || QUERY_TABLE_RELATIONSHIP_TYPE.into(),
            |relationship| relationship.reltype().to_string(),
        );
    if relationships_already_ordered(package, worksheet_part, &ordered, &relationship_type)? {
        return Ok(ordered);
    }
    crate::signature_policy::ensure_unsigned(package)?;
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
        result.push(Part::new(id, item.part_name, item.query_table).with_owner(worksheet_part));
    }
    Ok(result)
}

/// Whether the worksheet's query-table relationships already are exactly the
/// ones [`reorder_worksheet_query_tables`] would write for `ordered`: the
/// deterministic ID at each position, the selected relationship type and the
/// same relative target. Such a reorder publishes nothing.
fn relationships_already_ordered(
    package: &OpcPackage,
    worksheet_part: &PackURI,
    ordered: &[Part],
    relationship_type: &str,
) -> Result<bool> {
    let worksheet = package.get_part(worksheet_part)?;
    for (offset, item) in ordered.iter().enumerate() {
        if item.relationship_id != format!("rIdQueryTable{}", offset + 1) {
            return Ok(false);
        }
        let target = PackURI::new(item.part_name())
            .map_err(invalid)?
            .relative_ref(worksheet_part.base_uri());
        let Some(relationship) = worksheet.rels().get(&item.relationship_id) else {
            return Ok(false);
        };
        if relationship.reltype() != relationship_type
            || relationship.is_external()
            || relationship.target_ref() != target
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn validate_query_table_connection(package: &OpcPackage, connection_id: u32) -> Result<()> {
    let limits = package.read_limits();
    let connections = crate::connections::load_from_package(package, &limits)
        .map_err(|_source| invalid("connections graph is invalid"))?
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
        match package.get_part(&candidate) {
            Ok(_) => {},
            Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(candidate),
            Err(error) => return Err(error.into()),
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
