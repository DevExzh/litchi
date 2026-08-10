#![allow(
    clippy::expect_used,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! Bounded OPC traversal for XLSB Custom XML Maps.

use std::collections::HashSet;
use std::ops::Range;

use litchi_opc::constants::relationship_type;
use litchi_opc::{OpcPackage, PackURI, Part};

use super::codec::{parse_single_cells_value_with_connection_ids, parse_table_bindings_value};
use super::{
    MappedTable, SingleCellBinding, XmlMapConformance, XmlMapInfo, validate_binding_map_ids,
    validate_catalog,
};
use crate::package::error::{Error, Result};

pub(crate) const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";
const TABLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.table";
const SINGLE_CELLS_CONTENT_TYPE: &str = "application/vnd.ms-excel.tableSingleCells";
const SINGLE_CELLS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";
const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

#[derive(Debug)]
pub(crate) struct Loaded {
    pub(crate) map_info: Option<XmlMapInfo>,
    pub(crate) conformance: XmlMapConformance,
    pub(crate) mapped_tables: Vec<MappedTable>,
    pub(crate) single_cell_tables: Vec<SingleCellBinding>,
    pub(crate) single_cell_ranges: Vec<Option<Range<usize>>>,
    pub(crate) worksheet_parts: Vec<PackURI>,
    pub(crate) dependency_parts: Vec<PackURI>,
}

pub(crate) fn read_with_limits(
    package: &OpcPackage,
    worksheet_parts: Vec<PackURI>,
    limits: super::snapshot::ReadLimits,
) -> Result<Loaded> {
    preflight(package, limits)?;
    let workbook = package.main_document_part()?;
    let workbook_name = workbook.partname().clone();
    let connections =
        crate::package::connections::package::load_from_workbook(package, &workbook_name)?;
    let map_relationship = xml_maps_relationship(package, &workbook_name)?;
    let (map_info, conformance, map_part) = if let Some((part_name, conformance)) = map_relationship
    {
        let part = package.get_part(&part_name)?;
        require_content_type(
            part,
            litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE,
        )?;
        require_no_relationships(part, "Custom XML Maps part")?;
        if part.blob().len() > limits.xml_maps.max_part_bytes {
            return Err(invalid(
                "Custom XML Maps part exceeds its configured byte limit",
            ));
        }
        let parsed = litchi_ooxml_common::spreadsheet_xml_maps::parse_xml_map_info_with_conformance_and_limits(
            part.blob(),
            &limits.xml_maps,
        )
        .map_err(|error| invalid(error.to_string()))?;
        if parsed.conformance != conformance {
            return Err(invalid(
                "Custom XML Maps root namespace disagrees with its relationship conformance",
            ));
        }
        validate_catalog(&parsed.info, limits.xml_maps)?;
        validate_connection_references(connections.as_ref(), &parsed.info)?;
        (Some(parsed.info), conformance, Some(part_name))
    } else {
        (None, XmlMapConformance::Transitional, None)
    };

    let mut mapped_tables = Vec::new();
    let mut single_cell_tables = Vec::new();
    let mut single_cell_ranges = Vec::new();
    single_cell_ranges
        .try_reserve_exact(worksheet_parts.len())
        .map_err(|source| Error::Allocation {
            resource: "single-cell worksheet range index",
            source,
        })?;
    let mut dependencies = map_part.iter().cloned().collect::<Vec<_>>();
    let mut table_ids = HashSet::new();
    let mut total_bindings = 0usize;
    let mut total_xpath_units = 0usize;

    validate_single_cell_ownership(package, &worksheet_parts)?;

    for (sheet_index, worksheet_name) in worksheet_parts.iter().enumerate() {
        let worksheet = package.get_part(worksheet_name)?;
        require_content_type(worksheet, WORKSHEET_CONTENT_TYPE)?;
        let geometry = worksheet_geometry(worksheet.blob())?;
        let listed_tables = crate::package::table::parse_table_part_rel_ids(worksheet.blob())?;
        let listed = listed_tables.iter().cloned().collect::<HashSet<_>>();
        if listed.len() != listed_tables.len() {
            return Err(invalid(format!(
                "worksheet {sheet_index} repeats a BrtListPart relationship"
            )));
        }

        let mut sheet_singles = None;
        let mut seen_single = false;
        let mut ordinary_tables = Vec::new();
        for relationship in worksheet.rels().iter() {
            if matches!(
                relationship.reltype(),
                relationship_type::TABLE | relationship_type::STRICT_TABLE
            ) {
                if !listed.contains(relationship.r_id()) {
                    return Err(invalid(format!(
                        "worksheet {sheet_index} has an orphan table relationship {:?}",
                        relationship.r_id()
                    )));
                }
                if relationship.is_external() {
                    return Err(invalid("worksheet table relationship cannot be external"));
                }
                let part_name = relationship.target_partname()?;
                let part = package.get_part(&part_name)?;
                require_content_type(part, TABLE_CONTENT_TYPE)?;
                let typed = crate::package::table::parse_table_part(part.blob())?;
                validate_table_range(&typed.range)?;
                validate_table_height(&typed, sheet_index)?;
                validate_table_connection(&typed, connections.as_ref(), sheet_index)?;
                if !table_ids.insert(typed.id) {
                    return Err(invalid(format!(
                        "duplicate XML mapping table ID {}",
                        typed.id
                    )));
                }
                if geometry
                    .auto_filter
                    .is_some_and(|filter| ranges_overlap(typed.range, filter))
                {
                    return Err(invalid(format!(
                        "ordinary table {} on worksheet {sheet_index} overlaps the worksheet AutoFilter",
                        typed.id
                    )));
                }
                ordinary_tables.push((typed.id, typed.range));
                if let Some(dimension) = geometry.authoritative_dimension() {
                    require_range_within(typed.range, dimension, "ordinary table", sheet_index)?;
                }
                if matches!(typed.table_type, crate::package::table::Type::Xml) {
                    let table = parse_table_bindings_value(part.blob(), limits.bindings)?;
                    if typed.id != table.table_id() {
                        return Err(invalid(
                            "table XML binding list ID disagrees with BrtBeginList",
                        ));
                    }
                    for binding in table.columns() {
                        if !typed
                            .columns
                            .iter()
                            .any(|column| column.id == binding.column_id())
                        {
                            return Err(invalid(format!(
                                "table {} XML binding refers to missing table column {}",
                                typed.id,
                                binding.column_id()
                            )));
                        }
                    }
                    if !table.columns().is_empty() {
                        charge_bindings(
                            table
                                .columns()
                                .iter()
                                .map(|binding| binding.xpath().as_str()),
                            &mut total_bindings,
                            &mut total_xpath_units,
                            limits,
                        )?;
                        mapped_tables.push(table);
                    }
                }
                // Retain every ordinary table physically even when it has no
                // XML bindings. A transaction may remove the final binding
                // while preserving this exact table part.
                dependencies.push(part_name);
            } else if relationship.reltype() == SINGLE_CELLS_REL {
                if seen_single {
                    return Err(invalid(format!(
                        "worksheet {sheet_index} has multiple tableSingleCells relationships"
                    )));
                }
                seen_single = true;
                if relationship.is_external() {
                    return Err(invalid("tableSingleCells relationship cannot be external"));
                }
                let part_name = relationship.target_partname()?;
                let part = package.get_part(&part_name)?;
                require_content_type(part, SINGLE_CELLS_CONTENT_TYPE)?;
                require_no_relationships(part, "tableSingleCells part")?;
                let (values, connection_ids) =
                    parse_single_cells_value_with_connection_ids(part.blob(), limits.bindings)?;
                if connection_ids.len() != values.len() {
                    return Err(invalid(
                        "tableSingleCells connection ID projection is inconsistent",
                    ));
                }
                for (binding, connection_id) in values.iter().zip(&connection_ids) {
                    if *connection_id != 0 && !has_connection(connections.as_ref(), *connection_id)
                    {
                        return Err(invalid(format!(
                            "single-cell XML mapping {} on worksheet {sheet_index} references missing connection ID {connection_id}",
                            binding.table_id()
                        )));
                    }
                }
                charge_bindings(
                    values.iter().map(|binding| binding.xpath().as_str()),
                    &mut total_bindings,
                    &mut total_xpath_units,
                    limits,
                )?;
                for value in &values {
                    if !table_ids.insert(value.table_id()) {
                        return Err(invalid(format!(
                            "duplicate XML mapping table ID {}",
                            value.table_id()
                        )));
                    }
                }
                dependencies.push(part_name);
                let start = single_cell_tables.len();
                single_cell_tables
                    .try_reserve(values.len())
                    .map_err(|source| Error::Allocation {
                        resource: "aggregate single-cell XML bindings",
                        source,
                    })?;
                single_cell_tables.extend(values);
                let end = single_cell_tables.len();
                sheet_singles = Some(start..end);
            }
        }

        for rel_id in listed_tables {
            let relationship = worksheet.rels().get(&rel_id).ok_or_else(|| {
                invalid(format!(
                    "worksheet {sheet_index} BrtListPart relationship {rel_id:?} is missing"
                ))
            })?;
            if !matches!(
                relationship.reltype(),
                relationship_type::TABLE | relationship_type::STRICT_TABLE
            ) {
                return Err(invalid(format!(
                    "worksheet {sheet_index} BrtListPart relationship {rel_id:?} has the wrong type"
                )));
            }
        }
        validate_table_overlaps(&ordinary_tables, sheet_index)?;
        if let Some(range) = sheet_singles.as_ref() {
            for value in &single_cell_tables[range.clone()] {
                let cell = crate::package::table::Range {
                    first_row: value.cell().row(),
                    last_row: value.cell().row(),
                    first_column: value.cell().column(),
                    last_column: value.cell().column(),
                };
                if let Some(dimension) = geometry.authoritative_dimension() {
                    require_range_within(cell, dimension, "single-cell XML mapping", sheet_index)?;
                }
                if ordinary_tables
                    .iter()
                    .any(|(_, range)| ranges_overlap(cell, *range))
                {
                    return Err(invalid(format!(
                        "single-cell XML mapping {} overlaps an ordinary table",
                        value.table_id()
                    )));
                }
                if geometry
                    .auto_filter
                    .is_some_and(|filter| ranges_overlap(cell, filter))
                {
                    return Err(invalid(format!(
                        "single-cell XML mapping {} overlaps the worksheet AutoFilter",
                        value.table_id()
                    )));
                }
            }
        }
        single_cell_ranges.push(sheet_singles);
    }

    if !mapped_tables.is_empty() || !single_cell_tables.is_empty() {
        let info = map_info
            .as_ref()
            .ok_or_else(|| invalid("XML table bindings require a Custom XML Maps part"))?;
        validate_binding_map_ids(info, &mapped_tables, &single_cell_tables)?;
    }
    dependencies.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    dependencies.dedup();
    Ok(Loaded {
        map_info,
        conformance,
        mapped_tables,
        single_cell_tables,
        single_cell_ranges,
        worksheet_parts,
        dependency_parts: dependencies,
    })
}

fn validate_connection_references(
    connections: Option<&crate::package::connections::Connections>,
    info: &XmlMapInfo,
) -> Result<()> {
    for map in &info.maps {
        let Some(connection_id) = map
            .data_binding
            .as_ref()
            .and_then(|binding| binding.connection_id)
        else {
            continue;
        };
        if !has_connection(connections, connection_id) {
            return Err(invalid(format!(
                "MapInfo map {} references missing connection ID {connection_id}",
                map.id
            )));
        }
    }
    Ok(())
}

fn validate_table_connection(
    table: &crate::package::table::Table,
    connections: Option<&crate::package::connections::Connections>,
    sheet_index: usize,
) -> Result<()> {
    let Some(connection_id) = table.connection_id else {
        return Ok(());
    };
    if !matches!(table.table_type, crate::package::table::Type::Xml) {
        return Err(invalid(format!(
            "non-XML table {} on worksheet {sheet_index} has nonzero connection ID {connection_id}",
            table.id
        )));
    }
    if !has_connection(connections, connection_id) {
        return Err(invalid(format!(
            "XML table {} on worksheet {sheet_index} references missing connection ID {connection_id}",
            table.id
        )));
    }
    Ok(())
}

fn has_connection(
    connections: Option<&crate::package::connections::Connections>,
    connection_id: u32,
) -> bool {
    connections.is_some_and(|connections| connections.by_id(connection_id).is_some())
}

fn xml_maps_relationship(
    package: &OpcPackage,
    workbook_name: &PackURI,
) -> Result<Option<(PackURI, XmlMapConformance)>> {
    if package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_ooxml_common::spreadsheet_xml_maps::REL
                | litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
        )
    }) {
        return Err(invalid(
            "Custom XML Maps relationship cannot originate from the package root",
        ));
    }
    let mut found = Vec::new();
    for part in package.iter_parts() {
        for relationship in part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                litchi_ooxml_common::spreadsheet_xml_maps::REL
                    | litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
            )
        }) {
            if part.partname() != workbook_name {
                return Err(invalid(
                    "Custom XML Maps relationship must originate from the workbook",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("Custom XML Maps relationship cannot be external"));
            }
            let conformance =
                if relationship.reltype() == litchi_ooxml_common::spreadsheet_xml_maps::REL {
                    XmlMapConformance::Transitional
                } else {
                    XmlMapConformance::Strict
                };
            found.push((relationship.target_partname()?, conformance));
        }
    }
    match found.len() {
        0 => Ok(None),
        1 => Ok(found.pop()),
        _ => Err(invalid(
            "workbook has multiple Custom XML Maps relationships",
        )),
    }
}

/// Reject package-wide resource overages before callers clone or parse the workbook.
pub(crate) fn preflight(package: &OpcPackage, limits: super::snapshot::ReadLimits) -> Result<()> {
    if package.part_count() > limits.max_parts {
        return Err(invalid("package exceeds the XML Maps maximum part count"));
    }
    let mut relationships = package.rels().iter().count();
    let mut total = 0usize;
    for part in package.iter_parts() {
        relationships = relationships
            .checked_add(part.rels().iter().count())
            .ok_or_else(|| invalid("relationship count overflow"))?;
        total = total
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("package byte count overflow"))?;
    }
    if relationships > limits.max_relationships {
        return Err(invalid(
            "package exceeds the XML Maps maximum relationship count",
        ));
    }
    if total > limits.max_total_bytes {
        return Err(invalid(
            "package exceeds the XML Maps maximum total byte count",
        ));
    }
    let budget = litchi_core::Budget::root("xlsb.xml_maps", limits.core);
    for (resource, amount) in [
        (litchi_core::Resource::InputBytes, total as u64),
        (litchi_core::Resource::Memory, total as u64),
        (litchi_core::Resource::Objects, package.part_count() as u64),
        (litchi_core::Resource::Work, relationships as u64),
    ] {
        budget
            .consume(resource, amount)
            .map_err(|error| invalid(error.to_string()))?;
    }
    Ok(())
}

fn validate_single_cell_ownership(package: &OpcPackage, worksheets: &[PackURI]) -> Result<()> {
    let worksheet_names = worksheets.iter().collect::<HashSet<_>>();
    let mut targets = HashSet::new();
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == SINGLE_CELLS_REL)
        {
            if !worksheet_names.contains(&part.partname()) {
                return Err(invalid(
                    "tableSingleCells relationship must originate from a worksheet",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("tableSingleCells relationship cannot be external"));
            }
            let target = relationship.target_partname()?;
            if !targets.insert(target.clone()) {
                return Err(invalid(
                    "tableSingleCells part is shared by multiple worksheets",
                ));
            }
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == SINGLE_CELLS_CONTENT_TYPE)
    {
        if !targets.contains(part.partname()) {
            return Err(invalid("tableSingleCells part is orphaned"));
        }
    }
    Ok(())
}

fn validate_table_range(range: &crate::package::table::Range) -> Result<()> {
    if range.first_row > range.last_row
        || range.first_column > range.last_column
        || range.last_row > MAX_ROW
        || range.last_column > MAX_COLUMN
    {
        return Err(invalid("ordinary table range exceeds worksheet dimensions"));
    }
    Ok(())
}

fn validate_table_height(table: &crate::package::table::Table, sheet_index: usize) -> Result<()> {
    let height = u64::from(table.range.last_row) - u64::from(table.range.first_row) + 1;
    let reserved = u64::from(table.header_row_count) + u64::from(table.totals_row_count);
    if height <= reserved {
        return Err(invalid(format!(
            "ordinary table {} on worksheet {sheet_index} has no data rows beyond header and totals rows",
            table.id
        )));
    }
    Ok(())
}

fn validate_table_overlaps(
    tables: &[(u32, crate::package::table::Range)],
    sheet_index: usize,
) -> Result<()> {
    for (index, (left_id, left)) in tables.iter().enumerate() {
        for (right_id, right) in tables.iter().skip(index + 1) {
            if ranges_overlap(*left, *right) {
                return Err(invalid(format!(
                    "ordinary tables {left_id} and {right_id} overlap on worksheet {sheet_index}"
                )));
            }
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Default)]
struct WorksheetGeometry {
    dimension: Option<crate::package::table::Range>,
    auto_filter: Option<crate::package::table::Range>,
}

impl WorksheetGeometry {
    /// XLSB writers commonly use A1:A1 as an empty-sheet sentinel even when
    /// table metadata extends beyond stored cells, so it is not authoritative
    /// for package-scoped table ownership checks.
    fn authoritative_dimension(self) -> Option<crate::package::table::Range> {
        self.dimension.filter(|range| {
            range.first_row != 0
                || range.last_row != 0
                || range.first_column != 0
                || range.last_column != 0
        })
    }
}

fn worksheet_geometry(data: &[u8]) -> Result<WorksheetGeometry> {
    let mut geometry = WorksheetGeometry::default();
    for record in crate::raw::Records::new(data) {
        let record = record?;
        let slot = match record.kind() {
            crate::raw::kind::WS_DIM => &mut geometry.dimension,
            crate::raw::kind::BEGIN_A_FILTER => &mut geometry.auto_filter,
            _ => continue,
        };
        if slot.is_some() {
            return Err(invalid(
                "worksheet has duplicate BrtWsDim or BrtBeginAFilter records",
            ));
        }
        let payload = record.payload();
        if payload.len() != 16 {
            return Err(Error::InvalidLength {
                expected: 16,
                found: payload.len(),
            });
        }
        let read = |offset| {
            u32::from_le_bytes(
                payload[offset..offset + 4]
                    .try_into()
                    .expect("fixed range slice"),
            )
        };
        let range = crate::package::table::Range {
            first_row: read(0),
            last_row: read(4),
            first_column: read(8),
            last_column: read(12),
        };
        validate_table_range(&range)?;
        *slot = Some(range);
    }
    Ok(geometry)
}

fn require_range_within(
    inner: crate::package::table::Range,
    outer: crate::package::table::Range,
    name: &str,
    sheet_index: usize,
) -> Result<()> {
    if inner.first_row < outer.first_row
        || inner.last_row > outer.last_row
        || inner.first_column < outer.first_column
        || inner.last_column > outer.last_column
    {
        return Err(invalid(format!(
            "{name} on worksheet {sheet_index} exceeds BrtWsDim"
        )));
    }
    Ok(())
}

fn ranges_overlap(left: crate::package::table::Range, right: crate::package::table::Range) -> bool {
    left.first_row <= right.last_row
        && right.first_row <= left.last_row
        && left.first_column <= right.last_column
        && right.first_column <= left.last_column
}

fn charge_bindings<'a>(
    xpaths: impl Iterator<Item = &'a str>,
    total_bindings: &mut usize,
    total_xpath_units: &mut usize,
    limits: super::snapshot::ReadLimits,
) -> Result<()> {
    for xpath in xpaths {
        *total_bindings = total_bindings
            .checked_add(1)
            .ok_or_else(|| invalid("aggregate XML binding count overflow"))?;
        *total_xpath_units = total_xpath_units
            .checked_add(xpath.encode_utf16().count())
            .ok_or_else(|| invalid("aggregate XML XPath length overflow"))?;
        if *total_bindings > limits.max_total_bindings {
            return Err(invalid("XML Maps aggregate binding limit exceeded"));
        }
        if *total_xpath_units > limits.max_total_xpath_units {
            return Err(invalid("XML Maps aggregate XPath limit exceeded"));
        }
    }
    Ok(())
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: expected.to_string(),
            got: part.content_type().to_string(),
        })
    }
}

fn require_no_relationships(part: &dyn Part, owner: &str) -> Result<()> {
    if part.rels().iter().next().is_none() {
        Ok(())
    } else {
        Err(invalid(format!(
            "{owner} must not have outbound relationships"
        )))
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
