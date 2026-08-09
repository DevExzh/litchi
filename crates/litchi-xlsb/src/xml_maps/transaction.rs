//! Detached, bounded edits for XLSB Custom XML Maps.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::{PackURI, TargetMode};

use super::snapshot::{SourcePart, SourceRelationship, SourceState};
use super::{
    CellReference, ColumnBinding, Commit, MappedTable, Patch, SingleCellBinding, Snapshot, XmlMap,
    XmlMapConformance, XmlMapInfo, parse_single_cells, parse_xml_map_info_with_limits,
    patch_single_cells, patch_table_bindings, validate_binding_map_ids, validate_catalog,
};
use crate::package::error::{Error, Result};

const TABLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.table";
const SINGLE_CELLS_CONTENT_TYPE: &str = "application/vnd.ms-excel.tableSingleCells";
const SINGLE_CELLS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";

/// A validated draft whose commit remains bound to the exact source graph.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    catalog: Option<XmlMapInfo>,
    conformance: XmlMapConformance,
    mapped_tables: Vec<MappedTable>,
    single_cell_bindings: Vec<Option<Vec<SingleCellBinding>>>,
}

impl Transaction {
    pub(crate) fn new(before: Snapshot) -> Self {
        Self {
            catalog: before.map_info().cloned(),
            conformance: before.conformance(),
            mapped_tables: before.mapped_tables().to_vec(),
            single_cell_bindings: before.binding_groups(),
            before,
        }
    }

    /// Immutable source snapshot used by the eventual stale check.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the staged MapInfo catalog.
    pub fn catalog(&self) -> Option<&XmlMapInfo> {
        self.catalog.as_ref()
    }

    /// Replace or create the complete catalog after validating all bindings.
    pub fn set_catalog(&mut self, catalog: XmlMapInfo) -> Result<bool> {
        self.replace_catalog(Some(catalog))
    }

    /// Replace, create, or remove the complete catalog.
    pub fn replace_catalog(&mut self, catalog: Option<XmlMapInfo>) -> Result<bool> {
        validate_draft(
            catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        )?;
        if self.catalog == catalog {
            return Ok(false);
        }
        self.catalog = catalog;
        Ok(true)
    }

    /// Remove the catalog only when no binary binding depends on it.
    pub fn remove_catalog(&mut self) -> Result<Option<XmlMapInfo>> {
        if !self.mapped_tables.is_empty()
            || self
                .single_cell_bindings
                .iter()
                .filter_map(Option::as_ref)
                .any(|values| !values.is_empty())
        {
            return Err(invalid(
                "cannot remove the XML Maps catalog while binary bindings depend on it",
            ));
        }
        Ok(self.catalog.take())
    }

    /// Set one map name by stable map ID. Failed validation retains the draft.
    pub fn set_map_name(&mut self, map_id: u32, name: impl Into<String>) -> Result<bool> {
        let name = name.into();
        let catalog = self
            .catalog
            .as_mut()
            .ok_or_else(|| invalid("cannot rename a map in an absent catalog"))?;
        let map = catalog
            .maps
            .iter_mut()
            .find(|map| map.id == map_id)
            .ok_or_else(|| invalid(format!("XML map ID {map_id} was not found")))?;
        if map.name == name {
            return Ok(false);
        }
        let previous = std::mem::replace(&mut map.name, name);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(true),
            Err(error) => {
                self.catalog
                    .as_mut()
                    .and_then(|catalog| catalog.maps.iter_mut().find(|map| map.id == map_id))
                    .expect("renamed map remains present")
                    .name = previous;
                Err(error)
            },
        }
    }

    /// Edit one complete map by stable ID through a retry-safe cloned draft.
    pub fn edit_map(
        &mut self,
        map_id: u32,
        edit: impl FnOnce(&mut XmlMap) -> Result<()>,
    ) -> Result<bool> {
        let catalog = self
            .catalog
            .as_mut()
            .ok_or_else(|| invalid("cannot edit a map in an absent catalog"))?;
        let index = catalog
            .maps
            .iter()
            .position(|map| map.id == map_id)
            .ok_or_else(|| invalid(format!("XML map ID {map_id} was not found")))?;
        let before = catalog.maps[index].clone();
        if let Err(error) = edit(&mut catalog.maps[index]) {
            catalog.maps[index] = before;
            return Err(error);
        }
        if catalog.maps[index] == before {
            return Ok(false);
        }
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(true),
            Err(error) => {
                self.catalog.as_mut().expect("catalog remains present").maps[index] = before;
                Err(error)
            },
        }
    }

    /// Select the namespace conformance for the published catalog.
    pub fn set_conformance(&mut self, conformance: XmlMapConformance) -> Result<bool> {
        if self.conformance == conformance {
            return Ok(false);
        }
        let Some(catalog) = self.catalog.as_ref() else {
            return Ok(false);
        };
        serialize_catalog(catalog, conformance, self.before.limits())?;
        self.conformance = conformance;
        Ok(true)
    }

    /// Put one binding into an already existing worksheet single-cell part.
    pub fn put_single_cell_binding(
        &mut self,
        sheet_index: usize,
        binding: SingleCellBinding,
    ) -> Result<bool> {
        let values = self
            .single_cell_bindings
            .get(sheet_index)
            .ok_or_else(|| Error::WorksheetNotFound(sheet_index.to_string()))?
            .as_ref()
            .ok_or_else(|| {
                Error::UnsupportedFeature(
                    "creating a new tableSingleCells part is intentionally refused".to_string(),
                )
            })?;
        let mut values = values.clone();
        if let Some(index) = values
            .iter()
            .position(|value| value.table_id() == binding.table_id())
        {
            if values[index] == binding {
                return Ok(false);
            }
            values[index] = binding;
        } else {
            values.push(binding);
        }
        let previous = self.single_cell_bindings[sheet_index].replace(values);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(true),
            Err(error) => {
                self.single_cell_bindings[sheet_index] = previous;
                Err(error)
            },
        }
    }

    /// Remove one single-cell binding by its stable table ID.
    pub fn remove_single_cell_binding(
        &mut self,
        sheet_index: usize,
        table_id: u32,
    ) -> Result<Option<SingleCellBinding>> {
        let Some(values) = self
            .single_cell_bindings
            .get(sheet_index)
            .ok_or_else(|| Error::WorksheetNotFound(sheet_index.to_string()))?
            .as_ref()
        else {
            return Ok(None);
        };
        let mut values = values.clone();
        let Some(index) = values.iter().position(|value| value.table_id() == table_id) else {
            return Ok(None);
        };
        let removed = values.remove(index);
        let previous = self.single_cell_bindings[sheet_index].replace(values);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(Some(removed)),
            Err(error) => {
                self.single_cell_bindings[sheet_index] = previous;
                Err(error)
            },
        }
    }

    /// Put one mapped binding into an existing ordinary table column.
    pub fn put_table_column_binding(
        &mut self,
        table_id: u32,
        binding: ColumnBinding,
    ) -> Result<bool> {
        let column_id = binding.column_id();
        let table_index = self
            .mapped_tables
            .iter()
            .position(|table| table.table_id() == table_id);
        let mut columns = table_index.map_or_else(Vec::new, |index| {
            self.mapped_tables[index].columns().to_vec()
        });
        if let Some(index) = columns
            .iter()
            .position(|value| value.column_id() == binding.column_id())
        {
            if columns[index] == binding {
                return Ok(false);
            }
            columns[index] = binding;
        } else {
            columns.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "mapped table column transaction draft",
                source,
            })?;
            columns.push(binding);
        }
        let replacement =
            MappedTable::new_with_limits(table_id, columns, self.before.limits().bindings)?;
        let source = self.source_table(table_id, column_id)?;
        patch_table_bindings(&source, &replacement, self.before.limits().bindings)?;
        let previous = if let Some(index) = table_index {
            Some(std::mem::replace(
                &mut self.mapped_tables[index],
                replacement,
            ))
        } else {
            self.mapped_tables
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "mapped table transaction draft",
                    source,
                })?;
            self.mapped_tables.push(replacement);
            None
        };
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(true),
            Err(error) => {
                if let (Some(index), Some(previous)) = (table_index, previous) {
                    self.mapped_tables[index] = previous;
                } else {
                    self.mapped_tables.pop();
                }
                Err(error)
            },
        }
    }

    /// Remove one ordinary-table column binding without removing its table part.
    pub fn remove_table_column_binding(
        &mut self,
        table_id: u32,
        column_id: u32,
    ) -> Result<Option<ColumnBinding>> {
        let table_index = self
            .mapped_tables
            .iter()
            .position(|table| table.table_id() == table_id)
            .ok_or_else(|| invalid(format!("mapped table ID {table_id} was not found")))?;
        let mut columns = self.mapped_tables[table_index].columns().to_vec();
        let Some(column_index) = columns
            .iter()
            .position(|value| value.column_id() == column_id)
        else {
            return Ok(None);
        };
        let removed = columns.remove(column_index);
        let replacement =
            MappedTable::new_with_limits(table_id, columns, self.before.limits().bindings)?;
        let previous = std::mem::replace(&mut self.mapped_tables[table_index], replacement);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        ) {
            Ok(()) => Ok(Some(removed)),
            Err(error) => {
                self.mapped_tables[table_index] = previous;
                Err(error)
            },
        }
    }

    /// Build a reversible source-bound commit without touching a workbook.
    pub fn commit(self) -> Result<Commit> {
        validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.before.limits(),
        )?;
        if self.before.map_info() == self.catalog.as_ref()
            && self.before.conformance() == self.conformance
            && self.before.mapped_tables() == self.mapped_tables
            && self.before.binding_groups_equal(&self.single_cell_bindings)
        {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(patch, false));
        }

        let mut source = self.before.source().clone();
        self.materialize_catalog(&mut source)?;
        self.materialize_tables(&mut source)?;
        self.materialize_single_cells(&mut source)?;
        source.remove_signatures();
        let mut projected_tables = Vec::new();
        projected_tables
            .try_reserve(self.mapped_tables.len())
            .map_err(|source| Error::Allocation {
                resource: "mapped table snapshot projection",
                source,
            })?;
        projected_tables.extend(
            self.mapped_tables
                .into_iter()
                .filter(|table| !table.columns().is_empty()),
        );
        let projected_catalog = source
            .dependencies
            .iter()
            .find(|part| {
                part.content_type == litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE
            })
            .map(|part| {
                parse_xml_map_info_with_limits(part.bytes(), &self.before.limits().xml_maps)
                    .map_err(|error| invalid(error.to_string()))
            })
            .transpose()?;
        let projected_conformance = if projected_catalog.is_some() {
            self.conformance
        } else {
            XmlMapConformance::Transitional
        };
        let after = Snapshot::materialized(
            &self.before,
            projected_catalog,
            projected_conformance,
            projected_tables,
            self.single_cell_bindings,
            source,
        )?;
        let patch = Patch::new(self.before, after);
        Ok(Commit::new(patch, true))
    }

    fn materialize_catalog(&self, source: &mut SourceState) -> Result<()> {
        let content_type = litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE;
        let existing = source
            .dependencies
            .iter()
            .position(|part| part.content_type == content_type);
        match (existing, self.before.map_info(), self.catalog.as_ref()) {
            (Some(index), Some(before), Some(after)) => {
                let xml = litchi_ooxml_common::spreadsheet_xml_maps::patch_xml_map_info_source_with_limits(
                    source.dependencies[index].bytes(),
                    before,
                    after,
                    self.before.conformance(),
                    self.conformance,
                    &self.before.limits().xml_maps,
                )?;
                parse_xml_map_info_with_limits(&xml, &self.before.limits().xml_maps)
                    .map_err(|error| invalid(error.to_string()))?;
                source.dependencies[index].bytes = Arc::new(xml);
                let map_name = source.dependencies[index].part_name.clone();
                if self.before.conformance() != self.conformance {
                    let relationship = source
                        .workbook
                        .relationships
                        .iter_mut()
                        .find(|relationship| {
                            matches!(
                                relationship.relationship_type.as_str(),
                                litchi_ooxml_common::spreadsheet_xml_maps::REL
                                    | litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
                            )
                        })
                        .ok_or_else(|| invalid("MapInfo relationship is absent"))?;
                    relationship.relationship_type =
                        self.conformance.relationship_type().to_string();
                    relationship.target =
                        map_name.relative_ref(source.workbook.part_name.base_uri());
                }
            },
            (None, None, Some(after)) => {
                let part_name = next_map_part_name(source)?;
                let xml = serialize_catalog(after, self.conformance, self.before.limits())?;
                let relationship_id = next_relationship_id(&source.workbook.relationships);
                let target = part_name.relative_ref(source.workbook.part_name.base_uri());
                source.dependencies.push(SourcePart {
                    part_name: part_name.clone(),
                    content_type: content_type.to_string(),
                    bytes: Arc::new(xml),
                    relationships: Vec::new(),
                });
                source.part_names.push(part_name);
                source
                    .part_names
                    .sort_by(|left, right| left.as_str().cmp(right.as_str()));
                source.workbook.relationships.push(SourceRelationship {
                    id: relationship_id,
                    relationship_type: self.conformance.relationship_type().to_string(),
                    target,
                    mode: TargetMode::Internal,
                });
                source
                    .workbook
                    .relationships
                    .sort_by(|left, right| left.id.cmp(&right.id));
            },
            (Some(index), Some(_), None) => {
                let part_name = source.dependencies.remove(index).part_name;
                source.part_names.retain(|name| name != &part_name);
                source.workbook.relationships.retain(|relationship| {
                    !matches!(
                        relationship.relationship_type.as_str(),
                        litchi_ooxml_common::spreadsheet_xml_maps::REL
                            | litchi_ooxml_common::spreadsheet_xml_maps::STRICT_REL
                    )
                });
            },
            (None, None, None) => {},
            _ => {
                return Err(invalid(
                    "MapInfo source topology does not match its catalog",
                ));
            },
        }
        source
            .dependencies
            .sort_by(|left, right| left.part_name.as_str().cmp(right.part_name.as_str()));
        Ok(())
    }

    fn source_table(&self, table_id: u32, column_id: u32) -> Result<super::TableBindingsSource> {
        let mut found = None;
        for part in self
            .before
            .source()
            .dependencies
            .iter()
            .filter(|part| part.content_type == TABLE_CONTENT_TYPE)
        {
            let host = crate::package::table::parse_table_part(part.bytes())?;
            if host.id != table_id {
                continue;
            }
            if host.table_type != crate::package::table::Type::Xml || host.single_cell {
                return Err(invalid(format!(
                    "table ID {table_id} is not an ordinary LTXML Table"
                )));
            }
            if !host.columns.iter().any(|column| column.id == column_id) {
                return Err(invalid(format!(
                    "column {column_id} is absent from physical table ID {table_id}"
                )));
            }
            let parsed =
                match super::parse_table_bindings(part.bytes(), self.before.limits().bindings) {
                    Ok(parsed) => parsed,
                    Err(error) => return Err(error),
                };
            if parsed.value().table_id() != host.id {
                return Err(invalid(format!(
                    "host Table ID {} disagrees with XML binding table ID {}",
                    host.id,
                    parsed.value().table_id()
                )));
            }
            if found.is_some() {
                return Err(invalid(format!(
                    "mapped table ID {table_id} is ambiguous across LTXML Table parts"
                )));
            }
            found = Some(parsed);
        }
        found.ok_or_else(|| {
            invalid(format!(
                "mapped table ID {table_id} was not found in an existing LTXML Table part"
            ))
        })
    }

    fn materialize_tables(&self, source: &mut SourceState) -> Result<()> {
        let mut tables = HashMap::new();
        tables
            .try_reserve(self.mapped_tables.len())
            .map_err(|source| Error::Allocation {
                resource: "mapped table materialization index",
                source,
            })?;
        for table in &self.mapped_tables {
            tables.insert(table.table_id(), table);
        }
        for part in source
            .dependencies
            .iter_mut()
            .filter(|part| part.content_type == TABLE_CONTENT_TYPE)
        {
            let parsed =
                match super::parse_table_bindings(part.bytes(), self.before.limits().bindings) {
                    Ok(parsed) => parsed,
                    Err(Error::Unrecognized { typ, .. }) if typ == "mapped table type" => continue,
                    Err(error) => return Err(error),
                };
            let Some(table) = tables.get(&parsed.value().table_id()).copied() else {
                continue;
            };
            part.bytes = Arc::new(patch_table_bindings(
                &parsed,
                table,
                self.before.limits().bindings,
            )?);
        }
        Ok(())
    }

    fn materialize_single_cells(&self, source: &mut SourceState) -> Result<()> {
        for part in source
            .dependencies
            .iter_mut()
            .filter(|part| part.content_type == SINGLE_CELLS_CONTENT_TYPE)
        {
            let target_sheet = source
                .worksheets
                .iter()
                .position(|worksheet| {
                    let target = part.part_name.relative_ref(worksheet.part_name.base_uri());
                    worksheet.relationships.iter().any(|relationship| {
                        relationship.relationship_type == SINGLE_CELLS_REL
                            && relationship.mode == TargetMode::Internal
                            && relationship.target == target
                    })
                })
                .ok_or_else(|| invalid("tableSingleCells owner worksheet is absent"))?;
            let values = self
                .single_cell_bindings
                .get(target_sheet)
                .and_then(Option::as_deref)
                .ok_or_else(|| {
                    Error::UnsupportedFeature(
                        "removing a tableSingleCells part is intentionally refused".to_string(),
                    )
                })?;
            let parsed = parse_single_cells(part.bytes(), self.before.limits().bindings)?;
            part.bytes = Arc::new(patch_single_cells(
                &parsed,
                values,
                self.before.limits().bindings,
            )?);
        }
        Ok(())
    }
}

fn validate_draft(
    catalog: Option<&XmlMapInfo>,
    mapped_tables: &[MappedTable],
    groups: &[Option<Vec<SingleCellBinding>>],
    conformance: XmlMapConformance,
    limits: super::ReadLimits,
) -> Result<()> {
    let has_singles = groups
        .iter()
        .filter_map(Option::as_ref)
        .any(|values| !values.is_empty());
    let Some(catalog) = catalog else {
        if !mapped_tables.is_empty() || has_singles {
            return Err(invalid("binary XML bindings require a MapInfo catalog"));
        }
        return Ok(());
    };
    validate_catalog(catalog, limits.xml_maps)?;
    serialize_catalog(catalog, conformance, limits)?;
    validate_binding_map_ids(catalog, mapped_tables, &[])?;

    let mut total_bindings = 0usize;
    let mut total_xpath_units = 0usize;
    for table in mapped_tables {
        super::validation::mapped_table(table, limits.bindings)?;
        for binding in table.columns() {
            total_bindings = total_bindings
                .checked_add(1)
                .ok_or(Error::CapacityOverflow {
                    resource: "aggregate XML binding count",
                })?;
            total_xpath_units = total_xpath_units
                .checked_add(binding.xpath().as_str().encode_utf16().count())
                .ok_or(Error::CapacityOverflow {
                    resource: "aggregate XML binding XPath units",
                })?;
        }
    }
    for values in groups.iter().filter_map(Option::as_ref) {
        super::validation::single_cells(values, limits.bindings)?;
        for binding in values {
            total_bindings = total_bindings
                .checked_add(1)
                .ok_or(Error::CapacityOverflow {
                    resource: "aggregate XML binding count",
                })?;
            total_xpath_units = total_xpath_units
                .checked_add(binding.xpath().as_str().encode_utf16().count())
                .ok_or(Error::CapacityOverflow {
                    resource: "aggregate XML binding XPath units",
                })?;
        }
    }
    if total_bindings > limits.max_total_bindings {
        return Err(invalid(format!(
            "aggregate XML binding count {total_bindings} exceeds {}",
            limits.max_total_bindings
        )));
    }
    if total_xpath_units > limits.max_total_xpath_units {
        return Err(invalid(format!(
            "aggregate XML binding XPath units {total_xpath_units} exceeds {}",
            limits.max_total_xpath_units
        )));
    }

    let mut map_ids = HashSet::new();
    map_ids
        .try_reserve(catalog.maps.len())
        .map_err(|source| Error::Allocation {
            resource: "XML map ID validation index",
            source,
        })?;
    map_ids.extend(catalog.maps.iter().map(|map| map.id));
    let mut table_ids = HashSet::new();
    let singleton_count = groups
        .iter()
        .filter_map(Option::as_ref)
        .try_fold(0usize, |count, values| count.checked_add(values.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "single-cell binding validation count",
        })?;
    table_ids
        .try_reserve(mapped_tables.len().saturating_add(singleton_count))
        .map_err(|source| Error::Allocation {
            resource: "XML binding table ID validation index",
            source,
        })?;
    table_ids.extend(mapped_tables.iter().map(MappedTable::table_id));
    for binding in groups
        .iter()
        .filter_map(Option::as_ref)
        .flat_map(|values| values.iter())
    {
        if !map_ids.contains(&binding.map_id()) {
            return Err(invalid(format!(
                "single-cell binding references absent XML map ID {}",
                binding.map_id()
            )));
        }
        if !table_ids.insert(binding.table_id()) {
            return Err(invalid(format!(
                "duplicate XML mapping table ID {}",
                binding.table_id()
            )));
        }
    }
    Ok(())
}

fn next_map_part_name(source: &SourceState) -> Result<PackURI> {
    for index in 1usize..=source.part_names.len().saturating_add(1) {
        let candidate = if index == 1 {
            "/xl/xmlMaps.xml".to_string()
        } else {
            format!("/xl/xmlMaps{index}.xml")
        };
        let uri = PackURI::new(candidate)?;
        if !source.part_names.iter().any(|name| name == &uri) {
            return Ok(uri);
        }
    }
    Err(invalid("could not allocate a bounded MapInfo part name"))
}

fn serialize_catalog(
    catalog: &XmlMapInfo,
    conformance: XmlMapConformance,
    limits: super::ReadLimits,
) -> Result<Vec<u8>> {
    litchi_ooxml_common::spreadsheet_xml_maps::serialize_xml_map_info_with_limits(
        catalog,
        conformance,
        &limits.xml_maps,
    )
    .map_err(|error| invalid(error.to_string()))
}

fn next_relationship_id(relationships: &[SourceRelationship]) -> String {
    for index in 1usize..=relationships.len().saturating_add(1) {
        let candidate = format!("rId{index}");
        if !relationships
            .iter()
            .any(|relationship| relationship.id == candidate)
        {
            return candidate;
        }
    }
    unreachable!("len + 1 guarantees an unused relationship ID")
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
fn _cell_reference_is_public(_: CellReference) {}

#[cfg(test)]
mod transaction_tests {
    use super::*;
    use crate::xml_maps::{DataBinding, ReadLimits, XPath, XmlDataType, XmlSchema};

    #[test]
    fn absent_catalog_conformance_is_an_exact_signed_no_op() {
        let snapshot = Snapshot::empty_transaction_fixture(true);
        let before = snapshot.source().clone();
        let mut transaction = snapshot.edit();
        assert!(
            !transaction
                .set_conformance(XmlMapConformance::Strict)
                .expect("absent catalog conformance is inert")
        );
        let commit = transaction.commit().expect("commit no-op");
        assert!(!commit.changed());
        assert!(commit.patch().is_empty());
        assert_eq!(commit.patch().before().source(), &before);
        assert_eq!(commit.patch().after().source(), &before);
    }

    #[test]
    fn aggregate_binding_and_xpath_limits_are_exact() {
        let catalog = fixture_catalog();
        let binding = SingleCellBinding::new(
            1,
            1,
            CellReference::new(0, 0).expect("cell"),
            7,
            XmlDataType::new(1).expect("type"),
            XPath::new("/r").expect("XPath"),
        )
        .expect("binding");
        let groups = vec![Some(vec![binding])];
        let mut exact = ReadLimits::DEFAULT;
        exact.max_total_bindings = 1;
        exact.max_total_xpath_units = 2;
        validate_draft(
            Some(&catalog),
            &[],
            &groups,
            XmlMapConformance::Transitional,
            exact,
        )
        .expect("exact aggregate limits are inclusive");

        let mut below = exact;
        below.max_total_bindings = 0;
        assert!(
            validate_draft(
                Some(&catalog),
                &[],
                &groups,
                XmlMapConformance::Transitional,
                below,
            )
            .expect_err("binding limit")
            .to_string()
            .contains("binding")
        );
        below = exact;
        below.max_total_xpath_units = 1;
        assert!(
            validate_draft(
                Some(&catalog),
                &[],
                &groups,
                XmlMapConformance::Transitional,
                below,
            )
            .expect_err("XPath limit")
            .to_string()
            .contains("XPath")
        );

        let xml_bytes = serialize_catalog(
            &catalog,
            XmlMapConformance::Transitional,
            ReadLimits::DEFAULT,
        )
        .expect("default serialization")
        .len();
        let mut exact_xml = exact;
        exact_xml.xml_maps.max_part_bytes = xml_bytes;
        validate_draft(
            Some(&catalog),
            &[],
            &groups,
            XmlMapConformance::Transitional,
            exact_xml,
        )
        .expect("exact authored XML byte limit is inclusive");
        exact_xml.xml_maps.max_part_bytes = xml_bytes - 1;
        assert!(
            validate_draft(
                Some(&catalog),
                &[],
                &groups,
                XmlMapConformance::Transitional,
                exact_xml,
            )
            .is_err()
        );

        let opaque_bytes = catalog.schemas[0]
            .payload_xml
            .as_ref()
            .expect("opaque schema")
            .len();
        let mut exact_opaque = exact;
        exact_opaque.xml_maps.max_opaque_bytes = opaque_bytes;
        validate_draft(
            Some(&catalog),
            &[],
            &groups,
            XmlMapConformance::Transitional,
            exact_opaque,
        )
        .expect("exact opaque XML limit is inclusive");
        exact_opaque.xml_maps.max_opaque_bytes = opaque_bytes - 1;
        assert!(
            validate_draft(
                Some(&catalog),
                &[],
                &groups,
                XmlMapConformance::Transitional,
                exact_opaque,
            )
            .is_err()
        );
    }

    #[test]
    fn duplicate_singleton_cell_failure_retains_the_draft() {
        let before = Snapshot::empty_transaction_fixture(false);
        let first = SingleCellBinding::new(
            1,
            1,
            CellReference::new(0, 0).expect("cell"),
            7,
            XmlDataType::new(1).expect("type"),
            XPath::new("/r").expect("XPath"),
        )
        .expect("binding");
        let mut transaction = Transaction {
            before,
            catalog: Some(fixture_catalog()),
            conformance: XmlMapConformance::Transitional,
            mapped_tables: Vec::new(),
            single_cell_bindings: vec![Some(vec![first.clone()])],
        };
        let duplicate_cell = SingleCellBinding::new(
            2,
            1,
            CellReference::new(0, 0).expect("cell"),
            7,
            XmlDataType::new(1).expect("type"),
            XPath::new("/r").expect("XPath"),
        )
        .expect("binding");
        assert!(
            transaction
                .put_single_cell_binding(0, duplicate_cell)
                .expect_err("duplicate cell")
                .to_string()
                .contains("duplicate cell")
        );
        assert_eq!(
            transaction.single_cell_bindings[0].as_deref(),
            Some(&[first][..])
        );
    }

    #[test]
    fn additive_binding_refuses_a_non_xml_table_without_mutating_the_draft() {
        let snapshot = Snapshot::non_xml_table_transaction_fixture();
        let source = snapshot.source().clone();
        let mut transaction = snapshot.edit();
        let binding = ColumnBinding::new(
            1,
            7,
            XmlDataType::new(1).expect("type"),
            XPath::new("/r").expect("XPath"),
            false,
        )
        .expect("binding");
        let error = transaction
            .put_table_column_binding(2, binding)
            .expect_err("non-XML table refusal");
        assert!(error.to_string().contains("not an ordinary LTXML Table"));
        assert!(transaction.mapped_tables.is_empty());
        assert_eq!(transaction.before.source(), &source);
    }

    fn fixture_catalog() -> XmlMapInfo {
        XmlMapInfo {
            selection_namespaces: "xmlns:e='urn:test'".to_string(),
            schemas: vec![XmlSchema {
                id: "schema-7".to_string(),
                schema_reference: Some("urn:test".to_string()),
                namespace: Some("urn:test".to_string()),
                payload_xml: Some(b"<e:schema xmlns:e=\"urn:test\"/>".to_vec()),
            }],
            maps: vec![XmlMap {
                id: 7,
                name: "Map".to_string(),
                root_element: "root".to_string(),
                schema_id: "schema-7".to_string(),
                show_import_export_validation_errors: false,
                auto_fit: false,
                append: false,
                preserve_sort_auto_filter_layout: false,
                preserve_format: false,
                data_binding: None::<DataBinding>,
            }],
        }
    }
}
