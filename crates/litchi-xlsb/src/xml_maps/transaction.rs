#![allow(
    clippy::expect_used,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! Detached, bounded edits for XLSB Custom XML Maps.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::{PackURI, TargetMode};

use super::snapshot::{SourcePart, SourceRelationship, SourceState};
use super::{
    CellReference, ColumnBinding, Commit, MappedTable, Patch, SingleCellBinding, Snapshot, XmlMap,
    XmlMapConformance, XmlMapInfo, parse_single_cells, parse_xml_map_info, patch_single_cells,
    patch_table_bindings, serialize_single_cells, validate_binding_map_ids, validate_catalog,
};
use crate::package::error::{Error, Result};

const TABLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.table";
const SINGLE_CELLS_CONTENT_TYPE: &str = "application/vnd.ms-excel.tableSingleCells";
const SINGLE_CELLS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";

crate::part_edit::limits_options!(super::SemanticLimits, "Custom XML Maps");

/// A staged Custom XML Maps edit of one package, bound to its exact source
/// graph. Dropping it discards it.
#[derive(Clone, Debug)]
pub struct Edit {
    source: crate::Package,
    limits: super::SemanticLimits,
    before: Option<Snapshot>,
    catalog: Option<XmlMapInfo>,
    conformance: XmlMapConformance,
    mapped_tables: Vec<MappedTable>,
    single_cell_bindings: Vec<Option<Vec<SingleCellBinding>>>,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, limits: super::SemanticLimits) -> Self {
        Self {
            source,
            limits,
            before: None,
            catalog: None,
            conformance: XmlMapConformance::default(),
            mapped_tables: Vec::new(),
            single_cell_bindings: Vec::new(),
        }
    }

    #[cfg(test)]
    pub(crate) fn from_snapshot(source: crate::Package, before: Snapshot) -> Self {
        Self {
            source,
            limits: before.limits(),
            catalog: before.map_info().cloned(),
            conformance: before.conformance(),
            mapped_tables: before.mapped_tables().to_vec(),
            single_cell_bindings: before.binding_groups(),
            before: Some(before),
        }
    }

    fn ensure_prepared(&mut self) -> Result<()> {
        if self.before.is_some() {
            return Ok(());
        }
        let worksheets = self.source.workbook()?.xml_maps_worksheet_parts()?;
        let before =
            Snapshot::read_for_worksheets(self.source.opc_package(), worksheets, self.limits)?;
        self.catalog = before.map_info().cloned();
        self.conformance = before.conformance();
        self.mapped_tables = before.mapped_tables().to_vec();
        self.single_cell_bindings = before.binding_groups();
        self.before = Some(before);
        Ok(())
    }

    fn snapshot(&self) -> Result<&Snapshot> {
        self.before
            .as_ref()
            .ok_or_else(|| invalid("Custom XML Maps edit is not prepared"))
    }

    /// Immutable source snapshot used by the eventual stale check.
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()?;
        self.snapshot()
    }

    /// Borrow the staged MapInfo catalog.
    pub fn catalog(&mut self) -> Result<Option<&XmlMapInfo>> {
        self.ensure_prepared()?;
        Ok(self.catalog.as_ref())
    }

    /// Replace or create the complete catalog after validating all bindings.
    pub fn set_catalog(&mut self, catalog: XmlMapInfo) -> Result<bool> {
        self.ensure_prepared()?;
        self.replace_catalog(Some(catalog))
    }

    /// Replace, create, or remove the complete catalog.
    pub fn replace_catalog(&mut self, catalog: Option<XmlMapInfo>) -> Result<bool> {
        self.ensure_prepared()?;
        validate_draft(
            catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
        )?;
        if self.catalog == catalog {
            return Ok(false);
        }
        self.catalog = catalog;
        Ok(true)
    }

    /// Remove the catalog only when no binary binding depends on it.
    pub fn remove_catalog(&mut self) -> Result<Option<XmlMapInfo>> {
        self.ensure_prepared()?;
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
        self.ensure_prepared()?;
        let name = name.into();
        let catalog = self
            .catalog
            .as_mut()
            .ok_or_else(|| invalid("cannot rename a map in an absent catalog"))?;
        let map = catalog
            .maps
            .iter_mut()
            .find(|map| map.id == map_id)
            .ok_or_else(|| invalid("XML map ID was not"))?;
        if map.name == name {
            return Ok(false);
        }
        let previous = std::mem::replace(&mut map.name, name);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
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
        self.ensure_prepared()?;
        let catalog = self
            .catalog
            .as_mut()
            .ok_or_else(|| invalid("cannot edit a map in an absent catalog"))?;
        let index = catalog
            .maps
            .iter()
            .position(|map| map.id == map_id)
            .ok_or_else(|| invalid("XML map ID was not"))?;
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
            self.limits,
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
        self.ensure_prepared()?;
        if self.conformance == conformance {
            return Ok(false);
        }
        let Some(catalog) = self.catalog.as_ref() else {
            return Ok(false);
        };
        serialize_catalog(catalog, conformance, self.limits)?;
        self.conformance = conformance;
        Ok(true)
    }

    /// Put one binding into a worksheet single-cell part.
    pub fn put_single_cell_binding(
        &mut self,
        sheet_index: usize,
        binding: SingleCellBinding,
    ) -> Result<bool> {
        self.ensure_prepared()?;
        let values = self
            .single_cell_bindings
            .get(sheet_index)
            .ok_or(Error::WorksheetNotFound(Some(sheet_index)))?;
        let mut values = values.as_ref().cloned().unwrap_or_default();
        if let Some(index) = values
            .iter()
            .position(|value| value.table_id() == binding.table_id())
        {
            if values[index] == binding {
                return Ok(false);
            }
            values[index] = binding;
        } else {
            values.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "single-cell transaction draft",
                source,
            })?;
            values.push(binding);
        }
        let previous = self.single_cell_bindings[sheet_index].replace(values);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
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
        self.ensure_prepared()?;
        let Some(values) = self
            .single_cell_bindings
            .get(sheet_index)
            .ok_or(Error::WorksheetNotFound(Some(sheet_index)))?
            .as_ref()
        else {
            return Ok(None);
        };
        let mut values = values.clone();
        let Some(index) = values.iter().position(|value| value.table_id() == table_id) else {
            return Ok(None);
        };
        let removed = values.remove(index);
        let next = if values.is_empty() {
            None
        } else {
            Some(values)
        };
        let previous = std::mem::replace(&mut self.single_cell_bindings[sheet_index], next);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
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
        self.ensure_prepared()?;
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
        let replacement = MappedTable::bounded(table_id, columns, self.limits.bindings)?;
        let source = self.source_table(table_id, column_id)?;
        patch_table_bindings(&source, &replacement, self.limits.bindings)?;
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
            self.limits,
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
        self.ensure_prepared()?;
        let table_index = self
            .mapped_tables
            .iter()
            .position(|table| table.table_id() == table_id)
            .ok_or_else(|| invalid("mapped table ID was not"))?;
        let mut columns = self.mapped_tables[table_index].columns().to_vec();
        let Some(column_index) = columns
            .iter()
            .position(|value| value.column_id() == column_id)
        else {
            return Ok(None);
        };
        let removed = columns.remove(column_index);
        let replacement = MappedTable::bounded(table_id, columns, self.limits.bindings)?;
        let previous = std::mem::replace(&mut self.mapped_tables[table_index], replacement);
        match validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
        ) {
            Ok(()) => Ok(Some(removed)),
            Err(error) => {
                self.mapped_tables[table_index] = previous;
                Err(error)
            },
        }
    }

    /// Build a reversible source-bound commit without touching a workbook.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let source = self.source.clone();
        let patch = self.plan()?;
        source.apply_xml_maps(&patch)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let before = self.snapshot()?.clone();
        self.catalog = before.map_info().cloned();
        self.conformance = before.conformance();
        self.mapped_tables = before.mapped_tables().to_vec();
        self.single_cell_bindings = before.binding_groups();
        Ok(())
    }

    fn plan(mut self) -> Result<Patch> {
        self.ensure_prepared()?;
        let before = self.snapshot()?.clone();
        validate_draft(
            self.catalog.as_ref(),
            &self.mapped_tables,
            &self.single_cell_bindings,
            self.conformance,
            self.limits,
        )?;
        if before.map_info() == self.catalog.as_ref()
            && before.conformance() == self.conformance
            && before.mapped_tables() == self.mapped_tables
            && before.binding_groups_equal(&self.single_cell_bindings)
        {
            return Ok(Patch::new(before.clone(), before.clone()));
        }

        let mut source = before.source().clone();
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
                parse_xml_map_info(part.bytes(), &self.limits.xml_maps)
                    .map_err(|_error| invalid("the Custom XML Maps part is not valid"))
            })
            .transpose()?;
        let projected_conformance = if projected_catalog.is_some() {
            self.conformance
        } else {
            XmlMapConformance::Transitional
        };
        let after = Snapshot::materialized(
            &before,
            projected_catalog,
            projected_conformance,
            projected_tables,
            self.single_cell_bindings,
            source,
        )?;
        Ok(Patch::new(before, after))
    }

    fn materialize_catalog(&self, source: &mut SourceState) -> Result<()> {
        let content_type = litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE;
        let existing = source
            .dependencies
            .iter()
            .position(|part| part.content_type == content_type);
        match (existing, self.snapshot()?.map_info(), self.catalog.as_ref()) {
            (Some(index), Some(before), Some(after)) => {
                let xml = litchi_ooxml_common::spreadsheet_xml_maps::patch_xml_map_info_source(
                    source.dependencies[index].bytes(),
                    before,
                    after,
                    self.snapshot()?.conformance(),
                    self.conformance,
                    &self.limits.xml_maps,
                )?;
                parse_xml_map_info(&xml, &self.limits.xml_maps)
                    .map_err(|_error| invalid("the Custom XML Maps part is not valid"))?;
                source.dependencies[index].bytes = Arc::new(xml);
                let map_name = source.dependencies[index].part_name.clone();
                if self.snapshot()?.conformance() != self.conformance {
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
                let xml = serialize_catalog(after, self.conformance, self.limits)?;
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
            .snapshot()?
            .source()
            .dependencies
            .iter()
            .filter(|part| part.content_type == TABLE_CONTENT_TYPE)
        {
            let host = crate::package::table::parse_table_part(part.bytes(), self.limits.raw)?;
            if host.id != table_id {
                continue;
            }
            if host.table_type != crate::package::table::Type::Xml || host.single_cell {
                return Err(invalid("table ID is not an ordinary LTXML Table"));
            }
            if !host.columns.iter().any(|column| column.id == column_id) {
                return Err(invalid("column is absent from physical table ID"));
            }
            let parsed = match super::parse_table_bindings(part.bytes(), self.limits.bindings) {
                Ok(parsed) => parsed,
                Err(error) => return Err(error),
            };
            if parsed.value().table_id() != host.id {
                return Err(invalid("host Table ID disagrees with XML binding table ID"));
            }
            if found.is_some() {
                return Err(invalid(
                    "mapped table ID is ambiguous across LTXML Table parts",
                ));
            }
            found = Some(parsed);
        }
        found
            .ok_or_else(|| invalid("mapped table ID was not found in an existing LTXML Table part"))
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
            let parsed = match super::parse_table_bindings(part.bytes(), self.limits.bindings) {
                Ok(parsed) => parsed,
                Err(Error::Unrecognized {
                    typ: "mapped table type",
                    ..
                }) => continue,
                Err(error) => return Err(error),
            };
            let Some(table) = tables.get(&parsed.value().table_id()).copied() else {
                continue;
            };
            part.bytes = Arc::new(patch_table_bindings(&parsed, table, self.limits.bindings)?);
        }
        Ok(())
    }

    fn materialize_single_cells(&self, source: &mut SourceState) -> Result<()> {
        let mut owners: Vec<Option<(usize, String)>> = Vec::new();
        owners
            .try_reserve(source.worksheets.len())
            .map_err(|source| Error::Allocation {
                resource: "tableSingleCells owner index",
                source,
            })?;
        owners.resize(source.worksheets.len(), None);
        for worksheet in &source.worksheets {
            for relationship in &worksheet.relationships {
                if relationship.relationship_type == SINGLE_CELLS_REL
                    && relationship.mode == TargetMode::Internal
                {
                    PackURI::from_rel_ref(worksheet.part_name.base_uri(), &relationship.target)
                        .map_err(|_error| {
                            invalid("a single-cell relationship target is not a valid part name")
                        })?;
                }
            }
        }
        for (dependency_index, part) in source.dependencies.iter().enumerate() {
            if part.content_type != SINGLE_CELLS_CONTENT_TYPE {
                continue;
            }
            let mut owner = None;
            for (sheet_index, worksheet) in source.worksheets.iter().enumerate() {
                if let Some(relationship) = worksheet.relationships.iter().find(|relationship| {
                    source_relationship_targets(
                        worksheet.part_name.base_uri(),
                        relationship,
                        &part.part_name,
                    )
                }) {
                    if owner
                        .replace((sheet_index, relationship.id.clone()))
                        .is_some()
                    {
                        return Err(invalid(
                            "tableSingleCells part is referenced by multiple worksheets",
                        ));
                    }
                }
            }
            let Some((target_sheet, relationship_id)) = owner else {
                return Err(invalid("tableSingleCells owner worksheet is absent"));
            };
            if owners[target_sheet]
                .replace((dependency_index, relationship_id))
                .is_some()
            {
                return Err(invalid(
                    "tableSingleCells part is referenced by multiple worksheets",
                ));
            }
        }

        for (sheet_index, owner) in owners.iter().enumerate() {
            let Some((dependency_index, _)) = owner else {
                continue;
            };
            let Some(values) = self
                .single_cell_bindings
                .get(sheet_index)
                .and_then(Option::as_deref)
            else {
                continue;
            };
            let part = &mut source.dependencies[*dependency_index];
            let parsed = parse_single_cells(part.bytes(), self.limits.bindings)?;
            part.bytes = Arc::new(patch_single_cells(&parsed, values, self.limits.bindings)?);
        }

        let mut removals = Vec::new();
        removals
            .try_reserve(source.worksheets.len())
            .map_err(|source| Error::Allocation {
                resource: "tableSingleCells removal plan",
                source,
            })?;
        for (sheet_index, owner) in owners.iter().enumerate() {
            let Some((dependency_index, _)) = owner else {
                continue;
            };
            if self
                .single_cell_bindings
                .get(sheet_index)
                .and_then(Option::as_ref)
                .is_some()
            {
                continue;
            }
            let part = &source.dependencies[*dependency_index];
            let parsed = parse_single_cells(part.bytes(), self.limits.bindings)?;
            parsed.ensure_removal_lossless()?;
            patch_single_cells(&parsed, &[], self.limits.bindings)?;
            let relationship_id = owners[sheet_index]
                .as_ref()
                .map(|(_, id)| id.clone())
                .ok_or_else(|| invalid("tableSingleCells owner relationship is absent"))?;
            removals.push((sheet_index, relationship_id, part.part_name.clone()));
        }

        for (sheet_index, relationship_id, _) in &removals {
            source.worksheets[*sheet_index]
                .relationships
                .retain(|relationship| relationship.id != *relationship_id);
        }

        for (sheet_index, values) in self.single_cell_bindings.iter().enumerate() {
            if owners.get(sheet_index).and_then(Option::as_ref).is_some() {
                continue;
            }
            let Some(values) = values.as_deref() else {
                continue;
            };
            if values.is_empty() {
                continue;
            }
            let part_name = next_single_cells_part_name(source)?;
            let bytes = serialize_single_cells(values, self.limits.bindings)?;
            let relationship_id =
                next_relationship_id(&source.worksheets[sheet_index].relationships);
            let target =
                part_name.relative_ref(source.worksheets[sheet_index].part_name.base_uri());
            source
                .dependencies
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "tableSingleCells dependency",
                    source,
                })?;
            source
                .part_names
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "tableSingleCells part-name index",
                    source,
                })?;
            source.worksheets[sheet_index]
                .relationships
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "tableSingleCells worksheet relationship",
                    source,
                })?;
            source.dependencies.push(SourcePart {
                part_name: part_name.clone(),
                content_type: SINGLE_CELLS_CONTENT_TYPE.to_string(),
                bytes: Arc::new(bytes),
                relationships: Vec::new(),
            });
            source.part_names.push(part_name);
            source.worksheets[sheet_index]
                .relationships
                .push(SourceRelationship {
                    id: relationship_id,
                    relationship_type: SINGLE_CELLS_REL.to_string(),
                    target,
                    mode: TargetMode::Internal,
                });
        }

        for (_, _, part_name) in removals {
            source
                .dependencies
                .retain(|part| part.part_name != part_name);
            source.part_names.retain(|name| name != &part_name);
        }
        source
            .dependencies
            .sort_by(|left, right| left.part_name.as_str().cmp(right.part_name.as_str()));
        source
            .part_names
            .sort_by(|left, right| left.as_str().cmp(right.as_str()));
        for worksheet in &mut source.worksheets {
            worksheet
                .relationships
                .sort_by(|left, right| left.id.cmp(&right.id));
        }
        Ok(())
    }
}

fn validate_draft(
    catalog: Option<&XmlMapInfo>,
    mapped_tables: &[MappedTable],
    groups: &[Option<Vec<SingleCellBinding>>],
    conformance: XmlMapConformance,
    limits: super::SemanticLimits,
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
        return Err(invalid("aggregate XML binding count exceeds"));
    }
    if total_xpath_units > limits.max_total_xpath_units {
        return Err(invalid("aggregate XML binding XPath units exceeds"));
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
            return Err(invalid("single-cell binding references absent XML map ID"));
        }
        if !table_ids.insert(binding.table_id()) {
            return Err(invalid("duplicate XML mapping table ID"));
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
        if !source
            .part_names
            .iter()
            .any(|name| name.is_equivalent_to(&uri))
        {
            return Ok(uri);
        }
    }
    Err(invalid("could not allocate a bounded MapInfo part name"))
}

fn next_single_cells_part_name(source: &SourceState) -> Result<PackURI> {
    for index in 1usize..=source.part_names.len().saturating_add(1) {
        let uri = PackURI::new(format!("/xl/tables/tableSingleCells{index}.bin"))?;
        if !source
            .part_names
            .iter()
            .any(|name| name.is_equivalent_to(&uri))
        {
            return Ok(uri);
        }
    }
    Err(invalid(
        "could not allocate a bounded tableSingleCells part name",
    ))
}

fn source_relationship_targets(
    base_uri: &str,
    relationship: &SourceRelationship,
    part_name: &PackURI,
) -> bool {
    relationship.relationship_type == SINGLE_CELLS_REL
        && relationship.mode == TargetMode::Internal
        && PackURI::from_rel_ref(base_uri, &relationship.target)
            .ok()
            .is_some_and(|target| target.is_equivalent_to(part_name))
}

fn serialize_catalog(
    catalog: &XmlMapInfo,
    conformance: XmlMapConformance,
    limits: super::SemanticLimits,
) -> Result<Vec<u8>> {
    litchi_ooxml_common::spreadsheet_xml_maps::serialize_xml_map_info(
        catalog,
        conformance,
        &limits.xml_maps,
    )
    .map_err(Error::Common)
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

fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
fn _cell_reference_is_public(_: CellReference) {}

#[cfg(test)]
mod transaction_tests {
    use super::*;
    use crate::xml_maps::{DataBinding, SemanticLimits, XPath, XmlDataType, XmlSchema};

    #[test]
    fn absent_catalog_conformance_is_an_exact_signed_no_op() {
        let snapshot = Snapshot::empty_transaction_fixture(true);
        let before = snapshot.source().clone();
        let mut transaction = Edit::from_snapshot(
            crate::Package::from(litchi_opc::OpcPackage::new()),
            snapshot,
        );
        assert!(
            !transaction
                .set_conformance(XmlMapConformance::Strict)
                .expect("absent catalog conformance is inert")
        );
        let patch = transaction.plan().expect("plan no-op");
        assert!(patch.is_noop());
        assert_eq!(patch.before().source(), &before);
        assert_eq!(patch.after().source(), &before);
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
        let mut exact = SemanticLimits::DEFAULT;
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
            SemanticLimits::DEFAULT,
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
        let limits = before.limits();
        let first = SingleCellBinding::new(
            1,
            1,
            CellReference::new(0, 0).expect("cell"),
            7,
            XmlDataType::new(1).expect("type"),
            XPath::new("/r").expect("XPath"),
        )
        .expect("binding");
        let mut transaction = Edit {
            source: crate::Package::from(litchi_opc::OpcPackage::new()),
            before: Some(before),
            limits,
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
        let mut transaction = Edit::from_snapshot(
            crate::Package::from(litchi_opc::OpcPackage::new()),
            snapshot,
        );
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
        assert_eq!(
            transaction.before.as_ref().expect("prepared").source(),
            &source
        );
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
