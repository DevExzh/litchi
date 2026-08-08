//! Table Data Extraction from TST Protobuf Messages
//!
//! This module provides utilities for extracting cell data from Numbers table structures.
//! Numbers stores table data in a complex format using Tiles, TableDataList, and Cell messages.
//!
//! ## Architecture
//!
//! - **TableModelArchive**: Contains table metadata and references to data stores
//! - **DataStore**: Contains references to various data tables (strings, formulas, styles)
//! - **TableDataList**: Maps keys to actual cell content (strings, formulas, formats)
//! - **TileStorage**: Contains the actual cells in a sparse tile-based structure
//! - **Tile**: Contains rows of cells with their values
//!
//! ## Public boundary
//!
//! `TableDataExtractor`, `Components`, and `Index` are implementation details
//! of [`crate::Package`]. Applications should parse a native package through
//! [`crate::Package`] and consume its archive-free [`crate::Document`] result;
//! this decoder is intentionally not part of the public API.

use super::Components;
use super::table::Table;
use super::{
    Error, Result, SemanticLimitKind, SemanticLimits, SemanticPath, TABLE_MODEL_MESSAGE_TYPE,
    table_info_decode_options,
};
use super::{Index, Resolved};
use crate::DEFAULT_MAX_TEXT_BYTES;
use crate::cell::FiniteF64;
use crate::cell::Value as CellValue;
use crate::cell::wire::{BncCellView, CachedScalar, StoredValue};
use litchi_iwa_common::comment::{AuthorId, Comment, StorageId, Uuid};
use litchi_iwa_common::wire::{WireDescent, preflight_wire_tree_with_limits};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::group_node_category_codec::{self, CategoryValueView, GroupNodeView};
use litchi_iwa_protos::table_info_codec;
use litchi_iwa_protos::{tn, tsce, tsd, tst};
use prost::Message;
use std::borrow::Cow;
use std::cell::{Ref, RefCell};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

type CompactTable<T> = Box<[(u32, T)]>;
type StringTable = CompactTable<String>;
type FormulaTable = CompactTable<tsce::FormulaArchive>;
type FormulaErrorTable = CompactTable<String>;
type CommentTable = CompactTable<Comment>;
type FormulaOwnerKey = [u32; 4];
type FormulaCategoryKey = [u64; 2];

const TILE_MESSAGE_TYPE: u32 = 6_002;
const MAX_TABLE_ROWS: usize = 1 << 20;
const MAX_TABLE_COLUMNS: usize = 1 << 14;
const MAX_ADDRESSABLE_CELLS: usize = 1 << 24;
const MAX_TABLE_MATERIALIZED_CELLS: usize = 1 << 20;
const MAX_FORMULA_CATEGORY_DEPTH: usize = 64;
const MAX_FORMULA_WORK: usize = crate::MAX_REFERENCES;
const MAX_FORMULA_WIRE_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;

struct CellTables<'a> {
    strings: &'a StringTable,
    formulas: &'a FormulaTable,
    formula_errors: &'a FormulaErrorTable,
    rich_text: &'a StringTable,
    comments: &'a CommentTable,
    formula_references: &'a FormulaReferenceMaps,
}

struct ParsedCell {
    value: CellValue,
    comment_identifier: Option<u32>,
}

#[derive(Debug)]
struct CellBudget {
    remaining: usize,
}

impl CellBudget {
    fn new() -> Self {
        Self {
            remaining: MAX_TABLE_MATERIALIZED_CELLS,
        }
    }

    fn check(&self, requested: usize) -> Result<()> {
        if requested > self.remaining {
            return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::MaterializedCells,
                observed: MAX_TABLE_MATERIALIZED_CELLS
                    .saturating_sub(self.remaining)
                    .saturating_add(requested),
                limit: MAX_TABLE_MATERIALIZED_CELLS,
            }));
        }
        Ok(())
    }

    fn consume(&mut self, materialized: usize) -> Result<()> {
        self.check(materialized)?;
        self.remaining -= materialized;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct ProjectionBudget {
    materialized_cells: usize,
    output_text_bytes: usize,
    formula_render_work: usize,
    max_materialized_cells: usize,
    max_output_text_bytes: usize,
    max_formula_render_work: usize,
    max_formula_render_depth: usize,
}

impl ProjectionBudget {
    const fn new(limits: SemanticLimits) -> Self {
        Self {
            materialized_cells: 0,
            output_text_bytes: 0,
            formula_render_work: 0,
            max_materialized_cells: limits.max_materialized_cells(),
            max_output_text_bytes: limits.max_output_text_bytes(),
            max_formula_render_work: limits.max_formula_render_work(),
            max_formula_render_depth: limits.max_formula_render_depth(),
        }
    }

    fn charge_materialized_cells(&mut self, amount: usize) -> Result<()> {
        self.materialized_cells = projection_charge(
            self.materialized_cells,
            amount,
            self.max_materialized_cells,
            SemanticLimitKind::MaterializedCells,
        )?;
        Ok(())
    }

    fn check_output_text(&self, amount: usize) -> Result<()> {
        projection_charge(
            self.output_text_bytes,
            amount,
            self.max_output_text_bytes,
            SemanticLimitKind::OutputTextBytes,
        )?;
        Ok(())
    }

    fn charge_output_text(&mut self, amount: usize) -> Result<()> {
        self.output_text_bytes = projection_charge(
            self.output_text_bytes,
            amount,
            self.max_output_text_bytes,
            SemanticLimitKind::OutputTextBytes,
        )?;
        Ok(())
    }

    fn charge_formula_render_work(&mut self, amount: usize) -> Result<()> {
        let charged = projection_charge(
            self.formula_render_work,
            amount,
            self.max_formula_render_work,
            SemanticLimitKind::FormulaRenderWork,
        );
        match charged {
            Ok(observed) => self.formula_render_work = observed,
            Err(error) => {
                // Work already performed by a rejected candidate cannot be
                // reclaimed. Saturating the counter makes repeated hostile
                // candidates fail before receiving a fresh allowance.
                self.formula_render_work = self.max_formula_render_work;
                return Err(error);
            },
        }
        Ok(())
    }

    fn commit_attempt(&mut self, candidate: Self, published: bool) {
        if published {
            *self = candidate;
        } else {
            // Retained cells and text are transactional, but CPU work is a
            // package-wide admission cost even when the candidate is rejected.
            self.formula_render_work = self.formula_render_work.max(candidate.formula_render_work);
        }
    }

    fn check_formula_render_depth(&self, depth: usize) -> Result<()> {
        if depth > self.max_formula_render_depth {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderDepth,
                observed: depth,
                maximum: self.max_formula_render_depth,
                path: SemanticPath::StructuredTables,
            });
        }
        Ok(())
    }
}

fn projection_charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: SemanticLimitKind,
) -> Result<usize> {
    let observed = current.checked_add(amount).ok_or(Error::SemanticLimit {
        kind,
        observed: usize::MAX,
        maximum,
        path: SemanticPath::StructuredTables,
    })?;
    if observed > maximum {
        return Err(Error::SemanticLimit {
            kind,
            observed,
            maximum,
            path: SemanticPath::StructuredTables,
        });
    }
    Ok(observed)
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
}

fn table_limit_error(observed: usize, maximum: usize) -> Error {
    Error::SemanticLimit {
        kind: SemanticLimitKind::Tables,
        observed,
        maximum,
        path: SemanticPath::StructuredTables,
    }
}

fn compact_table<T>(entries: impl IntoIterator<Item = (u32, T)>) -> Result<CompactTable<T>> {
    let mut compacted = Vec::new();
    let entries = entries.into_iter();
    let (lower_bound, _) = entries.size_hint();
    compacted
        .try_reserve(lower_bound)
        .map_err(|_| allocation_error("Numbers table sidecar entries", lower_bound))?;
    for entry in entries {
        compacted
            .try_reserve(1)
            .map_err(|_| allocation_error("Numbers table sidecar entries", compacted.len() + 1))?;
        compacted.push(entry);
    }
    compact_table_vec(compacted)
}

fn compact_table_vec<T>(mut compacted: Vec<(u32, T)>) -> Result<CompactTable<T>> {
    compacted.sort_unstable_by_key(|(key, _)| *key);
    if compacted.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::InvalidFormat(
            "Numbers table sidecar contains duplicate keys".to_owned(),
        ));
    }
    Ok(compacted.into_boxed_slice())
}

fn compact_table_get<T>(table: &[(u32, T)], key: u32) -> Option<&T> {
    table
        .binary_search_by_key(&key, |(entry_key, _)| *entry_key)
        .ok()
        .map(|index| &table[index].1)
}

fn retain_text(value: &str, budget: &mut ProjectionBudget) -> Result<String> {
    budget.charge_output_text(value.len())?;
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_| allocation_error("Numbers retained semantic text", value.len()))?;
    retained.push_str(value);
    Ok(retained)
}

fn retained_table_text(
    table: &[(u32, String)],
    identifier: u32,
    budget: &mut ProjectionBudget,
) -> Result<Option<String>> {
    compact_table_get(table, identifier)
        .map(|value| retain_text(value, budget))
        .transpose()
}

fn checked_table_dimensions(row_count: u32, column_count: u32) -> Result<(usize, usize)> {
    let row_count = usize::try_from(row_count).map_err(|_| {
        Error::InvalidFormat("Numbers table row count does not fit the host usize".to_owned())
    })?;
    let column_count = usize::try_from(column_count).map_err(|_| {
        Error::InvalidFormat("Numbers table column count does not fit the host usize".to_owned())
    })?;

    if row_count > MAX_TABLE_ROWS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableRows,
            observed: row_count,
            limit: MAX_TABLE_ROWS,
        }));
    }
    if column_count > MAX_TABLE_COLUMNS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableColumns,
            observed: column_count,
            limit: MAX_TABLE_COLUMNS,
        }));
    }

    let addressable_cells = row_count.checked_mul(column_count).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table dimensions overflow host address space: {row_count}x{column_count}"
        ))
    })?;
    if addressable_cells > MAX_ADDRESSABLE_CELLS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::TableCells,
            observed: addressable_cells,
            limit: MAX_ADDRESSABLE_CELLS,
        }));
    }

    Ok((row_count, column_count))
}

fn validate_table_row(row: usize, row_count: usize) -> Result<()> {
    if row >= row_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers tile row {row} is outside the declared table height {row_count}"
        )));
    }
    Ok(())
}

fn validate_table_column(column: usize, column_count: usize) -> Result<()> {
    if column >= column_count {
        return Err(Error::InvalidFormat(format!(
            "Numbers cell column {column} is outside the declared table width {column_count}"
        )));
    }
    Ok(())
}

fn validate_table_data_list_segment(
    object_id: u64,
    list_type: tst::table_data_list::ListType,
    segment: &tst::TableDataListSegment,
) -> Result<()> {
    if segment.list_type != list_type as i32 {
        return Err(Error::InvalidFormat(format!(
            "Numbers table-data-list segment {object_id} has list type {}, expected {list_type:?}",
            segment.list_type
        )));
    }
    let end = segment
        .key_range
        .location
        .checked_add(segment.key_range.length)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment {object_id} key range overflows"
            ))
        })?;
    if segment
        .entries
        .iter()
        .any(|entry| entry.key < segment.key_range.location || entry.key >= end)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers table-data-list segment {object_id} contains an entry outside its key range"
        )));
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct FormulaReferenceName {
    sheet: Arc<str>,
    table: Arc<str>,
}

#[derive(Debug, Clone, Default)]
struct FormulaReferenceMaps {
    owners: HashMap<FormulaOwnerKey, FormulaReferenceName>,
    categories: HashMap<FormulaCategoryKey, String>,
}

/// Extractor for Numbers table data
pub(super) struct TableDataExtractor<'a> {
    bundle: &'a Components,
    object_index: &'a Index,
    formula_references: RefCell<Option<FormulaReferenceMaps>>,
    projection_budget: RefCell<ProjectionBudget>,
    max_formula_references: usize,
}

impl<'a> TableDataExtractor<'a> {
    /// Return whether the index contains a candidate table-model object.
    ///
    /// This cheap type probe lets generic structured extraction avoid building
    /// formula-reference sidecars for Pages and Keynote packages.
    pub(super) fn has_table_models(object_index: &Index) -> bool {
        [TABLE_MODEL_MESSAGE_TYPE, 6_000]
            .into_iter()
            .any(|message_type| {
                object_index
                    .iter_entries_by_type(message_type)
                    .next()
                    .is_some()
            })
    }

    /// Create a new table data extractor
    pub(super) fn new(
        bundle: &'a Components,
        object_index: &'a Index,
        limits: SemanticLimits,
    ) -> Self {
        Self {
            bundle,
            object_index,
            formula_references: RefCell::new(None),
            projection_budget: RefCell::new(ProjectionBudget::new(limits)),
            max_formula_references: limits.max_references(),
        }
    }

    fn formula_references(&self) -> Result<Ref<'_, FormulaReferenceMaps>> {
        if self.formula_references.borrow().is_none() {
            let references = build_formula_reference_maps(
                self.bundle,
                self.object_index,
                self.max_formula_references,
            )?;
            *self.formula_references.borrow_mut() = Some(references);
        }
        Ref::filter_map(self.formula_references.borrow(), Option::as_ref).map_err(|_references| {
            Error::InvalidFormat("Numbers formula-reference cache was not initialized".to_owned())
        })
    }

    /// Charge semantic text retained outside table projection, such as rooted
    /// sheet names, against the same package-wide output budget.
    pub(super) fn charge_output_text(&self, amount: usize) -> Result<()> {
        self.projection_budget
            .borrow_mut()
            .charge_output_text(amount)
    }

    /// Extract all tables from the document
    pub(super) fn extract_all_tables(&self) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        self.for_each_table(usize::MAX, |table| {
            tables.try_reserve(1).map_err(|_| {
                allocation_error("Numbers extracted table results", tables.len() + 1)
            })?;
            tables.push(table);
            Ok(())
        })?;
        Ok(tables)
    }

    /// Extract all tables directly into the canonical Numbers semantic model.
    ///
    /// The archive adapter's builder is consumed one table at a time. Its
    /// sparse cell and header buffers move into the leaf table, so this path
    /// avoids first allocating a `Vec<Table>` only to convert every
    /// element into a second result vector for structured extraction.
    pub(super) fn extract_all_semantic_tables(
        &self,
        max_tables: usize,
    ) -> Result<Vec<crate::Table>> {
        let mut tables = Vec::new();
        self.for_each_table(max_tables, |table| {
            tables.try_reserve(1).map_err(|_| {
                allocation_error("Numbers semantic table results", tables.len() + 1)
            })?;
            tables.push(table.into_semantic_table()?);
            Ok(())
        })?;
        Ok(tables)
    }

    fn for_each_table(
        &self,
        max_tables: usize,
        mut visit: impl FnMut(Table) -> Result<()>,
    ) -> Result<()> {
        let mut seen_objects = HashSet::new();
        let mut table_count = 0usize;

        // Real packages index TableModelArchive as 6001. Older generated
        // fixtures may store the same payload under 6000, so the object
        // adapter accepts 6000 only when its payload passes model extraction;
        // a genuine TableInfoArchive is ignored rather than mis-decoded.
        for message_type in [TABLE_MODEL_MESSAGE_TYPE, 6_000] {
            for entry in self.object_index.iter_entries_by_type(message_type) {
                if seen_objects.contains(&entry.id()) {
                    continue;
                }
                seen_objects.try_reserve(1).map_err(|_error| {
                    allocation_error(
                        "Numbers structured table identities",
                        seen_objects.len() + 1,
                    )
                })?;
                seen_objects.insert(entry.id());
                // Candidate admission is deliberately checked before protobuf
                // decoding. Once the caller-selected table budget is full, a
                // later malformed canonical candidate cannot force another
                // potentially large model allocation merely to choose an error.
                if message_type == TABLE_MODEL_MESSAGE_TYPE && table_count >= max_tables {
                    return Err(table_limit_error(table_count.saturating_add(1), max_tables));
                }
                if let Some(resolved) = self.object_index.resolve_ref(self.bundle, entry.id())?
                    && let Some(table) = self.extract_table_candidate(&resolved, message_type)?
                {
                    table_count = table_count
                        .checked_add(1)
                        .ok_or_else(|| table_limit_error(usize::MAX, max_tables))?;
                    if table_count > max_tables {
                        return Err(table_limit_error(table_count, max_tables));
                    }
                    visit(table)?;
                }
            }
        }
        Ok(())
    }

    /// Extract a single table from a resolved object
    fn extract_table_candidate(
        &self,
        object: &Resolved<'_>,
        candidate_type: u32,
    ) -> Result<Option<Table>> {
        if candidate_type == TABLE_MODEL_MESSAGE_TYPE {
            let mut messages = object
                .messages
                .iter()
                .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
            let Some(message) = messages.next() else {
                return Err(Error::InvalidFormat(
                    "Numbers canonical table candidate has no canonical payload".to_owned(),
                ));
            };
            if messages.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers canonical table candidate has duplicate canonical payloads".to_owned(),
                ));
            }
            let table_model = tst::TableModelArchive::decode(&*message.data).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers table-model message {} is malformed: {error}",
                    message.type_
                ))
            })?;
            return self.parse_table_model(table_model).map(Some);
        }

        // Protobuf is permissive, and legacy fixtures used 6000 for a model.
        // Decode only the primary candidate payload. A secondary canonical
        // payload must not promote an object classified as legacy metadata.
        let Some(message) = object
            .messages
            .iter()
            .next()
            .filter(|message| message.type_ == 6_000)
        else {
            return Err(Error::InvalidFormat(
                "Numbers legacy table candidate has no primary legacy payload".to_owned(),
            ));
        };
        if object
            .messages
            .iter()
            .skip(1)
            .any(|candidate| candidate.type_ == 6_000)
        {
            return Err(Error::InvalidFormat(
                "Numbers legacy table candidate has duplicate legacy payloads".to_owned(),
            ));
        }
        let Ok(table_model) = tst::TableModelArchive::decode(&*message.data) else {
            return Ok(None);
        };
        if let Ok(table) = self.parse_table_model(table_model) {
            return Ok(Some(table));
        }

        Ok(None)
    }

    /// Extract a table model reached through a schema-proven `TableInfo` edge.
    ///
    /// Rooted ownership is stricter than the global compatibility scan: one
    /// canonical type-6001 payload is preferred, the legacy type-6000 payload
    /// is accepted only when 6001 is absent, and duplicate candidates fail.
    pub(super) fn extract_reachable_table_from_object(
        &self,
        object: &Resolved<'_>,
        path: SemanticPath,
    ) -> Result<Table> {
        let mut typed = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
        if let Some(message) = typed.next() {
            if typed.next().is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table model contains duplicate canonical payloads"
                )));
            }
            let table_model =
                tst::TableModelArchive::decode(message.data.as_slice()).map_err(|error| {
                    Error::InvalidFormat(format!(
                        "Numbers {path} table-model payload is malformed: {error}"
                    ))
                })?;
            return self.parse_table_model(table_model);
        }

        let mut legacy = object
            .messages
            .iter()
            .filter(|message| message.type_ == 6_000);
        let Some(message) = legacy.next() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model has no recognized payload"
            )));
        };
        if legacy.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model contains duplicate legacy payloads"
            )));
        }
        let table_model =
            tst::TableModelArchive::decode(message.data.as_slice()).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers {path} legacy table-model payload is malformed: {error}"
                ))
            })?;
        self.parse_table_model(table_model)
    }

    /// Parse a TableModelArchive protobuf message
    fn parse_table_model(&self, table_model: tst::TableModelArchive) -> Result<Table> {
        // Projection is transactional at the table boundary. A rejected legacy
        // candidate must not consume retained-cell or retained-text capacity
        // that belongs to a later schema-proven table. Formula work remains a
        // monotonic package-wide cost across successful and rejected attempts.
        let mut projection_budget = *self.projection_budget.borrow();
        let result = (|| {
            let (row_count, column_count) = checked_table_dimensions(
                table_model.number_of_rows,
                table_model.number_of_columns,
            )?;
            projection_budget.charge_output_text(table_model.table_name.len())?;
            let mut table =
                Table::with_dimensions(table_model.table_name, row_count, column_count)?;

            // Extract string table for cell text values
            // string_table is a required field, not Optional
            let string_table =
                self.load_string_table(table_model.base_data_store.string_table.identifier)?;

            // Extract formula table for formula cells
            // formula_table is a required field, not Optional
            let formula_table =
                self.load_formula_table(table_model.base_data_store.formula_table.identifier)?;
            let formula_references = if formula_table.is_empty() {
                None
            } else {
                Some(self.formula_references()?)
            };
            let empty_formula_references = FormulaReferenceMaps::default();
            let formula_error_table = match table_model.base_data_store.formula_error_table {
                Some(reference) => self.load_formula_error_table(reference.identifier)?,
                None => Box::default(),
            };

            let rich_text_table = match table_model.base_data_store.rich_text_table {
                Some(reference) => self.load_rich_text_table(reference.identifier)?,
                None => Box::default(),
            };
            let comment_table = match table_model.base_data_store.comment_storage_table {
                Some(reference) => self.load_comment_table(reference.identifier)?,
                None => Box::default(),
            };

            // Parse tiles to extract cell data
            let cell_tables = CellTables {
                strings: &string_table,
                formulas: &formula_table,
                formula_errors: &formula_error_table,
                rich_text: &rich_text_table,
                comments: &comment_table,
                formula_references: formula_references
                    .as_deref()
                    .unwrap_or(&empty_formula_references),
            };
            self.parse_tiles(
                &table_model.base_data_store.tiles,
                &cell_tables,
                &mut projection_budget,
                &mut table,
            )?;

            Ok(table)
        })();

        self.projection_budget
            .borrow_mut()
            .commit_attempt(projection_budget, result.is_ok());
        result
    }

    /// Load a TableDataList from an object reference
    fn load_string_table(&self, object_id: u64) -> Result<StringTable> {
        compact_table(
            self.load_table_data_list_entries(object_id, tst::table_data_list::ListType::String)?
                .into_iter()
                .filter_map(|entry| entry.string.map(|value| (entry.key, value))),
        )
    }

    fn load_formula_table(&self, object_id: u64) -> Result<FormulaTable> {
        compact_table(
            self.load_table_data_list_entries(object_id, tst::table_data_list::ListType::Formula)?
                .into_iter()
                .filter_map(|entry| entry.formula.map(|value| (entry.key, value))),
        )
    }

    fn load_formula_error_table(&self, object_id: u64) -> Result<FormulaErrorTable> {
        compact_table(
            self.load_table_data_list_entries(
                object_id,
                tst::table_data_list::ListType::FormulaError,
            )?
            .into_iter()
            .filter_map(|entry| entry.string.map(|value| (entry.key, value))),
        )
    }

    fn load_rich_text_table(&self, object_id: u64) -> Result<StringTable> {
        let mut result = Vec::new();

        for entry in self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::RichTextPayload,
        )? {
            let Some(payload_reference) = entry.rich_text_payload else {
                continue;
            };
            let Some(payload_object) = self
                .object_index
                .resolve_ref_id(self.bundle, payload_reference.identifier)?
            else {
                continue;
            };
            for payload_message in payload_object.messages {
                let Ok(payload) =
                    tst::RichTextPayloadArchive::decode(payload_message.data.as_slice())
                else {
                    continue;
                };
                if let Some(text) = self.extract_rich_text(payload.storage.identifier)? {
                    result.try_reserve(1).map_err(|_| {
                        allocation_error("Numbers rich-text sidecar", result.len() + 1)
                    })?;
                    result.push((entry.key, text));
                    break;
                }
            }
        }

        compact_table_vec(result)
    }

    fn load_comment_table(&self, object_id: u64) -> Result<CommentTable> {
        let mut result = Vec::new();
        for entry in self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::CommentStorage,
        )? {
            if entry.refcount == 0 {
                return Err(Error::InvalidFormat(format!(
                    "Numbers comment entry {} has a zero reference count",
                    entry.key
                )));
            }
            let storage_id = entry
                .comment_storage
                .as_ref()
                .map(|reference| reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers comment entry {} has no storage reference",
                        entry.key
                    ))
                })?;
            let storage_object = self
                .object_index
                .resolve_ref_id(self.bundle, storage_id)?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers comment storage object {storage_id} is missing"
                    ))
                })?;
            let comments = storage_object
                .messages
                .iter()
                .filter(|message| message.type_ == 3056)
                .map(|message| tsd::CommentStorageArchive::decode(message.data.as_slice()))
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(Error::protobuf)?;
            let comment = comments.first().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {storage_id} has no TSD comment-storage payload"
                ))
            })?;
            if comments.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Object {storage_id} has multiple TSD comment-storage payloads"
                )));
            }
            result
                .try_reserve(1)
                .map_err(|_| allocation_error("Numbers comment sidecar", result.len() + 1))?;
            result.push((
                entry.key,
                Comment {
                    text: comment.text.clone().unwrap_or_default(),
                    creation_date_seconds: comment.creation_date.as_ref().map(|date| date.seconds),
                    author_id: comment
                        .author
                        .as_ref()
                        .map(|author| {
                            AuthorId::from_raw(author.identifier).map_err(map_comment_error)
                        })
                        .transpose()?,
                    reply_ids: comment
                        .replies
                        .iter()
                        .map(|reply| {
                            StorageId::from_raw(reply.identifier).map_err(map_comment_error)
                        })
                        .collect::<Result<Vec<_>>>()?
                        .into_boxed_slice(),
                    storage_uuid: comment
                        .storage_uuid
                        .as_ref()
                        .map(|uuid| {
                            Uuid::from_parts(uuid.lower, uuid.upper).map_err(map_comment_error)
                        })
                        .transpose()?,
                },
            ));
        }
        compact_table_vec(result)
    }

    fn load_table_data_list_entries(
        &self,
        object_id: u64,
        list_type: tst::table_data_list::ListType,
    ) -> Result<Vec<tst::table_data_list::ListEntry>> {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table-data-list object {object_id} is missing"
                ))
            })?;
        let lists = resolved
            .messages
            .iter()
            .filter(|message| message.type_ == 6005 || message.type_ == 6201)
            .map(|message| tst::TableDataList::decode(message.data.as_slice()))
            .collect::<std::result::Result<Vec<_>, _>>()
            .map_err(Error::protobuf)?;
        let mut matching = lists
            .into_iter()
            .filter(|list| list.list_type == list_type as i32);
        let list = matching.next().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {object_id} has no Numbers {list_type:?} TableDataList payload"
            ))
        })?;
        if matching.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Object {object_id} has multiple Numbers {list_type:?} TableDataList payloads"
            )));
        }

        let mut entries = list.entries;
        let mut keys = entries
            .iter()
            .map(|entry| entry.key)
            .collect::<HashSet<_>>();
        if keys.len() != entries.len() {
            return Err(Error::InvalidFormat(format!(
                "Numbers {list_type:?} table {object_id} contains duplicate root entry keys"
            )));
        }
        let mut segment_ids = HashSet::new();
        for reference in list.segments {
            if !segment_ids.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {list_type:?} table {object_id} repeats segment object {}",
                    reference.identifier
                )));
            }
            let segment_object = self
                .object_index
                .resolve_ref_id(self.bundle, reference.identifier)?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table-data-list segment object {} is missing",
                        reference.identifier
                    ))
                })?;
            let segment_messages = segment_object
                .messages
                .iter()
                .filter(|message| message.type_ == 6011)
                .collect::<Vec<_>>();
            let segment_message = segment_messages.first().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers TableDataListSegment payload",
                    reference.identifier
                ))
            })?;
            if segment_messages.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Object {} has multiple Numbers TableDataListSegment payloads",
                    reference.identifier
                )));
            }
            let segment = tst::TableDataListSegment::decode(segment_message.data.as_slice())
                .map_err(Error::protobuf)?;
            validate_table_data_list_segment(reference.identifier, list_type, &segment)?;
            for entry in segment.entries {
                if !keys.insert(entry.key) {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers {list_type:?} table {object_id} repeats entry key {} across root and segments",
                        entry.key
                    )));
                }
                entries.push(entry);
            }
        }
        Ok(entries)
    }

    /// Parse tile storage to extract cells
    fn parse_tiles(
        &self,
        tile_storage: &tst::TileStorage,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let tile_size = usize::try_from(tile_storage.tile_size.unwrap_or(256)).map_err(|_| {
            Error::InvalidFormat("Numbers tile size does not fit the host usize".to_owned())
        })?;
        if tile_size == 0 {
            return Err(Error::InvalidFormat(
                "Numbers table declares a zero tile size".to_owned(),
            ));
        }
        let tile_count = if table.row_count() == 0 {
            0
        } else {
            (table.row_count() - 1) / tile_size + 1
        };
        let mut seen_tile_ids = HashSet::new();
        seen_tile_ids
            .try_reserve(tile_storage.tiles.len())
            .map_err(|_| allocation_error("Numbers tile keys", tile_storage.tiles.len()))?;
        let mut budget = CellBudget::new();
        // Resolve each tile reference and parse its contents
        for tile_ref in &tile_storage.tiles {
            let tile_key = usize::try_from(tile_ref.tileid).map_err(|_| {
                Error::InvalidFormat("Numbers tile key does not fit the host usize".to_owned())
            })?;
            if tile_key >= tile_count {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile key {tile_key} is outside the declared table height {}",
                    table.row_count()
                )));
            }
            if !seen_tile_ids.insert(tile_ref.tileid) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table repeats tile key {tile_key}"
                )));
            }
            let row_origin = tile_key
                .checked_mul(tile_size)
                .ok_or_else(|| Error::ParseError("Numbers tile row origin overflow".to_owned()))?;
            // tile is a required field, not Optional
            let tile_reference = &tile_ref.tile;
            self.parse_tile(
                tile_reference.identifier,
                row_origin,
                tile_size,
                table.row_count(),
                table.column_count(),
                &mut budget,
                cell_tables,
                projection_budget,
                table,
            )?;
        }

        Ok(())
    }

    /// Parse a single tile object
    fn parse_tile(
        &self,
        tile_id: u64,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, tile_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} referenced by table is missing"
                ))
            })?;
        let mut decoded = false;
        for msg in resolved.messages {
            if msg.type_ != TILE_MESSAGE_TYPE {
                continue;
            }
            if decoded {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} contains multiple tile payloads"
                )));
            }
            let tile = tst::Tile::decode(&*msg.data).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} has a malformed tile payload: {error}"
                ))
            })?;
            self.parse_tile_rows(
                &tile,
                row_origin,
                tile_size,
                row_count,
                column_count,
                budget,
                cell_tables,
                projection_budget,
                table,
            )?;
            decoded = true;
        }

        if !decoded {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile object {tile_id} has no tile payload"
            )));
        }

        Ok(())
    }

    /// Parse rows within a tile
    fn parse_tile_rows(
        &self,
        tile: &tst::Tile,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        for row_info in &tile.row_infos {
            self.parse_tile_row(
                row_info,
                row_origin,
                tile_size,
                row_count,
                column_count,
                budget,
                cell_tables,
                projection_budget,
                table,
            )?;
        }

        Ok(())
    }

    /// Parse a single tile row
    fn parse_tile_row(
        &self,
        row_info: &tst::TileRowInfo,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
    ) -> Result<()> {
        let tile_row_index = usize::try_from(row_info.tile_row_index).map_err(|_| {
            Error::InvalidFormat("Numbers tile row index does not fit the host usize".to_owned())
        })?;
        if tile_row_index >= tile_size {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile row {} is outside tile size {tile_size}",
                row_info.tile_row_index
            )));
        }
        let row_index = row_origin
            .checked_add(tile_row_index)
            .ok_or_else(|| Error::ParseError("Numbers tile row index overflow".to_owned()))?;
        validate_table_row(row_index, row_count)?;

        // The cell_storage_buffer contains serialized Cell messages
        // The cell_offsets buffer contains the byte offsets for each cell

        let (cell_storage, cell_offsets) = match (
            row_info.cell_storage_buffer.as_deref(),
            row_info.cell_offsets.as_deref(),
        ) {
            (Some(storage), Some(offsets)) => (storage, offsets),
            _ => (
                row_info.cell_storage_buffer_pre_bnc.as_slice(),
                row_info.cell_offsets_pre_bnc.as_slice(),
            ),
        };

        let expected_cells = usize::try_from(row_info.cell_count).map_err(|_| {
            Error::InvalidFormat("Numbers cell count does not fit the host usize".to_owned())
        })?;
        budget.check(expected_cells)?;
        projection_budget.charge_materialized_cells(expected_cells)?;
        let cells = Self::parse_cell_offsets(
            cell_offsets,
            cell_storage.len(),
            row_info.has_wide_offsets.unwrap_or(false),
            expected_cells,
            column_count,
        )?;
        budget.consume(cells.len())?;

        for (column_index, range) in cells {
            validate_table_column(column_index, column_count)?;
            let parsed = Self::parse_cell_storage(
                &cell_storage[range],
                cell_tables,
                projection_budget,
                row_index,
                column_index,
            )?;
            table.try_set_cell(row_index, column_index, parsed.value)?;
            if let Some(identifier) = parsed.comment_identifier {
                let comment = compact_table_get(cell_tables.comments, identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers comment table has no entry {identifier} referenced by cell ({row_index}, {column_index})"
                    ))
                })?;
                table.try_set_comment(row_index, column_index, comment.clone())?;
            }
        }

        Ok(())
    }

    /// Parse cell offsets from the offsets buffer
    ///
    /// The offset table is an array of little-endian `u16` values. `0xffff`
    /// marks a missing column; wide rows store offsets in four-byte units.
    /// Native producers may pad the table past the semantic table width, but
    /// every padded slot must retain the missing-column sentinel.
    fn parse_cell_offsets(
        offsets_buffer: &[u8],
        storage_length: usize,
        wide_offsets: bool,
        expected_cells: usize,
        column_count: usize,
    ) -> Result<Vec<(usize, std::ops::Range<usize>)>> {
        if !offsets_buffer.len().is_multiple_of(2) {
            return Err(Error::ParseError(
                "Numbers cell offset table has an odd byte length".to_string(),
            ));
        }

        let slot_count = offsets_buffer.len() / 2;
        if expected_cells > slot_count {
            return Err(Error::ParseError(format!(
                "Numbers row declares {expected_cells} cells but has only {slot_count} offset slots"
            )));
        }
        if expected_cells > column_count {
            return Err(Error::InvalidFormat(format!(
                "Numbers row declares {expected_cells} cells but table width is {column_count}"
            )));
        }
        if let Some((column, _bytes)) = offsets_buffer
            .chunks_exact(2)
            .enumerate()
            .skip(column_count)
            .find(|(_column, bytes)| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers cell offset at column {column} is outside the declared table width {column_count}"
            )));
        }

        let present_cells = offsets_buffer
            .chunks_exact(2)
            .take(column_count)
            .filter(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
            .count();
        if present_cells != expected_cells {
            return Err(Error::ParseError(format!(
                "Numbers row declares {expected_cells} cells but has {present_cells} offsets"
            )));
        }

        let width = if wide_offsets { 4usize } else { 1usize };
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(expected_cells)
            .map_err(|_| allocation_error("Numbers cell ranges", expected_cells))?;
        let mut previous = None;
        for (column, bytes) in offsets_buffer
            .chunks_exact(2)
            .take(column_count)
            .enumerate()
        {
            let raw_offset = u16::from_le_bytes([bytes[0], bytes[1]]);
            if raw_offset == u16::MAX {
                continue;
            }
            let offset = usize::from(raw_offset)
                .checked_mul(width)
                .ok_or_else(|| Error::ParseError("Numbers cell offset overflow".to_string()))?;
            if offset >= storage_length {
                return Err(Error::ParseError(format!(
                    "Numbers cell offset {offset} exceeds storage length {storage_length}"
                )));
            }
            if let Some((previous_column, previous_offset)) = previous {
                if offset <= previous_offset {
                    return Err(Error::ParseError(format!(
                        "Numbers cell offsets are not strictly increasing: {previous_offset} then {offset}"
                    )));
                }
                cells.push((previous_column, previous_offset..offset));
            }
            previous = Some((column, offset));
        }
        if let Some((column, start)) = previous {
            if storage_length <= start {
                return Err(Error::ParseError(format!(
                    "Numbers cell offset range ends at {storage_length} after {start}"
                )));
            }
            cells.push((column, start..storage_length));
        }
        Ok(cells)
    }

    fn parse_cell_storage(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
    ) -> Result<ParsedCell> {
        let version = *data
            .first()
            .ok_or_else(|| Error::ParseError("Empty Numbers cell storage".to_string()))?;
        match version {
            0..=4 => Self::parse_pre_bnc_cell(data, cell_tables, projection_budget, row, column),
            5 => Self::parse_bnc_cell(data, cell_tables, projection_budget, row, column),
            other => Err(Error::ParseError(format!(
                "Unsupported Numbers cell storage version {other}"
            ))),
        }
    }

    fn parse_bnc_cell(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
    ) -> Result<ParsedCell> {
        let cell = BncCellView::parse(data).map_err(|error| {
            Error::ParseError(format!(
                "Numbers BNC cell ({row}, {column}) is invalid: {error}"
            ))
        })?;
        let comment_identifier = cell.comment_identifier();

        if let StoredValue::Formula(identifier) = cell.stored_value() {
            let formula = compact_table_get(cell_tables.formulas, identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula table has no entry {identifier} referenced by cell ({row}, {column})"
                ))
            })?;
            let rendered = Self::extract_formula_string(
                formula,
                row,
                column,
                cell_tables.formula_references,
                projection_budget,
            )
            .map_err(|error| {
                Error::ParseError(format!(
                    "Numbers formula {identifier} at cell ({row}, {column}) is invalid: {error}"
                ))
            })?;
            return Ok(ParsedCell {
                value: CellValue::Formula(rendered),
                comment_identifier,
            });
        }

        let zero = finite_zero()?;
        let scalar = cell.cached_scalar();
        let value = match cell.stored_value() {
            StoredValue::Empty => CellValue::Empty,
            StoredValue::Number => match scalar {
                Some(CachedScalar::Number(value)) => CellValue::Number(value),
                Some(
                    CachedScalar::Boolean(_) | CachedScalar::Date(_) | CachedScalar::Duration(_),
                ) => {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers numeric BNC cell ({row}, {column}) has a mismatched scalar encoding"
                    )));
                },
                Some(CachedScalar::Unsupported(_)) | None => CellValue::Number(zero),
            },
            StoredValue::Text(identifier) => {
                retained_table_text(cell_tables.strings, identifier, projection_budget)?
                    .map_or(CellValue::Empty, CellValue::Text)
            },
            StoredValue::RichText(identifier) => {
                retained_table_text(cell_tables.rich_text, identifier, projection_budget)?
                    .map_or(CellValue::Empty, CellValue::Text)
            },
            StoredValue::Date => match scalar {
                Some(CachedScalar::Date(value)) => CellValue::Date(value),
                Some(_) | None => CellValue::Date(zero),
            },
            StoredValue::Boolean => match scalar {
                Some(CachedScalar::Boolean(value)) => CellValue::Boolean(value),
                Some(_) | None => CellValue::Boolean(false),
            },
            StoredValue::Duration => match scalar {
                Some(CachedScalar::Duration(value)) => CellValue::Duration(value),
                Some(_) | None => CellValue::Duration(zero),
            },
            StoredValue::Error => {
                let error = cell
                    .formula_error_identifier()
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id))
                    .map_or("FORMULA", String::as_str);
                CellValue::Error(retain_text(error, projection_budget)?)
            },
            StoredValue::Formula(_) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers formula BNC cell ({row}, {column}) reached scalar decoding"
                )));
            },
            StoredValue::Unsupported(other) => {
                return Err(Error::ParseError(format!(
                    "Unsupported Numbers BNC cell type {other}"
                )));
            },
        };
        Ok(ParsedCell {
            value,
            comment_identifier,
        })
    }

    fn parse_pre_bnc_cell(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        row: usize,
        column: usize,
    ) -> Result<ParsedCell> {
        let version = data[0];
        let header_length = if version <= 1 { 8 } else { 12 };
        if data.len() < header_length {
            return Err(Error::ParseError(
                "Truncated Numbers pre-BNC cell header".to_string(),
            ));
        }
        let cell_type = data[if version == 4 { 1 } else { 2 }];
        let flags = if version <= 1 {
            u32::from(u16::from_le_bytes([data[4], data[5]]))
        } else {
            read_u32_le(&data[4..8])?
        };
        let mut cursor = header_length;
        let mut number: Option<FiniteF64> = None;
        let mut date: Option<FiniteF64> = None;
        let mut string_id = None;
        let mut rich_text_id = None;
        let mut formula_id = None;
        let mut formula_error_id = None;
        let mut comment_identifier = None;

        for (flag, size) in [
            (0x000002, 4),
            (0x000080, 4),
            (0x000400, 4),
            (0x000800, 4),
            (0x000004, 4),
            (0x000008, 4),
            (0x000100, 4),
            (0x000200, 4),
            (0x001000, 4),
            (0x002000, 4),
            (0x000010, 4),
            (0x000020, 8),
            (0x000040, 8),
            (0x010000, 4),
            (0x080000, 4),
            (0x020000, 4),
            (0x040000, 4),
            (0x100000, 4),
            (0x200000, 4),
            (0x400000, 4),
            (0x800000, 4),
        ] {
            if flags & flag == 0 {
                continue;
            }
            let field = take_field(data, &mut cursor, size)?;
            match flag {
                0x000008 => formula_id = Some(read_u32_le(field)?),
                0x000100 => formula_error_id = Some(read_u32_le(field)?),
                0x001000 => comment_identifier = Some(read_u32_le(field)?),
                0x000200 => rich_text_id = Some(read_u32_le(field)?),
                0x000010 => string_id = Some(read_u32_le(field)?),
                0x000020 => number = Some(read_f64_le(field)?),
                0x000040 => date = Some(read_f64_le(field)?),
                _ => {},
            }
        }

        if let Some(identifier) = formula_id {
            let formula = compact_table_get(cell_tables.formulas, identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula table has no entry {identifier} referenced by cell ({row}, {column})"
                ))
            })?;
            let rendered = Self::extract_formula_string(
                formula,
                row,
                column,
                cell_tables.formula_references,
                projection_budget,
            )
            .map_err(|error| {
                Error::ParseError(format!(
                    "Numbers formula {identifier} at cell ({row}, {column}) is invalid: {error}"
                ))
            })?;
            return Ok(ParsedCell {
                value: CellValue::Formula(rendered),
                comment_identifier,
            });
        }

        let zero = finite_zero()?;
        let value = match cell_type {
            0 => CellValue::Empty,
            2 => CellValue::Number(number.unwrap_or(zero)),
            3 => match string_id {
                Some(identifier) => {
                    retained_table_text(cell_tables.strings, identifier, projection_budget)?
                        .map_or(CellValue::Empty, CellValue::Text)
                },
                None => CellValue::Empty,
            },
            5 => CellValue::Date(date.unwrap_or(zero)),
            6 => CellValue::Boolean(number.unwrap_or(zero).get() != 0.0),
            7 => CellValue::Duration(number.unwrap_or(zero)),
            8 => {
                let error = formula_error_id
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id))
                    .map_or("FORMULA", String::as_str);
                CellValue::Error(retain_text(error, projection_budget)?)
            },
            9 => match rich_text_id {
                Some(identifier) => {
                    retained_table_text(cell_tables.rich_text, identifier, projection_budget)?
                        .map_or(CellValue::Empty, CellValue::Text)
                },
                None => CellValue::Empty,
            },
            other => {
                return Err(Error::ParseError(format!(
                    "Unsupported Numbers pre-BNC cell type {other}"
                )));
            },
        };
        Ok(ParsedCell {
            value,
            comment_identifier,
        })
    }

    /// Render a formula through the bounded, non-copying expression arena.
    fn extract_formula_string(
        formula: &tsce::FormulaArchive,
        host_row: usize,
        host_column: usize,
        formula_references: &FormulaReferenceMaps,
        projection_budget: &mut ProjectionBudget,
    ) -> Result<String> {
        render_formula(
            formula,
            host_row,
            host_column,
            formula_references,
            projection_budget,
        )
    }

    /// Test-only reference renderer retained for differential coverage while
    /// the streaming FormulaArchive reader is migrated independently.
    ///
    ///   - Reconstructs formula text from Abstract Syntax Tree
    ///   - Handles operators, functions, cell references, and constants
    ///   - Based on TSCE.ASTNodeArrayArchive protobuf structure
    ///   - Implements reverse-polish notation to infix conversion
    ///
    /// iWork stores formulas as Abstract Syntax Trees (AST) in reverse-polish
    /// notation (postfix). This function reconstructs the formula text by
    /// traversing the AST and converting it to standard infix notation.
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of AST nodes. Uses a stack-based algorithm
    /// for efficient conversion.
    #[cfg(test)]
    fn extract_formula_string_reference(
        formula: &tsce::FormulaArchive,
        host_row: usize,
        host_column: usize,
        formula_references: &FormulaReferenceMaps,
    ) -> Result<String> {
        use litchi_iwa_protos::tsce::ast_node_array_archive::AstNodeType;

        let ast_array = &formula.ast_node_array;

        // Formulas are stored in reverse-polish notation (postfix)
        // We need to convert to infix notation using a stack
        if ast_array.ast_node.is_empty() {
            return Ok("=".to_string());
        }

        // Stack to hold expression parts during reconstruction
        let mut expr_stack: Vec<String> = Vec::new();

        // Process each AST node
        for node in &ast_array.ast_node {
            let ast_node_type = node.ast_node_type();

            match ast_node_type {
                // Arithmetic operators (binary)
                AstNodeType::AdditionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "addition")?;
                    expr_stack.push(format!("({}+{})", left, right));
                },
                AstNodeType::SubtractionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "subtraction")?;
                    expr_stack.push(format!("({}-{})", left, right));
                },
                AstNodeType::MultiplicationNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "multiplication")?;
                    expr_stack.push(format!("({}*{})", left, right));
                },
                AstNodeType::DivisionNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "division")?;
                    expr_stack.push(format!("({}/{})", left, right));
                },
                AstNodeType::PowerNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "power")?;
                    expr_stack.push(format!("({}^{})", left, right));
                },
                AstNodeType::GreaterThanNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "greater than")?;
                    expr_stack.push(format!("({left}>{right})"));
                },
                AstNodeType::GreaterThanOrEqualToNode => {
                    let (left, right) =
                        pop_binary_operands(&mut expr_stack, "greater than or equal")?;
                    expr_stack.push(format!("({left}>={right})"));
                },
                AstNodeType::LessThanNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "less than")?;
                    expr_stack.push(format!("({left}<{right})"));
                },
                AstNodeType::LessThanOrEqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "less than or equal")?;
                    expr_stack.push(format!("({left}<={right})"));
                },
                AstNodeType::EqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "equality")?;
                    expr_stack.push(format!("({left}={right})"));
                },
                AstNodeType::NotEqualToNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "inequality")?;
                    expr_stack.push(format!("({left}<>{right})"));
                },

                // Constants
                AstNodeType::NumberNode => {
                    if let Some(number) = node.ast_number_node_number {
                        expr_stack.push(number.to_string());
                    }
                },
                AstNodeType::StringNode => {
                    if let Some(ref string) = node.ast_string_node_string {
                        expr_stack.push(format!("\"{}\"", string.replace('"', "\"\"")));
                    }
                },
                AstNodeType::BooleanNode => {
                    if let Some(boolean) = node.ast_boolean_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_string());
                    }
                },
                AstNodeType::TokenNode => {
                    if let Some(boolean) = node.ast_token_node_boolean {
                        expr_stack.push(if boolean { "TRUE" } else { "FALSE" }.to_owned());
                    }
                },
                AstNodeType::DateNode => {
                    if let Some(seconds) = node.ast_date_node_date_num {
                        expr_stack.push(format!("(DATE(2001,1,1)+{})", seconds / 86_400.0));
                    }
                },
                AstNodeType::DurationNode => {
                    if let Some(value) = node.ast_duration_node_unit_num {
                        expr_stack.push(value.to_string());
                    }
                },
                AstNodeType::EmptyArgumentNode => expr_stack.push(String::new()),

                // Cell references
                AstNodeType::CellReferenceNode => {
                    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
                        let column = resolve_formula_coordinate(
                            host_column,
                            ast_column.column,
                            ast_column.absolute.unwrap_or(false),
                            "column",
                        )?;
                        let row = resolve_formula_coordinate(
                            host_row,
                            ast_row.row,
                            ast_row.absolute.unwrap_or(false),
                            "row",
                        )?;
                        let column_absolute = ast_column.absolute.unwrap_or(false);
                        let row_absolute = ast_row.absolute.unwrap_or(false);
                        let prefix = node
                            .ast_cross_table_reference_extra_info
                            .as_ref()
                            .map(|extra| {
                                formula_reference_prefix(&extra.table_id, formula_references)
                            })
                            .unwrap_or_default();
                        expr_stack.push(format!(
                            "{prefix}{}{}{}{}",
                            if column_absolute { "$" } else { "" },
                            Self::column_index_to_letter(column),
                            if row_absolute { "$" } else { "" },
                            row + 1
                        ));
                    } else if let Some(ref cell_ref) = node.ast_local_cell_reference_node_reference
                    {
                        // Convert row/column handles to A1 notation
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        let row_num = cell_ref.row_handle + 1; // 0-based to 1-based
                        let col_sticky = if cell_ref.column_is_sticky != 0 {
                            "$"
                        } else {
                            ""
                        };
                        let row_sticky = if cell_ref.row_is_sticky != 0 { "$" } else { "" };
                        expr_stack.push(format!(
                            "{}{}{}{}",
                            col_sticky, col_letter, row_sticky, row_num
                        ));
                    } else if let Some(ref cross_ref) =
                        node.ast_cross_table_cell_reference_node_reference
                    {
                        // Cross-table reference
                        let col_letter = Self::column_index_to_letter(cross_ref.column_handle);
                        let row_num = cross_ref.row_handle + 1;
                        let prefix =
                            formula_reference_prefix(&cross_ref.table_id, formula_references);
                        expr_stack.push(format!("{prefix}{col_letter}{row_num}"));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },
                AstNodeType::LocalCellReferenceNode => {
                    if let Some(cell_ref) = &node.ast_local_cell_reference_node_reference {
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        expr_stack.push(format!("{}{}", col_letter, cell_ref.row_handle + 1));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },
                AstNodeType::CrossTableCellReferenceNode => {
                    if let Some(cell_ref) = &node.ast_cross_table_cell_reference_node_reference {
                        let col_letter = Self::column_index_to_letter(cell_ref.column_handle);
                        let prefix =
                            formula_reference_prefix(&cell_ref.table_id, formula_references);
                        expr_stack.push(format!("{prefix}{col_letter}{}", cell_ref.row_handle + 1));
                    } else {
                        expr_stack.push("#REF!".to_owned());
                    }
                },

                // Functions
                AstNodeType::FunctionNode => {
                    if let Some(function_index) = node.ast_function_node_index {
                        let num_args = node.ast_function_node_num_args.unwrap_or(0);
                        let function_name = Self::get_function_name(function_index);

                        // Pop arguments from stack (in reverse order)
                        let args = pop_formula_arguments(&mut expr_stack, num_args, "function")?;

                        let args_str = args.join(",");
                        expr_stack.push(format!("{}({})", function_name, args_str));
                    }
                },

                // List (for function arguments)
                AstNodeType::ListNode => {
                    if let Some(num_args) = node.ast_list_node_num_args {
                        // Collect arguments
                        let args = pop_formula_arguments(&mut expr_stack, num_args, "list")?;
                        expr_stack.push(args.join(","));
                    }
                },
                AstNodeType::ArrayNode => {
                    let columns = node.ast_array_node_num_col.unwrap_or(0);
                    let rows = node.ast_array_node_num_row.unwrap_or(0);
                    let count = columns.checked_mul(rows).ok_or_else(|| {
                        Error::ParseError("Numbers formula array size overflow".to_owned())
                    })?;
                    let values = pop_formula_arguments(&mut expr_stack, count, "array")?;
                    let columns = usize::try_from(columns).map_err(|_| {
                        Error::ParseError("Numbers formula array width exceeds usize".to_owned())
                    })?;
                    let rendered = if columns == 0 {
                        String::new()
                    } else {
                        values
                            .chunks(columns)
                            .map(|row| row.join(","))
                            .collect::<Vec<_>>()
                            .join(";")
                    };
                    expr_stack.push(format!("{{{rendered}}}"));
                },
                AstNodeType::ThunkNode => {
                    if let Some(array) = &node.ast_thunk_node_array {
                        let nested = tsce::FormulaArchive {
                            ast_node_array: array.clone(),
                            ..Default::default()
                        };
                        let rendered = Self::extract_formula_string_reference(
                            &nested,
                            host_row,
                            host_column,
                            formula_references,
                        )?;
                        expr_stack.push(rendered.trim_start_matches('=').to_owned());
                    }
                },

                // Unary operators - represented differently in the AST
                // Numbers uses NegationNode instead of UnaryMinusNode
                AstNodeType::NegationNode => {
                    if let Some(operand) = expr_stack.pop() {
                        expr_stack.push(format!("-({})", operand));
                    }
                },
                AstNodeType::PercentNode => {
                    let operand = expr_stack.pop().ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula percent operator is missing an operand".to_owned(),
                        )
                    })?;
                    expr_stack.push(format!("({operand})%"));
                },

                // Concatenation
                AstNodeType::ConcatenationNode => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "concatenation")?;
                    expr_stack.push(format!("({}&{})", left, right));
                },
                AstNodeType::ColonNode | AstNodeType::ColonNodeWithUids => {
                    let (left, right) = pop_binary_operands(&mut expr_stack, "range")?;
                    expr_stack.push(format!("{left}:{right}"));
                },
                AstNodeType::ColonTractNode => {
                    expr_stack.push(render_colon_tract(
                        node,
                        host_row,
                        host_column,
                        formula_references,
                    )?);
                },
                AstNodeType::ReferenceErrorNode | AstNodeType::ReferenceErrorWithUids => {
                    expr_stack.push("#REF!".to_owned());
                },
                AstNodeType::CategoryRefNode => {
                    expr_stack.push(render_category_reference(node, formula_references));
                },
                AstNodeType::UnknownFunctionNode => {
                    let count = node.ast_unknown_function_node_num_args.unwrap_or(0);
                    let arguments =
                        pop_formula_arguments(&mut expr_stack, count, "unknown function")?;
                    let name = node
                        .ast_unknown_function_node_string
                        .as_deref()
                        .unwrap_or("UNKNOWN");
                    expr_stack.push(format!("{name}({})", arguments.join(",")));
                },
                AstNodeType::PlusSignNode
                | AstNodeType::BeginThunkNode
                | AstNodeType::EndThunkNode
                | AstNodeType::AppendWhitespaceNode
                | AstNodeType::PrependWhitespaceNode => {},

                // Other node types - handle gracefully
                _ => {
                    // Unknown or special node types - keep processing
                    // (e.g., whitespace nodes, thunk nodes, etc.)
                },
            }
        }

        // The final result should be on top of the stack
        let result = expr_stack
            .pop()
            .map_or_else(|| "=FORMULA()".to_string(), |value| format!("={value}"));

        Ok(result)
    }

    /// Convert column index to Excel-style letter (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
    fn column_index_to_letter(index: u32) -> String {
        let mut result = String::new();
        let mut idx = index;

        loop {
            let remainder = idx % 26;
            result.insert(0, (b'A' + remainder as u8) as char);
            if idx < 26 {
                break;
            }
            idx = idx / 26 - 1;
        }

        result
    }

    /// Get function name from function index
    /// Based on Numbers built-in function list
    fn get_function_name(index: u32) -> String {
        super::function_map::function_name(index)
            .map(str::to_owned)
            .unwrap_or_else(|| format!("FUNC{index}"))
    }

    /// Extract rich text from a storage reference
    fn extract_rich_text(&self, storage_id: u64) -> Result<Option<String>> {
        if let Some(resolved) = self.object_index.resolve_ref_id(self.bundle, storage_id)? {
            // Look for TSWP.StorageArchive messages
            for msg in resolved.messages {
                if msg.type_ >= 2001
                    && msg.type_ <= 2022
                    && let Ok(storage) = litchi_iwa_protos::tswp::StorageArchive::decode(&*msg.data)
                    && !storage.text.is_empty()
                {
                    return Ok(Some(storage.text.join("\n")));
                }
            }
        }

        Ok(None)
    }
}

fn map_comment_error(_error: litchi_iwa_common::comment::Error) -> Error {
    Error::InvalidFormat("Numbers comment metadata is invalid".to_owned())
}

#[derive(Debug)]
struct FormulaReferenceBudget {
    retained_entries: usize,
    maximum_retained_entries: usize,
    work_items: usize,
    wire_bytes: usize,
    text_bytes: usize,
}

impl FormulaReferenceBudget {
    const fn new(maximum_retained_entries: usize) -> Self {
        Self {
            retained_entries: 0,
            maximum_retained_entries,
            work_items: 0,
            wire_bytes: 0,
            text_bytes: 0,
        }
    }

    fn charge_retained_entry(&mut self) -> Result<()> {
        self.retained_entries = self.retained_entries.checked_add(1).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::References,
                usize::MAX,
                self.maximum_retained_entries,
            )
        })?;
        if self.retained_entries > self.maximum_retained_entries {
            return Err(formula_semantic_limit(
                SemanticLimitKind::References,
                self.retained_entries,
                self.maximum_retained_entries,
            ));
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<()> {
        self.ensure_work_capacity(amount)?;
        self.work_items += amount;
        Ok(())
    }

    fn ensure_work_capacity(&self, additional: usize) -> Result<()> {
        let observed = self.work_items.checked_add(additional).ok_or_else(|| {
            formula_semantic_limit(SemanticLimitKind::FormulaWork, usize::MAX, MAX_FORMULA_WORK)
        })?;
        if observed > MAX_FORMULA_WORK {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                observed,
                MAX_FORMULA_WORK,
            ));
        }
        Ok(())
    }

    fn charge_wire_bytes(&mut self, bytes: usize) -> Result<()> {
        self.wire_bytes = self.wire_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                usize::MAX,
                MAX_FORMULA_WIRE_BYTES,
            )
        })?;
        if self.wire_bytes > MAX_FORMULA_WIRE_BYTES {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                self.wire_bytes,
                MAX_FORMULA_WIRE_BYTES,
            ));
        }
        Ok(())
    }

    fn charge_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                usize::MAX,
                DEFAULT_MAX_TEXT_BYTES,
            )
        })?;
        if self.text_bytes > DEFAULT_MAX_TEXT_BYTES {
            return Err(formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                self.text_bytes,
                DEFAULT_MAX_TEXT_BYTES,
            ));
        }
        Ok(())
    }
}

fn formula_semantic_limit(kind: SemanticLimitKind, observed: usize, maximum: usize) -> Error {
    Error::SemanticLimit {
        kind,
        observed,
        maximum,
        path: SemanticPath::StructuredTables,
    }
}

fn build_formula_reference_maps(
    bundle: &Components,
    object_index: &Index,
    max_formula_references: usize,
) -> Result<FormulaReferenceMaps> {
    let mut budget = FormulaReferenceBudget::new(max_formula_references);
    let mut result = FormulaReferenceMaps::default();
    result
        .categories
        .try_reserve(1)
        .map_err(|_error| allocation_error("Numbers formula categories", 1))?;
    result.categories.insert([1, 0], "Grand Total".to_owned());
    let mut table_info_names = HashMap::<u64, FormulaReferenceName>::new();
    let root_archive = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .and_then(|object| {
            object
                .messages
                .iter()
                .filter(|message| message.type_ == 1)
                .find_map(|message| tn::DocumentArchive::decode(message.data.as_slice()).ok())
        });

    if let Some(root) = root_archive {
        for sheet_reference in root.sheets {
            budget.charge_work(1)?;
            let Some(sheet_object) =
                object_index.resolve_ref_id(bundle, sheet_reference.identifier)?
            else {
                continue;
            };
            let Some(sheet) =
                sheet_object
                    .messages
                    .iter()
                    .find_map(|message| match message.type_ {
                        2 => tn::SheetArchive::decode(message.data.as_slice()).ok(),
                        3 => tn::FormBasedSheetArchive::decode(message.data.as_slice())
                            .ok()
                            .map(|form| form.super_),
                        _ => None,
                    })
            else {
                continue;
            };
            let sheet_name = sheet.name;
            let mut cached_sheet_name = None::<Arc<str>>;
            for drawable in sheet.drawable_infos {
                budget.charge_work(1)?;
                let Some(drawable_object) =
                    object_index.resolve_ref_id(bundle, drawable.identifier)?
                else {
                    continue;
                };
                let table_name = formula_table_name(
                    bundle,
                    object_index,
                    drawable_object.messages,
                    &mut budget,
                )?;
                if let Some(table) = table_name {
                    let is_new = !table_info_names.contains_key(&drawable.identifier);
                    if is_new {
                        budget.charge_retained_entry()?;
                    }
                    budget.charge_text(table.len())?;
                    if is_new {
                        table_info_names.try_reserve(1).map_err(|_error| {
                            allocation_error(
                                "Numbers formula table names",
                                table_info_names.len() + 1,
                            )
                        })?;
                    }
                    let retained_sheet_name = if let Some(name) = &cached_sheet_name {
                        Arc::clone(name)
                    } else {
                        budget.charge_text(sheet_name.len())?;
                        let name = Arc::<str>::from(sheet_name.as_str());
                        cached_sheet_name = Some(Arc::clone(&name));
                        name
                    };
                    table_info_names.insert(
                        drawable.identifier,
                        FormulaReferenceName {
                            sheet: retained_sheet_name,
                            table: Arc::from(table),
                        },
                    );
                }
            }
        }
    }

    for (_, archive) in bundle.iter_archives() {
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == 6383 {
                    collect_formula_category_payload(
                        message.data.as_slice(),
                        &mut result.categories,
                        &mut budget,
                    )?;
                    continue;
                }
                if message.type_ != 4008 {
                    continue;
                }
                budget.charge_work(1)?;
                let Ok(owner) =
                    tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                else {
                    continue;
                };
                let Some(table_info) = owner.formula_owner.as_ref() else {
                    continue;
                };
                let Some(name) = table_info_names.get(&table_info.identifier) else {
                    continue;
                };
                let key = formula_owner_key(&owner.formula_owner_uid);
                if !result.owners.contains_key(&key) {
                    budget.charge_retained_entry()?;
                    result.owners.try_reserve(1).map_err(|_error| {
                        allocation_error("Numbers formula owners", result.owners.len() + 1)
                    })?;
                }
                result.owners.insert(key, name.clone());
            }
        }
    }
    Ok(result)
}

fn formula_table_name(
    bundle: &Components,
    object_index: &Index,
    messages: &[litchi_iwa_core::RawMessage],
    budget: &mut FormulaReferenceBudget,
) -> Result<Option<String>> {
    for table_info_message in messages {
        if table_info_message.type_ != 6_000 && table_info_message.type_ != 6_003 {
            continue;
        }
        budget.charge_work(1)?;
        let Ok(model_reference) = table_info_codec::decode_table_model_reference(
            table_info_message.data.as_slice(),
            table_info_decode_options(table_info_message.data.as_slice()),
        ) else {
            continue;
        };
        let Some(model_object) =
            object_index.resolve_ref_id(bundle, model_reference.identifier().get())?
        else {
            continue;
        };
        let mut canonical = model_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
        let model_message = if let Some(message) = canonical.next() {
            if canonical.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers formula table model has duplicate canonical payloads".to_owned(),
                ));
            }
            Some(message)
        } else {
            let mut legacy = model_object
                .messages
                .iter()
                .filter(|message| message.type_ == 6_000);
            let message = legacy.next();
            if legacy.next().is_some() {
                return Err(Error::InvalidFormat(
                    "Numbers formula table model has duplicate legacy payloads".to_owned(),
                ));
            }
            message
        };
        if let Some(name) = model_message
            .and_then(|message| tst::TableModelArchive::decode(message.data.as_slice()).ok())
            .map(|model| model.table_name)
        {
            return Ok(Some(name));
        }
    }
    Ok(None)
}

fn charge_formula_preflight_work(
    work: &mut usize,
    amount: usize,
    maximum: usize,
) -> litchi_iwa_common::Result<()> {
    *work = work
        .checked_add(amount)
        .ok_or(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: usize::MAX,
            limit: maximum,
        })?;
    if *work > maximum {
        return Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: *work,
            limit: maximum,
        });
    }
    Ok(())
}

fn formula_projection_wire_type(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    expected: u8,
) -> litchi_iwa_common::Result<()> {
    if field.wire_type() == expected {
        Ok(())
    } else {
        Err(litchi_iwa_common::Error::InvalidFormat(
            "formula category projection field has the wrong wire type".to_owned(),
        ))
    }
}

fn preflight_formula_uuid(
    source: &[u8],
    work: &mut usize,
    maximum_work: usize,
) -> litchi_iwa_common::Result<()> {
    charge_formula_preflight_work(work, 1, maximum_work)?;
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(maximum_work.clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        charge_formula_preflight_work(work, 1, maximum_work)?;
        let field = visit.field();
        if visit.path().is_empty() && matches!(field.number(), 1 | 2) {
            formula_projection_wire_type(field, 0)?;
        }
        Ok(WireDescent::Skip)
    })?;
    Ok(())
}

fn preflight_formula_cell_value(
    source: &[u8],
    work: &mut usize,
    maximum_work: usize,
) -> litchi_iwa_common::Result<()> {
    charge_formula_preflight_work(work, 1, maximum_work)?;
    let input_bytes = source
        .len()
        .saturating_mul(2)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let limits = WireLimits::default()
        .with_input_bytes(input_bytes)?
        .with_fields(maximum_work.clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        charge_formula_preflight_work(work, 1, maximum_work)?;
        let field = visit.field();
        if visit.path().is_empty() && matches!(field.number(), 2..=5) {
            formula_projection_wire_type(field, 2)?;
            charge_formula_preflight_work(work, 1, maximum_work)?;
            return Ok(WireDescent::Descend);
        }
        let expected_wire_type = match (visit.path(), field.number()) {
            ([2], 1) => Some(0),
            ([3 | 4], 1) => Some(1),
            ([5], 1) => Some(2),
            _ => None,
        };
        if let Some(expected) = expected_wire_type {
            formula_projection_wire_type(field, expected)?;
        }
        if visit.path() == [5]
            && field.number() == 1
            && std::str::from_utf8(field.payload()).is_err()
        {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "formula category projection string is not UTF-8".to_owned(),
            ));
        }
        Ok(WireDescent::Skip)
    })?;
    Ok(())
}

fn preflight_formula_category_payload(
    source: &[u8],
    budget: &mut FormulaReferenceBudget,
) -> Result<Option<usize>> {
    // Charge source bytes before inspecting their framing so a package cannot
    // multiply malformed-candidate scan work without consuming a hard budget.
    budget.charge_wire_bytes(source.len())?;
    let remaining_work = MAX_FORMULA_WORK.saturating_sub(budget.work_items);
    if remaining_work == 0 {
        return Err(formula_semantic_limit(
            SemanticLimitKind::FormulaWork,
            budget.work_items.saturating_add(1),
            MAX_FORMULA_WORK,
        ));
    }
    let input_bytes = source
        .len()
        .saturating_mul(MAX_FORMULA_CATEGORY_DEPTH.saturating_add(1))
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let fields = remaining_work.clamp(1, WireLimits::MAX_FIELDS);
    let limits = WireLimits::default()
        .with_input_bytes(input_bytes)?
        .with_fields(fields)?
        .with_nesting(MAX_FORMULA_CATEGORY_DEPTH)?;
    let mut group_nodes = 1usize;
    let mut projection_work = 1usize;
    let preflight = preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        charge_formula_preflight_work(&mut projection_work, 1, remaining_work)?;
        if !visit.path().iter().all(|path_field| *path_field == 3) {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "formula category topology preflight left the child path".to_owned(),
            ));
        }
        match field.number() {
            1 => {
                formula_projection_wire_type(field, 2)?;
                preflight_formula_uuid(field.payload(), &mut projection_work, remaining_work)?;
                Ok(WireDescent::Skip)
            },
            3 => {
                formula_projection_wire_type(field, 2)?;
                let observed_depth = visit.path().len().saturating_add(1);
                if observed_depth > MAX_FORMULA_CATEGORY_DEPTH {
                    return Err(litchi_iwa_common::Error::LimitExceeded {
                        kind: LimitKind::Nesting,
                        observed: observed_depth,
                        limit: MAX_FORMULA_CATEGORY_DEPTH,
                    });
                }
                group_nodes =
                    group_nodes
                        .checked_add(1)
                        .ok_or(litchi_iwa_common::Error::LimitExceeded {
                            kind: LimitKind::Fields,
                            observed: usize::MAX,
                            limit: MAX_FORMULA_WORK,
                        })?;
                charge_formula_preflight_work(&mut projection_work, 1, remaining_work)?;
                Ok(WireDescent::Descend)
            },
            7 => {
                formula_projection_wire_type(field, 2)?;
                preflight_formula_cell_value(
                    field.payload(),
                    &mut projection_work,
                    remaining_work,
                )?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    });
    match preflight {
        Ok(_report) => {
            budget.charge_work(projection_work)?;
            Ok(Some(group_nodes))
        },
        Err(litchi_iwa_common::Error::InvalidFormat(_)) => {
            budget.charge_work(projection_work)?;
            Ok(None)
        },
        Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed,
            ..
        }) => Err(formula_semantic_limit(
            SemanticLimitKind::FormulaDepth,
            observed,
            MAX_FORMULA_CATEGORY_DEPTH,
        )),
        Err(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed,
            ..
        }) => Err(formula_semantic_limit(
            SemanticLimitKind::FormulaWork,
            budget.work_items.saturating_add(observed),
            MAX_FORMULA_WORK,
        )),
        Err(error) => Err(Error::Common(error)),
    }
}

fn collect_formula_category_payload(
    source: &[u8],
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let Some(expected_nodes) = preflight_formula_category_payload(source, budget)? else {
        return Ok(());
    };
    let decode_options = group_node_category_codec::DecodeOptions::new(
        source.len().max(1),
        u32::try_from(MAX_FORMULA_CATEGORY_DEPTH + 3).unwrap_or(u32::MAX),
    );
    let Ok(group_node) = group_node_category_codec::decode_group_node(source, decode_options)
    else {
        return Ok(());
    };
    collect_formula_category_names_with_budget(&group_node, expected_nodes, names, budget)
}

fn collect_formula_category_names_with_budget(
    root_node: &GroupNodeView<'_>,
    expected_nodes: usize,
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let mut visited = 1usize;
    retain_formula_category_name(root_node, names, budget)?;
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_error| allocation_error("Numbers formula category traversal", 1))?;
    pending.push(root_node.children());
    while let Some(children) = pending.last_mut() {
        let Some(child_result) = children.next() else {
            pending.pop();
            continue;
        };
        let child_node = child_result.map_err(|_error| {
            Error::InvalidFormat(
                "Numbers formula category projection diverged from its wire preflight".to_owned(),
            )
        })?;
        visited = visited.checked_add(1).ok_or_else(|| {
            formula_semantic_limit(SemanticLimitKind::FormulaWork, usize::MAX, MAX_FORMULA_WORK)
        })?;
        if visited > expected_nodes {
            return Err(Error::InvalidFormat(
                "Numbers formula category projection exceeded its wire preflight".to_owned(),
            ));
        }
        retain_formula_category_name(&child_node, names, budget)?;
        pending.try_reserve(1).map_err(|_error| {
            allocation_error("Numbers formula category traversal", pending.len() + 1)
        })?;
        pending.push(child_node.children());
    }
    if visited != expected_nodes {
        return Err(Error::InvalidFormat(
            "Numbers formula category projection did not reach its preflighted nodes".to_owned(),
        ));
    }
    Ok(())
}

fn retain_formula_category_name(
    node: &GroupNodeView<'_>,
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut FormulaReferenceBudget,
) -> Result<()> {
    let key = node
        .group_uid()
        .map_err(formula_category_projection_error)?
        .map_or([0, 0], |uid| [uid.lower(), uid.upper()]);
    let Some(value) = node
        .category_value()
        .map_err(formula_category_projection_error)?
    else {
        return Ok(());
    };
    let Some(label) = group_cell_value_label(&value).map_err(formula_category_projection_error)?
    else {
        return Ok(());
    };
    let is_new = !names.contains_key(&key);
    if is_new {
        budget.charge_retained_entry()?;
        names
            .try_reserve(1)
            .map_err(|_error| allocation_error("Numbers formula categories", names.len() + 1))?;
    }
    budget.charge_text(label.len())?;
    names.insert(key, label.into_owned());
    Ok(())
}

fn group_cell_value_label<'source>(
    value: &CategoryValueView<'source>,
) -> std::result::Result<Option<Cow<'source, str>>, group_node_category_codec::DecodeError> {
    if let Some(string) = value.string()? {
        return Ok(Some(Cow::Borrowed(string)));
    }
    if let Some(number) = value.number()? {
        return Ok(Some(Cow::Owned(number.to_string())));
    }
    if let Some(boolean) = value.boolean()? {
        return Ok(Some(Cow::Borrowed(if boolean { "TRUE" } else { "FALSE" })));
    }
    Ok(value.date()?.map(|date| Cow::Owned(date.to_string())))
}

fn formula_category_projection_error(_error: group_node_category_codec::DecodeError) -> Error {
    Error::InvalidFormat(
        "Numbers formula category projection diverged from its wire preflight".to_owned(),
    )
}

fn render_category_reference(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    references: &FormulaReferenceMaps,
) -> String {
    let category_uid = node
        .ast_category_ref
        .as_ref()
        .map(|ast| &ast.category_ref)
        .and_then(|category| {
            category
                .absolute_group_uid
                .as_ref()
                .or(category.relative_group_uid.as_ref())
                .or_else(|| category.group_uids.as_ref()?.uid.last())
        });
    category_uid
        .and_then(|uid| references.categories.get(&formula_category_key(uid)))
        .map(|label| {
            let escaped = label.replace('\\', "\\\\").replace(']', "\\]");
            format!("#CATEGORY![{escaped}]")
        })
        .unwrap_or_else(|| "#CATEGORY!".to_owned())
}

fn find_bundle_object(
    bundle: &Components,
    identifier: u64,
) -> Option<&litchi_iwa_core::ArchiveObject> {
    bundle
        .iter_archives()
        .map(|(_, archive)| archive)
        .find_map(|archive| archive.object(identifier))
}

fn formula_owner_key(owner: &litchi_iwa_protos::tsp::Uuid) -> FormulaOwnerKey {
    [
        owner.lower as u32,
        (owner.lower >> 32) as u32,
        owner.upper as u32,
        (owner.upper >> 32) as u32,
    ]
}

fn formula_category_key(category: &litchi_iwa_protos::tsp::Uuid) -> FormulaCategoryKey {
    [category.lower, category.upper]
}

fn cfuuid_key(owner: &litchi_iwa_protos::tsp::CfuuidArchive) -> Option<FormulaOwnerKey> {
    Some([
        owner.uuid_w0?,
        owner.uuid_w1?,
        owner.uuid_w2?,
        owner.uuid_w3?,
    ])
}

fn formula_reference_prefix(
    owner: &litchi_iwa_protos::tsp::CfuuidArchive,
    references: &FormulaReferenceMaps,
) -> String {
    cfuuid_key(owner)
        .and_then(|key| references.owners.get(&key))
        .map(|name| format!("{}::{}::", name.sheet, name.table))
        .unwrap_or_else(|| "Table::".to_owned())
}

type FormulaExpr = usize;

#[derive(Debug)]
enum FormulaPart {
    Static(&'static str),
    Owned(String),
    Expr(FormulaExpr),
}

#[derive(Debug)]
struct FormulaNode {
    parts: std::ops::Range<usize>,
    rendered_len: usize,
}

#[derive(Debug, Default)]
struct FormulaRenderer {
    nodes: Vec<FormulaNode>,
    parts: Vec<FormulaPart>,
    owned_bytes: usize,
}

impl FormulaRenderer {
    fn check_additional_owned(&self, additional: usize, budget: &ProjectionBudget) -> Result<()> {
        let retained = self
            .owned_bytes
            .checked_add(additional)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(retained)
    }

    fn static_expr(
        &mut self,
        value: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        self.fixed([FormulaPart::Static(value)], budget)
    }

    fn owned_expr(&mut self, value: String, budget: &ProjectionBudget) -> Result<FormulaExpr> {
        self.fixed([FormulaPart::Owned(value)], budget)
    }

    fn binary(
        &mut self,
        left: FormulaExpr,
        operator: &'static str,
        right: FormulaExpr,
        wrapped: bool,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        if wrapped {
            self.fixed(
                [
                    FormulaPart::Static("("),
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                    FormulaPart::Static(")"),
                ],
                budget,
            )
        } else {
            self.fixed(
                [
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                ],
                budget,
            )
        }
    }

    fn unary(
        &mut self,
        prefix: &'static str,
        expression: FormulaExpr,
        suffix: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        self.fixed(
            [
                FormulaPart::Static(prefix),
                FormulaPart::Expr(expression),
                FormulaPart::Static(suffix),
            ],
            budget,
        )
    }

    fn comma_joined(
        &mut self,
        function_prefix: Option<String>,
        arguments: Vec<FormulaExpr>,
        open: &'static str,
        close: &'static str,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let part_count = arguments
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(3))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| allocation_error("Numbers formula render parts", part_count))?;
        if let Some(label) = function_prefix {
            parts.push(FormulaPart::Owned(label));
        }
        parts.push(FormulaPart::Static(open));
        for (index, argument) in arguments.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(","));
            }
            parts.push(FormulaPart::Expr(argument));
        }
        parts.push(FormulaPart::Static(close));
        self.dynamic(parts, budget)
    }

    fn array(
        &mut self,
        values: Vec<FormulaExpr>,
        columns: usize,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let part_count = values
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(2))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| allocation_error("Numbers formula array parts", part_count))?;
        parts.push(FormulaPart::Static("{"));
        for (index, value) in values.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(
                    if columns != 0 && index % columns == 0 {
                        ";"
                    } else {
                        ","
                    },
                ));
            }
            parts.push(FormulaPart::Expr(value));
        }
        parts.push(FormulaPart::Static("}"));
        self.dynamic(parts, budget)
    }

    fn fixed<const N: usize>(
        &mut self,
        parts: [FormulaPart; N],
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(N)?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes)
    }

    fn dynamic(
        &mut self,
        parts: Vec<FormulaPart>,
        budget: &ProjectionBudget,
    ) -> Result<FormulaExpr> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(parts.len())?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes)
    }

    fn measure(&self, parts: &[FormulaPart], budget: &ProjectionBudget) -> Result<(usize, usize)> {
        let mut rendered_len = 0usize;
        let mut owned_bytes = 0usize;
        for part in parts {
            let part_len = match part {
                FormulaPart::Static(value) => value.len(),
                FormulaPart::Owned(value) => {
                    owned_bytes = owned_bytes
                        .checked_add(value.len())
                        .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
                    value.len()
                },
                FormulaPart::Expr(expression) => {
                    self.nodes
                        .get(*expression)
                        .ok_or_else(|| {
                            Error::ParseError(
                                "Numbers formula renderer contains an invalid expression"
                                    .to_owned(),
                            )
                        })?
                        .rendered_len
                },
            };
            rendered_len = rendered_len
                .checked_add(part_len)
                .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        }
        let retained_owned = self
            .owned_bytes
            .checked_add(owned_bytes)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(retained_owned)?;
        let output_len = rendered_len
            .checked_add(1)
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.check_output_text(output_len)?;
        Ok((rendered_len, owned_bytes))
    }

    fn reserve_node(&mut self, part_count: usize) -> Result<()> {
        self.nodes.try_reserve(1).map_err(|_error| {
            allocation_error("Numbers formula render nodes", self.nodes.len() + 1)
        })?;
        self.parts.try_reserve(part_count).map_err(|_error| {
            allocation_error(
                "Numbers formula render parts",
                self.parts.len().saturating_add(part_count),
            )
        })?;
        Ok(())
    }

    fn push_node(
        &mut self,
        start: usize,
        rendered_len: usize,
        owned_bytes: usize,
    ) -> Result<FormulaExpr> {
        self.owned_bytes = self
            .owned_bytes
            .checked_add(owned_bytes)
            .ok_or_else(|| allocation_error("Numbers formula owned text", usize::MAX))?;
        let end = self.parts.len();
        let expression = self.nodes.len();
        self.nodes.push(FormulaNode {
            parts: start..end,
            rendered_len,
        });
        Ok(expression)
    }

    fn render(&self, expression: FormulaExpr, budget: &mut ProjectionBudget) -> Result<String> {
        let node = self.nodes.get(expression).ok_or_else(|| {
            Error::ParseError("Numbers formula has no renderable expression".to_owned())
        })?;
        let output_len = node
            .rendered_len
            .checked_add(1)
            .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
        budget.charge_output_text(output_len)?;

        let mut output = String::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| allocation_error("Numbers rendered formula", output_len))?;
        output.push('=');

        let mut pending = Vec::new();
        self.push_parts_reversed(&mut pending, node.parts.clone())?;
        while let Some(part) = pending.pop() {
            match part {
                FormulaPart::Static(value) => output.push_str(value),
                FormulaPart::Owned(value) => output.push_str(value),
                FormulaPart::Expr(child) => {
                    let child_node = self.nodes.get(*child).ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula renderer contains an invalid child".to_owned(),
                        )
                    })?;
                    self.push_parts_reversed(&mut pending, child_node.parts.clone())?;
                },
            }
        }
        debug_assert_eq!(output.len(), output_len);
        Ok(output)
    }

    fn push_parts_reversed<'a>(
        &'a self,
        pending: &mut Vec<&'a FormulaPart>,
        range: std::ops::Range<usize>,
    ) -> Result<()> {
        let count = range.len();
        pending.try_reserve(count).map_err(|_error| {
            allocation_error(
                "Numbers formula render stack",
                pending.len().saturating_add(count),
            )
        })?;
        pending.extend(self.parts[range].iter().rev());
        Ok(())
    }
}

fn formula_output_limit_error(observed: usize, budget: &ProjectionBudget) -> Error {
    Error::SemanticLimit {
        kind: SemanticLimitKind::OutputTextBytes,
        observed,
        maximum: budget.max_output_text_bytes,
        path: SemanticPath::StructuredTables,
    }
}

fn render_formula(
    formula: &tsce::FormulaArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
) -> Result<String> {
    let ast = &formula.ast_node_array;
    if ast.ast_node.is_empty() {
        return retain_text("=", budget);
    }

    let mut renderer = FormulaRenderer::default();
    let root = match render_formula_ast_array(
        ast,
        host_row,
        host_column,
        formula_references,
        budget,
        &mut renderer,
        1,
    )? {
        Some(root) => root,
        None => renderer.static_expr("FORMULA()", budget)?,
    };
    renderer.render(root, budget)
}

#[allow(
    clippy::too_many_lines,
    reason = "the exhaustive AST match preserves native node semantics"
)]
fn render_formula_ast_array(
    ast: &tsce::AstNodeArrayArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    budget: &mut ProjectionBudget,
    renderer: &mut FormulaRenderer,
    depth: usize,
) -> Result<Option<FormulaExpr>> {
    use litchi_iwa_protos::tsce::ast_node_array_archive::AstNodeType;

    budget.check_formula_render_depth(depth)?;
    budget.charge_formula_render_work(ast.ast_node.len())?;
    let mut stack = Vec::new();
    stack.try_reserve(ast.ast_node.len()).map_err(|_error| {
        allocation_error("Numbers formula expression stack", ast.ast_node.len())
    })?;

    for node in &ast.ast_node {
        let expression = match node.ast_node_type() {
            AstNodeType::AdditionNode => Some(render_binary(
                &mut stack, renderer, "+", "addition", true, budget,
            )?),
            AstNodeType::SubtractionNode => Some(render_binary(
                &mut stack,
                renderer,
                "-",
                "subtraction",
                true,
                budget,
            )?),
            AstNodeType::MultiplicationNode => Some(render_binary(
                &mut stack,
                renderer,
                "*",
                "multiplication",
                true,
                budget,
            )?),
            AstNodeType::DivisionNode => Some(render_binary(
                &mut stack, renderer, "/", "division", true, budget,
            )?),
            AstNodeType::PowerNode => Some(render_binary(
                &mut stack, renderer, "^", "power", true, budget,
            )?),
            AstNodeType::GreaterThanNode => Some(render_binary(
                &mut stack,
                renderer,
                ">",
                "greater than",
                true,
                budget,
            )?),
            AstNodeType::GreaterThanOrEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                ">=",
                "greater than or equal",
                true,
                budget,
            )?),
            AstNodeType::LessThanNode => Some(render_binary(
                &mut stack,
                renderer,
                "<",
                "less than",
                true,
                budget,
            )?),
            AstNodeType::LessThanOrEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                "<=",
                "less than or equal",
                true,
                budget,
            )?),
            AstNodeType::EqualToNode => Some(render_binary(
                &mut stack, renderer, "=", "equality", true, budget,
            )?),
            AstNodeType::NotEqualToNode => Some(render_binary(
                &mut stack,
                renderer,
                "<>",
                "inequality",
                true,
                budget,
            )?),
            AstNodeType::NumberNode => node
                .ast_number_node_number
                .map(|number| renderer.owned_expr(number.to_string(), budget))
                .transpose()?,
            AstNodeType::StringNode => node
                .ast_string_node_string
                .as_deref()
                .map(|value| formula_string_literal(value, renderer, budget))
                .transpose()?
                .map(|value| renderer.owned_expr(value, budget))
                .transpose()?,
            AstNodeType::BooleanNode => node
                .ast_boolean_node_boolean
                .map(|value| renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget))
                .transpose()?,
            AstNodeType::TokenNode => node
                .ast_token_node_boolean
                .map(|value| renderer.static_expr(if value { "TRUE" } else { "FALSE" }, budget))
                .transpose()?,
            AstNodeType::DateNode => node
                .ast_date_node_date_num
                .map(|seconds| {
                    renderer.owned_expr(format!("(DATE(2001,1,1)+{})", seconds / 86_400.0), budget)
                })
                .transpose()?,
            AstNodeType::DurationNode => node
                .ast_duration_node_unit_num
                .map(|value| renderer.owned_expr(value.to_string(), budget))
                .transpose()?,
            AstNodeType::EmptyArgumentNode => Some(renderer.static_expr("", budget)?),
            AstNodeType::CellReferenceNode => Some(renderer.owned_expr(
                render_cell_reference(node, host_row, host_column, formula_references)?,
                budget,
            )?),
            AstNodeType::LocalCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_local_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || "#REF!".to_owned(),
                            |cell| {
                                format!(
                                    "{}{}",
                                    TableDataExtractor::column_index_to_letter(cell.column_handle),
                                    cell.row_handle + 1
                                )
                            },
                        ),
                    budget,
                )?,
            ),
            AstNodeType::CrossTableCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_cross_table_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || "#REF!".to_owned(),
                            |cell| {
                                format!(
                                    "{}{}{}",
                                    formula_reference_prefix(&cell.table_id, formula_references),
                                    TableDataExtractor::column_index_to_letter(cell.column_handle),
                                    cell.row_handle + 1
                                )
                            },
                        ),
                    budget,
                )?,
            ),
            AstNodeType::FunctionNode => {
                if let Some(index) = node.ast_function_node_index {
                    let arguments = pop_formula_arguments(
                        &mut stack,
                        node.ast_function_node_num_args.unwrap_or(0),
                        "function",
                    )?;
                    Some(renderer.comma_joined(
                        Some(TableDataExtractor::get_function_name(index)),
                        arguments,
                        "(",
                        ")",
                        budget,
                    )?)
                } else {
                    None
                }
            },
            AstNodeType::ListNode => {
                if let Some(count) = node.ast_list_node_num_args {
                    let arguments = pop_formula_arguments(&mut stack, count, "list")?;
                    Some(renderer.comma_joined(None, arguments, "", "", budget)?)
                } else {
                    None
                }
            },
            AstNodeType::ArrayNode => {
                let column_count = node.ast_array_node_num_col.unwrap_or(0);
                let rows = node.ast_array_node_num_row.unwrap_or(0);
                let count = column_count.checked_mul(rows).ok_or_else(|| {
                    Error::ParseError("Numbers formula array size overflow".to_owned())
                })?;
                let values = pop_formula_arguments(&mut stack, count, "array")?;
                let column_count_usize = usize::try_from(column_count).map_err(|_error| {
                    Error::ParseError("Numbers formula array width exceeds usize".to_owned())
                })?;
                Some(renderer.array(values, column_count_usize, budget)?)
            },
            AstNodeType::ThunkNode => {
                if let Some(nested) = &node.ast_thunk_node_array {
                    let nested_expression = match render_formula_ast_array(
                        nested,
                        host_row,
                        host_column,
                        formula_references,
                        budget,
                        renderer,
                        depth.checked_add(1).ok_or(Error::SemanticLimit {
                            kind: SemanticLimitKind::FormulaRenderDepth,
                            observed: usize::MAX,
                            maximum: budget.max_formula_render_depth,
                            path: SemanticPath::StructuredTables,
                        })?,
                    )? {
                        Some(expression) => expression,
                        None => renderer.static_expr(
                            if nested.ast_node.is_empty() {
                                ""
                            } else {
                                "FORMULA()"
                            },
                            budget,
                        )?,
                    };
                    Some(nested_expression)
                } else {
                    None
                }
            },
            AstNodeType::NegationNode => stack
                .pop()
                .map(|operand| renderer.unary("-(", operand, ")", budget))
                .transpose()?,
            AstNodeType::PercentNode => {
                let operand = stack.pop().ok_or_else(|| {
                    Error::ParseError(
                        "Numbers formula percent operator is missing an operand".to_owned(),
                    )
                })?;
                Some(renderer.unary("(", operand, ")%", budget)?)
            },
            AstNodeType::ConcatenationNode => Some(render_binary(
                &mut stack,
                renderer,
                "&",
                "concatenation",
                true,
                budget,
            )?),
            AstNodeType::ColonNode | AstNodeType::ColonNodeWithUids => Some(render_binary(
                &mut stack, renderer, ":", "range", false, budget,
            )?),
            AstNodeType::ColonTractNode => Some(renderer.owned_expr(
                render_colon_tract(node, host_row, host_column, formula_references)?,
                budget,
            )?),
            AstNodeType::ReferenceErrorNode | AstNodeType::ReferenceErrorWithUids => {
                Some(renderer.static_expr("#REF!", budget)?)
            },
            AstNodeType::CategoryRefNode => Some(
                renderer.owned_expr(render_category_reference(node, formula_references), budget)?,
            ),
            AstNodeType::UnknownFunctionNode => {
                let arguments = pop_formula_arguments(
                    &mut stack,
                    node.ast_unknown_function_node_num_args.unwrap_or(0),
                    "unknown function",
                )?;
                Some(
                    renderer.comma_joined(
                        Some(
                            node.ast_unknown_function_node_string
                                .clone()
                                .unwrap_or_else(|| "UNKNOWN".to_owned()),
                        ),
                        arguments,
                        "(",
                        ")",
                        budget,
                    )?,
                )
            },
            AstNodeType::PlusSignNode
            | AstNodeType::BeginThunkNode
            | AstNodeType::EndThunkNode
            | AstNodeType::AppendWhitespaceNode
            | AstNodeType::PrependWhitespaceNode
            | AstNodeType::UidReferenceNode
            | AstNodeType::LetBindNode
            | AstNodeType::VarNode
            | AstNodeType::EndScopeNode
            | AstNodeType::LambdaNode
            | AstNodeType::BeginLambdaThunkNode
            | AstNodeType::EndLambdaThunkNode
            | AstNodeType::LinkedCellRefNode
            | AstNodeType::LinkedColumnRefNode
            | AstNodeType::LinkedRowRefNode
            | AstNodeType::ViewTractRefNode
            | AstNodeType::IntersectionNode
            | AstNodeType::SpillRangeNode => None,
        };
        if let Some(rendered_expression) = expression {
            stack.push(rendered_expression);
        }
    }
    Ok(stack.pop())
}

fn render_binary(
    stack: &mut Vec<FormulaExpr>,
    renderer: &mut FormulaRenderer,
    operator: &'static str,
    operation: &str,
    wrapped: bool,
    budget: &ProjectionBudget,
) -> Result<FormulaExpr> {
    let (left, right) = pop_binary_operands(stack, operation)?;
    renderer.binary(left, operator, right, wrapped, budget)
}

fn formula_string_literal(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    let quote_count = value.bytes().filter(|byte| *byte == b'"').count();
    let length = value
        .len()
        .checked_add(quote_count)
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| allocation_error("Numbers formula string literal", usize::MAX))?;
    renderer.check_additional_owned(length, budget)?;
    let mut literal = String::new();
    literal
        .try_reserve_exact(length)
        .map_err(|_error| allocation_error("Numbers formula string literal", length))?;
    literal.push('"');
    for character in value.chars() {
        if character == '"' {
            literal.push('"');
        }
        literal.push(character);
    }
    literal.push('"');
    Ok(literal)
}

fn render_cell_reference(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
) -> Result<String> {
    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
        let column = resolve_formula_coordinate(
            host_column,
            ast_column.column,
            ast_column.absolute.unwrap_or(false),
            "column",
        )?;
        let row = resolve_formula_coordinate(
            host_row,
            ast_row.row,
            ast_row.absolute.unwrap_or(false),
            "row",
        )?;
        let prefix = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .map(|extra| formula_reference_prefix(&extra.table_id, formula_references))
            .unwrap_or_default();
        return Ok(format!(
            "{prefix}{}{}{}{}",
            if ast_column.absolute.unwrap_or(false) {
                "$"
            } else {
                ""
            },
            TableDataExtractor::column_index_to_letter(column),
            if ast_row.absolute.unwrap_or(false) {
                "$"
            } else {
                ""
            },
            row + 1
        ));
    }
    if let Some(cell) = &node.ast_local_cell_reference_node_reference {
        return Ok(format!(
            "{}{}{}{}",
            if cell.column_is_sticky != 0 { "$" } else { "" },
            TableDataExtractor::column_index_to_letter(cell.column_handle),
            if cell.row_is_sticky != 0 { "$" } else { "" },
            cell.row_handle + 1
        ));
    }
    if let Some(cell) = &node.ast_cross_table_cell_reference_node_reference {
        return Ok(format!(
            "{}{}{}",
            formula_reference_prefix(&cell.table_id, formula_references),
            TableDataExtractor::column_index_to_letter(cell.column_handle),
            cell.row_handle + 1
        ));
    }
    Ok("#REF!".to_owned())
}

fn resolve_formula_coordinate(host: usize, stored: i32, absolute: bool, axis: &str) -> Result<u32> {
    let coordinate = if absolute {
        i64::from(stored)
    } else {
        i64::try_from(host)
            .map_err(|_error| {
                Error::ParseError(format!("Numbers formula host {axis} exceeds i64"))
            })?
            .checked_add(i64::from(stored))
            .ok_or_else(|| Error::ParseError(format!("Numbers formula {axis} overflow")))?
    };
    u32::try_from(coordinate).map_err(|_error| {
        Error::ParseError(format!(
            "Numbers formula {axis} coordinate {coordinate} is out of range"
        ))
    })
}

fn render_colon_tract(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
) -> Result<String> {
    let tract = node.ast_colon_tract.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its coordinates".to_owned())
    })?;
    let sticky = node.ast_sticky_bits.as_ref().ok_or_else(|| {
        Error::ParseError("Numbers formula colon tract is missing its sticky bits".to_owned())
    })?;
    let prefix = node
        .ast_cross_table_reference_extra_info
        .as_ref()
        .map(|extra| formula_reference_prefix(&extra.table_id, formula_references))
        .unwrap_or_default();
    // Numbers uses these maximum-handle sentinels for the unbounded axis of
    // whole-row and whole-column references (for example `1:2` and `B:C`).
    let whole_rows = tract.relative_column.is_empty()
        && tract.absolute_column.len() == 1
        && tract.absolute_column[0].range_begin == i16::MAX as u32
        && tract.absolute_column[0].range_end.is_none();
    let whole_columns = tract.relative_row.is_empty()
        && tract.absolute_row.len() == 1
        && tract.absolute_row[0].range_begin == i32::MAX as u32
        && tract.absolute_row[0].range_end.is_none();
    let has_columns =
        !whole_rows && (!tract.relative_column.is_empty() || !tract.absolute_column.is_empty());
    let has_rows =
        !whole_columns && (!tract.relative_row.is_empty() || !tract.absolute_row.is_empty());
    match (has_columns, has_rows) {
        (true, true) => {
            let (begin_column, end_column) = resolve_colon_axis(
                &tract.relative_column,
                &tract.absolute_column,
                sticky.begin_column_is_absolute,
                sticky.end_column_is_absolute,
                host_column,
                "column",
            )?;
            let (begin_row, end_row) = resolve_colon_axis(
                &tract.relative_row,
                &tract.absolute_row,
                sticky.begin_row_is_absolute,
                sticky.end_row_is_absolute,
                host_row,
                "row",
            )?;
            Ok(format!(
                "{prefix}{}{}{}{}:{}{}{}{}",
                if sticky.begin_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(begin_column),
                if sticky.begin_row_is_absolute {
                    "$"
                } else {
                    ""
                },
                u64::from(begin_row) + 1,
                if sticky.end_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(end_column),
                if sticky.end_row_is_absolute { "$" } else { "" },
                u64::from(end_row) + 1,
            ))
        },
        (false, true) => {
            let (begin, end) = resolve_colon_axis(
                &tract.relative_row,
                &tract.absolute_row,
                sticky.begin_row_is_absolute,
                sticky.end_row_is_absolute,
                host_row,
                "row",
            )?;
            Ok(format!(
                "{prefix}{}{}:{}{}",
                if sticky.begin_row_is_absolute {
                    "$"
                } else {
                    ""
                },
                u64::from(begin) + 1,
                if sticky.end_row_is_absolute { "$" } else { "" },
                u64::from(end) + 1,
            ))
        },
        (true, false) => {
            let (begin, end) = resolve_colon_axis(
                &tract.relative_column,
                &tract.absolute_column,
                sticky.begin_column_is_absolute,
                sticky.end_column_is_absolute,
                host_column,
                "column",
            )?;
            Ok(format!(
                "{prefix}{}{}:{}{}",
                if sticky.begin_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(begin),
                if sticky.end_column_is_absolute {
                    "$"
                } else {
                    ""
                },
                TableDataExtractor::column_index_to_letter(end),
            ))
        },
        (false, false) => Err(Error::ParseError(
            "Numbers formula colon tract has no row or column coordinates".to_owned(),
        )),
    }
}

fn resolve_colon_axis(
    relative: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive],
    absolute: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive],
    begin_is_absolute: bool,
    end_is_absolute: bool,
    host: usize,
    axis: &str,
) -> Result<(u32, u32)> {
    let resolve = |is_absolute: bool, is_end: bool| -> Result<u32> {
        if is_absolute {
            let range = absolute.first().ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers formula colon tract has no absolute {axis} coordinate"
                ))
            })?;
            Ok(if is_end {
                range.range_end.unwrap_or(range.range_begin)
            } else {
                range.range_begin
            })
        } else {
            let range = relative.first().ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers formula colon tract has no relative {axis} coordinate"
                ))
            })?;
            let stored = if is_end {
                range.range_end.unwrap_or(range.range_begin)
            } else {
                range.range_begin
            };
            resolve_formula_coordinate(host, stored, false, axis)
        }
    };
    Ok((
        resolve(begin_is_absolute, false)?,
        resolve(end_is_absolute, true)?,
    ))
}

fn pop_binary_operands<T>(stack: &mut Vec<T>, operation: &str) -> Result<(T, T)> {
    let right = stack.pop().ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {operation} is missing its right operand"
        ))
    })?;
    let left = stack.pop().ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {operation} is missing its left operand"
        ))
    })?;
    Ok((left, right))
}

fn pop_formula_arguments<T>(stack: &mut Vec<T>, count: u32, node_kind: &str) -> Result<Vec<T>> {
    let count = usize::try_from(count).map_err(|_| {
        Error::ParseError(format!(
            "Numbers formula {node_kind} argument count exceeds usize"
        ))
    })?;
    let start = stack.len().checked_sub(count).ok_or_else(|| {
        Error::ParseError(format!(
            "Malformed Numbers formula: {node_kind} requires {count} arguments but only {} are available",
            stack.len()
        ))
    })?;
    Ok(stack.split_off(start))
}

fn take_field<'a>(data: &'a [u8], cursor: &mut usize, length: usize) -> Result<&'a [u8]> {
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| Error::ParseError("Numbers cell field offset overflow".to_string()))?;
    let field = data.get(*cursor..end).ok_or_else(|| {
        Error::ParseError(format!(
            "Truncated Numbers cell field at offset {} (need {length} bytes)",
            *cursor
        ))
    })?;
    *cursor = end;
    Ok(field)
}

fn read_u32_le(data: &[u8]) -> Result<u32> {
    let bytes: [u8; 4] = data
        .try_into()
        .map_err(|_| Error::ParseError("Expected a four-byte Numbers field".to_string()))?;
    Ok(u32::from_le_bytes(bytes))
}

fn read_f64_le(data: &[u8]) -> Result<FiniteF64> {
    let bytes: [u8; 8] = data
        .try_into()
        .map_err(|_| Error::ParseError("Expected an eight-byte Numbers field".to_string()))?;
    FiniteF64::new(f64::from_le_bytes(bytes)).map_err(|_| {
        Error::ParseError("Numbers scalar field must contain a finite value".to_string())
    })
}

fn finite_zero() -> Result<FiniteF64> {
    FiniteF64::new(0.0).map_err(|_| {
        Error::InvalidFormat("Numbers zero scalar is unexpectedly non-finite".to_string())
    })
}

#[cfg(test)]
mod tests {
    use super::{
        CellTables, Error, FormulaReferenceBudget, FormulaReferenceMaps, FormulaRenderer,
        MAX_FORMULA_CATEGORY_DEPTH, MAX_FORMULA_WIRE_BYTES, MAX_FORMULA_WORK, ProjectionBudget,
        TableDataExtractor, collect_formula_category_payload, render_formula,
        render_formula_ast_array,
    };
    use crate::cell::Value as CellValue;
    use crate::cell::wire::{BncCell, decimal128_le};
    use crate::package::SemanticPath;
    use crate::{PackageSemanticLimits as SemanticLimits, SemanticLimitKind};
    use litchi_iwa_common::comment::Comment;
    use litchi_iwa_common::wire::append_length_delimited_field;
    use litchi_iwa_protos::tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};
    use litchi_iwa_protos::{tsce, tsp, tst};
    use prost::Message as _;
    use std::collections::HashMap;

    const TEST_DECIMAL_FLAG: u32 = 0x0000_0001;

    fn formula_node(kind: AstNodeType) -> AstNodeArchive {
        AstNodeArchive {
            ast_node_type: kind as i32,
            ..Default::default()
        }
    }

    fn number_node(value: f64) -> AstNodeArchive {
        AstNodeArchive {
            ast_number_node_number: Some(value),
            ..formula_node(AstNodeType::NumberNode)
        }
    }

    fn formula(nodes: Vec<AstNodeArchive>) -> tsce::FormulaArchive {
        tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
            ..Default::default()
        }
    }

    #[test]
    fn padded_missing_cell_offset_slots_are_accepted() {
        let offsets = [
            0, 0, // column 0 starts at byte 0
            0xff, 0xff, // native tile-width padding
            0xff, 0xff,
        ];
        let cells = TableDataExtractor::parse_cell_offsets(&offsets, 1, false, 1, 1)
            .unwrap_or_else(|error| panic!("missing padded slots were rejected: {error}"));
        assert_eq!(cells, vec![(0, 0..1)]);
    }

    #[test]
    fn populated_cell_offset_slots_outside_table_width_are_rejected() {
        let offsets = [0, 0, 0, 0];
        let error = match TableDataExtractor::parse_cell_offsets(&offsets, 1, false, 1, 1) {
            Ok(cells) => panic!("populated padded slot produced cells: {cells:?}"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::InvalidFormat(message) if message.contains("outside the declared table width")
        ));
    }

    #[test]
    fn numeric_type_nine_bnc_cell_is_not_misclassified_as_empty_rich_text() {
        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, tsce::FormulaArchive)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };

        let mut encoded = vec![5, 9, 0, 0, 0, 0, 0, 0];
        encoded.extend_from_slice(&TEST_DECIMAL_FLAG.to_le_bytes());
        encoded.extend_from_slice(
            &decimal128_le(-1_234.5)
                .unwrap_or_else(|error| panic!("test decimal did not encode: {error}")),
        );
        let round_tripped = BncCell::parse(&encoded)
            .unwrap_or_else(|error| panic!("type-nine cell did not parse: {error}"))
            .encode();

        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let parsed = TableDataExtractor::parse_bnc_cell(&round_tripped, &tables, &mut budget, 2, 3)
            .unwrap_or_else(|error| panic!("type-nine cell did not extract: {error}"));
        let CellValue::Number(value) = parsed.value else {
            panic!("type-nine decimal was not extracted as a number");
        };
        assert_eq!(value.get(), -1_234.5);
    }

    #[test]
    fn arena_formula_renderer_matches_reference_output() -> super::Result<()> {
        let mut string = formula_node(AstNodeType::StringNode);
        string.ast_string_node_string = Some("a\"b".to_owned());
        let input = formula(vec![
            number_node(1.0),
            number_node(2.0),
            formula_node(AstNodeType::AdditionNode),
            string,
            formula_node(AstNodeType::ConcatenationNode),
        ]);
        let references = FormulaReferenceMaps::default();
        let expected =
            TableDataExtractor::extract_formula_string_reference(&input, 0, 0, &references)?;
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let actual = render_formula(&input, 0, 0, &references, &mut budget)?;
        assert_eq!(actual, expected);
        assert_eq!(actual, "=((1+2)&\"a\"\"b\")");
        Ok(())
    }

    #[test]
    fn skewed_concatenation_uses_linear_arena_storage() -> super::Result<()> {
        const VALUES: usize = 4_096;
        let mut nodes = Vec::new();
        nodes
            .try_reserve_exact(VALUES * 2 - 1)
            .map_err(|_| super::allocation_error("test formula nodes", VALUES * 2 - 1))?;
        nodes.push(number_node(1.0));
        for _ in 1..VALUES {
            nodes.push(number_node(1.0));
            nodes.push(formula_node(AstNodeType::ConcatenationNode));
        }
        let input = formula(nodes);
        let references = FormulaReferenceMaps::default();
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let mut renderer = FormulaRenderer::default();
        let root = render_formula_ast_array(
            &input.ast_node_array,
            0,
            0,
            &references,
            &mut budget,
            &mut renderer,
            1,
        )?
        .unwrap_or_else(|| panic!("skewed formula did not produce an expression"));
        assert_eq!(renderer.nodes.len(), VALUES * 2 - 1);
        assert_eq!(renderer.parts.len(), VALUES + (VALUES - 1) * 5);
        let output = renderer.render(root, &mut budget)?;
        assert_eq!(output.len(), VALUES * 4 - 2);
        Ok(())
    }

    #[test]
    fn formula_work_text_and_depth_limits_are_inclusive() -> super::Result<()> {
        let references = FormulaReferenceMaps::default();
        let input = formula(vec![
            number_node(1.0),
            number_node(2.0),
            formula_node(AstNodeType::AdditionNode),
        ]);

        let exact_limits = SemanticLimits::default()
            .with_formula_render_limits(3, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?
            .with_projection_limits(crate::MAX_MATERIALIZED_CELLS, 6)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut exact = ProjectionBudget::new(exact_limits);
        assert_eq!(
            render_formula(&input, 0, 0, &references, &mut exact)?,
            "=(1+2)"
        );

        let tight_work = SemanticLimits::default()
            .with_formula_render_limits(2, 64)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut work_budget = ProjectionBudget::new(tight_work);
        assert!(matches!(
            render_formula(&input, 0, 0, &references, &mut work_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 3,
                maximum: 2,
                ..
            })
        ));

        let tight_text = SemanticLimits::default()
            .with_projection_limits(crate::MAX_MATERIALIZED_CELLS, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut text_budget = ProjectionBudget::new(tight_text);
        assert!(matches!(
            render_formula(&input, 0, 0, &references, &mut text_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 6,
                maximum: 5,
                ..
            })
        ));

        let mut nested = tsce::AstNodeArrayArchive {
            ast_node: vec![number_node(1.0)],
        };
        for _ in 1..=3 {
            let mut thunk = formula_node(AstNodeType::ThunkNode);
            thunk.ast_thunk_node_array = Some(nested);
            nested = tsce::AstNodeArrayArchive {
                ast_node: vec![thunk],
            };
        }
        let nested = tsce::FormulaArchive {
            ast_node_array: nested,
            ..Default::default()
        };
        let depth_limits = SemanticLimits::default()
            .with_formula_render_limits(4, 3)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut depth_budget = ProjectionBudget::new(depth_limits);
        assert!(matches!(
            render_formula(&nested, 0, 0, &references, &mut depth_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderDepth,
                observed: 4,
                maximum: 3,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn projection_budget_is_package_aggregate() -> super::Result<()> {
        let limits = SemanticLimits::default()
            .with_projection_limits(3, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut budget = ProjectionBudget::new(limits);
        budget.charge_materialized_cells(2)?;
        budget.charge_materialized_cells(1)?;
        assert!(matches!(
            budget.charge_materialized_cells(1),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 4,
                maximum: 3,
                ..
            })
        ));
        budget.charge_output_text(2)?;
        budget.charge_output_text(3)?;
        assert!(matches!(
            budget.charge_output_text(1),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 6,
                maximum: 5,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn rejected_attempt_rolls_back_retained_values_but_not_formula_work() -> super::Result<()> {
        let limits = SemanticLimits::default()
            .with_projection_limits(3, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?
            .with_formula_render_limits(2, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut published = ProjectionBudget::new(limits);
        published.charge_materialized_cells(1)?;
        published.charge_output_text(1)?;

        let mut rejected = published;
        rejected.charge_materialized_cells(2)?;
        rejected.charge_output_text(4)?;
        rejected.charge_formula_render_work(1)?;
        published.commit_attempt(rejected, false);

        assert_eq!(published.materialized_cells, 1);
        assert_eq!(published.output_text_bytes, 1);
        assert_eq!(published.formula_render_work, 1);

        let mut over_budget = published;
        assert!(matches!(
            over_budget.charge_formula_render_work(2),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaRenderWork,
                observed: 3,
                maximum: 2,
                ..
            })
        ));
        published.commit_attempt(over_budget, false);
        assert_eq!(published.formula_render_work, 2);
        Ok(())
    }

    #[test]
    fn formula_category_walk_is_lazy_iterative_and_bounded() -> super::Result<()> {
        let mut deep_wire = Vec::new();
        for _ in 0..=MAX_FORMULA_CATEGORY_DEPTH {
            let mut parent = Vec::new();
            append_length_delimited_field(&mut parent, 3, &deep_wire)?;
            deep_wire = parent;
        }
        let mut depth_names = HashMap::new();
        let mut depth_budget = FormulaReferenceBudget::new(crate::MAX_REFERENCES);
        let depth_result =
            collect_formula_category_payload(&deep_wire, &mut depth_names, &mut depth_budget);
        assert!(
            matches!(
                &depth_result,
                Err(Error::SemanticLimit {
                    kind: SemanticLimitKind::FormulaDepth,
                    observed,
                    maximum: MAX_FORMULA_CATEGORY_DEPTH,
                    path: SemanticPath::StructuredTables,
                }) if *observed == MAX_FORMULA_CATEGORY_DEPTH + 1
            ),
            "unexpected depth result: {depth_result:?}"
        );

        let category = |lower, value: &str, child| tst::group_by_archive::GroupNodeArchive {
            group_uid: tsp::Uuid { lower, upper: 0 },
            group_cell_value: Some(tsce::CellValueArchive {
                string_value: Some(tsce::StringCellValueArchive {
                    value: value.to_owned(),
                    ..Default::default()
                }),
                ..Default::default()
            }),
            child,
            ..Default::default()
        };
        let shallow = category(1, "root", vec![category(2, "child", Vec::new())]);
        let mut tight_names = HashMap::new();
        let mut tight_budget = FormulaReferenceBudget::new(1);
        assert!(matches!(
            collect_formula_category_payload(
                &shallow.encode_to_vec(),
                &mut tight_names,
                &mut tight_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::References,
                observed: 2,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            })
        ));

        let empty_fanout = tst::group_by_archive::GroupNodeArchive {
            child: vec![tst::group_by_archive::GroupNodeArchive::default(); 32],
            ..Default::default()
        };
        let mut fanout_names = HashMap::new();
        let mut false_positive_budget = FormulaReferenceBudget::new(1);
        collect_formula_category_payload(
            &empty_fanout.encode_to_vec(),
            &mut fanout_names,
            &mut false_positive_budget,
        )?;
        assert_eq!(false_positive_budget.retained_entries, 0);

        let one_child = tst::group_by_archive::GroupNodeArchive {
            child: vec![tst::group_by_archive::GroupNodeArchive::default()],
            ..Default::default()
        };
        let mut work_names = HashMap::new();
        let mut full_work_budget = FormulaReferenceBudget::new(crate::MAX_REFERENCES);
        full_work_budget.work_items = MAX_FORMULA_WORK - 1;
        assert!(matches!(
            collect_formula_category_payload(
                &one_child.encode_to_vec(),
                &mut work_names,
                &mut full_work_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed > MAX_FORMULA_WORK
        ));

        let boolean_wrapper = [0x08, 0x01];
        let mut cell_value = Vec::new();
        append_length_delimited_field(&mut cell_value, 2, &boolean_wrapper)?;
        let mut nested_projection = Vec::new();
        append_length_delimited_field(&mut nested_projection, 7, &cell_value)?;
        let mut nested_names = HashMap::new();
        let mut nested_work_budget = FormulaReferenceBudget::new(crate::MAX_REFERENCES);
        nested_work_budget.work_items = MAX_FORMULA_WORK - 5;
        assert!(matches!(
            collect_formula_category_payload(
                &nested_projection,
                &mut nested_names,
                &mut nested_work_budget,
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed > MAX_FORMULA_WORK
        ));

        let mut string_wrapper = Vec::new();
        append_length_delimited_field(&mut string_wrapper, 1, b"valid")?;
        let mut malformed_cell = Vec::new();
        append_length_delimited_field(&mut malformed_cell, 5, &string_wrapper)?;
        malformed_cell.extend_from_slice(&[0x20, 0x01]);
        let mut malformed_projection = Vec::new();
        append_length_delimited_field(&mut malformed_projection, 7, &malformed_cell)?;
        let mut malformed_names = HashMap::new();
        let mut malformed_budget = FormulaReferenceBudget::new(crate::MAX_REFERENCES);
        collect_formula_category_payload(
            &malformed_projection,
            &mut malformed_names,
            &mut malformed_budget,
        )?;
        assert!(malformed_names.is_empty());
        let malformed_work = malformed_budget.work_items;
        assert!(malformed_work > 0);
        malformed_budget.work_items = MAX_FORMULA_WORK - malformed_work;
        collect_formula_category_payload(
            &malformed_projection,
            &mut malformed_names,
            &mut malformed_budget,
        )?;
        assert_eq!(malformed_budget.work_items, MAX_FORMULA_WORK);
        assert!(matches!(
            collect_formula_category_payload(&[], &mut malformed_names, &mut malformed_budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWork,
                observed,
                maximum: MAX_FORMULA_WORK,
                path: SemanticPath::StructuredTables,
            }) if observed == MAX_FORMULA_WORK + 1
        ));

        let duplicate = category(1, "first", vec![category(1, "second", Vec::new())]);
        let mut duplicate_names = HashMap::new();
        let mut duplicate_budget = FormulaReferenceBudget::new(1);
        collect_formula_category_payload(
            &duplicate.encode_to_vec(),
            &mut duplicate_names,
            &mut duplicate_budget,
        )?;
        assert_eq!(
            duplicate_names.get(&[1, 0]).map(String::as_str),
            Some("second")
        );
        assert_eq!(duplicate_budget.retained_entries, 1);

        duplicate_names.insert([1, 0], "Grand Total".to_owned());
        let localized = category(1, "Localized Total", Vec::new());
        collect_formula_category_payload(
            &localized.encode_to_vec(),
            &mut duplicate_names,
            &mut duplicate_budget,
        )?;
        assert_eq!(
            duplicate_names.get(&[1, 0]).map(String::as_str),
            Some("Localized Total")
        );
        assert_eq!(duplicate_budget.retained_entries, 1);
        Ok(())
    }

    #[test]
    fn formula_category_wire_bytes_are_aggregate_and_inclusive() {
        let mut names = HashMap::new();
        let mut budget = FormulaReferenceBudget::new(crate::MAX_REFERENCES);
        budget.wire_bytes = MAX_FORMULA_WIRE_BYTES;

        assert!(matches!(
            collect_formula_category_payload(&[0x08], &mut names, &mut budget),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::FormulaWireBytes,
                observed,
                maximum: MAX_FORMULA_WIRE_BYTES,
                path: SemanticPath::StructuredTables,
            }) if observed == MAX_FORMULA_WIRE_BYTES + 1
        ));
    }
}
