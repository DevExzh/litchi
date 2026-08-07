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
use super::{Error, Result};
use super::{Index, Resolved};
use crate::cell::FiniteF64;
use crate::cell::Value as CellValue;
use crate::cell::wire::{BncCellView, CachedScalar, StoredValue};
use litchi_iwa_common::comment::{AuthorId, Comment, StorageId, Uuid};
use litchi_iwa_protos::{tn, tsce, tsd, tst};
use prost::Message;
use std::collections::{HashMap, HashSet};

type CompactTable<T> = Box<[(u32, T)]>;
type StringTable = CompactTable<String>;
type FormulaTable = CompactTable<tsce::FormulaArchive>;
type FormulaErrorTable = CompactTable<String>;
type CommentTable = CompactTable<Comment>;
type FormulaOwnerKey = [u32; 4];
type FormulaCategoryKey = [u64; 2];

const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const MAX_TABLE_ROWS: usize = 1 << 20;
const MAX_TABLE_COLUMNS: usize = 1 << 14;
const MAX_ADDRESSABLE_CELLS: usize = 1 << 24;
const MAX_MATERIALIZED_CELLS: usize = 1 << 20;

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
            remaining: MAX_MATERIALIZED_CELLS,
        }
    }

    fn check(&self, requested: usize) -> Result<()> {
        if requested > self.remaining {
            return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::MaterializedCells,
                observed: MAX_MATERIALIZED_CELLS
                    .saturating_sub(self.remaining)
                    .saturating_add(requested),
                limit: MAX_MATERIALIZED_CELLS,
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

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
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

fn checked_table_dimensions(row_count: u32, column_count: u32) -> Result<(usize, usize)> {
    let row_count = usize::try_from(row_count).map_err(|_| {
        Error::InvalidFormat("Numbers table row count does not fit the host usize".to_owned())
    })?;
    let column_count = usize::try_from(column_count).map_err(|_| {
        Error::InvalidFormat("Numbers table column count does not fit the host usize".to_owned())
    })?;

    if row_count > MAX_TABLE_ROWS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::TableRows,
            observed: row_count,
            limit: MAX_TABLE_ROWS,
        }));
    }
    if column_count > MAX_TABLE_COLUMNS {
        return Err(Error::Common(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::TableColumns,
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
            kind: litchi_iwa_common::LimitKind::TableCells,
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
    sheet: String,
    table: String,
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
    formula_references: FormulaReferenceMaps,
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
    pub(super) fn new(bundle: &'a Components, object_index: &'a Index) -> Self {
        Self {
            bundle,
            object_index,
            formula_references: build_formula_reference_maps(bundle),
        }
    }

    /// Extract all tables from the document
    pub(super) fn extract_all_tables(&self) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        self.for_each_table(|table| {
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
    pub(super) fn extract_all_semantic_tables(&self) -> Result<Vec<crate::Table>> {
        let mut tables = Vec::new();
        self.for_each_table(|table| {
            tables.try_reserve(1).map_err(|_| {
                allocation_error("Numbers semantic table results", tables.len() + 1)
            })?;
            tables.push(table.into_semantic_table()?);
            Ok(())
        })?;
        Ok(tables)
    }

    fn for_each_table(&self, mut visit: impl FnMut(Table) -> Result<()>) -> Result<()> {
        let mut seen_objects = HashSet::new();

        // Real packages index TableModelArchive as 6001. Older generated
        // fixtures may store the same payload under 6000, so the object
        // adapter accepts 6000 only when its payload passes model extraction;
        // a genuine TableInfoArchive is ignored rather than mis-decoded.
        for message_type in [TABLE_MODEL_MESSAGE_TYPE, 6_000] {
            for entry in self.object_index.iter_entries_by_type(message_type) {
                if !seen_objects.insert(entry.id()) {
                    continue;
                }
                if let Some(resolved) = self.object_index.resolve_ref(self.bundle, entry.id())?
                    && let Some(table) = self.extract_table_from_object(&resolved)?
                {
                    visit(table)?;
                }
            }
        }
        Ok(())
    }

    /// Extract a single table from a resolved object
    pub(super) fn extract_table_from_object(&self, object: &Resolved<'_>) -> Result<Option<Table>> {
        // Prefer the typed TableModelArchive message in real packages.
        if let Some(message) = object
            .messages
            .iter()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        {
            let table_model = tst::TableModelArchive::decode(&*message.data).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers table-model message {} is malformed: {error}",
                    message.type_
                ))
            })?;
            return self.parse_table_model(table_model).map(Some);
        }

        // Protobuf is permissive, and legacy fixtures used 6000 for a model.
        // Only return that fallback after the complete table extraction
        // succeeds; a genuine TableInfoArchive has no cell data stores and is
        // therefore skipped safely.
        for message in object
            .messages
            .iter()
            .filter(|message| message.type_ == 6_000)
        {
            let Ok(table_model) = tst::TableModelArchive::decode(&*message.data) else {
                continue;
            };
            if let Ok(table) = self.parse_table_model(table_model) {
                return Ok(Some(table));
            }
        }

        Ok(None)
    }

    /// Parse a TableModelArchive protobuf message
    fn parse_table_model(&self, table_model: tst::TableModelArchive) -> Result<Table> {
        let (row_count, column_count) =
            checked_table_dimensions(table_model.number_of_rows, table_model.number_of_columns)?;
        let mut table = Table::with_dimensions(table_model.table_name, row_count, column_count)?;

        // Extract string table for cell text values
        // string_table is a required field, not Optional
        let string_table =
            self.load_string_table(table_model.base_data_store.string_table.identifier)?;

        // Extract formula table for formula cells
        // formula_table is a required field, not Optional
        let formula_table =
            self.load_formula_table(table_model.base_data_store.formula_table.identifier)?;
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
            formula_references: &self.formula_references,
        };
        self.parse_tiles(&table_model.base_data_store.tiles, &cell_tables, &mut table)?;

        Ok(table)
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
                .collect::<std::result::Result<Vec<_>, _>>()?;
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
                        .map(|author| AuthorId::from_raw(author.identifier))
                        .transpose()?,
                    reply_ids: comment
                        .replies
                        .iter()
                        .map(|reply| StorageId::from_raw(reply.identifier).map_err(Error::from))
                        .collect::<Result<Vec<_>>>()?
                        .into_boxed_slice(),
                    storage_uuid: comment
                        .storage_uuid
                        .as_ref()
                        .map(|uuid| Uuid::from_parts(uuid.lower, uuid.upper))
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
            .collect::<std::result::Result<Vec<_>, _>>()?;
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
            let segment = tst::TableDataListSegment::decode(segment_message.data.as_slice())?;
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
        row: usize,
        column: usize,
    ) -> Result<ParsedCell> {
        let version = *data
            .first()
            .ok_or_else(|| Error::ParseError("Empty Numbers cell storage".to_string()))?;
        match version {
            0..=4 => Self::parse_pre_bnc_cell(data, cell_tables, row, column),
            5 => Self::parse_bnc_cell(data, cell_tables, row, column),
            other => Err(Error::ParseError(format!(
                "Unsupported Numbers cell storage version {other}"
            ))),
        }
    }

    fn parse_bnc_cell(
        data: &[u8],
        cell_tables: &CellTables<'_>,
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
            StoredValue::Text(identifier) => compact_table_get(cell_tables.strings, identifier)
                .cloned()
                .map_or(CellValue::Empty, CellValue::Text),
            StoredValue::RichText(identifier) => {
                compact_table_get(cell_tables.rich_text, identifier)
                    .cloned()
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
            StoredValue::Error => CellValue::Error(
                cell.formula_error_identifier()
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id).cloned())
                    .unwrap_or_else(|| "FORMULA".to_owned()),
            ),
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
            3 => string_id
                .and_then(|id| compact_table_get(cell_tables.strings, id).cloned())
                .map_or(CellValue::Empty, CellValue::Text),
            5 => CellValue::Date(date.unwrap_or(zero)),
            6 => CellValue::Boolean(number.unwrap_or(zero).get() != 0.0),
            7 => CellValue::Duration(number.unwrap_or(zero)),
            8 => CellValue::Error(
                formula_error_id
                    .and_then(|id| compact_table_get(cell_tables.formula_errors, id).cloned())
                    .unwrap_or_else(|| "FORMULA".to_owned()),
            ),
            9 => rich_text_id
                .and_then(|id| compact_table_get(cell_tables.rich_text, id).cloned())
                .map_or(CellValue::Empty, CellValue::Text),
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

    /// Extract formula string from FormulaArchive
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
    fn extract_formula_string(
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
                        let rendered = Self::extract_formula_string(
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

fn build_formula_reference_maps(bundle: &Components) -> FormulaReferenceMaps {
    let mut result = FormulaReferenceMaps::default();
    result.categories.insert([1, 0], "Grand Total".to_owned());
    let mut table_info_names = HashMap::<u64, FormulaReferenceName>::new();
    let root = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find_map(|message| tn::DocumentArchive::decode(message.data.as_slice()).ok())
        });

    if let Some(root) = root {
        for sheet_reference in root.sheets {
            let Some(sheet_object) = find_bundle_object(bundle, sheet_reference.identifier) else {
                continue;
            };
            let Some(sheet) = sheet_object.messages.iter().find_map(|message| {
                tn::SheetArchive::decode(message.data.as_slice())
                    .ok()
                    .or_else(|| {
                        tn::FormBasedSheetArchive::decode(message.data.as_slice())
                            .ok()
                            .map(|form| form.super_)
                    })
            }) else {
                continue;
            };
            for drawable in sheet.drawable_infos {
                let Some(drawable_object) = find_bundle_object(bundle, drawable.identifier) else {
                    continue;
                };
                let table_name = drawable_object.messages.iter().find_map(|message| {
                    let table_info = tst::TableInfoArchive::decode(message.data.as_slice()).ok()?;
                    let model_object =
                        find_bundle_object(bundle, table_info.table_model.identifier)?;
                    model_object.messages.iter().find_map(|message| {
                        (message.type_ == 6000 || message.type_ == 6001)
                            .then(|| tst::TableModelArchive::decode(message.data.as_slice()).ok())
                            .flatten()
                            .map(|model| model.table_name)
                    })
                });
                if let Some(table) = table_name {
                    table_info_names.insert(
                        drawable.identifier,
                        FormulaReferenceName {
                            sheet: sheet.name.clone(),
                            table,
                        },
                    );
                }
            }
        }
    }

    for (_, archive) in bundle.iter_archives() {
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == 6383
                    && let Ok(group_node) =
                        tst::group_by_archive::GroupNodeArchive::decode(message.data.as_slice())
                {
                    collect_formula_category_names(&group_node, &mut result.categories);
                    continue;
                }
                if message.type_ != 4008 {
                    continue;
                }
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
                result
                    .owners
                    .insert(formula_owner_key(&owner.formula_owner_uid), name.clone());
            }
        }
    }
    result
}

fn collect_formula_category_names(
    node: &tst::group_by_archive::GroupNodeArchive,
    names: &mut HashMap<FormulaCategoryKey, String>,
) {
    if let Some(value) = node
        .group_cell_value
        .as_ref()
        .and_then(group_cell_value_label)
    {
        names.insert(formula_category_key(&node.group_uid), value);
    }
    for child in &node.child {
        collect_formula_category_names(child, names);
    }
}

fn group_cell_value_label(value: &tsce::CellValueArchive) -> Option<String> {
    if let Some(string) = &value.string_value {
        return Some(string.value.clone());
    }
    if let Some(number) = &value.number_value
        && let Some(number) = number.value
    {
        return Some(number.to_string());
    }
    if let Some(boolean) = &value.boolean_value {
        return Some(if boolean.value { "TRUE" } else { "FALSE" }.to_owned());
    }
    value.date_value.as_ref().map(|date| date.value.to_string())
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

fn resolve_formula_coordinate(host: usize, stored: i32, absolute: bool, axis: &str) -> Result<u32> {
    let coordinate = if absolute {
        i64::from(stored)
    } else {
        i64::try_from(host)
            .map_err(|_| Error::ParseError(format!("Numbers formula host {axis} exceeds i64")))?
            .checked_add(i64::from(stored))
            .ok_or_else(|| Error::ParseError(format!("Numbers formula {axis} overflow")))?
    };
    u32::try_from(coordinate).map_err(|_| {
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

fn pop_binary_operands(stack: &mut Vec<String>, operation: &str) -> Result<(String, String)> {
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

fn pop_formula_arguments(
    stack: &mut Vec<String>,
    count: u32,
    node_kind: &str,
) -> Result<Vec<String>> {
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
    use super::{CellTables, Error, FormulaReferenceMaps, TableDataExtractor};
    use crate::cell::Value as CellValue;
    use crate::cell::wire::{BncCell, DECIMAL_FLAG, decimal128_le};
    use litchi_iwa_common::comment::Comment;
    use litchi_iwa_protos::tsce;

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
        encoded.extend_from_slice(&DECIMAL_FLAG.to_le_bytes());
        encoded.extend_from_slice(
            &decimal128_le(-1_234.5)
                .unwrap_or_else(|error| panic!("test decimal did not encode: {error}")),
        );
        let round_tripped = BncCell::parse(&encoded)
            .unwrap_or_else(|error| panic!("type-nine cell did not parse: {error}"))
            .encode();

        let parsed = TableDataExtractor::parse_bnc_cell(&round_tripped, &tables, 2, 3)
            .unwrap_or_else(|error| panic!("type-nine cell did not extract: {error}"));
        let CellValue::Number(value) = parsed.value else {
            panic!("type-nine decimal was not extracted as a number");
        };
        assert_eq!(value.get(), -1_234.5);
    }
}
