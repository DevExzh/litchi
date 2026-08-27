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
//! ## Example
//!
//! ```rust,ignore
//! use litchi_iwa::numbers::table_extractor::TableDataExtractor;
//! use litchi_iwa::bundle::Bundle;
//! use litchi_iwa::object_index::ObjectIndex;
//!
//! let bundle = Bundle::open("document.numbers")?;
//! let index = ObjectIndex::from_bundle(&bundle)?;
//! let extractor = TableDataExtractor::new(&bundle, &index);
//!
//! let tables = extractor.extract_all_tables()?;
//! for table in tables {
//!     println!("Table: {}", table.name());
//!     println!("{}", table.to_csv());
//! }
//! ```

use super::bnc::{BncCellView, CachedScalar, StoredValue};
#[cfg(test)]
use super::bnc::{decimal128_le, read_decimal128_le};
use super::cell::CellValue;
use super::editor::table_model_projection::{ProbeBudget, map_resource_error, select_candidate};
use super::formula_renderer::{FormulaArchiveBytes, render_formula_string};
use super::table::NumbersTable;
use crate::bundle::Bundle;
use crate::object_index::{ObjectIndex, ResolvedObjectRef};
use crate::protobuf::{tn, tsce, tst};
use crate::{Error, Result};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::comment::{AuthorId, Comment, StorageId, Uuid};
use litchi_iwa_common::formula::FiniteF64 as CommonFiniteF64;
use litchi_iwa_protos::comment_storage_codec;
use litchi_iwa_protos::numbers_formula_codec;
use litchi_iwa_protos::numbers_names_codec;
use litchi_iwa_protos::numbers_table_cell_storage_codec;
use litchi_numbers::cell::FiniteF64;
use litchi_numbers::table::Dimensions;
use prost::Message;
use std::collections::{HashMap, HashSet};

type CompactTable<T> = Box<[(u32, T)]>;
type StringTable = CompactTable<String>;
type FormulaTable = CompactTable<FormulaArchiveBytes>;
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
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const COMMENT_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;

// Repeated sidecar entries are bounded independently of the generated
// archive's vector capacity.  The strict codecs use the same ceilings for
// their aggregate fields/references; keep the host-side staging collections
// on those finite budgets too.
const MAX_TABLE_LIST_ENTRIES: usize = WireLimits::MAX_FIELDS;
const MAX_TABLE_LIST_SEGMENT_REFERENCES: usize = litchi_numbers::MAX_REFERENCES;
const MAX_COMMENT_STORAGE_CANDIDATES: usize = WireLimits::MAX_FIELDS;
const MAX_TABLE_LIST_PAYLOAD_FIELDS: usize = WireLimits::MAX_FIELDS;
const MAX_TABLE_LIST_PAYLOAD_WORK: usize = WireLimits::MAX_REWRITE_WORK;
const MAX_TABLE_LIST_REFERENCES: usize = litchi_numbers::MAX_REFERENCES;
const MAX_TABLE_LIST_TEXT_BYTES: usize = litchi_numbers::DEFAULT_MAX_TEXT_BYTES;
const MAX_FORMULA_WIRE_BYTES: usize = litchi_numbers::DEFAULT_MAX_TEXT_BYTES;
const MAX_FORMULA_RENDER_WORK: usize = litchi_numbers::MAX_REFERENCES;
const MAX_FORMULA_RENDER_DEPTH: usize = WireLimits::MAX_NESTING;

/// Aggregate admission state for one host table-data-list projection.
///
/// The strict codec's `DecodeOptions` are per source message.  A root list
/// can name many segment objects, so applying those options independently
/// would allow a package to multiply the same fields/work/references/text
/// allowance once per segment.  This small host-owned budget carries the
/// successful root and segment reports together.  Probe and duplicate/wrong
/// candidates charge only fields/work: those candidates are wire-checked for
/// compatibility, but their selected references/text are not published by
/// the legacy best-effort route.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct ProjectionBudget {
    pub(super) references: usize,
    pub(super) payload_fields: usize,
    pub(super) payload_work: usize,
    staging_text_bytes: usize,
    entries: usize,
    pub(super) formula_wire_bytes: usize,
    pub(super) output_text_bytes: usize,
    pub(super) formula_render_work: usize,
    pub(super) max_output_text_bytes: usize,
    pub(super) max_formula_render_depth: usize,
}

impl ProjectionBudget {
    const fn new() -> Self {
        Self {
            references: 0,
            payload_fields: 0,
            payload_work: 0,
            staging_text_bytes: 0,
            entries: 0,
            formula_wire_bytes: 0,
            output_text_bytes: 0,
            formula_render_work: 0,
            max_output_text_bytes: MAX_TABLE_LIST_TEXT_BYTES,
            max_formula_render_depth: MAX_FORMULA_RENDER_DEPTH,
        }
    }

    const fn remaining_references(self) -> usize {
        MAX_TABLE_LIST_REFERENCES.saturating_sub(self.references)
    }

    pub(super) const fn remaining_payload_fields(self) -> usize {
        MAX_TABLE_LIST_PAYLOAD_FIELDS.saturating_sub(self.payload_fields)
    }

    pub(super) const fn remaining_payload_work(self) -> usize {
        MAX_TABLE_LIST_PAYLOAD_WORK.saturating_sub(self.payload_work)
    }

    pub(super) const fn remaining_staging_text_bytes(self) -> usize {
        MAX_TABLE_LIST_TEXT_BYTES.saturating_sub(self.staging_text_bytes)
    }

    const fn remaining_entries(self) -> usize {
        MAX_TABLE_LIST_ENTRIES.saturating_sub(self.entries)
    }

    fn charge(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: litchi_iwa_common::LimitKind,
    ) -> Result<()> {
        let observed = current.saturating_add(amount);
        if observed > maximum {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit: maximum,
            }));
        }
        *current = observed;
        Ok(())
    }

    fn charge_references(&mut self, amount: usize) -> Result<()> {
        let observed = self.references.saturating_add(amount);
        if observed > MAX_TABLE_LIST_REFERENCES {
            return Err(Error::InvalidFormat(format!(
                "Numbers table-data-list references exceeded their aggregate limit: observed {observed}, limit {MAX_TABLE_LIST_REFERENCES}"
            )));
        }
        self.references = observed;
        Ok(())
    }

    fn charge_payload_fields(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.payload_fields,
            amount,
            MAX_TABLE_LIST_PAYLOAD_FIELDS,
            litchi_iwa_common::LimitKind::Fields,
        )
    }

    fn charge_payload_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.payload_work,
            amount,
            MAX_TABLE_LIST_PAYLOAD_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )
    }

    fn charge_staging_text(&mut self, amount: usize) -> Result<()> {
        let observed = self.staging_text_bytes.saturating_add(amount);
        if observed > MAX_TABLE_LIST_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "Numbers table-data-list text exceeded its aggregate limit: observed {observed}, limit {MAX_TABLE_LIST_TEXT_BYTES}"
            )));
        }
        self.staging_text_bytes = observed;
        Ok(())
    }

    fn charge_entries(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.entries,
            amount,
            MAX_TABLE_LIST_ENTRIES,
            litchi_iwa_common::LimitKind::Fields,
        )
    }

    pub(super) fn charge_formula_wire(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.formula_wire_bytes,
            amount,
            MAX_FORMULA_WIRE_BYTES,
            litchi_iwa_common::LimitKind::InputBytes,
        )
    }

    pub(super) fn charge_wire_preflight(
        &mut self,
        report: litchi_iwa_common::wire::WirePreflight,
    ) -> Result<()> {
        let mut next = *self;
        Self::charge(
            &mut next.payload_fields,
            report.fields(),
            MAX_TABLE_LIST_PAYLOAD_FIELDS,
            litchi_iwa_common::LimitKind::Fields,
        )?;
        Self::charge(
            &mut next.payload_work,
            report.scanned_bytes(),
            MAX_TABLE_LIST_PAYLOAD_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )?;
        *self = next;
        Ok(())
    }

    pub(super) fn retain_formula_preflight_cost(&mut self, fields: usize, work: usize) {
        self.payload_fields = self
            .payload_fields
            .saturating_add(fields)
            .min(MAX_TABLE_LIST_PAYLOAD_FIELDS);
        self.payload_work = self
            .payload_work
            .saturating_add(work)
            .min(MAX_TABLE_LIST_PAYLOAD_WORK);
    }

    pub(super) fn charge_formula_lazy_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.payload_work,
            amount,
            MAX_TABLE_LIST_PAYLOAD_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )
    }

    pub(super) fn check_output_text(&self, amount: usize) -> Result<()> {
        let observed = self.output_text_bytes.saturating_add(amount);
        if observed > MAX_TABLE_LIST_TEXT_BYTES {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed,
                limit: MAX_TABLE_LIST_TEXT_BYTES,
            }));
        }
        Ok(())
    }

    pub(super) fn charge_output_text(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.output_text_bytes,
            amount,
            MAX_TABLE_LIST_TEXT_BYTES,
            litchi_iwa_common::LimitKind::OutputBytes,
        )
    }

    pub(super) fn charge_formula_render_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.formula_render_work,
            amount,
            MAX_FORMULA_RENDER_WORK,
            litchi_iwa_common::LimitKind::RewriteWork,
        )
    }

    pub(super) fn charge_formula_decode_report(
        &mut self,
        report: numbers_formula_codec::DecodeReport,
    ) -> Result<()> {
        // Formula decoding performs a strict preflight pass followed by the
        // caller callback pass. Merge the complete successful report into a
        // copy so a later aggregate-axis refusal cannot publish a partially
        // charged traversal.
        let mut next = *self;
        next.charge_payload_fields(report.fields())?;
        next.charge_payload_work(report.work())?;
        next.charge_staging_text(report.text_bytes())?;
        *self = next;
        Ok(())
    }

    pub(super) fn check_formula_render_depth(&self, depth: usize) -> Result<()> {
        if depth > self.max_formula_render_depth {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Nesting,
                observed: depth,
                limit: self.max_formula_render_depth,
            }));
        }
        Ok(())
    }

    fn charge_decode_work(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        // Validate the complete report before publishing any of its counters.
        // A segment may be the first source that crosses the aggregate work
        // ceiling; retaining its fields while refusing its work would leave
        // a partially charged candidate if a caller inspects or reuses the
        // budget after the atomic refusal.
        let mut next = *self;
        next.charge_payload_fields(report.fields())?;
        next.charge_payload_work(report.work_bytes())?;
        *self = next;
        Ok(())
    }

    fn charge_decode_report(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        // Keep report admission atomic across references, fields, work, and
        // selected text.  Root/segment values are staged until the complete
        // traversal succeeds, so a rejected segment must not leave a prefix
        // of its aggregate counters behind.
        let mut next = *self;
        next.charge_references(report.references())?;
        next.charge_payload_fields(report.fields())?;
        next.charge_payload_work(report.work_bytes())?;
        next.charge_staging_text(report.text_bytes())?;
        *self = next;
        Ok(())
    }
}

fn table_list_collection_limits(source: &[u8]) -> (usize, usize) {
    (
        source.len().clamp(1, MAX_TABLE_LIST_ENTRIES),
        source.len().clamp(1, MAX_TABLE_LIST_SEGMENT_REFERENCES),
    )
}

fn table_list_entry_limit_error(observed: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
        kind: litchi_iwa_common::LimitKind::Fields,
        observed,
        limit: MAX_TABLE_LIST_ENTRIES,
    })
}

fn table_list_segment_reference_limit_error(observed: usize) -> Error {
    // LimitKind intentionally has no reference variant.  Match the strict
    // table-list error mapping used for DecodeLimit::References instead of
    // silently turning this bounded host collection into an allocation error.
    Error::InvalidFormat(format!(
        "Numbers table-list segment references exceeded their aggregate limit: observed {observed}, limit {MAX_TABLE_LIST_SEGMENT_REFERENCES}"
    ))
}

fn comment_storage_candidate_limit_error(observed: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
        kind: litchi_iwa_common::LimitKind::Fields,
        observed,
        limit: MAX_COMMENT_STORAGE_CANDIDATES,
    })
}

fn comment_storage_decode_options(source: &[u8]) -> comment_storage_codec::DecodeOptions {
    comment_storage_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        COMMENT_STORAGE_CODEC_RECURSION_LIMIT,
        source.len().clamp(1, litchi_numbers::MAX_REFERENCES),
        source
            .len()
            .clamp(1, litchi_numbers::DEFAULT_MAX_TEXT_BYTES),
    )
}

fn table_name_decode_options(source: &[u8]) -> numbers_names_codec::DecodeOptions {
    numbers_names_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(4)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
    )
}

fn comment_storage_allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
}

#[derive(Debug, Default)]
struct CommentStorageReplyIds {
    identifiers: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl CommentStorageReplyIds {
    fn into_identifiers(self) -> Result<Vec<u64>> {
        match self.allocation_failed {
            Some(amount) => Err(comment_storage_allocation_error(
                "Numbers comment reply identifiers",
                amount,
            )),
            None => Ok(self.identifiers),
        }
    }
}

impl comment_storage_codec::CommentStorageVisitor for CommentStorageReplyIds {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), comment_storage_codec::DecodeError> {
        if self.allocation_failed.is_some() {
            return Ok(());
        }
        if self.identifiers.try_reserve(1).is_err() {
            // Keep traversing so a later malformed field still wins over a
            // candidate-local allocation failure. IDs are published only
            // after the enclosing payload has decoded successfully.
            self.allocation_failed = Some(self.identifiers.len().saturating_add(1));
            return Ok(());
        }
        self.identifiers.push(reply.identifier());
        Ok(())
    }
}

fn strict_comment_storage_error(
    storage_id: u64,
    error: comment_storage_codec::DecodeError,
) -> Error {
    Error::InvalidFormat(format!(
        "Numbers comment storage object {storage_id} failed strict validation: {error}"
    ))
}

fn decode_comment_storage_payload<'source>(
    storage_id: u64,
    source: &'source [u8],
) -> Result<(
    comment_storage_codec::CommentStorageSnapshot<'source>,
    Vec<u64>,
)> {
    let mut replies = CommentStorageReplyIds::default();
    let (comment, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        source,
        comment_storage_decode_options(source),
        &mut replies,
    )
    .map_err(|error| strict_comment_storage_error(storage_id, error))?;
    let reply_ids = replies.into_identifiers()?;
    if reply_ids.len() != report.reply_references() {
        return Err(Error::InvalidFormat(format!(
            "Numbers comment storage object {storage_id} streamed {} replies but reported {}",
            reply_ids.len(),
            report.reply_references(),
        )));
    }
    Ok((comment, reply_ids))
}

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

#[derive(Debug, Clone, Copy)]
struct TileRoute {
    tile_id: u32,
    object_identifier: u64,
}

#[derive(Default)]
struct TileRoutes {
    routes: Vec<TileRoute>,
}

impl numbers_table_cell_storage_codec::StorageVisitor for TileRoutes {
    fn visit_tile_reference(
        &mut self,
        record: numbers_table_cell_storage_codec::TileReferenceRecord<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        self.routes.try_reserve(1).map_err(|_| {
            numbers_table_cell_storage_codec::DecodeError::allocation(
                self.routes.len().saturating_add(1),
            )
        })?;
        self.routes.push(TileRoute {
            tile_id: record.tile_id(),
            object_identifier: record.reference().identifier(),
        });
        Ok(())
    }
}

struct StagedTableCell {
    row: usize,
    column: usize,
    parsed: ParsedCell,
}

struct TileRowStage<'context, 'tables> {
    row_origin: usize,
    tile_size: usize,
    dimensions: Dimensions,
    budget: &'context mut CellBudget,
    formula_budget: &'context mut ProjectionBudget,
    cell_tables: &'context CellTables<'tables>,
    rows: HashSet<u32>,
    cells: Vec<StagedTableCell>,
    semantic_error: Option<Error>,
}

impl<'context, 'tables> TileRowStage<'context, 'tables> {
    fn new(
        row_origin: usize,
        tile_size: usize,
        dimensions: Dimensions,
        budget: &'context mut CellBudget,
        formula_budget: &'context mut ProjectionBudget,
        cell_tables: &'context CellTables<'tables>,
    ) -> Self {
        Self {
            row_origin,
            tile_size,
            dimensions,
            budget,
            formula_budget,
            cell_tables,
            rows: HashSet::new(),
            cells: Vec::new(),
            semantic_error: None,
        }
    }

    fn finish(self) -> Result<Vec<StagedTableCell>> {
        match self.semantic_error {
            Some(error) => Err(error),
            None => Ok(self.cells),
        }
    }
}

impl numbers_table_cell_storage_codec::StorageVisitor for TileRowStage<'_, '_> {
    fn visit_tile_row(
        &mut self,
        row: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if self.semantic_error.is_some() {
            return Ok(());
        }
        if self.rows.try_reserve(1).is_err() {
            self.semantic_error = Some(allocation_error(
                "Numbers tile row keys",
                self.rows.len().saturating_add(1),
            ));
            return Ok(());
        }
        if !self.rows.insert(row.tile_row_index()) {
            self.semantic_error = Some(Error::InvalidFormat(format!(
                "Numbers tile repeats row {}",
                row.tile_row_index()
            )));
            return Ok(());
        }
        if let Err(error) = stage_tile_row(
            row,
            self.row_origin,
            self.tile_size,
            self.dimensions,
            self.budget,
            self.formula_budget,
            self.cell_tables,
            &mut self.cells,
        ) {
            self.semantic_error = Some(error);
        }
        Ok(())
    }
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
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
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

fn stage_tile_row(
    row_info: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    row_origin: usize,
    tile_size: usize,
    dimensions: Dimensions,
    budget: &mut CellBudget,
    formula_budget: &mut ProjectionBudget,
    cell_tables: &CellTables<'_>,
    staged: &mut Vec<StagedTableCell>,
) -> Result<()> {
    let tile_row_index = usize::try_from(row_info.tile_row_index()).map_err(|_| {
        Error::InvalidFormat("Numbers tile row index does not fit the host usize".to_owned())
    })?;
    if tile_row_index >= tile_size {
        return Err(Error::InvalidFormat(format!(
            "Numbers tile row {} is outside tile size {tile_size}",
            row_info.tile_row_index()
        )));
    }
    let row_index = row_origin
        .checked_add(tile_row_index)
        .ok_or_else(|| Error::ParseError("Numbers tile row index overflow".to_owned()))?;
    dimensions.check_row(row_index).map_err(|_| {
        Error::InvalidFormat(format!(
            "Numbers tile row {row_index} is outside the declared table height {}",
            dimensions.rows()
        ))
    })?;

    let (cell_storage, cell_offsets) =
        match (row_info.cell_storage_buffer(), row_info.cell_offsets()) {
            (Some(storage), Some(offsets)) => (storage, offsets),
            _ => (
                row_info.cell_storage_buffer_pre_bnc(),
                row_info.cell_offsets_pre_bnc(),
            ),
        };

    let expected_cells = usize::try_from(row_info.cell_count()).map_err(|_| {
        Error::InvalidFormat("Numbers cell count does not fit the host usize".to_owned())
    })?;
    budget.check(expected_cells)?;
    let cells = TableDataExtractor::parse_cell_offsets(
        cell_offsets,
        cell_storage.len(),
        row_info.has_wide_offsets().unwrap_or(false),
        expected_cells,
        dimensions.columns() as usize,
    )?;
    budget.consume(cells.len())?;
    staged
        .try_reserve(cells.len())
        .map_err(|_| allocation_error("Numbers staged tile cells", staged.len() + cells.len()))?;

    for (column_index, range) in cells {
        dimensions.check_column(column_index).map_err(|_| {
            Error::InvalidFormat(format!(
                "Numbers cell column {column_index} is outside the declared table width {}",
                dimensions.columns()
            ))
        })?;
        let parsed = TableDataExtractor::parse_cell_storage_with_budget(
            &cell_storage[range],
            cell_tables,
            row_index,
            column_index,
            dimensions.rows() as usize,
            dimensions.columns() as usize,
            formula_budget,
        )?;
        staged.push(StagedTableCell {
            row: row_index,
            column: column_index,
            parsed,
        });
    }
    Ok(())
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
}

#[cfg(test)]
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
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::TableRows,
            observed: row_count,
            limit: MAX_TABLE_ROWS,
        }));
    }
    if column_count > MAX_TABLE_COLUMNS {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
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
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::TableCells,
            observed: addressable_cells,
            limit: MAX_ADDRESSABLE_CELLS,
        }));
    }

    Ok((row_count, column_count))
}

#[cfg(test)]
fn table_data_list_decode_options(
    source: &[u8],
) -> numbers_table_cell_storage_codec::DecodeOptions {
    numbers_table_cell_storage_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
        source.len().clamp(1, litchi_numbers::MAX_REFERENCES),
        source
            .len()
            .clamp(1, litchi_numbers::DEFAULT_MAX_TEXT_BYTES),
    )
}

/// Build per-message codec options from the aggregate host budget.  Zero
/// remaining counters are passed through deliberately: the strict codec
/// accepts a zero finite ceiling and reports the first attempted unit, while
/// the offset-aware mapper below turns that local report into the aggregate
/// observation.
fn table_data_list_decode_options_with_budget(
    source: &[u8],
    budget: ProjectionBudget,
    admit_selected_values: bool,
) -> numbers_table_cell_storage_codec::DecodeOptions {
    numbers_table_cell_storage_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        budget.remaining_payload_fields(),
        budget.remaining_payload_work(),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
        if admit_selected_values {
            budget.remaining_references()
        } else {
            usize::MAX
        },
        if admit_selected_values {
            budget.remaining_staging_text_bytes()
        } else {
            usize::MAX
        },
    )
}

fn table_data_list_decode_error_with_offsets(
    object_id: u64,
    list_type: tst::table_data_list::ListType,
    error: numbers_table_cell_storage_codec::DecodeError,
    reference_offset: usize,
    field_offset: usize,
    work_offset: usize,
    text_offset: usize,
) -> Error {
    if let Some(limit) = error.resource_limit() {
        use numbers_table_cell_storage_codec::DecodeLimit;
        return match limit {
            DecodeLimit::Bytes { observed, maximum } => {
                Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::InputBytes,
                    observed,
                    limit: maximum,
                })
            },
            DecodeLimit::Fields { observed, maximum } => {
                Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::Fields,
                    observed: observed.saturating_add(field_offset),
                    limit: maximum.saturating_add(field_offset),
                })
            },
            DecodeLimit::Work { observed, maximum } => {
                Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::RewriteWork,
                    observed: observed.saturating_add(work_offset),
                    limit: maximum.saturating_add(work_offset),
                })
            },
            DecodeLimit::Nesting { observed, maximum } => {
                Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::Nesting,
                    observed: observed as usize,
                    limit: maximum as usize,
                })
            },
            DecodeLimit::References { observed, maximum }
            | DecodeLimit::Text { observed, maximum } => {
                let (kind, offset) = match limit {
                    DecodeLimit::References { .. } => ("references", reference_offset),
                    DecodeLimit::Text { .. } => ("text bytes", text_offset),
                    _ => unreachable!("matched reference/text limit"),
                };
                Error::InvalidFormat(format!(
                    "Numbers {list_type:?} table {object_id} strict projection exceeded its aggregate {kind} limit: observed {}, limit {}",
                    observed.saturating_add(offset),
                    maximum.saturating_add(offset),
                ))
            },
            DecodeLimit::Allocation { requested } => {
                allocation_error("Numbers table storage projection", requested)
            },
            _ => Error::InvalidFormat(format!(
                "Numbers {list_type:?} table {object_id} strict projection exceeded an unsupported resource limit"
            )),
        };
    }
    Error::InvalidFormat(format!(
        "Numbers {list_type:?} table {object_id} failed strict validation: {error}"
    ))
}

fn record_first_list_error(slot: &mut Option<Error>, error: Error) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

trait ListValueConverter<T> {
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<Option<T>>;
}

impl<T, F> ListValueConverter<T> for F
where
    F: for<'source, 'budget> FnMut(
        numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'source>,
        &'budget mut ProjectionBudget,
    ) -> Result<Option<T>>,
{
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<Option<T>> {
        self(entry, budget)
    }
}

/// Stage one list candidate while the strict codec validates its complete
/// source.  Visitor callbacks can run before a later wire error, so semantic
/// conversion/allocation failures are retained until the enclosing decode has
/// finished.  Only a successful, admitted candidate is published.
struct TypedListVisitor<'converter, 'budget, T, C> {
    converter: &'converter mut C,
    budget: &'budget mut ProjectionBudget,
    stage_semantics: bool,
    // Retained for the constructor's call-site/error context. The legacy
    // extractor selects the requested union member in each converter rather
    // than rejecting entries that carry another (or multiple) union members.
    _expected_list_type: i32,
    max_entries: usize,
    max_segment_references: usize,
    values: Vec<(u32, T)>,
    keys: HashSet<u32>,
    segment_ids: Vec<u64>,
    segment_id_set: HashSet<u64>,
    segment: bool,
    segment_key_min: Option<u32>,
    segment_key_max: Option<u32>,
    structural_error: Option<Error>,
    semantic_error: Option<Error>,
}

impl<'converter, 'budget, T, C> TypedListVisitor<'converter, 'budget, T, C> {
    fn new(
        converter: &'converter mut C,
        budget: &'budget mut ProjectionBudget,
        expected_list_type: i32,
        segment: bool,
        stage_semantics: bool,
        max_entries: usize,
        max_segment_references: usize,
    ) -> Self {
        Self {
            converter,
            budget,
            stage_semantics,
            _expected_list_type: expected_list_type,
            max_entries,
            max_segment_references,
            values: Vec::new(),
            keys: HashSet::new(),
            segment_ids: Vec::new(),
            segment_id_set: HashSet::new(),
            segment,
            segment_key_min: None,
            segment_key_max: None,
            structural_error: None,
            semantic_error: None,
        }
    }

    fn record_structural_error(&mut self, error: Error) {
        if self.structural_error.is_none() {
            self.structural_error = Some(error);
        }
    }

    fn record_semantic_error(&mut self, error: Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn take_parts(
        self,
    ) -> (
        Vec<(u32, T)>,
        HashSet<u32>,
        Vec<u64>,
        Option<Error>,
        Option<Error>,
    ) {
        (
            self.values,
            self.keys,
            self.segment_ids,
            self.structural_error,
            self.semantic_error,
        )
    }

    fn take_segment_bounds(&self) -> (Option<u32>, Option<u32>) {
        (self.segment_key_min, self.segment_key_max)
    }
}

impl<T, C> numbers_table_cell_storage_codec::StorageVisitor for TypedListVisitor<'_, '_, T, C>
where
    C: ListValueConverter<T>,
{
    fn visit_list_entry(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if self.segment {
            match &mut self.segment_key_min {
                Some(minimum) => *minimum = (*minimum).min(entry.key()),
                minimum @ None => *minimum = Some(entry.key()),
            }
            match &mut self.segment_key_max {
                Some(maximum) => *maximum = (*maximum).max(entry.key()),
                maximum @ None => *maximum = Some(entry.key()),
            }
        }
        if self.structural_error.is_some() {
            return Ok(());
        }
        if self.keys.contains(&entry.key()) {
            self.record_structural_error(Error::InvalidFormat(
                "Numbers table sidecar contains duplicate keys".to_owned(),
            ));
            return Ok(());
        }
        if self.keys.len() >= self.max_entries {
            self.record_semantic_error(table_list_entry_limit_error(
                self.keys.len().saturating_add(1),
            ));
            return Ok(());
        }
        if self.keys.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list entry keys",
                self.keys.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.keys.insert(entry.key());
        if self.semantic_error.is_some() || !self.stage_semantics {
            return Ok(());
        }
        match self.converter.convert(entry, self.budget) {
            Ok(Some(value)) => {
                if self.values.try_reserve(1).is_err() {
                    self.record_semantic_error(allocation_error(
                        "Numbers table-list entries",
                        self.values.len().saturating_add(1),
                    ));
                    return Ok(());
                }
                self.values.push((entry.key(), value));
            },
            Ok(None) => {},
            Err(error) => self.record_semantic_error(error),
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        reference: numbers_table_cell_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if !self.stage_semantics {
            return Ok(());
        }
        let identifier = reference.reference().identifier();
        if self.segment_id_set.contains(&identifier) {
            self.record_structural_error(Error::InvalidFormat(format!(
                "Numbers table-data-list repeats segment object {identifier}"
            )));
            return Ok(());
        }
        if self.segment_id_set.len() >= self.max_segment_references {
            self.record_semantic_error(table_list_segment_reference_limit_error(
                self.segment_id_set.len().saturating_add(1),
            ));
            return Ok(());
        }
        if self.segment_id_set.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list segment identities",
                self.segment_id_set.len().saturating_add(1),
            ));
            return Ok(());
        }
        if self.segment_ids.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list segment identities",
                self.segment_ids.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.segment_id_set.insert(identifier);
        self.segment_ids.push(identifier);
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(super) struct FormulaReferenceName {
    pub(super) sheet: String,
    pub(super) table: String,
}

#[derive(Debug, Clone, Default)]
pub(super) struct FormulaReferenceMaps {
    pub(super) owners: HashMap<FormulaOwnerKey, FormulaReferenceName>,
    pub(super) categories: HashMap<FormulaCategoryKey, String>,
}

/// Extractor for Numbers table data
pub struct TableDataExtractor<'a> {
    bundle: &'a Bundle,
    object_index: &'a ObjectIndex,
}

impl<'a> TableDataExtractor<'a> {
    /// Create a new table data extractor
    pub fn new(bundle: &'a Bundle, object_index: &'a ObjectIndex) -> Self {
        Self {
            bundle,
            object_index,
        }
    }

    /// Extract all tables from the document
    pub fn extract_all_tables(&self) -> Result<Vec<NumbersTable>> {
        let mut tables = Vec::new();
        let mut formula_budget = ProjectionBudget::new();
        let formula_references = build_formula_reference_maps(self.bundle, &mut formula_budget)?;
        self.for_each_table(&mut formula_budget, &formula_references, |table| {
            tables.try_reserve(1).map_err(|_| {
                allocation_error("Numbers extracted table results", tables.len() + 1)
            })?;
            tables.push(table);
            Ok(())
        })?;
        Ok(tables)
    }

    fn for_each_table(
        &self,
        formula_budget: &mut ProjectionBudget,
        formula_references: &FormulaReferenceMaps,
        mut visit: impl FnMut(NumbersTable) -> Result<()>,
    ) -> Result<()> {
        let mut seen_objects = HashSet::new();
        let candidate_count = [TABLE_MODEL_MESSAGE_TYPE, 6_000].into_iter().try_fold(
            0usize,
            |count, message_type| {
                count
                    .checked_add(self.object_index.iter_entries_by_type(message_type).count())
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Numbers table-model candidate count overflows usize".to_owned(),
                        )
                    })
            },
        )?;
        formula_budget.charge_entries(candidate_count)?;
        seen_objects.try_reserve(candidate_count).map_err(|_| {
            allocation_error("Numbers table-model candidate identities", candidate_count)
        })?;

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
                    && let Some(table) = self.extract_table_from_object_with_formula_budget(
                        &resolved,
                        formula_budget,
                        formula_references,
                    )?
                {
                    visit(table)?;
                }
            }
        }
        Ok(())
    }

    /// Extract a single table from a resolved object
    pub fn extract_table_from_object(
        &self,
        object: &ResolvedObjectRef<'_>,
    ) -> Result<Option<NumbersTable>> {
        let mut formula_budget = ProjectionBudget::new();
        let formula_references = build_formula_reference_maps(self.bundle, &mut formula_budget)?;
        self.extract_table_from_object_with_formula_budget(
            object,
            &mut formula_budget,
            &formula_references,
        )
    }

    fn extract_table_from_object_with_formula_budget(
        &self,
        object: &ResolvedObjectRef<'_>,
        formula_budget: &mut ProjectionBudget,
        formula_references: &FormulaReferenceMaps,
    ) -> Result<Option<NumbersTable>> {
        let mut projection_budget = ProbeBudget::new();
        let selected = select_candidate(object.messages, &mut projection_budget, |message| {
            Error::InvalidFormat(format!("Object {:?} {message}", object.id()))
        })?;
        let Some(selected) = selected else {
            return Ok(None);
        };
        let source = object.messages[selected].data.as_slice();

        // Candidate admission above proves whether the selected source uses
        // the strict or historical sparse envelope. Replaying the selected
        // source through the compatibility projection is therefore only a
        // borrowed extraction adapter: no unvalidated candidate can reach
        // this path, and the strict selector remains the authority for modern
        // known fields and framing.
        let decoded =
            numbers_table_cell_storage_codec::decode_table_model_compatibility_with_data_store_and_visitor(
                source,
                projection_budget.options(source),
                &mut (),
            );
        let (projection, report) = match decoded {
            Ok(decoded) => decoded,
            Err(error) => return Err(map_resource_error(error, projection_budget)?),
        };
        projection_budget.charge_report(report)?;
        self.parse_table_model_projection(
            projection.model(),
            projection.data_store(),
            &mut projection_budget,
            formula_budget,
            formula_references,
        )
        .map(Some)
    }

    /// Parse one strictly admitted borrowed model/DataStore projection.
    fn parse_table_model_projection(
        &self,
        table_model: numbers_table_cell_storage_codec::TableModelSnapshot<'_>,
        data_store: numbers_table_cell_storage_codec::DataStoreSnapshot<'_>,
        projection_budget: &mut ProbeBudget,
        formula_budget: &mut ProjectionBudget,
        formula_references: &FormulaReferenceMaps,
    ) -> Result<NumbersTable> {
        let (row_count, column_count) = checked_table_dimensions(
            table_model.number_of_rows(),
            table_model.number_of_columns(),
        )?;
        let mut table =
            NumbersTable::with_dimensions(table_model.table_name(), row_count, column_count)?;

        // Extract string table for cell text values
        let string_table = self.load_string_table(data_store.string_table().identifier())?;

        // Extract formula table for formula cells
        let formula_table =
            self.load_formula_table(data_store.formula_table().identifier(), formula_budget)?;
        let formula_error_table = match data_store.formula_error_table() {
            Some(reference) => self.load_formula_error_table(reference.identifier())?,
            None => Box::default(),
        };

        let rich_text_table = match data_store.rich_text_table() {
            Some(reference) => self.load_rich_text_table(reference.identifier())?,
            None => Box::default(),
        };
        let comment_table = match data_store.comment_storage_table() {
            Some(reference) => self.load_comment_table(reference.identifier())?,
            None => Box::default(),
        };

        // Parse tiles to extract cell data
        let cell_tables = CellTables {
            strings: &string_table,
            formulas: &formula_table,
            formula_errors: &formula_error_table,
            rich_text: &rich_text_table,
            comments: &comment_table,
            formula_references,
        };
        self.parse_tiles(
            data_store.tiles(),
            projection_budget,
            formula_budget,
            &cell_tables,
            &mut table,
        )?;

        Ok(table)
    }

    /// Load a TableDataList from an object reference
    fn load_string_table(&self, object_id: u64) -> Result<StringTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let Some(value) = entry.string_value() else {
                    // Preserve the legacy filter_map behavior: a string-list
                    // entry carrying another union member is wire-valid but
                    // contributes no string sidecar value.
                    return Ok(None);
                };
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers string sidecar", value.len()))?;
                retained.push_str(value);
                Ok(Some(retained))
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::String,
            &mut converter,
        )
    }

    fn load_formula_table(
        &self,
        object_id: u64,
        formula_budget: &mut ProjectionBudget,
    ) -> Result<FormulaTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                let Some(value) = entry.formula() else {
                    // The compatibility route historically ignored entries
                    // without the requested formula union member.
                    return Ok(None);
                };
                FormulaArchiveBytes::from_wire(value, budget).map(Some)
            };
        self.load_table_data_list_entries_with_budget(
            object_id,
            tst::table_data_list::ListType::Formula,
            formula_budget,
            &mut converter,
        )
    }

    fn load_formula_error_table(&self, object_id: u64) -> Result<FormulaErrorTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let Some(value) = entry.string_value() else {
                    // Keep the old filter_map semantics for a sparse or
                    // mixed-union formula-error sidecar.
                    return Ok(None);
                };
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers formula-error sidecar", value.len()))?;
                retained.push_str(value);
                Ok(Some(retained))
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::FormulaError,
            &mut converter,
        )
    }

    fn load_rich_text_table(&self, object_id: u64) -> Result<StringTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let Some(payload_reference) = entry.rich_text_payload() else {
                    // Rich-text sidecars historically ignored entries whose
                    // union did not carry a payload reference. The strict
                    // table-list codec still validated the complete entry
                    // before this compatibility conversion runs.
                    return Ok(None);
                };
                let Some(payload_object) = self
                    .object_index
                    .resolve_ref_id(self.bundle, payload_reference.identifier())?
                else {
                    return Ok(None);
                };

                // A payload object may contain unrelated messages, malformed
                // historical envelopes, or more than one candidate. Preserve
                // the old first-decodable/first-text behavior and continue to
                // the next message when a candidate is unusable.
                let mut extract_text = |storage_id| self.extract_rich_text(storage_id);
                first_rich_text_payload_text(payload_object.messages, &mut extract_text)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::RichTextPayload,
            &mut converter,
        )
    }

    fn load_comment_table(&self, object_id: u64) -> Result<CommentTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                if entry.ref_count() == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers comment entry {} has a zero reference count",
                        entry.key()
                    )));
                }
                let storage_id = entry
                    .comment_storage()
                    .map(|reference| reference.identifier())
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers comment entry {} has no storage reference",
                            entry.key()
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
                // Decode every candidate before publishing the first one.
                // Candidate-local allocation failures are staged so a later
                // malformed payload still retains precedence, while the
                // bounded vector prevents an object with many duplicate
                // payload messages from growing without a finite ceiling.
                let mut comments = Vec::new();
                let mut candidate_count = 0usize;
                let mut candidate_error = None;
                for message in storage_object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
                {
                    candidate_count = candidate_count.saturating_add(1);
                    let decoded =
                        match decode_comment_storage_payload(storage_id, message.data.as_slice()) {
                            Ok(decoded) => decoded,
                            Err(error)
                                if matches!(
                                    &error,
                                    Error::IwaCommon(litchi_iwa_common::Error::Allocation { .. })
                                ) =>
                            {
                                record_first_list_error(&mut candidate_error, error);
                                continue;
                            },
                            Err(error) => return Err(error),
                        };
                    if candidate_count > MAX_COMMENT_STORAGE_CANDIDATES {
                        record_first_list_error(
                            &mut candidate_error,
                            comment_storage_candidate_limit_error(candidate_count),
                        );
                        continue;
                    }
                    if comments.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut candidate_error,
                            comment_storage_allocation_error(
                                "Numbers comment storage candidates",
                                comments.len().saturating_add(1),
                            ),
                        );
                        continue;
                    }
                    comments.push(decoded);
                }
                if candidate_count == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Object {storage_id} has no TSD comment-storage payload"
                    )));
                } else if candidate_count != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "Object {storage_id} has multiple TSD comment-storage payloads"
                    )));
                } else if let Some(error) = candidate_error {
                    return Err(error);
                }
                let (comment, reply_ids) = comments.first().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {storage_id} has no retained TSD comment-storage payload"
                    ))
                })?;
                let author_id = comment
                    .author()
                    .map(|author| AuthorId::from_raw(author.identifier()))
                    .transpose()?;
                // Validate IDs before reserving the output vector so malformed
                // typed IDs retain precedence over candidate-local allocation.
                for identifier in reply_ids {
                    StorageId::from_raw(*identifier).map_err(crate::Error::from)?;
                }
                let mut typed_reply_ids = Vec::new();
                typed_reply_ids
                    .try_reserve_exact(reply_ids.len())
                    .map_err(|_| {
                        comment_storage_allocation_error(
                            "Numbers comment reply identifiers",
                            reply_ids.len(),
                        )
                    })?;
                for identifier in reply_ids {
                    typed_reply_ids
                        .push(StorageId::from_raw(*identifier).map_err(crate::Error::from)?);
                }
                let storage_uuid = comment
                    .storage_uuid()
                    .map(|uuid| Uuid::from_parts(uuid.lower(), uuid.upper()))
                    .transpose()?;
                let source_text = comment.text().unwrap_or_default();
                let mut text = String::new();
                text.try_reserve_exact(source_text.len()).map_err(|_| {
                    comment_storage_allocation_error("Numbers comment text", source_text.len())
                })?;
                text.push_str(source_text);
                Ok(Some(Comment {
                    text,
                    creation_date_seconds: comment.creation_date().map(|date| date.seconds()),
                    author_id,
                    reply_ids: typed_reply_ids.into_boxed_slice(),
                    storage_uuid,
                }))
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::CommentStorage,
            &mut converter,
        )
    }

    fn load_table_data_list_entries<T, C>(
        &self,
        object_id: u64,
        list_type: tst::table_data_list::ListType,
        converter: &mut C,
    ) -> Result<CompactTable<T>>
    where
        C: ListValueConverter<T>,
    {
        let mut budget = ProjectionBudget::new();
        self.load_table_data_list_entries_with_budget(object_id, list_type, &mut budget, converter)
    }

    fn load_table_data_list_entries_with_budget<T, C>(
        &self,
        object_id: u64,
        list_type: tst::table_data_list::ListType,
        budget: &mut ProjectionBudget,
        converter: &mut C,
    ) -> Result<CompactTable<T>>
    where
        C: ListValueConverter<T>,
    {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, object_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table-data-list object {object_id} is missing"
                ))
            })?;
        let expected = list_type as i32;
        let mut selected_values = None;
        let mut selected_keys = None;
        let mut segment_ids = None;
        let mut structural_error = None;
        let mut semantic_error = None;
        for message in resolved
            .messages
            .iter()
            .filter(|message| message.type_ == 6005 || message.type_ == 6201)
        {
            // Probe only the scalar envelope before admitting any entry
            // conversion or referenced-object resolution. The full strict
            // decode below still validates every repeated child and unknown.
            let field_offset = budget.payload_fields;
            let work_offset = budget.payload_work;
            let (probe, probe_report) =
                numbers_table_cell_storage_codec::decode_table_data_list_type_with_report(
                    &message.data,
                    table_data_list_decode_options_with_budget(&message.data, *budget, false),
                )
                .map_err(|error| {
                    table_data_list_decode_error_with_offsets(
                        object_id,
                        list_type,
                        error,
                        budget.references,
                        field_offset,
                        work_offset,
                        budget.staging_text_bytes,
                    )
                })?;
            budget.charge_decode_work(probe_report)?;
            let duplicate_candidate = selected_values.is_some();
            let admitting_candidate = !duplicate_candidate && probe.list_type() == expected;
            let (local_max_entries, max_segment_references) =
                table_list_collection_limits(&message.data);
            let max_entries = if admitting_candidate {
                local_max_entries.min(budget.remaining_entries())
            } else {
                local_max_entries
            };
            let mut staged_budget = *budget;
            let mut reserved_report = None;
            let decode_options = if admitting_candidate {
                // The storage codec invokes converters while it scans list
                // entries. Obtain and charge an exact no-callback report,
                // then reserve the identical callback-pass report before any
                // formula wire preflight or owned copy can run.
                let (_, report) =
                    numbers_table_cell_storage_codec::decode_table_data_list_with_report(
                        &message.data,
                        table_data_list_decode_options_with_budget(&message.data, *budget, true),
                    )
                    .map_err(|error| {
                        table_data_list_decode_error_with_offsets(
                            object_id,
                            list_type,
                            error,
                            budget.references,
                            budget.payload_fields,
                            budget.payload_work,
                            budget.staging_text_bytes,
                        )
                    })?;
                budget.charge_decode_report(report)?;
                let options =
                    table_data_list_decode_options_with_budget(&message.data, *budget, true);
                staged_budget = *budget;
                staged_budget.charge_decode_report(report)?;
                reserved_report = Some(report);
                options
            } else {
                table_data_list_decode_options_with_budget(&message.data, *budget, false)
            };
            let field_offset = budget.payload_fields;
            let work_offset = budget.payload_work;
            let reference_offset = budget.references;
            let text_offset = budget.staging_text_bytes;
            let visitor_budget = if admitting_candidate {
                &mut staged_budget
            } else {
                &mut *budget
            };
            let mut visitor = TypedListVisitor::new(
                converter,
                visitor_budget,
                expected,
                false,
                admitting_candidate,
                max_entries,
                max_segment_references,
            );
            let (snapshot, report) =
                numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
                    &message.data,
                    decode_options,
                    &mut visitor,
                )
                .map_err(|error| {
                    table_data_list_decode_error_with_offsets(
                        object_id,
                        list_type,
                        error,
                        reference_offset,
                        field_offset,
                        work_offset,
                        text_offset,
                    )
                })?;
            let (values, keys, references, callback_structural, callback_semantic) =
                visitor.take_parts();
            if snapshot.list_type() != expected {
                budget.charge_decode_work(report)?;
                continue;
            }
            if selected_values.is_some() {
                budget.charge_decode_work(report)?;
                record_first_list_error(
                    &mut structural_error,
                    Error::InvalidFormat(format!(
                        "Object {object_id} has multiple Numbers {list_type:?} TableDataList payloads"
                    )),
                );
                continue;
            }
            if let Some(reserved) = reserved_report {
                if report != reserved {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers {list_type:?} TableDataList report changed between preflight and callback"
                    )));
                }
                staged_budget.charge_entries(keys.len())?;
                *budget = staged_budget;
            } else {
                budget.charge_decode_report(report)?;
                budget.charge_entries(keys.len())?;
            }
            if let Some(error) = callback_structural {
                record_first_list_error(&mut structural_error, error);
            }
            if let Some(error) = callback_semantic {
                record_first_list_error(&mut semantic_error, error);
            }
            selected_values = Some(values);
            selected_keys = Some(keys);
            segment_ids = Some(references);
        }

        let mut values = selected_values.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {object_id} has no Numbers {list_type:?} TableDataList payload"
            ))
        })?;
        let mut keys = selected_keys.unwrap_or_default();
        let segment_ids = segment_ids.unwrap_or_default();

        for segment_id in segment_ids {
            let segment_object = match self.object_index.resolve_ref_id(self.bundle, segment_id) {
                Ok(Some(object)) => object,
                Ok(None) => {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment object {segment_id} is missing"
                        )),
                    );
                    continue;
                },
                Err(error) => {
                    record_first_list_error(&mut structural_error, error);
                    continue;
                },
            };
            let mut segment_count = 0usize;
            for segment_message in segment_object
                .messages
                .iter()
                .filter(|message| message.type_ == 6011)
            {
                segment_count = segment_count.saturating_add(1);
                let field_offset = budget.payload_fields;
                let work_offset = budget.payload_work;
                let (probe, probe_report) = numbers_table_cell_storage_codec::decode_table_data_list_segment_type_with_report(
                    &segment_message.data,
                    table_data_list_decode_options_with_budget(
                        &segment_message.data,
                        *budget,
                        false,
                    ),
                )
                .map_err(|error| {
                    table_data_list_decode_error_with_offsets(
                        segment_id,
                        list_type,
                        error,
                        budget.references,
                        field_offset,
                        work_offset,
                        budget.staging_text_bytes,
                    )
                })?;
                budget.charge_decode_work(probe_report)?;
                let admitting_segment = segment_count == 1 && probe.list_type() == expected;
                let (local_max_entries, max_segment_references) =
                    table_list_collection_limits(&segment_message.data);
                let max_entries = if admitting_segment {
                    local_max_entries.min(budget.remaining_entries())
                } else {
                    local_max_entries
                };
                let mut staged_budget = *budget;
                let mut reserved_report = None;
                let decode_options = if admitting_segment {
                    let (_, report) = numbers_table_cell_storage_codec::decode_table_data_list_segment_with_report(
                        &segment_message.data,
                        table_data_list_decode_options_with_budget(
                            &segment_message.data,
                            *budget,
                            true,
                        ),
                    )
                    .map_err(|error| {
                        table_data_list_decode_error_with_offsets(
                            segment_id,
                            list_type,
                            error,
                            budget.references,
                            budget.payload_fields,
                            budget.payload_work,
                            budget.staging_text_bytes,
                        )
                    })?;
                    budget.charge_decode_report(report)?;
                    let options = table_data_list_decode_options_with_budget(
                        &segment_message.data,
                        *budget,
                        true,
                    );
                    staged_budget = *budget;
                    staged_budget.charge_decode_report(report)?;
                    reserved_report = Some(report);
                    options
                } else {
                    table_data_list_decode_options_with_budget(
                        &segment_message.data,
                        *budget,
                        false,
                    )
                };
                let field_offset = budget.payload_fields;
                let work_offset = budget.payload_work;
                let reference_offset = budget.references;
                let text_offset = budget.staging_text_bytes;
                let visitor_budget = if admitting_segment {
                    &mut staged_budget
                } else {
                    &mut *budget
                };
                let mut visitor = TypedListVisitor::new(
                    converter,
                    visitor_budget,
                    expected,
                    true,
                    admitting_segment,
                    max_entries,
                    max_segment_references,
                );
                let (snapshot, report) =
                    numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
                        &segment_message.data,
                        decode_options,
                        &mut visitor,
                    )
                    .map_err(|error| {
                        table_data_list_decode_error_with_offsets(
                            segment_id,
                            list_type,
                            error,
                            reference_offset,
                            field_offset,
                            work_offset,
                            text_offset,
                        )
                    })?;
                let bounds = visitor.take_segment_bounds();
                let (
                    segment_values,
                    segment_keys,
                    _segment_refs,
                    callback_structural,
                    callback_semantic,
                ) = visitor.take_parts();
                if segment_count > 1 {
                    budget.charge_decode_work(report)?;
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Object {segment_id} has multiple Numbers TableDataListSegment payloads"
                        )),
                    );
                    continue;
                }
                if snapshot.list_type() != expected {
                    budget.charge_decode_work(report)?;
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} has list type {}, expected {list_type:?}",
                            snapshot.list_type()
                        )),
                    );
                    continue;
                }
                if let Some(reserved) = reserved_report {
                    if report != reserved {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers {list_type:?} TableDataListSegment report changed between preflight and callback"
                        )));
                    }
                    staged_budget.charge_entries(segment_keys.len())?;
                    *budget = staged_budget;
                } else {
                    budget.charge_decode_report(report)?;
                    budget.charge_entries(segment_keys.len())?;
                }
                if let Some(error) = callback_structural {
                    record_first_list_error(&mut structural_error, error);
                }
                if let Some(error) = callback_semantic {
                    record_first_list_error(&mut semantic_error, error);
                }
                let end = snapshot
                    .key_range_location()
                    .checked_add(snapshot.key_range_length())
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} key range overflows"
                        ))
                    })?;
                if let (Some(minimum), Some(maximum)) = bounds
                    && (minimum < snapshot.key_range_location() || maximum >= end)
                {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} contains an entry outside its key range"
                        )),
                    );
                    continue;
                }
                for (key, value) in segment_values {
                    if keys.contains(&key) {
                        record_first_list_error(
                            &mut structural_error,
                            Error::InvalidFormat(format!(
                                "Numbers {list_type:?} table {object_id} repeats entry key {key} across root and segments"
                            )),
                        );
                        continue;
                    }
                    if keys.len() >= MAX_TABLE_LIST_ENTRIES {
                        record_first_list_error(
                            &mut semantic_error,
                            table_list_entry_limit_error(keys.len().saturating_add(1)),
                        );
                        continue;
                    }
                    if keys.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error(
                                "Numbers table-list entry keys",
                                keys.len().saturating_add(1),
                            ),
                        );
                        continue;
                    }
                    keys.insert(key);
                    if values.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error(
                                "Numbers table-list entries",
                                values.len().saturating_add(1),
                            ),
                        );
                        continue;
                    }
                    values.push((key, value));
                }
            }
            if segment_count == 0 {
                record_first_list_error(
                    &mut structural_error,
                    Error::InvalidFormat(format!(
                        "Object {segment_id} has no Numbers TableDataListSegment payload"
                    )),
                );
            }
        }
        if let Some(error) = structural_error {
            return Err(error);
        }
        if let Some(error) = semantic_error {
            return Err(error);
        }
        compact_table_vec(values)
    }

    /// Parse tile storage to extract cells
    fn parse_tiles(
        &self,
        tile_storage_source: &[u8],
        projection_budget: &mut ProbeBudget,
        formula_budget: &mut ProjectionBudget,
        cell_tables: &CellTables<'_>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        let mut tile_routes = TileRoutes::default();
        let decoded = numbers_table_cell_storage_codec::decode_tile_storage_with_visitor(
            tile_storage_source,
            projection_budget.options(tile_storage_source),
            &mut tile_routes,
        );
        let (tile_storage, report) = match decoded {
            Ok(decoded) => decoded,
            Err(error) => return Err(map_resource_error(error, *projection_budget)?),
        };
        projection_budget.charge_report(report)?;

        let tile_size = usize::try_from(tile_storage.tile_size().unwrap_or(256)).map_err(|_| {
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
        let dimensions = Dimensions::try_from_usize(table.row_count(), table.column_count())
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut seen_tile_ids = HashSet::new();
        seen_tile_ids
            .try_reserve(tile_routes.routes.len())
            .map_err(|_| allocation_error("Numbers tile keys", tile_routes.routes.len()))?;
        let mut budget = CellBudget::new();
        // Resolve each tile reference and parse its contents
        for tile_route in tile_routes.routes {
            let tile_key = usize::try_from(tile_route.tile_id).map_err(|_| {
                Error::InvalidFormat("Numbers tile key does not fit the host usize".to_owned())
            })?;
            if tile_key >= tile_count {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile key {tile_key} is outside the declared table height {}",
                    table.row_count()
                )));
            }
            if !seen_tile_ids.insert(tile_route.tile_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table repeats tile key {tile_key}"
                )));
            }
            let row_origin = tile_key
                .checked_mul(tile_size)
                .ok_or_else(|| Error::ParseError("Numbers tile row origin overflow".to_owned()))?;
            self.parse_tile(
                tile_route.object_identifier,
                row_origin,
                tile_size,
                dimensions,
                &mut budget,
                projection_budget,
                formula_budget,
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
        dimensions: Dimensions,
        budget: &mut CellBudget,
        projection_budget: &mut ProbeBudget,
        formula_budget: &mut ProjectionBudget,
        cell_tables: &CellTables<'_>,
        table: &mut NumbersTable,
    ) -> Result<()> {
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, tile_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} referenced by table is missing"
                ))
            })?;
        let mut selected = None;
        for msg in resolved.messages {
            if msg.type_ != TILE_MESSAGE_TYPE {
                continue;
            }
            if selected.replace(msg).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile object {tile_id} contains multiple tile payloads"
                )));
            }
        }
        let source = selected
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers tile object {tile_id} has no tile payload"))
            })?
            .data
            .as_slice();
        let mut stage = TileRowStage::new(
            row_origin,
            tile_size,
            dimensions,
            budget,
            formula_budget,
            cell_tables,
        );
        let decoded = numbers_table_cell_storage_codec::decode_tile_with_visitor(
            source,
            projection_budget.options(source),
            &mut stage,
        );
        let (_tile, report) = match decoded {
            Ok(decoded) => decoded,
            Err(error) => return Err(map_resource_error(error, *projection_budget)?),
        };
        projection_budget.charge_report(report)?;

        for cell in stage.finish()? {
            table.try_set_cell(cell.row, cell.column, cell.parsed.value)?;
            if let Some(identifier) = cell.parsed.comment_identifier {
                let comment = compact_table_get(cell_tables.comments, identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers comment table has no entry {identifier} referenced by cell ({}, {})",
                        cell.row, cell.column
                    ))
                })?;
                table.try_set_comment(cell.row, cell.column, comment.clone())?;
            }
        }

        Ok(())
    }

    /// Parse cell offsets from the offsets buffer
    ///
    /// The offset table is an array of little-endian `u16` values. `0xffff`
    /// marks a missing column; wide rows store offsets in four-byte units.
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
        if slot_count > column_count {
            let populated_outside_width = offsets_buffer[column_count.saturating_mul(2)..]
                .chunks_exact(2)
                .any(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX);
            if populated_outside_width {
                return Err(Error::InvalidFormat(format!(
                    "Numbers row has a populated offset slot outside table width {column_count}"
                )));
            }
        }
        if expected_cells > slot_count {
            return Err(Error::ParseError(format!(
                "Numbers row declares {expected_cells} cells but has only {slot_count} offset slots"
            )));
        }

        let present_cells = offsets_buffer
            .chunks_exact(2)
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
        for (column, bytes) in offsets_buffer.chunks_exact(2).enumerate() {
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

    #[cfg(test)]
    fn parse_cell_storage(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        row: usize,
        column: usize,
    ) -> Result<ParsedCell> {
        let mut formula_budget = ProjectionBudget::new();
        Self::parse_cell_storage_with_budget(
            data,
            cell_tables,
            row,
            column,
            row.saturating_add(1),
            column.saturating_add(1),
            &mut formula_budget,
        )
    }

    fn parse_cell_storage_with_budget(
        data: &[u8],
        cell_tables: &CellTables<'_>,
        row: usize,
        column: usize,
        row_count: usize,
        column_count: usize,
        formula_budget: &mut ProjectionBudget,
    ) -> Result<ParsedCell> {
        let version = *data
            .first()
            .ok_or_else(|| Error::ParseError("Empty Numbers cell storage".to_string()))?;
        match version {
            0..=4 => Self::parse_pre_bnc_cell(
                data,
                cell_tables,
                row,
                column,
                row_count,
                column_count,
                formula_budget,
            ),
            5 => Self::parse_bnc_cell(
                data,
                cell_tables,
                row,
                column,
                row_count,
                column_count,
                formula_budget,
            ),
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
        row_count: usize,
        column_count: usize,
        formula_budget: &mut ProjectionBudget,
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
            let rendered = render_formula_string(
                formula,
                row,
                column,
                row_count,
                column_count,
                cell_tables.formula_references,
                formula_budget,
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
                Some(CachedScalar::Number(value)) => CellValue::Number(to_cell_finite(value)?),
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
                Some(CachedScalar::Date(value)) => CellValue::Date(to_cell_finite(value)?),
                Some(_) | None => CellValue::Date(zero),
            },
            StoredValue::Boolean => match scalar {
                Some(CachedScalar::Boolean(value)) => CellValue::Boolean(value),
                Some(_) | None => CellValue::Boolean(false),
            },
            StoredValue::Duration => match scalar {
                Some(CachedScalar::Duration(value)) => CellValue::Duration(to_cell_finite(value)?),
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
        row_count: usize,
        column_count: usize,
        formula_budget: &mut ProjectionBudget,
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
            let rendered = render_formula_string(
                formula,
                row,
                column,
                row_count,
                column_count,
                cell_tables.formula_references,
                formula_budget,
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
    #[cfg(test)]
    fn extract_formula_string(
        formula: &tsce::FormulaArchive,
        host_row: usize,
        host_column: usize,
        formula_references: &FormulaReferenceMaps,
    ) -> Result<String> {
        use crate::protobuf::tsce::ast_node_array_archive::AstNodeType;

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
    #[cfg(test)]
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
    #[cfg(test)]
    fn get_function_name(index: u32) -> String {
        super::function_map::function_name(index)
            .map(str::to_owned)
            .unwrap_or_else(|| format!("FUNC{index}"))
    }

    /// Extract rich text from a storage reference
    fn extract_rich_text(&self, storage_id: u64) -> Result<Option<String>> {
        if let Some(resolved) = self.object_index.resolve_ref_id(self.bundle, storage_id)? {
            // Look for TSWP.StorageArchive messages.  The compatibility
            // extractor intentionally keeps the first usable candidate
            // behavior, but the selected storage envelope itself goes through
            // the bounded raw/Buffa projection rather than constructing the
            // generated archive just to read field 3.
            for msg in resolved.messages {
                if msg.type_ >= 2001
                    && msg.type_ <= 2022
                    && let Ok(Some(text)) = compatibility_storage_text(&msg.data)
                {
                    return Ok(Some(text));
                }
            }
        }

        Ok(None)
    }
}

/// Build the finite profile used by the legacy host's rich-text storage
/// compatibility path.  The source byte length is a conservative bound for
/// each selected resource, while the text-wire adapter retains non-bypassable
/// hard ceilings for hostile inputs.
fn compatibility_storage_limits(
    source: &[u8],
) -> litchi_iwa_text_wire::Result<litchi_iwa_text_wire::Limits> {
    let source_bytes = source.len().max(1);
    litchi_iwa_text_wire::Limits::new(
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_MESSAGE_BYTES),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_FIELDS),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_FRAGMENTS),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_TEXT_BYTES),
    )
}

/// Project one host rich-text storage envelope without retaining a generated
/// `TSWP.StorageArchive`.  Caller-owned bytes remain the preservation
/// authority; the semantic output retains the legacy newline between every
/// native field-3 fragment, including empty fragments.
fn compatibility_storage_text_with_limits(
    source: &[u8],
    limits: litchi_iwa_text_wire::Limits,
) -> litchi_iwa_text_wire::Result<Option<String>> {
    let storage = litchi_iwa_text_wire::from_bytes_with_limits(source, limits)?;
    // Prost's legacy path distinguishes an absent repeated field from a
    // present-but-empty fragment (`Vec::is_empty`, not concatenated text
    // emptiness). Keep that compatibility distinction in the projection.
    if storage.runs().is_empty() {
        return Ok(None);
    }

    let separator_count = storage.runs().len().saturating_sub(1);
    let output_len = storage
        .text()
        .len()
        .checked_add(separator_count)
        .ok_or(litchi_iwa_text_wire::Error::TextLengthOverflow)?;
    if output_len > limits.max_text_bytes() {
        return Err(litchi_iwa_text_wire::Error::TooManyTextBytes {
            actual: output_len,
            limit: limits.max_text_bytes(),
        });
    }

    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_allocation| litchi_iwa_common::Error::Allocation {
            resource: "Numbers compatibility rich-text output",
            amount: output_len,
        })?;

    for (index, run) in storage.runs().iter().copied().enumerate() {
        if index != 0 {
            output.push('\n');
        }
        let end = run.end().ok_or(litchi_iwa_text_wire::Error::Storage(
            litchi_iwa_text::storage::Error::RunOutOfBounds { index },
        ))?;
        let fragment =
            storage
                .text()
                .get(run.start()..end)
                .ok_or(litchi_iwa_text_wire::Error::Storage(
                    litchi_iwa_text::storage::Error::RunNotOnBoundary { index },
                ))?;
        output.push_str(fragment);
    }

    debug_assert_eq!(output.len(), output_len);
    Ok(Some(output))
}

/// Compatibility wrapper that maps strict projection failures to the host's
/// existing content-free format error.  Rich-text candidates are best effort,
/// so the caller may continue to a later candidate after a failure.
fn compatibility_storage_text(source: &[u8]) -> Result<Option<String>> {
    let limits = compatibility_storage_limits(source).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers rich-text storage limits are invalid: {error}"
        ))
    })?;
    compatibility_storage_text_with_limits(source, limits).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers rich-text storage failed strict projection: {error}"
        ))
    })
}

fn first_rich_text_payload_text(
    messages: &[crate::archive::RawMessage],
    extract_text: &mut impl FnMut(u64) -> Result<Option<String>>,
) -> Result<Option<String>> {
    for payload_message in messages {
        let Ok(payload) = tst::RichTextPayloadArchive::decode(payload_message.data.as_slice())
        else {
            continue;
        };
        match extract_text(payload.storage.identifier) {
            Ok(Some(text)) => return Ok(Some(text)),
            // Rich-text payloads are compatibility sidecars. A malformed or
            // missing storage candidate must not prevent a later payload from
            // supplying the cell text.
            Ok(None) | Err(_) => continue,
        }
    }
    Ok(None)
}

fn build_formula_reference_maps(
    bundle: &Bundle,
    budget: &mut ProjectionBudget,
) -> Result<FormulaReferenceMaps> {
    // Build one operation-local identifier index instead of rescanning every
    // archive for each sheet, drawable, and model edge. Charge and reserve the
    // complete object inventory before retaining any borrowed entries.
    let object_count = bundle
        .iter_archives()
        .try_fold(0usize, |count, (_, archive)| {
            count.checked_add(archive.objects.len()).ok_or_else(|| {
                Error::InvalidFormat("Numbers formula object inventory overflows usize".to_owned())
            })
        })?;
    budget.charge_entries(object_count)?;
    budget.charge_payload_work(object_count)?;
    let mut objects = HashMap::<u64, &crate::archive::ArchiveObject>::new();
    objects
        .try_reserve(object_count)
        .map_err(|_| allocation_error("Numbers formula object inventory", object_count))?;
    for (_, archive) in bundle.iter_archives() {
        for object in &archive.objects {
            if let Some(identifier) = object.archive_info.identifier {
                objects.entry(identifier).or_insert(object);
            }
        }
    }

    let mut result = FormulaReferenceMaps::default();
    budget.charge_entries(1)?;
    budget.charge_staging_text("Grand Total".len())?;
    result
        .categories
        .try_reserve(1)
        .map_err(|_| allocation_error("Numbers formula category names", 1))?;
    result.categories.insert([1, 0], "Grand Total".to_owned());
    let mut table_info_names = HashMap::<u64, FormulaReferenceName>::new();
    let mut root = None;
    if let Some(object) = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
    {
        for message in &object.messages {
            budget.charge_payload_fields(message.data.len())?;
            budget.charge_payload_work(message.data.len())?;
            if let Ok(document) = tn::DocumentArchive::decode(message.data.as_slice()) {
                root = Some(document);
                break;
            }
        }
    }

    if let Some(root) = root {
        for sheet_reference in root.sheets {
            let Some(sheet_object) = objects.get(&sheet_reference.identifier).copied() else {
                continue;
            };
            let mut sheet = None;
            for message in &sheet_object.messages {
                budget.charge_payload_fields(message.data.len())?;
                budget.charge_payload_work(message.data.len())?;
                if let Ok(decoded) = tn::SheetArchive::decode(message.data.as_slice()) {
                    sheet = Some(decoded);
                    break;
                }
                if let Ok(form) = tn::FormBasedSheetArchive::decode(message.data.as_slice()) {
                    sheet = Some(form.super_);
                    break;
                }
            }
            let Some(sheet) = sheet else {
                continue;
            };
            for drawable in sheet.drawable_infos {
                let Some(drawable_object) = objects.get(&drawable.identifier).copied() else {
                    continue;
                };
                let mut table_name = None;
                for message in &drawable_object.messages {
                    budget.charge_payload_fields(message.data.len())?;
                    budget.charge_payload_work(message.data.len())?;
                    let Ok(table_info) = tst::TableInfoArchive::decode(message.data.as_slice())
                    else {
                        continue;
                    };
                    let Some(model_object) =
                        objects.get(&table_info.table_model.identifier).copied()
                    else {
                        continue;
                    };
                    for model_message in &model_object.messages {
                        budget.charge_payload_work(model_message.data.len())?;
                        if model_message.type_ != 6000 && model_message.type_ != 6001 {
                            continue;
                        }
                        budget.charge_payload_fields(model_message.data.len())?;
                        budget.charge_payload_work(model_message.data.len().saturating_mul(3))?;
                        let Ok(model) = numbers_names_codec::decode_table_names(
                            model_message.data.as_slice(),
                            table_name_decode_options(model_message.data.as_slice()),
                        ) else {
                            continue;
                        };
                        let name = model.table_name();
                        budget.charge_staging_text(name.len())?;
                        let mut owned = String::new();
                        owned.try_reserve_exact(name.len()).map_err(|_| {
                            allocation_error("Numbers formula table name", name.len())
                        })?;
                        owned.push_str(name);
                        table_name = Some(owned);
                        break;
                    }
                    if table_name.is_some() {
                        break;
                    }
                }
                if let Some(table) = table_name {
                    budget.charge_entries(1)?;
                    budget.charge_staging_text(sheet.name.len())?;
                    table_info_names.try_reserve(1).map_err(|_| {
                        allocation_error(
                            "Numbers formula table reference names",
                            table_info_names.len().saturating_add(1),
                        )
                    })?;
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
                budget.charge_payload_work(1)?;
                if message.type_ == 6383 {
                    budget.charge_payload_fields(message.data.len())?;
                    budget.charge_payload_work(message.data.len())?;
                    if let Ok(group_node) =
                        tst::group_by_archive::GroupNodeArchive::decode(message.data.as_slice())
                    {
                        collect_formula_category_names(
                            &group_node,
                            &mut result.categories,
                            budget,
                            1,
                        )?;
                    }
                    continue;
                }
                if message.type_ != 4008 {
                    continue;
                }
                budget.charge_payload_fields(message.data.len())?;
                budget.charge_payload_work(message.data.len())?;
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
                budget.charge_entries(1)?;
                budget.charge_staging_text(name.sheet.len())?;
                budget.charge_staging_text(name.table.len())?;
                result.owners.try_reserve(1).map_err(|_| {
                    allocation_error(
                        "Numbers formula owner names",
                        result.owners.len().saturating_add(1),
                    )
                })?;
                result
                    .owners
                    .insert(formula_owner_key(&owner.formula_owner_uid), name.clone());
            }
        }
    }
    Ok(result)
}

fn collect_formula_category_names(
    node: &tst::group_by_archive::GroupNodeArchive,
    names: &mut HashMap<FormulaCategoryKey, String>,
    budget: &mut ProjectionBudget,
    depth: usize,
) -> Result<()> {
    budget.check_formula_render_depth(depth)?;
    budget.charge_payload_work(1)?;
    if let Some(value) = node
        .group_cell_value
        .as_ref()
        .and_then(group_cell_value_label)
    {
        budget.charge_entries(1)?;
        budget.charge_staging_text(value.len())?;
        names.try_reserve(1).map_err(|_| {
            allocation_error(
                "Numbers formula category names",
                names.len().saturating_add(1),
            )
        })?;
        names.insert(formula_category_key(&node.group_uid), value);
    }
    for child in &node.child {
        let child_depth = depth.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("Numbers formula category depth overflows usize".to_owned())
        })?;
        collect_formula_category_names(child, names, budget, child_depth)?;
    }
    Ok(())
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

#[cfg(test)]
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

fn formula_owner_key(owner: &crate::protobuf::tsp::Uuid) -> FormulaOwnerKey {
    [
        owner.lower as u32,
        (owner.lower >> 32) as u32,
        owner.upper as u32,
        (owner.upper >> 32) as u32,
    ]
}

fn formula_category_key(category: &crate::protobuf::tsp::Uuid) -> FormulaCategoryKey {
    [category.lower, category.upper]
}

#[cfg(test)]
fn cfuuid_key(owner: &crate::protobuf::tsp::CfuuidArchive) -> Option<FormulaOwnerKey> {
    Some([
        owner.uuid_w0?,
        owner.uuid_w1?,
        owner.uuid_w2?,
        owner.uuid_w3?,
    ])
}

#[cfg(test)]
fn formula_reference_prefix(
    owner: &crate::protobuf::tsp::CfuuidArchive,
    references: &FormulaReferenceMaps,
) -> String {
    cfuuid_key(owner)
        .and_then(|key| references.owners.get(&key))
        .map(|name| format!("{}::{}::", name.sheet, name.table))
        .unwrap_or_else(|| "Table::".to_owned())
}

#[cfg(test)]
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

#[cfg(test)]
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

#[cfg(test)]
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

#[cfg(test)]
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

#[cfg(test)]
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

fn to_cell_finite(value: CommonFiniteF64) -> Result<FiniteF64> {
    FiniteF64::new(value.get()).map_err(|_| {
        Error::ParseError("Numbers BNC cached scalar must contain a finite value".to_owned())
    })
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
    use super::*;

    #[test]
    fn tracked_native_numbers_fixture_streams_model_store_and_tiles() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers");
        let bundle = Bundle::open(path).unwrap();
        let index = ObjectIndex::from_bundle(&bundle).unwrap();
        let tables = TableDataExtractor::new(&bundle, &index)
            .extract_all_tables()
            .unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name(), "Table 1");
        assert_eq!(tables[0].dimensions(), (22, 7));
        assert!(tables[0].cell_count() > 0);
    }

    #[test]
    fn formula_stack_helpers_reject_underflow() {
        let mut stack = vec!["1".to_owned()];
        assert!(pop_binary_operands(&mut stack, "addition").is_err());
        let mut stack = vec!["1".to_owned()];
        assert!(pop_formula_arguments(&mut stack, 2, "function").is_err());
    }

    #[test]
    fn pivot_category_references_preserve_expression_stack_shape() {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::CategoryRefNode as i32,
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::CategoryRefNode as i32,
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::DivisionNode as i32,
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        };
        assert_eq!(
            TableDataExtractor::extract_formula_string(
                &formula,
                0,
                0,
                &FormulaReferenceMaps::default(),
            )
            .unwrap(),
            "=(#CATEGORY!/#CATEGORY!)"
        );
    }

    #[test]
    fn pivot_category_references_render_group_node_labels() {
        use crate::protobuf::tsp::Uuid;
        use tsce::ast_node_array_archive::{
            AstCategoryReferenceArchive, AstNodeArchive, AstNodeType,
        };

        let category_node = |uid: Uuid| AstNodeArchive {
            ast_node_type: AstNodeType::CategoryRefNode as i32,
            ast_category_ref: Some(AstCategoryReferenceArchive {
                category_ref: tsce::CategoryReferenceArchive {
                    absolute_group_uid: Some(uid),
                    ..Default::default()
                },
            }),
            ..Default::default()
        };
        let north = Uuid {
            lower: 7,
            upper: 11,
        };
        let grand_total = Uuid { lower: 1, upper: 0 };
        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    category_node(north),
                    category_node(grand_total),
                    AstNodeArchive {
                        ast_node_type: AstNodeType::DivisionNode as i32,
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        };
        let references = FormulaReferenceMaps {
            categories: HashMap::from([
                (formula_category_key(&north), "North]west".to_owned()),
                (formula_category_key(&grand_total), "Grand Total".to_owned()),
            ]),
            ..Default::default()
        };
        assert_eq!(
            TableDataExtractor::extract_formula_string(&formula, 0, 0, &references).unwrap(),
            "=(#CATEGORY![North\\]west]/#CATEGORY![Grand Total])"
        );
    }

    #[test]
    fn modern_formula_coordinates_are_rendered_at_the_host_cell() {
        use tsce::ast_node_array_archive::{
            AstColumnCoordinateArchive, AstNodeArchive, AstNodeType, AstRowCoordinateArchive,
        };

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![AstNodeArchive {
                    ast_node_type: AstNodeType::CellReferenceNode as i32,
                    ast_column: Some(AstColumnCoordinateArchive {
                        column: -1,
                        absolute: Some(false),
                    }),
                    ast_row: Some(AstRowCoordinateArchive {
                        row: 0,
                        absolute: Some(false),
                    }),
                    ..Default::default()
                }],
            },
            ..Default::default()
        };
        assert_eq!(
            TableDataExtractor::extract_formula_string(
                &formula,
                2,
                1,
                &FormulaReferenceMaps::default(),
            )
            .unwrap(),
            "=A3"
        );
    }

    #[test]
    fn colon_tract_ranges_render_relative_absolute_and_cross_table_coordinates() {
        use tsce::ast_node_array_archive::ast_colon_tract_archive::{
            AstColonTractAbsoluteRangeArchive, AstColonTractRelativeRangeArchive,
        };
        use tsce::ast_node_array_archive::{
            AstColonTractArchive, AstCrossTableReferenceExtraInfoArchive, AstNodeArchive,
            AstNodeType, AstStickyBits,
        };

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::ColonTractNode as i32,
                        ast_sticky_bits: Some(AstStickyBits {
                            begin_row_is_absolute: false,
                            begin_column_is_absolute: true,
                            end_row_is_absolute: true,
                            end_column_is_absolute: false,
                        }),
                        ast_colon_tract: Some(AstColonTractArchive {
                            relative_column: vec![AstColonTractRelativeRangeArchive {
                                range_begin: -1,
                                range_end: None,
                            }],
                            relative_row: vec![AstColonTractRelativeRangeArchive {
                                range_begin: -3,
                                range_end: None,
                            }],
                            absolute_column: vec![AstColonTractAbsoluteRangeArchive {
                                range_begin: 0,
                                range_end: None,
                            }],
                            absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                                range_begin: 1,
                                range_end: None,
                            }],
                            preserve_rectangular: Some(true),
                        }),
                        ast_cross_table_reference_extra_info: Some(
                            AstCrossTableReferenceExtraInfoArchive::default(),
                        ),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::FunctionNode as i32,
                        ast_function_node_index: Some(168),
                        ast_function_node_num_args: Some(1),
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        };
        assert_eq!(
            TableDataExtractor::extract_formula_string(
                &formula,
                3,
                2,
                &FormulaReferenceMaps::default(),
            )
            .unwrap(),
            "=SUM(Table::$A1:B$2)"
        );

        let whole_rows = AstNodeArchive {
            ast_node_type: AstNodeType::ColonTractNode as i32,
            ast_sticky_bits: Some(AstStickyBits {
                begin_row_is_absolute: false,
                begin_column_is_absolute: false,
                end_row_is_absolute: false,
                end_column_is_absolute: false,
            }),
            ast_colon_tract: Some(AstColonTractArchive {
                relative_row: vec![AstColonTractRelativeRangeArchive {
                    range_begin: -1,
                    range_end: Some(0),
                }],
                absolute_column: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: i16::MAX as u32,
                    range_end: None,
                }],
                preserve_rectangular: Some(true),
                ..Default::default()
            }),
            ast_cross_table_reference_extra_info: Some(
                AstCrossTableReferenceExtraInfoArchive::default(),
            ),
            ..Default::default()
        };
        assert_eq!(
            render_colon_tract(&whole_rows, 1, 0, &FormulaReferenceMaps::default()).unwrap(),
            "Table::1:2"
        );

        let whole_columns = AstNodeArchive {
            ast_colon_tract: Some(AstColonTractArchive {
                relative_column: vec![AstColonTractRelativeRangeArchive {
                    range_begin: 1,
                    range_end: Some(2),
                }],
                absolute_row: vec![AstColonTractAbsoluteRangeArchive {
                    range_begin: i32::MAX as u32,
                    range_end: None,
                }],
                preserve_rectangular: Some(true),
                ..Default::default()
            }),
            ..whole_rows
        };
        assert_eq!(
            render_colon_tract(&whole_columns, 0, 0, &FormulaReferenceMaps::default()).unwrap(),
            "Table::B:C"
        );
    }

    #[test]
    fn standalone_local_reference_uses_compatibility_sticky_semantics() {
        use tsce::ast_node_array_archive::{
            AstLocalCellReferenceNodeArchive, AstNodeArchive, AstNodeType,
        };

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![AstNodeArchive {
                    ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
                    ast_local_cell_reference_node_reference: Some(
                        AstLocalCellReferenceNodeArchive {
                            row_handle: 2,
                            column_handle: 3,
                            row_is_sticky: 1,
                            column_is_sticky: 1,
                        },
                    ),
                    ..Default::default()
                }],
            },
            ..Default::default()
        };
        let source = formula.encode_to_vec();
        let mut admission_budget = ProjectionBudget::new();
        let retained = FormulaArchiveBytes::from_wire(&source, &mut admission_budget).unwrap();

        // The scalar codec renders this node with sticky flags, whereas the
        // historical generated renderer intentionally drops them. The
        // generated-free compatibility visitor must therefore handle it.
        let mut render_budget = ProjectionBudget::new();
        let actual = render_formula_string(
            &retained,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut render_budget,
        )
        .unwrap();
        let expected = TableDataExtractor::extract_formula_string(
            &formula,
            0,
            0,
            &FormulaReferenceMaps::default(),
        )
        .unwrap();
        assert_eq!(actual, expected);
        assert_eq!(actual, "=D3");
    }

    #[test]
    fn empty_formula_wire_preserves_legacy_default_render() {
        let mut admission_budget = ProjectionBudget::new();
        let retained = FormulaArchiveBytes::from_wire(&[], &mut admission_budget).unwrap();
        let mut render_budget = ProjectionBudget::new();
        let rendered = render_formula_string(
            &retained,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut render_budget,
        )
        .unwrap();
        assert_eq!(rendered, "=");

        // Prost omits an empty nested message, so the generated default
        // archive serializes to the same zero-byte compatibility form.  Keep
        // a separate wire case with the required root field present but an
        // empty AST array to ensure the strict generated-free path does not
        // confuse "present and empty" with "missing".
        let encoded_empty = [0x0a, 0x00];
        let mut encoded_budget = ProjectionBudget::new();
        let retained = FormulaArchiveBytes::from_wire(&encoded_empty, &mut encoded_budget).unwrap();
        let mut encoded_render_budget = ProjectionBudget::new();
        let rendered = render_formula_string(
            &retained,
            0,
            0,
            10,
            10,
            &FormulaReferenceMaps::default(),
            &mut encoded_render_budget,
        )
        .unwrap();
        assert_eq!(rendered, "=");
    }

    #[test]
    fn incomplete_postfix_preserves_explicit_legacy_placeholder_semantics() {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let cases = [
            (
                vec![AstNodeArchive {
                    ast_node_type: AstNodeType::NegationNode as i32,
                    ..Default::default()
                }],
                "=FORMULA()",
            ),
            (
                vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(1.0),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(2.0),
                        ..Default::default()
                    },
                ],
                "=2",
            ),
        ];

        for (nodes, expected) in cases {
            let formula = tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
                ..Default::default()
            };
            let source = formula.encode_to_vec();
            let source_before = source.clone();
            let mut admission_budget = ProjectionBudget::new();
            let retained = FormulaArchiveBytes::from_wire(&source, &mut admission_budget).unwrap();
            let mut render_budget = ProjectionBudget::new();
            let actual = render_formula_string(
                &retained,
                0,
                0,
                10,
                10,
                &FormulaReferenceMaps::default(),
                &mut render_budget,
            )
            .unwrap();
            let legacy = TableDataExtractor::extract_formula_string(
                &formula,
                0,
                0,
                &FormulaReferenceMaps::default(),
            )
            .unwrap();
            assert_eq!(source, source_before);
            assert_eq!(actual, legacy);
            assert_eq!(actual, expected);
        }
    }

    #[test]
    fn scalar_formula_decode_report_is_merged_across_repeated_renders() -> Result<()> {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![AstNodeArchive {
                    ast_node_type: AstNodeType::NumberNode as i32,
                    ast_number_node_number: Some(7.5),
                    ..Default::default()
                }],
            },
            ..Default::default()
        };
        let source = formula.encode_to_vec();
        let source_before = source.clone();
        let mut admission_budget = ProjectionBudget::new();
        let retained = FormulaArchiveBytes::from_wire(&source, &mut admission_budget)?;
        assert_eq!(source, source_before);

        let render = |budget: &mut ProjectionBudget| {
            render_formula_string(
                &retained,
                0,
                0,
                10,
                10,
                &FormulaReferenceMaps::default(),
                budget,
            )
        };

        let mut probe = ProjectionBudget::new();
        let expected = render(&mut probe)?;
        let fields = probe.payload_fields;
        let work = probe.payload_work;
        assert!(fields > 0, "successful scalar decode must report fields");
        assert!(work > 0, "successful scalar decode must report work");
        assert!(probe.formula_render_work > 0);

        let mut fields_budget = ProjectionBudget::new();
        fields_budget.payload_fields = MAX_TABLE_LIST_PAYLOAD_FIELDS - fields;
        assert_eq!(render(&mut fields_budget)?, expected);
        assert_eq!(fields_budget.payload_fields, MAX_TABLE_LIST_PAYLOAD_FIELDS);
        let fields_before_refusal = fields_budget;
        let fields_error = render(&mut fields_budget).unwrap_err();
        assert!(matches!(
            fields_error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                ..
            })
        ));
        assert_eq!(
            fields_budget.payload_fields,
            fields_before_refusal.payload_fields
        );

        let mut work_budget = ProjectionBudget::new();
        work_budget.payload_work = MAX_TABLE_LIST_PAYLOAD_WORK - work;
        assert_eq!(render(&mut work_budget)?, expected);
        assert_eq!(work_budget.payload_work, MAX_TABLE_LIST_PAYLOAD_WORK);
        let work_before_refusal = work_budget;
        let work_error = render(&mut work_budget).unwrap_err();
        assert!(matches!(
            work_error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                ..
            })
        ));
        assert_eq!(work_budget.payload_work, work_before_refusal.payload_work);
        Ok(())
    }

    #[test]
    fn compatibility_formula_decode_report_text_is_merged_across_repeated_renders() -> Result<()> {
        use tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};

        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![AstNodeArchive {
                    ast_node_type: AstNodeType::StringNode as i32,
                    ast_string_node_string: Some("aggregate".to_owned()),
                    ..Default::default()
                }],
            },
            ..Default::default()
        };
        let source = formula.encode_to_vec();
        let source_before = source.clone();
        let mut admission_budget = ProjectionBudget::new();
        let retained = FormulaArchiveBytes::from_wire(&source, &mut admission_budget)?;
        assert_eq!(source, source_before);

        let render = |budget: &mut ProjectionBudget| {
            render_formula_string(
                &retained,
                0,
                0,
                10,
                10,
                &FormulaReferenceMaps::default(),
                budget,
            )
        };

        let mut probe = ProjectionBudget::new();
        let expected = render(&mut probe)?;
        let fields = probe.payload_fields;
        let work = probe.payload_work;
        let text = probe.staging_text_bytes;
        assert!(
            fields > 0,
            "successful compatibility decode must report fields"
        );
        assert!(work > 0, "successful compatibility decode must report work");
        assert!(text > 0, "successful compatibility decode must report text");
        assert!(!expected.is_empty());

        let mut text_budget = ProjectionBudget::new();
        text_budget.staging_text_bytes = MAX_TABLE_LIST_TEXT_BYTES - text;
        assert_eq!(render(&mut text_budget)?, expected);
        assert_eq!(text_budget.staging_text_bytes, MAX_TABLE_LIST_TEXT_BYTES);
        let text_before_refusal = text_budget;
        let text_error = render(&mut text_budget).unwrap_err();
        assert!(
            matches!(
                &text_error,
                Error::InvalidFormat(message) if message.contains("text exceeded")
            ) || matches!(
                text_error,
                Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::OutputBytes,
                    ..
                })
            ),
            "unexpected second-render error: {text_error:?}"
        );
        assert_eq!(
            text_budget.staging_text_bytes,
            text_before_refusal.staging_text_bytes
        );
        Ok(())
    }

    #[test]
    fn test_fixed_width_cell_offsets() {
        let offsets = [
            0xff, 0xff, // column 0 missing
            0x00, 0x00, // column 1 starts at 0
            0x18, 0x00, // column 2 starts at 24
            0x30, 0x00, // column 3 starts at 48
            0xff, 0xff,
        ];
        let cells = TableDataExtractor::parse_cell_offsets(&offsets, 72, false, 3, 5).unwrap();
        assert_eq!(cells, vec![(1, 0..24), (2, 24..48), (3, 48..72)]);
    }

    #[test]
    fn table_dimensions_are_bounded_before_archive_references_are_loaded() {
        assert_eq!(checked_table_dimensions(0, 0).unwrap(), (0, 0));
        assert_eq!(checked_table_dimensions(3, 5).unwrap(), (3, 5));

        let error = checked_table_dimensions((MAX_TABLE_ROWS + 1) as u32, 1).unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::TableRows,
                ..
            })
        ));

        let error = checked_table_dimensions(1, (MAX_TABLE_COLUMNS + 1) as u32).unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::TableColumns,
                ..
            })
        ));

        let error = checked_table_dimensions(1_025, MAX_TABLE_COLUMNS as u32).unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::TableCells,
                ..
            })
        ));
    }

    #[test]
    fn table_coordinates_are_checked_against_declared_dimensions() {
        let dimensions = Dimensions::new(3, 5);
        dimensions.check_row(2).unwrap();
        dimensions.check_column(4).unwrap();
        assert!(dimensions.check_row(3).is_err());
        assert!(dimensions.check_column(5).is_err());
    }

    #[test]
    fn cell_count_is_validated_before_offset_reservation() {
        let error =
            TableDataExtractor::parse_cell_offsets(&[], 0, false, usize::MAX, 0).unwrap_err();
        assert!(matches!(error, Error::ParseError(message) if message.contains("offset slots")));

        let offsets = [0, 0, 1, 0];
        let error = TableDataExtractor::parse_cell_offsets(&offsets, 1, false, 1, 1).unwrap_err();
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("outside table width"))
        );
    }

    #[test]
    fn malformed_and_sparse_cell_offsets_are_handled_strictly() {
        let error = TableDataExtractor::parse_cell_offsets(&[0, 0, 1], 2, false, 1, 2).unwrap_err();
        assert!(matches!(error, Error::ParseError(message) if message.contains("odd")));

        let descending = [0, 0, 2, 0, 1, 0];
        let error =
            TableDataExtractor::parse_cell_offsets(&descending, 3, false, 3, 3).unwrap_err();
        assert!(
            matches!(error, Error::ParseError(message) if message.contains("strictly increasing"))
        );

        let sparse = [0xff, 0xff, 0, 0, 2, 0, 0xff, 0xff];
        let cells = TableDataExtractor::parse_cell_offsets(&sparse, 3, false, 2, 4).unwrap();
        assert_eq!(cells, vec![(1, 0..2), (2, 2..3)]);
    }

    #[test]
    fn test_decimal128_decoding() {
        let mut encoded = [0u8; 16];
        encoded[0] = 35;
        encoded[14] = 0x3e;
        encoded[15] = 0x30;
        assert_eq!(read_decimal128_le(&encoded).unwrap(), 3.5);
    }

    #[test]
    fn compact_sidecars_sort_and_binary_search_without_hash_buckets() {
        let table = compact_table([(9, "nine"), (1, "one"), (5, "five")]).unwrap();
        assert_eq!(
            table.iter().map(|(key, _)| *key).collect::<Vec<_>>(),
            [1, 5, 9]
        );
        assert_eq!(compact_table_get(&table, 5), Some(&"five"));
        assert_eq!(compact_table_get(&table, 7), None);
    }

    #[test]
    fn compact_sidecars_reject_duplicate_keys() {
        let error = compact_table([(7, "first"), (7, "second")]).unwrap_err();
        assert!(matches!(error, Error::InvalidFormat(message) if message.contains("duplicate")));
    }

    #[test]
    fn rich_text_payload_candidates_skip_malformed_and_missing_text() {
        use crate::archive::RawMessage;
        use crate::protobuf::tsp;

        let payload = |storage_id| {
            tst::RichTextPayloadArchive {
                storage: tsp::Reference {
                    identifier: storage_id,
                    ..Default::default()
                },
                range: None,
                cellid: tst::CellId {
                    packed_data: 0,
                    expanded_coord: None,
                },
            }
            .encode_to_vec()
        };
        let messages = vec![
            RawMessage {
                type_: 6_218,
                data: vec![0xff],
            },
            RawMessage {
                type_: 6_218,
                data: payload(8),
            },
            RawMessage {
                type_: 6_218,
                data: payload(9),
            },
        ];
        let mut visited = Vec::new();
        let mut extract_text = |storage_id| {
            visited.push(storage_id);
            Ok::<Option<String>, Error>((storage_id == 9).then(|| "rich".to_owned()))
        };

        let actual = first_rich_text_payload_text(&messages, &mut extract_text).unwrap();
        assert_eq!(actual.as_deref(), Some("rich"));
        assert_eq!(visited, [8, 9]);
    }

    #[test]
    fn rich_text_payload_candidates_skip_storage_errors_and_continue() {
        use crate::archive::RawMessage;
        use crate::protobuf::tsp;

        let payload = |storage_id| {
            tst::RichTextPayloadArchive {
                storage: tsp::Reference {
                    identifier: storage_id,
                    ..Default::default()
                },
                range: None,
                cellid: tst::CellId {
                    packed_data: 0,
                    expanded_coord: None,
                },
            }
            .encode_to_vec()
        };
        let messages = vec![
            RawMessage {
                type_: 6_218,
                data: payload(8),
            },
            RawMessage {
                type_: 6_218,
                data: payload(9),
            },
        ];
        let mut extract_text = |storage_id| {
            if storage_id == 8 {
                return Err(Error::InvalidFormat(
                    "malformed rich-text storage".to_owned(),
                ));
            }
            Ok::<Option<String>, Error>((storage_id == 9).then(|| "rich".to_owned()))
        };

        let actual = first_rich_text_payload_text(&messages, &mut extract_text).unwrap();
        assert_eq!(actual.as_deref(), Some("rich"));
    }

    #[test]
    fn compatibility_storage_projection_preserves_fragment_order_empty_runs_and_unknowns() {
        use crate::protobuf::tswp;

        let archive = tswp::StorageArchive {
            text: vec!["first".to_owned(), String::new(), "last".to_owned()],
            ..Default::default()
        };
        let mut source = archive.encode_to_vec();
        // Unknown fields stay in the caller-owned source and are not exposed
        // by the private text projection.
        source.extend_from_slice(&[0xa0, 0x06, 0x01]);
        let before = source.clone();
        let limits = compatibility_storage_limits(&source).unwrap();

        let actual = compatibility_storage_text_with_limits(&source, limits)
            .unwrap()
            .unwrap_or_else(|| panic!("non-empty storage should produce text"));

        assert_eq!(actual, archive.text.join("\n"));
        assert_eq!(source, before);

        let present_empty = tswp::StorageArchive {
            text: vec![String::new()],
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            compatibility_storage_text_with_limits(&present_empty, limits)
                .unwrap()
                .as_deref(),
            Some("")
        );
    }

    #[test]
    fn compatibility_storage_projection_rejects_limits_before_output_allocation() {
        let source = [0x1a, 0x01, b'a', 0x1a, 0x01, b'b'];
        let limits = litchi_iwa_text_wire::Limits::new(1_024, 32, 1, 16)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));

        let error = compatibility_storage_text_with_limits(&source, limits).unwrap_err();
        assert!(matches!(
            error,
            litchi_iwa_text_wire::Error::TooManyFragments {
                actual: 2,
                limit: 1
            }
        ));
    }

    #[test]
    fn rich_text_list_stages_valid_entries_and_skips_missing_payloads() {
        use crate::protobuf::tsp;

        let entry = |key, storage_id: Option<u64>| tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            rich_text_payload: storage_id.map(|identifier| tsp::Reference {
                identifier,
                ..Default::default()
            }),
            ..Default::default()
        };
        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 1,
            entries: vec![entry(3, None), entry(7, Some(42))],
            segments: Vec::new(),
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                Ok::<Option<u64>, Error>(
                    entry
                        .rich_text_payload()
                        .map(|reference| reference.identifier()),
                )
            };
        let mut list_budget = ProjectionBudget::new();
        let mut visitor = TypedListVisitor::new(
            &mut converter,
            &mut list_budget,
            tst::table_data_list::ListType::RichTextPayload as i32,
            false,
            true,
            table_list_collection_limits(&source).0,
            table_list_collection_limits(&source).1,
        );
        let (snapshot, _) = numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
            &source,
            table_data_list_decode_options(&source),
            &mut visitor,
        )
        .unwrap();
        let (values, keys, _, structural_error, semantic_error) = visitor.take_parts();

        assert_eq!(
            snapshot.list_type(),
            tst::table_data_list::ListType::RichTextPayload as i32
        );
        assert_eq!(values, [(7, 42)]);
        assert_eq!(keys.len(), 2);
        assert!(structural_error.is_none());
        assert!(semantic_error.is_none());
    }

    #[test]
    fn table_list_visitor_preserves_mixed_union_filtering() {
        // The strict wire owner validates every union field, while the legacy
        // host projection selects only the requested member. An entry may
        // therefore carry a string and an unrelated formula without being
        // rejected or duplicated in the string sidecar.
        let mixed = tst::table_data_list::ListEntry {
            key: 3,
            refcount: 1,
            string: Some("kept".to_owned()),
            formula: Some(tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive::default(),
                host_column: Some(0),
                ..Default::default()
            }),
            ..Default::default()
        };
        let formula_only = tst::table_data_list::ListEntry {
            key: 7,
            refcount: 1,
            formula: Some(tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive::default(),
                host_column: Some(0),
                ..Default::default()
            }),
            ..Default::default()
        };
        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: 1,
            entries: vec![mixed, formula_only],
            ..Default::default()
        }
        .encode_to_vec();
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                Ok::<Option<String>, Error>(entry.string_value().map(ToOwned::to_owned))
            };
        let mut list_budget = ProjectionBudget::new();
        let mut visitor = TypedListVisitor::new(
            &mut converter,
            &mut list_budget,
            tst::table_data_list::ListType::String as i32,
            false,
            true,
            table_list_collection_limits(&source).0,
            table_list_collection_limits(&source).1,
        );
        numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
            &source,
            table_data_list_decode_options(&source),
            &mut visitor,
        )
        .unwrap();
        let (values, keys, _, structural_error, semantic_error) = visitor.take_parts();

        assert_eq!(values, [(3, "kept".to_owned())]);
        assert_eq!(keys.len(), 2);
        assert!(structural_error.is_none());
        assert!(semantic_error.is_none());
    }

    #[test]
    fn table_list_staging_bounds_entry_and_segment_sets() {
        use crate::protobuf::tsp;

        let entry = |key, storage_id| tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            rich_text_payload: Some(tsp::Reference {
                identifier: storage_id,
                ..Default::default()
            }),
            ..Default::default()
        };
        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 1,
            entries: vec![entry(3, 42), entry(7, 43)],
            segments: vec![
                tsp::Reference {
                    identifier: 100,
                    ..Default::default()
                },
                tsp::Reference {
                    identifier: 101,
                    ..Default::default()
                },
            ],
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                Ok::<Option<u64>, Error>(
                    entry
                        .rich_text_payload()
                        .map(|reference| reference.identifier()),
                )
            };
        let mut list_budget = ProjectionBudget::new();
        let mut visitor = TypedListVisitor::new(
            &mut converter,
            &mut list_budget,
            tst::table_data_list::ListType::RichTextPayload as i32,
            false,
            true,
            2,
            1,
        );
        numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
            &source,
            table_data_list_decode_options(&source),
            &mut visitor,
        )
        .unwrap();
        let (values, keys, segment_ids, structural_error, semantic_error) = visitor.take_parts();

        assert_eq!(values, [(3, 42), (7, 43)]);
        assert_eq!(keys.len(), 2);
        assert_eq!(segment_ids, [100]);
        assert!(structural_error.is_none());
        assert!(matches!(
            semantic_error,
            Some(Error::InvalidFormat(message)) if message.contains("segment references")
        ));
    }

    #[test]
    fn table_list_staging_bounds_entry_set_before_growth() {
        let entry = |key: u32, value: &str| tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            string: Some(value.to_owned()),
            ..Default::default()
        };
        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: 1,
            entries: vec![entry(3, "first"), entry(7, "second")],
            segments: Vec::new(),
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                Ok::<Option<String>, Error>(entry.string_value().map(ToOwned::to_owned))
            };
        let mut list_budget = ProjectionBudget::new();
        let mut visitor = TypedListVisitor::new(
            &mut converter,
            &mut list_budget,
            tst::table_data_list::ListType::String as i32,
            false,
            true,
            1,
            1,
        );
        numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
            &source,
            table_data_list_decode_options(&source),
            &mut visitor,
        )
        .unwrap();
        let (values, keys, _, structural_error, semantic_error) = visitor.take_parts();

        assert_eq!(values, [(3, "first".to_owned())]);
        assert_eq!(keys.len(), 1);
        assert!(structural_error.is_none());
        assert!(matches!(
            semantic_error,
            Some(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed: 2,
                limit: MAX_TABLE_LIST_ENTRIES,
            }))
        ));
    }

    #[test]
    fn table_list_projection_budget_charges_root_and_segment_reports() {
        use crate::protobuf::tsp;

        let root = tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: 1,
            entries: vec![tst::table_data_list::ListEntry {
                key: 3,
                refcount: 1,
                string: Some("root".to_owned()),
                ..Default::default()
            }],
            segments: vec![tsp::Reference {
                identifier: 99,
                ..Default::default()
            }],
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::String as i32,
            key_range: tsp::Range {
                location: 7,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 7,
                refcount: 1,
                string: Some("segment".to_owned()),
                ..Default::default()
            }],
        }
        .encode_to_vec();

        let (_, root_report) =
            numbers_table_cell_storage_codec::decode_table_data_list_with_report(
                &root,
                table_data_list_decode_options(&root),
            )
            .unwrap();
        let (_, segment_report) =
            numbers_table_cell_storage_codec::decode_table_data_list_segment_with_report(
                &segment,
                table_data_list_decode_options(&segment),
            )
            .unwrap();

        let mut budget = ProjectionBudget::new();
        budget.charge_decode_report(root_report).unwrap();
        budget.charge_decode_report(segment_report).unwrap();
        assert_eq!(
            budget.references,
            root_report.references() + segment_report.references()
        );
        assert_eq!(
            budget.payload_fields,
            root_report.fields() + segment_report.fields()
        );
        assert_eq!(
            budget.payload_work,
            root_report.work_bytes() + segment_report.work_bytes()
        );
        assert_eq!(
            budget.staging_text_bytes,
            root_report.text_bytes() + segment_report.text_bytes()
        );
    }

    #[test]
    fn table_list_projection_budget_rejects_work_across_root_and_segment() {
        use crate::protobuf::tsp;

        let root = tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: 1,
            entries: vec![tst::table_data_list::ListEntry {
                key: 3,
                refcount: 1,
                string: Some("root".to_owned()),
                ..Default::default()
            }],
            segments: vec![tsp::Reference {
                identifier: 99,
                ..Default::default()
            }],
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::String as i32,
            key_range: tsp::Range {
                location: 7,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 7,
                refcount: 1,
                string: Some("segment".to_owned()),
                ..Default::default()
            }],
        }
        .encode_to_vec();
        let (_, root_report) =
            numbers_table_cell_storage_codec::decode_table_data_list_with_report(
                &root,
                table_data_list_decode_options(&root),
            )
            .unwrap();
        let (_, segment_report) =
            numbers_table_cell_storage_codec::decode_table_data_list_segment_with_report(
                &segment,
                table_data_list_decode_options(&segment),
            )
            .unwrap();

        let mut budget = ProjectionBudget::new();
        budget.payload_work = MAX_TABLE_LIST_PAYLOAD_WORK - root_report.work_bytes();
        budget.charge_decode_work(root_report).unwrap();
        let before_segment = budget;
        let error = budget.charge_decode_work(segment_report).unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                ..
            })
        ));
        assert_eq!(budget, before_segment);
    }

    #[test]
    fn table_list_projection_budget_rejects_aggregate_text_and_references() {
        use crate::protobuf::tsp;

        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::String as i32,
            next_list_id: 1,
            entries: vec![tst::table_data_list::ListEntry {
                key: 3,
                refcount: 1,
                string: Some("x".to_owned()),
                reference: Some(tsp::Reference {
                    identifier: 9,
                    ..Default::default()
                }),
                ..Default::default()
            }],
            ..Default::default()
        }
        .encode_to_vec();
        let (_, report) = numbers_table_cell_storage_codec::decode_table_data_list_with_report(
            &source,
            table_data_list_decode_options(&source),
        )
        .unwrap();
        assert!(report.text_bytes() > 0);
        assert!(report.references() > 0);

        let mut text_budget = ProjectionBudget::new();
        text_budget.staging_text_bytes = MAX_TABLE_LIST_TEXT_BYTES - report.text_bytes() + 1;
        let before_text = text_budget;
        let text_error = text_budget.charge_decode_report(report).unwrap_err();
        assert!(matches!(
            text_error,
            Error::InvalidFormat(message) if message.contains("text exceeded")
        ));
        assert_eq!(text_budget, before_text);

        let mut reference_budget = ProjectionBudget::new();
        reference_budget.references = MAX_TABLE_LIST_REFERENCES - report.references() + 1;
        let before_references = reference_budget;
        let reference_error = reference_budget.charge_decode_report(report).unwrap_err();
        assert!(matches!(
            reference_error,
            Error::InvalidFormat(message) if message.contains("references exceeded")
        ));
        assert_eq!(reference_budget, before_references);
    }

    #[test]
    fn wrong_list_candidates_charge_traversal_without_publishing_text() {
        let source = tst::TableDataList {
            list_type: tst::table_data_list::ListType::Formula as i32,
            next_list_id: 1,
            entries: vec![tst::table_data_list::ListEntry {
                key: 3,
                refcount: 1,
                string: Some("wrong candidate".to_owned()),
                ..Default::default()
            }],
            ..Default::default()
        }
        .encode_to_vec();
        let mut budget = ProjectionBudget::new();
        let (probe, probe_report) =
            numbers_table_cell_storage_codec::decode_table_data_list_type_with_report(
                &source,
                table_data_list_decode_options_with_budget(&source, budget, false),
            )
            .unwrap();
        budget.charge_decode_work(probe_report).unwrap();
        assert_ne!(
            probe.list_type(),
            tst::table_data_list::ListType::String as i32
        );

        let (_, report) = numbers_table_cell_storage_codec::decode_table_data_list_with_report(
            &source,
            table_data_list_decode_options_with_budget(&source, budget, false),
        )
        .unwrap();
        budget.charge_decode_work(report).unwrap();
        assert_eq!(budget.references, 0);
        assert_eq!(budget.staging_text_bytes, 0);
        assert!(budget.payload_fields > 0);
        assert!(budget.payload_work > 0);
    }

    #[test]
    fn test_bnc_string_cell() {
        let data = [
            5, 3, 0, 0, 0, 0, 0, 0, // version and type
            0x08, 0x00, 0x00, 0x00, // string-id flag
            7, 0, 0, 0, // string id
        ];
        let strings: StringTable = compact_table([(7, "hello".to_string())]).unwrap();
        let formulas: FormulaTable = Box::default();
        let errors: FormulaErrorTable = Box::default();
        let rich_text: StringTable = Box::default();
        let comments: CommentTable = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };
        let value = TableDataExtractor::parse_cell_storage(&data, &tables, 0, 0).unwrap();
        assert_eq!(value.value.as_text(), "hello");
    }

    #[test]
    fn type_nine_bnc_numeric_cell_uses_union_semantics() {
        let mut data = vec![5, 9, 0, 0, 0, 0, 0, 0];
        data.extend_from_slice(&1_u32.to_le_bytes());
        data.extend_from_slice(&decimal128_le(-1_234.5).unwrap());

        let strings: StringTable = Box::default();
        let formulas: FormulaTable = Box::default();
        let errors: FormulaErrorTable = Box::default();
        let rich_text: StringTable = Box::default();
        let comments: CommentTable = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };

        let parsed = TableDataExtractor::parse_cell_storage(&data, &tables, 2, 3).unwrap();
        let CellValue::Number(value) = parsed.value else {
            panic!("type-nine decimal was not extracted as a number");
        };
        assert_eq!(value.get(), -1_234.5);
    }

    #[test]
    fn type_nine_bnc_cell_prefers_rich_text_then_plain_text_over_numeric_cache() {
        let strings: StringTable = compact_table([(7, "plain".to_owned())]).unwrap();
        let formulas: FormulaTable = Box::default();
        let errors: FormulaErrorTable = Box::default();
        let rich_text: StringTable = compact_table([(8, "rich".to_owned())]).unwrap();
        let comments: CommentTable = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };

        let mut plain = vec![5, 9, 0, 0, 0, 0, 0, 0];
        plain.extend_from_slice(&(0x000001_u32 | 0x000008).to_le_bytes());
        plain.extend_from_slice(&decimal128_le(42.0).unwrap());
        plain.extend_from_slice(&7_u32.to_le_bytes());
        let parsed = TableDataExtractor::parse_cell_storage(&plain, &tables, 0, 0).unwrap();
        assert_eq!(parsed.value, CellValue::Text("plain".to_owned()));

        let mut rich = vec![5, 9, 0, 0, 0, 0, 0, 0];
        rich.extend_from_slice(&(0x000001_u32 | 0x000008 | 0x000010).to_le_bytes());
        rich.extend_from_slice(&decimal128_le(42.0).unwrap());
        rich.extend_from_slice(&7_u32.to_le_bytes());
        rich.extend_from_slice(&8_u32.to_le_bytes());
        let parsed = TableDataExtractor::parse_cell_storage(&rich, &tables, 0, 0).unwrap();
        assert_eq!(parsed.value, CellValue::Text("rich".to_owned()));
    }

    #[test]
    fn bnc_and_pre_bnc_error_cells_resolve_formula_error_text() {
        let errors: FormulaErrorTable = compact_table([(7, "Syntax Error".to_owned())]).unwrap();
        let bnc = [
            5, 8, 0, 0, 0, 0, 0, 0, // version and type
            0x00, 0x08, 0x00, 0x00, // formula-error flag
            7, 0, 0, 0,
        ];
        let strings: StringTable = Box::default();
        let formulas: FormulaTable = Box::default();
        let rich_text: StringTable = Box::default();
        let comments: CommentTable = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };
        let value = TableDataExtractor::parse_cell_storage(&bnc, &tables, 0, 0).unwrap();
        assert_eq!(value.value, CellValue::Error("Syntax Error".to_owned()));

        let pre_bnc = [
            4, 8, 0, 0, // version and type
            0x00, 0x01, 0x00, 0x00, // formula-error flag
            0, 0, 0, 0, // V4 header padding
            7, 0, 0, 0,
        ];
        let value = TableDataExtractor::parse_cell_storage(&pre_bnc, &tables, 0, 0).unwrap();
        assert_eq!(value.value, CellValue::Error("Syntax Error".to_owned()));
    }

    #[test]
    fn bnc_and_pre_bnc_cells_expose_comment_identifiers() {
        let strings: StringTable = Box::default();
        let formulas: FormulaTable = Box::default();
        let errors: FormulaErrorTable = Box::default();
        let rich_text: StringTable = Box::default();
        let comments: CommentTable = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &errors,
            rich_text: &rich_text,
            comments: &comments,
            formula_references: &formula_references,
        };
        let bnc = [
            5, 0, 0, 0, 0, 0, 0, 0, // version and empty type
            0x00, 0x00, 0x08, 0x00, // comment flag
            9, 0, 0, 0,
        ];
        let parsed = TableDataExtractor::parse_cell_storage(&bnc, &tables, 0, 0).unwrap();
        assert_eq!(parsed.comment_identifier, Some(9));

        let pre_bnc = [
            4, 0, 0, 0, // version and empty type
            0x00, 0x10, 0x00, 0x00, // comment flag
            0, 0, 0, 0, // V4 header padding
            9, 0, 0, 0,
        ];
        let parsed = TableDataExtractor::parse_cell_storage(&pre_bnc, &tables, 0, 0).unwrap();
        assert_eq!(parsed.comment_identifier, Some(9));
    }

    #[test]
    fn strict_comment_storage_decode_matches_legacy_projection_and_preserves_unknowns() {
        use crate::protobuf::{tsd, tsp};

        let legacy = tsd::CommentStorageArchive {
            text: Some("comment".to_owned()),
            creation_date: Some(tsp::Date { seconds: -7.5 }),
            author: Some(tsp::Reference {
                identifier: 20,
                ..Default::default()
            }),
            replies: vec![
                tsp::Reference {
                    identifier: 30,
                    ..Default::default()
                },
                tsp::Reference {
                    identifier: 31,
                    ..Default::default()
                },
            ],
            storage_uuid: Some(tsp::Uuid { lower: 1, upper: 2 }),
        };
        let mut source = legacy.encode_to_vec();
        // Unknown fields remain in the caller-owned payload; the strict
        // projection must accept them without re-encoding the message.
        source.extend_from_slice(&[0x78, 0x01]);
        let before = source.clone();

        let (strict, reply_ids) = decode_comment_storage_payload(99, &source).unwrap();

        assert_eq!(source, before);
        assert_eq!(strict.text(), legacy.text.as_deref());
        assert_eq!(
            strict.creation_date().map(|date| date.seconds()),
            legacy.creation_date.as_ref().map(|date| date.seconds)
        );
        assert_eq!(
            strict.author().map(|author| author.identifier()),
            legacy.author.as_ref().map(|author| author.identifier)
        );
        assert_eq!(reply_ids, [30, 31]);
        assert_eq!(
            strict
                .storage_uuid()
                .map(|uuid| (uuid.lower(), uuid.upper())),
            legacy
                .storage_uuid
                .as_ref()
                .map(|uuid| (uuid.lower, uuid.upper))
        );
    }

    #[test]
    fn strict_comment_storage_decode_rejects_malformed_utf8_duplicate_and_missing_reference() {
        let invalid_utf8 = [0x0a, 0x01, 0xff];
        let error = decode_comment_storage_payload(99, &invalid_utf8).unwrap_err();
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("invalid UTF-8"))
        );

        let duplicate_text = [0x0a, 0x01, b'a', 0x0a, 0x01, b'b'];
        let error = decode_comment_storage_payload(99, &duplicate_text).unwrap_err();
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("duplicate singular field") && message.contains("text"))
        );

        let missing_reference = [0x1a, 0x00];
        let error = decode_comment_storage_payload(99, &missing_reference).unwrap_err();
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("missing required field") && message.contains("identifier"))
        );
    }
}
