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
use super::names;
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
use litchi_iwa_protos::{numbers_table_cell_storage_codec, tsce, tsd, tst};
use prost::Message;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::sync::Arc;

type CompactTable<T> = Box<[(u32, T)]>;
type StringTable = CompactTable<String>;
type FormulaTable = CompactTable<tsce::FormulaArchive>;
type FormulaErrorTable = CompactTable<String>;
type CommentTable = CompactTable<Comment>;
type FormulaOwnerKey = [u32; 4];
type FormulaCategoryKey = [u64; 2];
const RICH_TEXT_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const MAX_TABLE_ROWS: usize = 1 << 20;
const MAX_TABLE_COLUMNS: usize = 1 << 14;
const MAX_ADDRESSABLE_CELLS: usize = 1 << 24;
const MAX_TABLE_MATERIALIZED_CELLS: usize = 1 << 20;
const MAX_FORMULA_CATEGORY_DEPTH: usize = 64;
const MAX_FORMULA_WORK: usize = crate::MAX_REFERENCES;
const MAX_FORMULA_WIRE_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;
const MAX_PAYLOAD_WORK: usize = WireLimits::MAX_REWRITE_WORK;

fn table_cell_decode_options(
    source: &[u8],
    max_references: usize,
    max_text_bytes: usize,
    max_fields: usize,
    max_work: usize,
) -> numbers_table_cell_storage_codec::DecodeOptions {
    numbers_table_cell_storage_codec::DecodeOptions::new(
        source.len().max(1),
        max_fields,
        max_work,
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
        max_references,
        max_text_bytes,
    )
}

fn map_table_cell_codec_error(error: numbers_table_cell_storage_codec::DecodeError) -> Error {
    map_table_cell_codec_error_with_reference_offset(error, 0)
}

fn map_table_cell_codec_error_with_reference_offset(
    error: numbers_table_cell_storage_codec::DecodeError,
    reference_offset: usize,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidFormat("Numbers table storage projection is invalid".to_owned());
    };
    map_table_cell_decode_limit_with_reference_offset(limit, reference_offset)
}

fn map_table_cell_decode_limit_with_reference_offset(
    limit: numbers_table_cell_storage_codec::DecodeLimit,
    reference_offset: usize,
) -> Error {
    use numbers_table_cell_storage_codec::DecodeLimit;
    let (kind, observed, maximum) = match limit {
        DecodeLimit::Bytes { observed, maximum } => {
            (SemanticLimitKind::FormulaWireBytes, observed, maximum)
        },
        DecodeLimit::References { observed, maximum } => (
            SemanticLimitKind::References,
            observed.saturating_add(reference_offset),
            maximum.saturating_add(reference_offset),
        ),
        DecodeLimit::Text { observed, maximum } => {
            (SemanticLimitKind::TextBytes, observed, maximum)
        },
        DecodeLimit::Fields { observed, maximum } => {
            (SemanticLimitKind::Objects, observed, maximum)
        },
        DecodeLimit::Work { observed, maximum } => {
            (SemanticLimitKind::FormulaWork, observed, maximum)
        },
        DecodeLimit::Nesting { observed, maximum } => (
            SemanticLimitKind::FormulaDepth,
            observed as usize,
            maximum as usize,
        ),
        DecodeLimit::Allocation { requested } => {
            return Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table storage projection",
                amount: requested,
            });
        },
        _ => {
            return Error::InvalidFormat("Numbers table storage projection is invalid".to_owned());
        },
    };
    Error::SemanticLimit {
        kind,
        observed,
        maximum,
        path: SemanticPath::Package,
    }
}

struct CellTables<'a> {
    strings: &'a StringTable,
    formulas: &'a FormulaTable,
    formula_errors: &'a FormulaErrorTable,
    rich_text: &'a StringTable,
    comments: Option<&'a CommentTable>,
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
    references: usize,
    payload_fields: usize,
    payload_work: usize,
    staging_text_bytes: usize,
    formula_wire_bytes: usize,
    materialized_cells: usize,
    output_text_bytes: usize,
    formula_render_work: usize,
    max_materialized_cells: usize,
    max_output_text_bytes: usize,
    max_formula_render_work: usize,
    max_formula_render_depth: usize,
    max_references: usize,
}

impl ProjectionBudget {
    const fn new(limits: SemanticLimits) -> Self {
        Self {
            references: 0,
            payload_fields: 0,
            payload_work: 0,
            staging_text_bytes: 0,
            formula_wire_bytes: 0,
            materialized_cells: 0,
            output_text_bytes: 0,
            formula_render_work: 0,
            max_materialized_cells: limits.max_materialized_cells(),
            max_output_text_bytes: limits.max_output_text_bytes(),
            max_formula_render_work: limits.max_formula_render_work(),
            max_formula_render_depth: limits.max_formula_render_depth(),
            max_references: limits.max_references(),
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

    fn check_materialized_cells(&self, amount: usize) -> Result<()> {
        projection_charge(
            self.materialized_cells,
            amount,
            self.max_materialized_cells,
            SemanticLimitKind::MaterializedCells,
        )?;
        Ok(())
    }

    const fn remaining_materialized_cells(&self) -> usize {
        self.max_materialized_cells
            .saturating_sub(self.materialized_cells)
    }

    const fn remaining_output_text_bytes(&self) -> usize {
        self.max_output_text_bytes
            .saturating_sub(self.output_text_bytes)
    }

    const fn remaining_references(&self) -> usize {
        self.max_references.saturating_sub(self.references)
    }

    fn charge_references(&mut self, amount: usize) -> Result<()> {
        self.references = projection_charge(
            self.references,
            amount,
            self.max_references,
            SemanticLimitKind::References,
        )?;
        Ok(())
    }

    const fn remaining_payload_fields(&self) -> usize {
        crate::MAX_REFERENCES.saturating_sub(self.payload_fields)
    }

    const fn remaining_payload_work(&self) -> usize {
        MAX_PAYLOAD_WORK.saturating_sub(self.payload_work)
    }

    const fn remaining_staging_text_bytes(&self) -> usize {
        DEFAULT_MAX_TEXT_BYTES.saturating_sub(self.staging_text_bytes)
    }

    const fn remaining_formula_wire_bytes(&self) -> usize {
        MAX_FORMULA_WIRE_BYTES.saturating_sub(self.formula_wire_bytes)
    }

    fn charge_staging_text(&mut self, amount: usize) -> Result<()> {
        self.staging_text_bytes = projection_charge(
            self.staging_text_bytes,
            amount,
            DEFAULT_MAX_TEXT_BYTES,
            SemanticLimitKind::TextBytes,
        )?;
        Ok(())
    }

    fn charge_formula_wire(&mut self, amount: usize) -> Result<()> {
        self.formula_wire_bytes = projection_charge(
            self.formula_wire_bytes,
            amount,
            MAX_FORMULA_WIRE_BYTES,
            SemanticLimitKind::FormulaWireBytes,
        )?;
        Ok(())
    }

    fn charge_wire_preflight(
        &mut self,
        report: litchi_iwa_common::wire::WirePreflight,
    ) -> Result<()> {
        self.payload_fields = projection_charge(
            self.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        self.payload_work = projection_charge(
            self.payload_work,
            report.scanned_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        Ok(())
    }

    fn charge_decode_report(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        self.charge_references(report.references())?;
        self.payload_fields = projection_charge(
            self.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        self.payload_work = projection_charge(
            self.payload_work,
            report.work_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        Ok(())
    }

    fn charge_decode_work(
        &mut self,
        report: numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<()> {
        self.payload_fields = projection_charge(
            self.payload_fields,
            report.fields(),
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        self.payload_work = projection_charge(
            self.payload_work,
            report.work_bytes(),
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
        )?;
        Ok(())
    }

    fn charge_list_type_probe(&mut self, fields: usize, work_bytes: usize) -> Result<()> {
        self.payload_fields = projection_charge(
            self.payload_fields,
            fields,
            crate::MAX_REFERENCES,
            SemanticLimitKind::Objects,
        )?;
        self.payload_work = projection_charge(
            self.payload_work,
            work_bytes,
            MAX_PAYLOAD_WORK,
            SemanticLimitKind::FormulaWork,
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
            self.payload_work = self.payload_work.max(candidate.payload_work);
            self.formula_wire_bytes = self.formula_wire_bytes.max(candidate.formula_wire_bytes);
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

trait ListValueConverter<T> {
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<T>;
}

impl<T, F> ListValueConverter<T> for F
where
    F: for<'source> FnMut(
        numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'source>,
        &mut ProjectionBudget,
    ) -> Result<T>,
{
    fn convert(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
        budget: &mut ProjectionBudget,
    ) -> Result<T> {
        self(entry, budget)
    }
}

/// Stage only final semantic values while the strict list codec walks the
/// source.  The visitor never returns a conversion/allocation error to the
/// codec: callbacks can run before a later wire or Buffa parity failure, so a
/// candidate-local semantic error is retained and the complete source is
/// still traversed.
struct TypedListVisitor<'converter, T, C> {
    converter: &'converter mut C,
    projection_budget: &'converter mut ProjectionBudget,
    stage_semantics: bool,
    expected_list_type: i32,
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

impl<'converter, T, C> TypedListVisitor<'converter, T, C> {
    fn new(
        converter: &'converter mut C,
        projection_budget: &'converter mut ProjectionBudget,
        expected_list_type: i32,
        segment: bool,
        stage_semantics: bool,
    ) -> Self {
        Self {
            converter,
            projection_budget,
            stage_semantics,
            expected_list_type,
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

impl<T, C> numbers_table_cell_storage_codec::StorageVisitor for TypedListVisitor<'_, T, C>
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
        if !entry_matches_list_type_snapshot(&entry, self.expected_list_type) {
            self.record_structural_error(Error::InvalidFormat(
                "Numbers table-data-list entry has no selected payload".to_owned(),
            ));
            return Ok(());
        }
        if self.keys.contains(&entry.key()) {
            self.record_structural_error(Error::InvalidFormat(
                "Numbers table sidecar contains duplicate keys".to_owned(),
            ));
            return Ok(());
        }
        if self.keys.try_reserve(1).is_err() {
            // Keep walking after a fallible staging failure. The error is
            // candidate-local and must not prevent a later structural error
            // from winning publication.
            self.record_semantic_error(allocation_error(
                "Numbers table-list entry keys",
                self.keys.len().saturating_add(1),
            ));
            return Ok(());
        }
        self.keys.insert(entry.key());
        // Once a semantic conversion/allocation error has been retained, keep
        // checking the wire-level shape and duplicate-key invariants above,
        // but do not invoke the converter or allocate another retained value.
        // Non-admitted candidates use the same structural path while never
        // staging semantic values or charging references/text.
        if self.semantic_error.is_some() || !self.stage_semantics {
            return Ok(());
        }
        if self.values.try_reserve(1).is_err() {
            self.record_semantic_error(allocation_error(
                "Numbers table-list entries",
                self.values.len().saturating_add(1),
            ));
            return Ok(());
        }
        match self.converter.convert(entry, self.projection_budget) {
            Ok(value) => self.values.push((entry.key(), value)),
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

struct TileRowVisitor<'a, 'tables> {
    row_origin: usize,
    tile_size: usize,
    row_count: usize,
    column_count: usize,
    budget: &'a mut CellBudget,
    cell_tables: &'a CellTables<'tables>,
    projection_budget: &'a mut ProjectionBudget,
    table: &'a mut Table,
    materialized_cells: usize,
    semantic_error: Option<Error>,
}

impl<'a, 'tables> numbers_table_cell_storage_codec::StorageVisitor for TileRowVisitor<'a, 'tables> {
    fn visit_tile_row(
        &mut self,
        row: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        self.materialized_cells = self
            .materialized_cells
            .saturating_add(usize::try_from(row.cell_count()).unwrap_or(usize::MAX));
        if self.semantic_error.is_some() {
            return Ok(());
        }
        if let Err(error) = self
            .projection_budget
            .check_materialized_cells(self.materialized_cells)
        {
            self.semantic_error = Some(error);
            return Ok(());
        }
        if let Err(error) = TableDataExtractor::parse_tile_row(
            &row,
            self.row_origin,
            self.tile_size,
            self.row_count,
            self.column_count,
            self.budget,
            self.cell_tables,
            self.projection_budget,
            self.table,
        ) {
            self.semantic_error = Some(error);
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

fn decode_legacy_table_candidate<T>(
    data: &[u8],
    admit: impl FnOnce() -> Result<()>,
    parse: impl FnOnce(tst::TableModelArchive) -> Result<T>,
) -> Result<Option<T>> {
    match has_legacy_table_model_wire_shape(data) {
        Ok(true) => {},
        Ok(false) => return Ok(None),
        Err(error) => return Err(error),
    }
    admit()?;
    let table_model = tst::TableModelArchive::decode(data).map_err(Error::protobuf)?;
    parse(table_model).map(Some)
}

/// Require the parse-relevant required fields that distinguish a historical
/// type-6000 `TableModelArchive` from its modern `TableInfoArchive` owner.
///
/// Prost supplies defaults for absent proto2 required fields, so a successful
/// generated decode alone is not evidence that the payload is a table model.
/// In particular, protobuf's permissive unknown-field handling lets ordinary
/// table-info metadata decode into a mostly default model. Scanning only the
/// flat envelope retains the compatibility fallback without allocating a
/// second field index. Once this shape is present, the candidate is admitted
/// and every model-decoding failure must remain observable to the caller.
fn has_legacy_table_model_wire_shape(data: &[u8]) -> Result<bool> {
    const DATA_STORE: u8 = 1 << 0;
    const ROW_COUNT: u8 = 1 << 1;
    const COLUMN_COUNT: u8 = 1 << 2;
    const TABLE_NAME: u8 = 1 << 3;
    const REQUIRED: u8 = DATA_STORE | ROW_COUNT | COLUMN_COUNT | TABLE_NAME;

    let mut present = 0_u8;
    let mut ambiguous = false;
    let preflight = preflight_wire_tree_with_limits(data, WireLimits::default(), |visit| {
        let field = visit.field();
        let required = match field.number() {
            4 => Some((DATA_STORE, 2)),
            6 => Some((ROW_COUNT, 0)),
            7 => Some((COLUMN_COUNT, 0)),
            8 => Some((TABLE_NAME, 2)),
            _ => None,
        };
        if let Some((bit, expected_wire_type)) = required {
            if field.wire_type() != expected_wire_type || present & bit != 0 {
                ambiguous = true;
            } else {
                present |= bit;
            }
        }
        Ok(WireDescent::Skip)
    });
    if let Err(error) = preflight {
        return match error {
            litchi_iwa_common::Error::InvalidFormat(_) if present == REQUIRED => {
                Err(Error::MalformedPayload {
                    path: SemanticPath::StructuredTables,
                })
            },
            litchi_iwa_common::Error::InvalidFormat(_) => Ok(false),
            common_error @ (litchi_iwa_common::Error::LimitExceeded { .. }
            | litchi_iwa_common::Error::Allocation { .. }
            | litchi_iwa_common::Error::InvalidLimit { .. }) => Err(Error::Common(common_error)),
        };
    }
    if present != REQUIRED {
        return Ok(false);
    }
    if ambiguous {
        return Err(Error::MalformedPayload {
            path: SemanticPath::StructuredTables,
        });
    }
    Ok(true)
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

/// Read only the root list-type scalar to decide whether entry conversion is
/// worth staging. The strict storage codec remains authoritative for every
/// candidate; this bounded probe merely avoids resolving rich/comment payloads
/// from a known wrong-list candidate before that codec reports the mismatch.
///
/// The caller charges one source-byte of projection work for this complete
/// envelope pass. Length-delimited child payloads are skipped without
/// descending, while unknown groups are balanced and field-counted.
#[derive(Debug, Clone, Copy)]
struct ListTypeProbe {
    list_type: Option<i32>,
    fields: usize,
}

fn probe_table_data_list_type(source: &[u8]) -> ListTypeProbe {
    let mut remaining = source;
    let mut list_type = None;
    let mut fields = 0usize;
    while !remaining.is_empty() {
        let Some(tag) = probe_varint(&mut remaining) else {
            return ListTypeProbe { list_type, fields };
        };
        fields = fields.saturating_add(1);
        let Ok(number) = u32::try_from(tag >> 3) else {
            return ListTypeProbe { list_type, fields };
        };
        let Ok(wire_type) = u8::try_from(tag & 7) else {
            return ListTypeProbe { list_type, fields };
        };
        if number == 0 || number > 0x1fff_ffff {
            return ListTypeProbe { list_type, fields };
        }
        match wire_type {
            0 => {
                let Some(value) = probe_varint(&mut remaining) else {
                    return ListTypeProbe { list_type, fields };
                };
                if number == 1 {
                    let Some(value) = probe_int32(value) else {
                        return ListTypeProbe { list_type, fields };
                    };
                    if list_type.replace(value).is_some() {
                        return ListTypeProbe { list_type, fields };
                    }
                }
            },
            1 => {
                let Some(next) = remaining.get(8..) else {
                    return ListTypeProbe { list_type, fields };
                };
                remaining = next;
            },
            2 => {
                let Some(length) =
                    probe_varint(&mut remaining).and_then(|value| usize::try_from(value).ok())
                else {
                    return ListTypeProbe { list_type, fields };
                };
                let Some(next) = remaining.get(length..) else {
                    return ListTypeProbe { list_type, fields };
                };
                remaining = next;
            },
            3 => {
                if !probe_skip_group(&mut remaining, number, 1, &mut fields) {
                    return ListTypeProbe { list_type, fields };
                }
            },
            4 => return ListTypeProbe { list_type, fields },
            5 => {
                let Some(next) = remaining.get(4..) else {
                    return ListTypeProbe { list_type, fields };
                };
                remaining = next;
            },
            _ => return ListTypeProbe { list_type, fields },
        }
    }
    ListTypeProbe { list_type, fields }
}

fn probe_skip_group(source: &mut &[u8], expected: u32, depth: usize, fields: &mut usize) -> bool {
    if depth > WireLimits::MAX_NESTING {
        return false;
    }
    while !source.is_empty() {
        let Some(tag) = probe_varint(source) else {
            return false;
        };
        *fields = fields.saturating_add(1);
        let Ok(number) = u32::try_from(tag >> 3) else {
            return false;
        };
        let Ok(wire_type) = u8::try_from(tag & 7) else {
            return false;
        };
        if number == 0 || number > 0x1fff_ffff {
            return false;
        }
        match wire_type {
            0 => {
                if probe_varint(source).is_none() {
                    return false;
                }
            },
            1 => {
                if source.get(8..).is_none() {
                    return false;
                }
                *source = &source[8..];
            },
            2 => {
                let Some(length) =
                    probe_varint(source).and_then(|value| usize::try_from(value).ok())
                else {
                    return false;
                };
                if source.get(length..).is_none() {
                    return false;
                }
                *source = &source[length..];
            },
            3 => {
                if !probe_skip_group(source, number, depth + 1, fields) {
                    return false;
                }
            },
            4 => return number == expected,
            5 => {
                if source.get(4..).is_none() {
                    return false;
                }
                *source = &source[4..];
            },
            _ => return false,
        }
    }
    false
}

fn probe_varint(source: &mut &[u8]) -> Option<u64> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index)?;
        if index == 9 && byte > 1 {
            return None;
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let encoded_length = if value == 0 {
                1
            } else {
                (64usize - value.leading_zeros() as usize).div_ceil(7)
            };
            if encoded_length != index + 1 {
                return None;
            }
            *source = &original[index + 1..];
            return Some(value);
        }
    }
    None
}

fn probe_int32(value: u64) -> Option<i32> {
    if let Ok(value) = i32::try_from(value) {
        return Some(value);
    }
    if value < 0xffff_ffff_8000_0000 {
        return None;
    }
    Some(i32::from_ne_bytes((value as u32).to_ne_bytes()))
}

fn record_first_list_error(slot: &mut Option<Error>, error: Error) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn entry_matches_list_type_snapshot(
    entry: &numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    list_type: i32,
) -> bool {
    let present = [
        entry.string_value().is_some(),
        entry.reference().is_some(),
        entry.formula().is_some(),
        entry.format().is_some(),
        entry.custom_format().is_some(),
        entry.rich_text_payload().is_some(),
        entry.comment_storage().is_some(),
        entry.import_warning_set().is_some(),
        entry.cell_spec().is_some(),
    ];
    if present.into_iter().filter(|value| *value).count() != 1 {
        return false;
    }
    match list_type {
        value
            if value == tst::table_data_list::ListType::String as i32
                || value == tst::table_data_list::ListType::FormulaError as i32 =>
        {
            entry.string_value().is_some()
        },
        value if value == tst::table_data_list::ListType::Formula as i32 => {
            entry.formula().is_some()
        },
        value if value == tst::table_data_list::ListType::RichTextPayload as i32 => {
            entry.rich_text_payload().is_some()
        },
        value if value == tst::table_data_list::ListType::CommentStorage as i32 => {
            entry.comment_storage().is_some()
        },
        _ => true,
    }
}

#[derive(Debug, Clone)]
struct FormulaReferenceName {
    sheet: Arc<String>,
    table: Arc<String>,
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
    projection_budget: RefCell<ProjectionBudget>,
    retain_comments: bool,
    document_projection: bool,
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
            projection_budget: RefCell::new(ProjectionBudget::new(limits)),
            retain_comments: true,
            document_projection: false,
        }
    }

    pub(super) fn without_comments(mut self) -> Self {
        self.retain_comments = false;
        self.document_projection = true;
        self
    }

    /// Charge semantic text retained outside table projection, such as rooted
    /// sheet names, against the same package-wide output budget.
    pub(super) fn charge_output_text(&self, amount: usize) -> Result<()> {
        self.projection_budget
            .borrow_mut()
            .charge_output_text(amount)
    }

    /// Merge a rooted reference admission into the table-sidecar budget.
    pub(super) fn charge_references(&self, amount: usize) -> Result<()> {
        self.projection_budget
            .borrow_mut()
            .charge_references(amount)
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
                    && let Some(table) =
                        self.extract_table_candidate(&resolved, message_type, || {
                            if table_count >= max_tables {
                                Err(table_limit_error(table_count.saturating_add(1), max_tables))
                            } else {
                                Ok(())
                            }
                        })?
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
        legacy_admit: impl FnOnce() -> Result<()>,
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
            return self.parse_table_model(table_model, false, None).map(Some);
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
        decode_legacy_table_candidate(&message.data, legacy_admit, |table_model| {
            self.parse_table_model(table_model, false, None)
        })
    }

    /// Extract a table model reached through a schema-proven `TableInfo` edge.
    ///
    /// Rooted ownership is stricter than the global compatibility scan:
    /// canonical and legacy payloads are mutually exclusive and duplicates
    /// fail independently of message order.
    pub(super) fn extract_reachable_table_from_object(
        &self,
        object: &Resolved<'_>,
        path: SemanticPath,
    ) -> Result<Table> {
        let mut typed = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
        let canonical = typed.next();
        if canonical.is_some() && typed.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model contains duplicate canonical payloads"
            )));
        }
        let mut legacy = object
            .messages
            .iter()
            .filter(|message| message.type_ == 6_000);
        let legacy_first = legacy.next();
        let legacy_duplicate = legacy.next().is_some();
        if legacy_duplicate && (self.document_projection || canonical.is_none()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} table model contains duplicate legacy payloads"
            )));
        }
        let message = match (canonical, legacy_first) {
            (Some(message), Some(_)) if !self.document_projection => message,
            (Some(_), Some(_)) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table model has ambiguous payload ownership"
                )));
            },
            (Some(message), None) | (None, Some(message)) => message,
            (None, None) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table model has no recognized payload"
                )));
            },
        };
        if !self.document_projection {
            let table_model =
                tst::TableModelArchive::decode(message.data.as_slice()).map_err(|error| {
                    Error::InvalidFormat(format!(
                        "Numbers {path} table-model payload is malformed: {error}"
                    ))
                })?;
            return self.parse_table_model(table_model, false, None);
        }
        let projected_name = names::preflight_table_name(message.data.as_slice())
            .map_err(|error| super::map_sheet_preflight_error(error, path, 0))?;
        let mut candidate_budget = *self.projection_budget.borrow();
        candidate_budget.charge_output_text(projected_name.len())?;
        let projected_model = if self.document_projection {
            let options = table_cell_decode_options(
                &message.data,
                candidate_budget.remaining_references(),
                MAX_FORMULA_WIRE_BYTES,
                candidate_budget.remaining_payload_fields(),
                candidate_budget.remaining_payload_work(),
            );
            let (projected, report) =
                numbers_table_cell_storage_codec::decode_table_model_with_report(
                    &message.data,
                    options,
                )
                .map_err(|error| {
                    map_table_cell_codec_error_with_reference_offset(
                        error,
                        candidate_budget.references,
                    )
                })?;
            if let Err(error) = candidate_budget.charge_decode_report(report) {
                self.projection_budget
                    .borrow_mut()
                    .commit_attempt(candidate_budget, false);
                return Err(error);
            }
            Some(projected)
        } else {
            None
        };
        let table_model = match tst::TableModelArchive::decode(message.data.as_slice()) {
            Ok(table_model) => table_model,
            Err(error) => {
                self.projection_budget
                    .borrow_mut()
                    .commit_attempt(candidate_budget, false);
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} table-model payload is malformed: {error}"
                )));
            },
        };
        if table_model.table_name != projected_name {
            self.projection_budget
                .borrow_mut()
                .commit_attempt(candidate_budget, false);
            return Err(Error::MalformedPayload { path });
        }
        if let Some(projected_model) = projected_model
            && (table_model.number_of_rows != projected_model.number_of_rows()
                || table_model.number_of_columns != projected_model.number_of_columns()
                || table_model.table_id != projected_model.table_id())
        {
            self.projection_budget
                .borrow_mut()
                .commit_attempt(candidate_budget, false);
            return Err(Error::MalformedPayload { path });
        }
        self.parse_table_model(table_model, true, Some(candidate_budget))
    }

    /// Parse a TableModelArchive protobuf message
    fn parse_table_model(
        &self,
        table_model: tst::TableModelArchive,
        name_precharged: bool,
        initial_budget: Option<ProjectionBudget>,
    ) -> Result<Table> {
        // Projection is transactional at the table boundary. A rejected legacy
        // candidate must not consume retained-cell or retained-text capacity
        // that belongs to a later schema-proven table. Formula work remains a
        // monotonic package-wide cost across successful and rejected attempts.
        let mut projection_budget =
            initial_budget.unwrap_or_else(|| *self.projection_budget.borrow());
        let result = (|| {
            let (row_count, column_count) = checked_table_dimensions(
                table_model.number_of_rows,
                table_model.number_of_columns,
            )?;
            if !name_precharged {
                projection_budget.charge_output_text(table_model.table_name.len())?;
            }
            let mut table =
                Table::with_dimensions(table_model.table_name, row_count, column_count)?;

            // Extract string table for cell text values
            // string_table is a required field, not Optional
            let string_table = self.load_string_table(
                table_model.base_data_store.string_table.identifier,
                &mut projection_budget,
            )?;

            // Extract formula table for formula cells
            // formula_table is a required field, not Optional
            let formula_table = self.load_formula_table(
                table_model.base_data_store.formula_table.identifier,
                &mut projection_budget,
            )?;
            let formula_references = if formula_table.is_empty() {
                None
            } else {
                let (references, cost) = build_formula_reference_maps(
                    self.bundle,
                    self.object_index,
                    projection_budget.remaining_references(),
                    projection_budget.remaining_payload_work(),
                    MAX_FORMULA_WIRE_BYTES.saturating_sub(projection_budget.payload_work),
                    projection_budget.remaining_staging_text_bytes(),
                )?;
                projection_budget.charge_references(cost.retained_entries)?;
                projection_budget.charge_staging_text(cost.text_bytes)?;
                projection_budget.payload_work = projection_charge(
                    projection_budget.payload_work,
                    cost.work_items,
                    MAX_PAYLOAD_WORK,
                    SemanticLimitKind::FormulaWork,
                )?;
                Some(references)
            };
            let empty_formula_references = FormulaReferenceMaps::default();
            let formula_error_table = match table_model.base_data_store.formula_error_table {
                Some(reference) => {
                    self.load_formula_error_table(reference.identifier, &mut projection_budget)?
                },
                None => Box::default(),
            };

            let rich_text_table = match table_model.base_data_store.rich_text_table {
                Some(reference) => {
                    self.load_rich_text_table(reference.identifier, &mut projection_budget)?
                },
                None => Box::default(),
            };
            let comment_table = if self.retain_comments {
                match table_model.base_data_store.comment_storage_table {
                    Some(reference) => {
                        Some(self.load_comment_table(reference.identifier, &mut projection_budget)?)
                    },
                    None => None,
                }
            } else {
                None
            };

            // Parse tiles to extract cell data
            let cell_tables = CellTables {
                strings: &string_table,
                formulas: &formula_table,
                formula_errors: &formula_error_table,
                rich_text: &rich_text_table,
                comments: comment_table.as_ref(),
                formula_references: formula_references
                    .as_ref()
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
    fn load_string_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<StringTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat("Numbers string entry has no string value".to_owned())
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers string sidecar", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::String,
            budget,
            &mut converter,
        )
    }

    fn load_formula_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<FormulaTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                let value = entry.formula().ok_or_else(|| {
                    Error::InvalidFormat("Numbers formula entry has no formula payload".to_owned())
                })?;
                budget.charge_formula_wire(value.len())?;
                tsce::FormulaArchive::decode(value).map_err(Error::protobuf)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::Formula,
            budget,
            &mut converter,
        )
    }

    fn load_formula_error_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<FormulaErrorTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers formula-error entry has no string value".to_owned(),
                    )
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| allocation_error("Numbers formula-error sidecar", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::FormulaError,
            budget,
            &mut converter,
        )
    }

    fn load_rich_text_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<StringTable> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             budget: &mut ProjectionBudget| {
                let payload_reference = entry.rich_text_payload().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers rich-text entry has no payload reference".to_owned(),
                    )
                })?;
                let payload_object = self
                    .object_index
                    .resolve_ref_id(self.bundle, payload_reference.identifier())?
                    .ok_or_else(|| {
                        Error::InvalidFormat("Numbers rich-text payload is missing".to_owned())
                    })?;
                let mut messages = payload_object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == RICH_TEXT_PAYLOAD_MESSAGE_TYPE);
                let payload_message = messages.next().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Numbers rich-text payload has no canonical message".to_owned(),
                    )
                })?;
                if messages.next().is_some() {
                    return Err(Error::InvalidFormat(
                        "Numbers rich-text payload has duplicate canonical messages".to_owned(),
                    ));
                }
                let (projected_storage, payload_report) =
                    preflight_rich_text_payload(&payload_message.data)?;
                budget.charge_wire_preflight(payload_report)?;
                budget.charge_references(1)?;
                // The bounded preflight above is the authoritative projection of this
                // payload.  Do not decode the complete RichTextPayloadArchive here:
                // it is a tiny envelope whose only parse-relevant value is the local
                // storage reference, and a generated decode would allocate an entire
                // archive before the text-wire budgets have been applied.
                self.extract_rich_text(projected_storage, budget)
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::RichTextPayload,
            budget,
            &mut converter,
        )
    }

    fn load_comment_table(
        &self,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> Result<CommentTable> {
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
                let source_text = comment.text.as_deref().unwrap_or_default();
                let mut text = String::new();
                text.try_reserve_exact(source_text.len())
                    .map_err(|_| allocation_error("Numbers comment text", source_text.len()))?;
                text.push_str(source_text);
                let mut reply_ids = Vec::new();
                reply_ids
                    .try_reserve_exact(comment.replies.len())
                    .map_err(|_| {
                        allocation_error("Numbers comment replies", comment.replies.len())
                    })?;
                for reply in &comment.replies {
                    reply_ids
                        .push(StorageId::from_raw(reply.identifier).map_err(map_comment_error)?);
                }
                Ok(Comment {
                    text,
                    creation_date_seconds: comment.creation_date.as_ref().map(|date| date.seconds),
                    author_id: comment
                        .author
                        .as_ref()
                        .map(|author| {
                            AuthorId::from_raw(author.identifier).map_err(map_comment_error)
                        })
                        .transpose()?,
                    reply_ids: reply_ids.into_boxed_slice(),
                    storage_uuid: comment
                        .storage_uuid
                        .as_ref()
                        .map(|uuid| {
                            Uuid::from_parts(uuid.lower, uuid.upper).map_err(map_comment_error)
                        })
                        .transpose()?,
                })
            };
        self.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::CommentStorage,
            budget,
            &mut converter,
        )
    }

    fn load_table_data_list_entries<T, C>(
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
            let list_type_probe = probe_table_data_list_type(&message.data);
            budget.charge_list_type_probe(list_type_probe.fields, message.data.len())?;
            let duplicate_candidate = selected_values.is_some();
            let admitting_candidate = !duplicate_candidate
                && list_type_probe
                    .list_type
                    .is_none_or(|value| value == expected);
            let options = table_cell_decode_options(
                &message.data,
                if admitting_candidate {
                    budget.remaining_references()
                } else {
                    usize::MAX
                },
                if admitting_candidate {
                    budget.remaining_staging_text_bytes()
                } else {
                    usize::MAX
                },
                budget.remaining_payload_fields(),
                budget.remaining_payload_work(),
            );
            let reference_offset = budget.references;
            let mut visitor =
                TypedListVisitor::new(converter, budget, expected, false, admitting_candidate);
            let decoded = numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
                &message.data,
                options,
                &mut visitor,
            )
            .map_err(|error| {
                map_table_cell_codec_error_with_reference_offset(error, reference_offset)
            })?;
            let (snapshot, report) = decoded;
            let (values, keys, references, callback_structural, callback_semantic) =
                visitor.take_parts();
            if snapshot.list_type() != expected {
                // A candidate for another list type is still strictly walked,
                // but its references and text are not admitted to this table.
                budget.charge_decode_work(report)?;
                continue;
            }
            if selected_values.is_some() {
                // A duplicate selected root is structurally inspected but is
                // never admitted to the aggregate reference/text budgets.
                budget.charge_decode_work(report)?;
                record_first_list_error(
                    &mut structural_error,
                    Error::InvalidFormat(format!(
                        "Object {object_id} has multiple Numbers {list_type:?} TableDataList payloads"
                    )),
                );
                continue;
            }
            budget.charge_decode_report(report)?;
            budget.charge_staging_text(report.text_bytes())?;
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
                let list_type_probe = probe_table_data_list_type(&segment_message.data);
                budget
                    .charge_list_type_probe(list_type_probe.fields, segment_message.data.len())?;
                let admitting_segment = segment_count == 1
                    && list_type_probe
                        .list_type
                        .is_none_or(|value| value == expected);
                let options = table_cell_decode_options(
                    &segment_message.data,
                    if admitting_segment {
                        budget.remaining_references()
                    } else {
                        usize::MAX
                    },
                    if admitting_segment {
                        budget.remaining_staging_text_bytes()
                    } else {
                        usize::MAX
                    },
                    budget.remaining_payload_fields(),
                    budget.remaining_payload_work(),
                );
                let reference_offset = budget.references;
                let mut visitor =
                    TypedListVisitor::new(converter, budget, expected, true, admitting_segment);
                let decoded =
                    numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
                        &segment_message.data,
                        options,
                        &mut visitor,
                    )
                    .map_err(|error| {
                        map_table_cell_codec_error_with_reference_offset(error, reference_offset)
                    })?;
                let (snapshot, report) = decoded;
                let bounds = visitor.take_segment_bounds();
                let (
                    segment_values,
                    _segment_keys,
                    _segment_refs,
                    callback_structural,
                    callback_semantic,
                ) = visitor.take_parts();
                if segment_count > 1 {
                    // Duplicate segment payloads are fully wire-checked but
                    // cannot contribute references or text to the admitted
                    // segment's aggregate budget.
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
                budget.charge_decode_report(report)?;
                budget.charge_staging_text(report.text_bytes())?;
                if let Some(error) = callback_structural {
                    record_first_list_error(&mut structural_error, error);
                }
                if let Some(error) = callback_semantic {
                    record_first_list_error(&mut semantic_error, error);
                }
                let end = snapshot
                    .key_range_location()
                    .checked_add(snapshot.key_range_length());
                let Some(end) = end else {
                    record_first_list_error(
                        &mut structural_error,
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list segment {segment_id} key range overflows"
                        )),
                    );
                    continue;
                };
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
                    if keys.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error("Numbers table-list entry keys", keys.len() + 1),
                        );
                        continue;
                    }
                    keys.insert(key);
                    if values.try_reserve(1).is_err() {
                        record_first_list_error(
                            &mut semantic_error,
                            allocation_error("Numbers table-list entries", values.len() + 1),
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
            let options = table_cell_decode_options(
                &msg.data,
                projection_budget.remaining_references(),
                projection_budget.remaining_output_text_bytes(),
                projection_budget.remaining_payload_fields(),
                projection_budget.remaining_payload_work(),
            );
            let reference_offset = projection_budget.references;
            let (materialized_cells, semantic_error, report) = {
                let mut visitor = TileRowVisitor {
                    row_origin,
                    tile_size,
                    row_count,
                    column_count,
                    budget,
                    cell_tables,
                    projection_budget,
                    table,
                    materialized_cells: 0,
                    semantic_error: None,
                };
                let (_, report) = numbers_table_cell_storage_codec::decode_tile_with_visitor(
                    &msg.data,
                    options,
                    &mut visitor,
                )
                .map_err(|error| {
                    map_table_cell_codec_error_with_reference_offset(error, reference_offset)
                })?;
                (
                    visitor.materialized_cells,
                    visitor.semantic_error.take(),
                    report,
                )
            };
            projection_budget.charge_decode_report(report)?;
            projection_budget.charge_materialized_cells(materialized_cells)?;
            if let Some(error) = semantic_error {
                return Err(error);
            }
            decoded = true;
        }

        if !decoded {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile object {tile_id} has no tile payload"
            )));
        }

        Ok(())
    }

    /// Parse a single tile row
    fn parse_tile_row(
        row_info: &numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
        row_origin: usize,
        tile_size: usize,
        row_count: usize,
        column_count: usize,
        budget: &mut CellBudget,
        cell_tables: &CellTables<'_>,
        projection_budget: &mut ProjectionBudget,
        table: &mut Table,
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
        validate_table_row(row_index, row_count)?;

        // The cell_storage_buffer contains serialized Cell messages
        // The cell_offsets buffer contains the byte offsets for each cell

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
        // The strict borrowed-row projection charged the aggregate declared
        // cell count before row-level cell decoding.
        let cells = Self::parse_cell_offsets(
            cell_offsets,
            cell_storage.len(),
            row_info.has_wide_offsets().unwrap_or(false),
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
            if let Some(identifier) = parsed.comment_identifier
                && let Some(comments) = cell_tables.comments
            {
                let comment = compact_table_get(comments, identifier).ok_or_else(|| {
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
    fn extract_rich_text(&self, storage_id: u64, budget: &mut ProjectionBudget) -> Result<String> {
        if storage_id == 0 {
            return Err(Error::InvalidFormat(
                "Numbers rich-text storage has a null reference".to_owned(),
            ));
        }
        let resolved = self
            .object_index
            .resolve_ref_id(self.bundle, storage_id)?
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers rich-text storage is missing".to_owned())
            })?;
        let mut messages = resolved
            .messages
            .iter()
            .filter(|message| message.type_ == STORAGE_MESSAGE_TYPE);
        let message = messages.next().ok_or_else(|| {
            Error::InvalidFormat("Numbers rich-text storage has no canonical message".to_owned())
        })?;
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(
                "Numbers rich-text storage has duplicate canonical messages".to_owned(),
            ));
        }
        let remaining = if self.document_projection {
            budget.remaining_staging_text_bytes()
        } else {
            budget.remaining_output_text_bytes()
        };
        if remaining == 0 {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 1,
                maximum: 0,
                path: SemanticPath::Package,
            });
        }
        let limits = litchi_iwa_text_wire::Limits::new(
            message.data.len().max(1),
            message.data.len().max(1),
            message.data.len().max(1),
            remaining,
        )
        .map_err(|_error| {
            Error::InvalidFormat("Numbers rich-text limits are invalid".to_owned())
        })?;
        let validated_text_len = if self.document_projection {
            let rewrite_limits = litchi_iwa_text_wire::RewriteLimits::new(
                message.data.len().max(1),
                budget.remaining_payload_fields().max(1),
                8,
                crate::MAX_REFERENCES,
                remaining,
                crate::MAX_REFERENCES,
                budget.remaining_references().max(1),
                message.data.len().max(1),
                budget.remaining_payload_work().max(1),
            )
            .map_err(|_error| {
                Error::InvalidFormat("Numbers rich-text limits are invalid".to_owned())
            })?;
            let validation =
                litchi_iwa_text_wire::validate_storage_with_limits(&message.data, rewrite_limits)
                    .map_err(map_rich_text_rewrite_error)?;
            budget.payload_fields = projection_charge(
                budget.payload_fields,
                validation.fields(),
                crate::MAX_REFERENCES,
                SemanticLimitKind::Objects,
            )?;
            budget.charge_references(validation.reference_occurrences())?;
            let second_pass_work = message
                .data
                .len()
                .checked_mul(2)
                .and_then(|work| work.checked_add(validation.utf8_len().saturating_mul(2)))
                .ok_or_else(|| {
                    formula_semantic_limit(
                        SemanticLimitKind::FormulaWork,
                        usize::MAX,
                        MAX_PAYLOAD_WORK,
                    )
                })?;
            budget.payload_work = projection_charge(
                budget.payload_work,
                validation
                    .validation_work()
                    .saturating_add(second_pass_work),
                MAX_PAYLOAD_WORK,
                SemanticLimitKind::FormulaWork,
            )?;
            Some(validation.utf8_len())
        } else {
            None
        };
        let storage = litchi_iwa_text_wire::from_bytes_with_limits(&message.data, limits)
            .map_err(map_rich_text_error)?;
        if self.document_projection {
            if validated_text_len != Some(storage.len()) {
                return Err(Error::InvalidFormat(
                    "Numbers rich-text storage failed strict text parity".to_owned(),
                ));
            }
            budget.charge_staging_text(storage.len())?;
        } else {
            budget.charge_output_text(storage.len())?;
        }
        Ok(storage.into_text())
    }
}

fn preflight_rich_text_payload(
    source: &[u8],
) -> Result<(u64, litchi_iwa_common::wire::WirePreflight)> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(1)?;
    let mut storage = None;
    let mut cell = false;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        match visit.field().number() {
            1 => {
                if storage.is_some() || visit.field().wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text storage reference".to_owned(),
                    ));
                }
                visit.field().validate_canonical_framing()?;
                storage = Some(names::preflight_local_reference(visit.field().payload())?);
            },
            3 => {
                if cell || visit.field().wire_type() != 2 {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "invalid rich-text cell owner".to_owned(),
                    ));
                }
                visit.field().validate_canonical_framing()?;
                cell = true;
            },
            _ => {},
        }
        Ok(WireDescent::Skip)
    })?;
    if !cell {
        return Err(Error::InvalidFormat(
            "Numbers rich-text payload has no cell owner".to_owned(),
        ));
    }
    let storage = storage.ok_or_else(|| {
        Error::InvalidFormat("Numbers rich-text payload has no storage reference".to_owned())
    })?;
    Ok((storage, report))
}

fn map_rich_text_error(error: litchi_iwa_text_wire::Error) -> Error {
    match error {
        litchi_iwa_text_wire::Error::TooManyTextBytes { actual, limit } => Error::SemanticLimit {
            kind: SemanticLimitKind::OutputTextBytes,
            observed: actual,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::Error::TooManyFragments { actual, limit } => Error::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed: actual,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::Error::Common(error) => Error::Common(error),
        _ => Error::InvalidFormat("Numbers rich-text storage is invalid".to_owned()),
    }
}

fn map_rich_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::SemanticLimit {
            kind: if resource.contains("reference") {
                SemanticLimitKind::References
            } else if resource.contains("field") || resource.contains("entry") {
                SemanticLimitKind::Objects
            } else if resource.contains("text") {
                SemanticLimitKind::TextBytes
            } else if resource.contains("work") {
                SemanticLimitKind::FormulaWork
            } else {
                SemanticLimitKind::FormulaWireBytes
            },
            observed,
            maximum: limit,
            path: SemanticPath::Package,
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            allocation_error("Numbers rich-text storage validation", amount)
        },
        _ => Error::InvalidFormat("Numbers rich-text storage is invalid".to_owned()),
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
    maximum_work: usize,
    maximum_wire_bytes: usize,
    maximum_text_bytes: usize,
}

impl FormulaReferenceBudget {
    const fn new(
        maximum_retained_entries: usize,
        maximum_work: usize,
        maximum_wire_bytes: usize,
        maximum_text_bytes: usize,
    ) -> Self {
        Self {
            retained_entries: 0,
            maximum_retained_entries,
            work_items: 0,
            wire_bytes: 0,
            text_bytes: 0,
            maximum_work,
            maximum_wire_bytes,
            maximum_text_bytes,
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
            formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                usize::MAX,
                self.maximum_work,
            )
        })?;
        if observed > self.maximum_work {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWork,
                observed,
                self.maximum_work,
            ));
        }
        Ok(())
    }

    fn charge_wire_bytes(&mut self, bytes: usize) -> Result<()> {
        self.charge_work(bytes)?;
        self.wire_bytes = self.wire_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                usize::MAX,
                self.maximum_wire_bytes,
            )
        })?;
        if self.wire_bytes > self.maximum_wire_bytes {
            return Err(formula_semantic_limit(
                SemanticLimitKind::FormulaWireBytes,
                self.wire_bytes,
                self.maximum_wire_bytes,
            ));
        }
        Ok(())
    }

    fn charge_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                usize::MAX,
                self.maximum_text_bytes,
            )
        })?;
        if self.text_bytes > self.maximum_text_bytes {
            return Err(formula_semantic_limit(
                SemanticLimitKind::TextBytes,
                self.text_bytes,
                self.maximum_text_bytes,
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
    max_work: usize,
    max_wire_bytes: usize,
    max_text_bytes: usize,
) -> Result<(FormulaReferenceMaps, FormulaReferenceBudget)> {
    let mut budget = FormulaReferenceBudget::new(
        max_formula_references,
        max_work,
        max_wire_bytes,
        max_text_bytes,
    );
    let mut result = FormulaReferenceMaps::default();
    result
        .categories
        .try_reserve(1)
        .map_err(|_error| allocation_error("Numbers formula categories", 1))?;
    let mut grand_total = String::new();
    grand_total
        .try_reserve_exact("Grand Total".len())
        .map_err(|_error| allocation_error("Numbers formula categories", "Grand Total".len()))?;
    grand_total.push_str("Grand Total");
    result.categories.insert([1, 0], grand_total);
    let mut table_info_names = HashMap::<u64, FormulaReferenceName>::new();
    let root_message = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .and_then(|object| object.messages.iter().find(|message| message.type_ == 1));

    if let Some(root) = root_message {
        budget.charge_wire_bytes(root.data.len())?;
        budget.charge_work(root.data.len().saturating_mul(7))?;
        let options = litchi_iwa_protos::numbers_sheet_order_codec::DecodeOptions::new(
            root.data.len().max(1),
            root.data.len().saturating_mul(2).max(1),
            budget.maximum_work.max(1),
            2,
            budget.maximum_retained_entries.saturating_add(1).max(1),
        );
        let sheet_order =
            litchi_iwa_protos::numbers_sheet_order_codec::decode_document_sheet_order(
                &root.data, options,
            )
            .map_err(|_error| Error::MalformedPayload {
                path: SemanticPath::Document,
            })?;
        for sheet_reference in sheet_order.sheet_references() {
            budget.charge_work(1)?;
            let Some(sheet_object) =
                object_index.resolve_ref_id(bundle, sheet_reference.identifier())?
            else {
                continue;
            };
            let Some(sheet_message) = sheet_object.messages.iter().find(|message| {
                message.type_ == super::SHEET_MESSAGE_TYPE
                    || message.type_ == super::FORM_BASED_SHEET_MESSAGE_TYPE
            }) else {
                continue;
            };
            budget.charge_wire_bytes(sheet_message.data.len())?;
            budget.charge_work(sheet_message.data.len().saturating_mul(7))?;
            let (sheet_name, drawables) = names::preflight_sheet_payload(
                sheet_message.type_,
                &sheet_message.data,
                budget.maximum_retained_entries,
            )
            .map_err(|_error| Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })?;
            let mut cached_sheet_name = None::<Arc<String>>;
            for drawable in drawables {
                budget.charge_work(1)?;
                let Some(drawable_object) = object_index.resolve_ref_id(bundle, drawable)? else {
                    continue;
                };
                let table_name = formula_table_name(
                    bundle,
                    object_index,
                    drawable_object.messages,
                    &mut budget,
                )?;
                if let Some(table) = table_name {
                    let is_new = !table_info_names.contains_key(&drawable);
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
                        let mut name = String::new();
                        name.try_reserve_exact(sheet_name.len()).map_err(|_error| {
                            allocation_error("Numbers formula sheet name", sheet_name.len())
                        })?;
                        name.push_str(sheet_name);
                        let name = Arc::new(name);
                        cached_sheet_name = Some(Arc::clone(&name));
                        name
                    };
                    table_info_names.insert(
                        drawable,
                        FormulaReferenceName {
                            sheet: retained_sheet_name,
                            table: Arc::new(table),
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
                budget.charge_wire_bytes(message.data.len())?;
                let (key, table_identifier, report) = match preflight_formula_owner(&message.data) {
                    Ok(projection) => projection,
                    Err(_) => continue,
                };
                budget.charge_work(report.scanned_bytes().saturating_add(report.fields()))?;
                let Some(name) = table_info_names.get(&table_identifier) else {
                    continue;
                };
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
    Ok((result, budget))
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
        if let Some(message) = model_message {
            budget.charge_wire_bytes(message.data.len())?;
            budget.charge_work(message.data.len().saturating_mul(7))?;
            let name = names::preflight_table_name(&message.data).map_err(|_error| {
                Error::MalformedPayload {
                    path: SemanticPath::StructuredTables,
                }
            })?;
            budget.charge_text(name.len())?;
            let mut retained = String::new();
            retained
                .try_reserve_exact(name.len())
                .map_err(|_error| allocation_error("Numbers formula table name", name.len()))?;
            retained.push_str(name);
            return Ok(Some(retained));
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
    let remaining_work = budget.maximum_work.saturating_sub(budget.work_items);
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
    let label = match label {
        Cow::Owned(label) => label,
        Cow::Borrowed(label) => {
            let mut retained = String::new();
            retained.try_reserve_exact(label.len()).map_err(|_error| {
                allocation_error("Numbers formula category label", label.len())
            })?;
            retained.push_str(label);
            retained
        },
    };
    names.insert(key, label);
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

fn preflight_formula_owner(
    source: &[u8],
) -> Result<(FormulaOwnerKey, u64, litchi_iwa_common::wire::WirePreflight)> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().clamp(1, WireLimits::MAX_INPUT_BYTES))?
        .with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS))?
        .with_nesting(2)?;
    let mut owner_key = None;
    let mut table = None;
    let report = preflight_wire_tree_with_limits(source, limits, |visit| {
        if visit.path().is_empty() && visit.field().number() == 1 {
            if owner_key.is_some() || visit.field().wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid formula owner".into(),
                ));
            }
            let mut lower = None;
            let mut upper = None;
            let nested = preflight_wire_tree_with_limits(
                visit.field().payload(),
                WireLimits::default()
                    .with_input_bytes(
                        visit
                            .field()
                            .payload()
                            .len()
                            .clamp(1, WireLimits::MAX_INPUT_BYTES),
                    )?
                    .with_fields(8)?
                    .with_nesting(1)?,
                |uuid| {
                    if !uuid.path().is_empty() || uuid.field().wire_type() != 0 {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    let (value, length) =
                        litchi_iwa_common::varint::decode_varint_from_bytes(uuid.field().payload())
                            .map_err(|_error| {
                                litchi_iwa_common::Error::InvalidFormat(
                                    "invalid formula owner UUID".into(),
                                )
                            })?;
                    if length != uuid.field().payload().len() {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    let canonical_length = if value == 0 {
                        1
                    } else {
                        usize::try_from((64 - value.leading_zeros()).div_ceil(7)).map_err(
                            |_error| {
                                litchi_iwa_common::Error::InvalidFormat(
                                    "invalid formula owner UUID".into(),
                                )
                            },
                        )?
                    };
                    if length != canonical_length {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid formula owner UUID".into(),
                        ));
                    }
                    match uuid.field().number() {
                        1 if lower.replace(value).is_none() => {},
                        2 if upper.replace(value).is_none() => {},
                        _ => {
                            return Err(litchi_iwa_common::Error::InvalidFormat(
                                "invalid formula owner UUID".into(),
                            ));
                        },
                    }
                    Ok(WireDescent::Skip)
                },
            )?;
            let _ = nested;
            let lower = lower.ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat("missing formula owner UUID".into())
            })?;
            let upper = upper.ok_or_else(|| {
                litchi_iwa_common::Error::InvalidFormat("missing formula owner UUID".into())
            })?;
            owner_key = Some([
                lower as u32,
                (lower >> 32) as u32,
                upper as u32,
                (upper >> 32) as u32,
            ]);
        } else if visit.path().is_empty() && visit.field().number() == 11 {
            if table.is_some() || visit.field().wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid formula owner table".into(),
                ));
            }
            table = Some(names::preflight_local_reference(visit.field().payload())?);
        }
        Ok(WireDescent::Skip)
    })?;
    Ok((
        owner_key
            .ok_or_else(|| Error::InvalidFormat("Numbers formula owner has no UUID".to_owned()))?,
        table
            .ok_or_else(|| Error::InvalidFormat("Numbers formula owner has no table".to_owned()))?,
        report,
    ))
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
                .map(|number| {
                    let value = fallible_formula_display(number, renderer, budget)?;
                    renderer.owned_expr(value, budget)
                })
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
                    let days = seconds / 86_400.0;
                    let value = fallible_formula_format(renderer, budget, |output| {
                        write!(output, "(DATE(2001,1,1)+{days})")
                    })?;
                    renderer.owned_expr(value, budget)
                })
                .transpose()?,
            AstNodeType::DurationNode => node
                .ast_duration_node_unit_num
                .map(|value| {
                    let value = fallible_formula_display(value, renderer, budget)?;
                    renderer.owned_expr(value, budget)
                })
                .transpose()?,
            AstNodeType::EmptyArgumentNode => Some(renderer.static_expr("", budget)?),
            AstNodeType::CellReferenceNode => Some(renderer.owned_expr(
                render_cell_reference_checked(
                    node,
                    host_row,
                    host_column,
                    formula_references,
                    renderer,
                    budget,
                )?,
                budget,
            )?),
            AstNodeType::LocalCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_local_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || fallible_formula_owned("#REF!", renderer, budget),
                            |cell| {
                                let column = FormulaColumn(cell.column_handle);
                                let row = cell.row_handle + 1;
                                fallible_formula_format(renderer, budget, |output| {
                                    write!(output, "{column}{row}")
                                })
                            },
                        )?,
                    budget,
                )?,
            ),
            AstNodeType::CrossTableCellReferenceNode => Some(
                renderer.owned_expr(
                    node.ast_cross_table_cell_reference_node_reference
                        .as_ref()
                        .map_or_else(
                            || fallible_formula_owned("#REF!", renderer, budget),
                            |cell| {
                                let prefix = formula_reference_prefix_parts(
                                    &cell.table_id,
                                    formula_references,
                                );
                                let column = FormulaColumn(cell.column_handle);
                                let row = cell.row_handle + 1;
                                fallible_formula_format(renderer, budget, |output| {
                                    write_formula_reference_prefix(output, prefix)?;
                                    write!(output, "{column}{row}")
                                })
                            },
                        )?,
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
                        Some(fallible_function_name(index, renderer, budget)?),
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
                render_colon_tract_checked(
                    node,
                    host_row,
                    host_column,
                    formula_references,
                    renderer,
                    budget,
                )?,
                budget,
            )?),
            AstNodeType::ReferenceErrorNode | AstNodeType::ReferenceErrorWithUids => {
                Some(renderer.static_expr("#REF!", budget)?)
            },
            AstNodeType::CategoryRefNode => Some(renderer.owned_expr(
                render_category_reference_checked(node, formula_references, renderer, budget)?,
                budget,
            )?),
            AstNodeType::UnknownFunctionNode => {
                let arguments = pop_formula_arguments(
                    &mut stack,
                    node.ast_unknown_function_node_num_args.unwrap_or(0),
                    "unknown function",
                )?;
                Some(
                    renderer.comma_joined(
                        Some(fallible_formula_owned(
                            node.ast_unknown_function_node_string
                                .as_deref()
                                .unwrap_or("UNKNOWN"),
                            renderer,
                            budget,
                        )?),
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

fn fallible_formula_owned(
    value: &str,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    renderer.check_additional_owned(value.len(), budget)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation_error("Numbers formula owned text", value.len()))?;
    owned.push_str(value);
    Ok(owned)
}

fn fallible_formula_display(
    value: impl std::fmt::Display,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    #[derive(Default)]
    struct Counter {
        bytes: usize,
    }
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.bytes = self.bytes.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write!(&mut counter, "{value}")
        .map_err(|_error| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(counter.bytes, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.bytes)
        .map_err(|_error| allocation_error("Numbers formula owned text", counter.bytes))?;
    write!(&mut output, "{value}")
        .map_err(|_error| Error::InvalidFormat("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

fn fallible_formula_format(
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
    write_value: impl Fn(&mut dyn std::fmt::Write) -> std::fmt::Result,
) -> Result<String> {
    #[derive(Default)]
    struct Counter(usize);
    impl std::fmt::Write for Counter {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            self.0 = self.0.checked_add(value.len()).ok_or(std::fmt::Error)?;
            Ok(())
        }
    }
    let mut counter = Counter::default();
    write_value(&mut counter).map_err(|_error| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(counter.0, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(counter.0)
        .map_err(|_error| allocation_error("Numbers formula owned text", counter.0))?;
    write_value(&mut output)
        .map_err(|_error| Error::InvalidFormat("Numbers formula formatting failed".to_owned()))?;
    Ok(output)
}

#[derive(Clone, Copy)]
struct FormulaColumn(u32);

impl std::fmt::Display for FormulaColumn {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut bytes = [0_u8; 7];
        let mut cursor = bytes.len();
        let mut value = self.0;
        loop {
            cursor -= 1;
            bytes[cursor] = b'A' + u8::try_from(value % 26).map_err(|_error| std::fmt::Error)?;
            if value < 26 {
                break;
            }
            value = value / 26 - 1;
        }
        formatter
            .write_str(std::str::from_utf8(&bytes[cursor..]).map_err(|_error| std::fmt::Error)?)
    }
}

fn fallible_function_name(
    index: u32,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    if let Some(name) = super::function_map::function_name(index) {
        fallible_formula_owned(name, renderer, budget)
    } else {
        fallible_formula_format(renderer, budget, |output| write!(output, "FUNC{index}"))
    }
}

type FormulaPrefix<'a> = Option<&'a FormulaReferenceName>;

fn formula_reference_prefix_parts<'a>(
    owner: &litchi_iwa_protos::tsp::CfuuidArchive,
    references: &'a FormulaReferenceMaps,
) -> FormulaPrefix<'a> {
    cfuuid_key(owner).and_then(|key| references.owners.get(&key))
}

fn write_formula_reference_prefix(
    output: &mut dyn std::fmt::Write,
    prefix: FormulaPrefix<'_>,
) -> std::fmt::Result {
    if let Some(name) = prefix {
        write!(output, "{}::{}::", name.sheet, name.table)
    } else {
        output.write_str("Table::")
    }
}

fn render_category_reference_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
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
    let Some(label) =
        category_uid.and_then(|uid| references.categories.get(&formula_category_key(uid)))
    else {
        return fallible_formula_owned("#CATEGORY!", renderer, budget);
    };
    let escaped_extra = label
        .bytes()
        .filter(|byte| *byte == b'\\' || *byte == b']')
        .count();
    let required = "#CATEGORY!["
        .len()
        .checked_add(label.len())
        .and_then(|length| length.checked_add(escaped_extra))
        .and_then(|length| length.checked_add(1))
        .ok_or_else(|| formula_output_limit_error(usize::MAX, budget))?;
    renderer.check_additional_owned(required, budget)?;
    let mut output = String::new();
    output
        .try_reserve_exact(required)
        .map_err(|_error| allocation_error("Numbers formula category text", required))?;
    output.push_str("#CATEGORY![");
    for character in label.chars() {
        if character == '\\' || character == ']' {
            output.push('\\');
        }
        output.push(character);
    }
    output.push(']');
    Ok(output)
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

fn render_cell_reference_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
) -> Result<String> {
    if let (Some(ast_column), Some(ast_row)) = (&node.ast_column, &node.ast_row) {
        let column = FormulaColumn(resolve_formula_coordinate(
            host_column,
            ast_column.column,
            ast_column.absolute.unwrap_or(false),
            "column",
        )?);
        let row = resolve_formula_coordinate(
            host_row,
            ast_row.row,
            ast_row.absolute.unwrap_or(false),
            "row",
        )? + 1;
        let prefix = node
            .ast_cross_table_reference_extra_info
            .as_ref()
            .and_then(|extra| formula_reference_prefix_parts(&extra.table_id, formula_references));
        return fallible_formula_format(renderer, budget, |output| {
            if node.ast_cross_table_reference_extra_info.is_some() {
                write_formula_reference_prefix(output, prefix)?;
            }
            write!(
                output,
                "{}{column}{}{row}",
                if ast_column.absolute.unwrap_or(false) {
                    "$"
                } else {
                    ""
                },
                if ast_row.absolute.unwrap_or(false) {
                    "$"
                } else {
                    ""
                }
            )
        });
    }
    if let Some(cell) = &node.ast_local_cell_reference_node_reference {
        let column = FormulaColumn(cell.column_handle);
        let row = cell.row_handle + 1;
        return fallible_formula_format(renderer, budget, |output| {
            write!(
                output,
                "{}{column}{}{row}",
                if cell.column_is_sticky != 0 { "$" } else { "" },
                if cell.row_is_sticky != 0 { "$" } else { "" }
            )
        });
    }
    if let Some(cell) = &node.ast_cross_table_cell_reference_node_reference {
        let prefix = formula_reference_prefix_parts(&cell.table_id, formula_references);
        let column = FormulaColumn(cell.column_handle);
        let row = cell.row_handle + 1;
        return fallible_formula_format(renderer, budget, |output| {
            write_formula_reference_prefix(output, prefix)?;
            write!(output, "{column}{row}")
        });
    }
    fallible_formula_owned("#REF!", renderer, budget)
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

fn render_colon_tract_checked(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: usize,
    host_column: usize,
    formula_references: &FormulaReferenceMaps,
    renderer: &FormulaRenderer,
    budget: &ProjectionBudget,
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
        .and_then(|extra| formula_reference_prefix_parts(&extra.table_id, formula_references));
    let has_prefix = node.ast_cross_table_reference_extra_info.is_some();
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
    let (begin_column, end_column) = if has_columns {
        let (begin, end) = resolve_colon_axis(
            &tract.relative_column,
            &tract.absolute_column,
            sticky.begin_column_is_absolute,
            sticky.end_column_is_absolute,
            host_column,
            "column",
        )?;
        (Some(FormulaColumn(begin)), Some(FormulaColumn(end)))
    } else {
        (None, None)
    };
    let (begin_row, end_row) = if has_rows {
        let (begin, end) = resolve_colon_axis(
            &tract.relative_row,
            &tract.absolute_row,
            sticky.begin_row_is_absolute,
            sticky.end_row_is_absolute,
            host_row,
            "row",
        )?;
        (Some(u64::from(begin) + 1), Some(u64::from(end) + 1))
    } else {
        (None, None)
    };
    if !has_columns && !has_rows {
        return Err(Error::ParseError(
            "Numbers formula colon tract has no row or column coordinates".to_owned(),
        ));
    }
    fallible_formula_format(renderer, budget, |output| {
        if has_prefix {
            write_formula_reference_prefix(output, prefix)?;
        }
        if let Some(column) = begin_column {
            if sticky.begin_column_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{column}")?;
        }
        if let Some(row) = begin_row {
            if sticky.begin_row_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{row}")?;
        }
        output.write_char(':')?;
        if let Some(column) = end_column {
            if sticky.end_column_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{column}")?;
        }
        if let Some(row) = end_row {
            if sticky.end_row_is_absolute {
                output.write_char('$')?;
            }
            write!(output, "{row}")?;
        }
        Ok(())
    })
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
        CellBudget, CellTables, Error, FormulaReferenceBudget, FormulaReferenceMaps,
        FormulaRenderer, MAX_FORMULA_CATEGORY_DEPTH, MAX_FORMULA_WIRE_BYTES, MAX_FORMULA_WORK,
        ProjectionBudget, Table, TableDataExtractor, TileRowVisitor,
        collect_formula_category_payload, decode_legacy_table_candidate,
        has_legacy_table_model_wire_shape, map_table_cell_decode_limit_with_reference_offset,
        render_formula, render_formula_ast_array,
    };
    use crate::cell::Value as CellValue;
    use crate::cell::wire::{BncCell, decimal128_le};
    use crate::package::{
        Components, Index, ReadOptions, SemanticPath, compatibility_tables_from_bytes_with_options,
    };
    use crate::{
        DEFAULT_MAX_TEXT_BYTES, Package, PackageSemanticLimits as SemanticLimits, SemanticLimitKind,
    };
    use litchi_iwa_archive::Limits;
    use litchi_iwa_common::comment::Comment;
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};
    use litchi_iwa_protos::{numbers_table_cell_storage_codec, tn, tsce, tsp, tst};
    use prost::Message as _;
    use std::collections::HashMap;
    use std::path::PathBuf;

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

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn archive_object(identifier: u64, messages: Vec<RawMessage>) -> super::Result<ArchiveObject> {
        ArchiveObject::new(identifier, messages)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    fn compatibility_package(objects: Vec<ArchiveObject>) -> super::Result<Vec<u8>> {
        let archive = Archive { objects }
            .to_bytes()
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let stream = SnappyStream::compress(&archive)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        Ok(litchi_iwa_archive::package::to_bytes(
            [("Index/Document.iwa", stream.as_slice())],
            Limits::default(),
        )?)
    }

    fn legacy_model(name: &str, sidecar_id: u64) -> tst::TableModelArchive {
        tst::TableModelArchive {
            table_name: name.to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(sidecar_id),
                formula_table: reference(sidecar_id),
                ..Default::default()
            },
            ..Default::default()
        }
    }

    fn native_fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    fn empty_list(list_type: tst::table_data_list::ListType) -> RawMessage {
        RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        }
    }

    fn string_list_entry(key: u32, value: &str) -> tst::table_data_list::ListEntry {
        tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            string: Some(value.to_owned()),
            ..Default::default()
        }
    }

    fn list_message(
        list_type: tst::table_data_list::ListType,
        entries: Vec<tst::table_data_list::ListEntry>,
        segments: Vec<u64>,
    ) -> RawMessage {
        RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                entries,
                segments: segments.into_iter().map(reference).collect(),
                ..Default::default()
            }
            .encode_to_vec(),
        }
    }

    fn segment_message(
        list_type: tst::table_data_list::ListType,
        location: u32,
        length: u32,
        entries: Vec<tst::table_data_list::ListEntry>,
    ) -> RawMessage {
        RawMessage {
            type_: 6_011,
            data: tst::TableDataListSegment {
                list_type: list_type as i32,
                key_range: tsp::Range { location, length },
                entries,
            }
            .encode_to_vec(),
        }
    }

    fn nested_unknown_group_prefix(depth: usize) -> Vec<u8> {
        fn varint(output: &mut Vec<u8>, mut value: u64) {
            while value >= 0x80 {
                output.push((value as u8 & 0x7f) | 0x80);
                value >>= 7;
            }
            output.push(value as u8);
        }

        let mut prefix = Vec::new();
        for _ in 0..depth {
            varint(&mut prefix, (90_u64 << 3) | 3);
        }
        for _ in 0..depth {
            varint(&mut prefix, (90_u64 << 3) | 4);
        }
        prefix
    }

    fn with_list_extractor<T>(
        objects: Vec<ArchiveObject>,
        document_projection: bool,
        visit: impl FnOnce(&TableDataExtractor<'_>) -> super::Result<T>,
    ) -> super::Result<T> {
        let bytes = compatibility_package(objects)?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());
        if document_projection {
            visit(&extractor.without_comments())
        } else {
            visit(&extractor)
        }
    }

    fn load_string_list(
        extractor: &TableDataExtractor<'_>,
        object_id: u64,
    ) -> super::Result<(super::CompactTable<String>, ProjectionBudget)> {
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let table = load_string_list_with_budget(extractor, object_id, &mut budget)?;
        Ok((table, budget))
    }

    fn load_string_list_with_budget(
        extractor: &TableDataExtractor<'_>,
        object_id: u64,
        budget: &mut ProjectionBudget,
    ) -> super::Result<super::CompactTable<String>> {
        let mut converter =
            |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
             _budget: &mut ProjectionBudget| {
                let value = entry.string_value().ok_or_else(|| {
                    Error::InvalidFormat("test string entry has no value".to_owned())
                })?;
                let mut retained = String::new();
                retained
                    .try_reserve_exact(value.len())
                    .map_err(|_| super::allocation_error("test string", value.len()))?;
                retained.push_str(value);
                Ok(retained)
            };
        let table = extractor.load_table_data_list_entries(
            object_id,
            tst::table_data_list::ListType::String,
            budget,
            &mut converter,
        )?;
        Ok(table)
    }

    fn tile_row(
        tile_row_index: u32,
        cell_count: u32,
        pre_bnc_storage: Vec<u8>,
        pre_bnc_offsets: Vec<u8>,
        modern_storage: Option<Vec<u8>>,
        modern_offsets: Option<Vec<u8>>,
        has_wide_offsets: Option<bool>,
    ) -> tst::TileRowInfo {
        tst::TileRowInfo {
            tile_row_index,
            cell_count,
            cell_storage_buffer_pre_bnc: pre_bnc_storage,
            cell_offsets_pre_bnc: pre_bnc_offsets,
            cell_storage_buffer: modern_storage,
            cell_offsets: modern_offsets,
            has_wide_offsets,
            ..Default::default()
        }
    }

    fn tile_source(rows: Vec<tst::TileRowInfo>, num_cells: u32, num_rows: u32) -> Vec<u8> {
        tst::Tile {
            max_column: 1,
            max_row: 1,
            num_cells,
            numrows: num_rows,
            row_infos: rows,
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn decode_tile_scalars(source: &[u8]) -> numbers_table_cell_storage_codec::TileSnapshot {
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        numbers_table_cell_storage_codec::decode_tile_with_visitor(source, options, &mut ())
            .unwrap_or_else(|error| panic!("tile projection failed: {error:?}"))
            .0
    }

    fn run_tile_visitor(
        source: &[u8],
        column_count: usize,
        limits: SemanticLimits,
    ) -> super::Result<(
        Result<
            numbers_table_cell_storage_codec::DecodeReport,
            numbers_table_cell_storage_codec::DecodeError,
        >,
        Table,
        ProjectionBudget,
        Option<Error>,
        usize,
    )> {
        let mut table = Table::with_dimensions("tile", 2, column_count)?;
        let strings: Box<[(u32, String)]> = Box::default();
        let formulas: Box<[(u32, tsce::FormulaArchive)]> = Box::default();
        let formula_errors: Box<[(u32, String)]> = Box::default();
        let rich_text: Box<[(u32, String)]> = Box::default();
        let comments: Box<[(u32, Comment)]> = Box::default();
        let formula_references = FormulaReferenceMaps::default();
        let cell_tables = CellTables {
            strings: &strings,
            formulas: &formulas,
            formula_errors: &formula_errors,
            rich_text: &rich_text,
            comments: Some(&comments),
            formula_references: &formula_references,
        };
        let mut cell_budget = CellBudget::new();
        let mut projection_budget = ProjectionBudget::new(limits);
        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        let (materialized_cells, semantic_error, report) = {
            let mut visitor = TileRowVisitor {
                row_origin: 0,
                tile_size: 2,
                row_count: table.row_count(),
                column_count: table.column_count(),
                budget: &mut cell_budget,
                cell_tables: &cell_tables,
                projection_budget: &mut projection_budget,
                table: &mut table,
                materialized_cells: 0,
                semantic_error: None,
            };
            let decode_result = numbers_table_cell_storage_codec::decode_tile_with_visitor(
                source,
                options,
                &mut visitor,
            )
            .map(|(_, report)| report);
            (
                visitor.materialized_cells,
                visitor.semantic_error.take(),
                decode_result,
            )
        };
        Ok((
            report,
            table,
            projection_budget,
            semantic_error,
            materialized_cells,
        ))
    }

    fn parse_projected_rows(source: &[u8], column_count: usize) -> super::Result<Table> {
        let (report, table, mut projection_budget, semantic_error, materialized_cells) =
            run_tile_visitor(source, column_count, SemanticLimits::default())?;
        let report = report
            .map_err(|error| Error::InvalidFormat(format!("tile projection failed: {error:?}")))?;
        projection_budget.charge_decode_report(report)?;
        projection_budget.charge_materialized_cells(materialized_cells)?;
        if let Some(error) = semantic_error {
            return Err(error);
        }
        Ok(table)
    }

    #[test]
    fn table_info_wire_is_a_legacy_classification_miss() -> super::Result<()> {
        let table_info = tst::TableInfoArchive::default().encode_to_vec();
        assert!(!has_legacy_table_model_wire_shape(&table_info)?);
        assert!(
            decode_legacy_table_candidate(
                &table_info,
                || panic!("table-info false positive invoked admission"),
                |_model| -> super::Result<()> {
                    panic!("table-info false positive reached model extraction")
                },
            )?
            .is_none()
        );
        Ok(())
    }

    #[test]
    fn legacy_table_info_false_positive_is_ignored_when_table_budget_is_full() {
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: tst::TableInfoArchive::default().encode_to_vec(),
                }],
            )
            .expect("table-info archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());

        let tables = extractor
            .extract_all_semantic_tables(0)
            .expect("table-info false positive is not a table");
        assert!(tables.is_empty());
        let budget = *extractor.projection_budget.borrow();
        assert_eq!(budget.references, 0);
        assert_eq!(budget.payload_fields, 0);
        assert_eq!(budget.payload_work, 0);
        assert_eq!(budget.staging_text_bytes, 0);
        assert_eq!(budget.materialized_cells, 0);
        assert_eq!(budget.output_text_bytes, 0);
        assert_eq!(budget.formula_render_work, 0);
    }

    #[test]
    fn legacy_candidate_admission_happens_before_decode_or_parse() {
        let encoded = legacy_model("over-budget", 90).encode_to_vec();
        let mut admission_called = false;
        let mut parse_called = false;
        let result = decode_legacy_table_candidate(
            &encoded,
            || {
                admission_called = true;
                Err(super::table_limit_error(1, 0))
            },
            |_model| {
                parse_called = true;
                Ok(())
            },
        );

        assert!(matches!(
            result,
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Tables,
                observed: 1,
                maximum: 0,
                path: SemanticPath::StructuredTables,
            })
        ));
        assert!(admission_called);
        assert!(!parse_called);
    }

    #[test]
    fn legacy_model_at_full_table_budget_leaves_projection_budget_unchanged() {
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: legacy_model("over-budget", 90).encode_to_vec(),
                }],
            )
            .expect("legacy table archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());
        let before = *extractor.projection_budget.borrow();

        let error = extractor
            .extract_all_semantic_tables(0)
            .expect_err("over-budget legacy model");
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::Tables,
                observed: 1,
                maximum: 0,
                path: SemanticPath::StructuredTables,
            }
        ));
        let after = *extractor.projection_budget.borrow();
        assert_eq!(before.references, after.references);
        assert_eq!(before.payload_fields, after.payload_fields);
        assert_eq!(before.payload_work, after.payload_work);
        assert_eq!(before.staging_text_bytes, after.staging_text_bytes);
        assert_eq!(before.materialized_cells, after.materialized_cells);
        assert_eq!(before.output_text_bytes, after.output_text_bytes);
        assert_eq!(before.formula_render_work, after.formula_render_work);
    }

    #[test]
    fn malformed_schema_shaped_legacy_candidate_wins_when_table_budget_is_available() {
        let malformed = [0x22, 0x01, 0xff, 0x30, 0x01, 0x38, 0x01, 0x42, 0x01, b'x'];
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )
            .expect("document archive"),
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: malformed.to_vec(),
                }],
            )
            .expect("malformed legacy archive"),
        ])
        .expect("compatibility package");
        let components = Components::from_bytes(&bytes, Limits::default()).expect("components");
        let index =
            Index::from_components(&components, SemanticLimits::MAX_OBJECTS).expect("object index");
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default());

        let error = extractor
            .extract_all_semantic_tables(1)
            .expect_err("malformed schema-shaped legacy payload");
        assert!(matches!(
            error,
            Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            }
        ));
    }

    #[test]
    fn admitted_legacy_candidate_preserves_common_allocation_error() -> super::Result<()> {
        let encoded = legacy_model("model", 90).encode_to_vec();
        assert!(has_legacy_table_model_wire_shape(&encoded)?);
        let result: super::Result<Option<()>> = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| {
                Err(Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers retained semantic text",
                    amount: 5,
                }))
            },
        );
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("allocation error was swallowed".to_owned()))?;
        assert!(matches!(
            &error,
            Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers retained semantic text",
                amount: 5,
            })
        ));
        assert_eq!(
            error.to_string(),
            "IWA wire allocation failed for Numbers retained semantic text: 5"
        );
        Ok(())
    }

    #[test]
    fn table_cell_codec_allocation_limit_maps_to_typed_common_error() {
        let error = map_table_cell_decode_limit_with_reference_offset(
            numbers_table_cell_storage_codec::DecodeLimit::Allocation { requested: 17 },
            0,
        );
        assert!(matches!(
            error,
            Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table storage projection",
                amount: 17,
            })
        ));
    }

    #[test]
    fn admitted_legacy_decode_failure_reports_exact_content_free_path() -> super::Result<()> {
        // Required model fields 4, 6, 7, and 8 are present with their schema
        // wire types, but the nested DataStore payload is malformed.
        let encoded = [0x22, 0x01, 0xff, 0x30, 0x01, 0x38, 0x01, 0x42, 0x01, b'x'];
        assert!(has_legacy_table_model_wire_shape(&encoded)?);
        let result = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| -> super::Result<()> {
                panic!("malformed admitted model reached semantic extraction")
            },
        );
        let error = result.err().ok_or_else(|| {
            Error::InvalidFormat("admitted model decode error was swallowed".to_owned())
        })?;
        assert!(matches!(
            &error,
            Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(
            error.to_string(),
            "malformed Numbers payload at structured tables"
        );
        Ok(())
    }

    #[test]
    fn admitted_legacy_duplicate_required_field_fails_closed() -> super::Result<()> {
        let mut encoded = legacy_model("model", 90).encode_to_vec();
        encoded.extend_from_slice(&[0x30, 0x01]);
        let result = decode_legacy_table_candidate(
            &encoded,
            || Ok(()),
            |_model| -> super::Result<()> {
                panic!("ambiguous admitted model reached semantic extraction")
            },
        );
        assert!(matches!(
            result,
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
        ));
        Ok(())
    }

    #[test]
    fn admitted_legacy_malformed_trailing_wire_fails_closed() -> super::Result<()> {
        let mut encoded = legacy_model("model", 90).encode_to_vec();
        encoded.push(0xff);
        assert!(matches!(
            decode_legacy_table_candidate(
                &encoded,
                || Ok(()),
                |_model| -> super::Result<()> {
                    panic!("malformed admitted model reached semantic extraction")
                },
            ),
            Err(Error::MalformedPayload {
                path: SemanticPath::StructuredTables,
            })
        ));
        Ok(())
    }

    #[test]
    fn admitted_legacy_name_limit_reports_exact_content_free_error() -> super::Result<()> {
        let sidecars = archive_object(
            90,
            [
                tst::table_data_list::ListType::String,
                tst::table_data_list::ListType::Formula,
            ]
            .into_iter()
            .map(|list_type| RawMessage {
                type_: 6_005,
                data: tst::TableDataList {
                    list_type: list_type as i32,
                    next_list_id: 1,
                    ..Default::default()
                }
                .encode_to_vec(),
            })
            .collect(),
        )?;
        let bytes = compatibility_package(vec![
            archive_object(
                1,
                vec![RawMessage {
                    type_: 1,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                }],
            )?,
            sidecars,
            archive_object(
                10,
                vec![RawMessage {
                    type_: 6_000,
                    data: legacy_model("\u{e9}", 90).encode_to_vec(),
                }],
            )?,
        ])?;
        let semantic = SemanticLimits::default()
            .with_projection_limits(SemanticLimits::MAX_MATERIALIZED_CELLS, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let result = compatibility_tables_from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), semantic),
        );
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidFormat("text limit error was swallowed".to_owned()))?;
        assert!(matches!(
            &error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 2,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            }
        ));
        assert_eq!(
            error.to_string(),
            "Numbers semantic output text bytes limit exceeded at structured tables: observed 2, maximum 1"
        );
        Ok(())
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
    fn tile_rows_use_pre_bnc_buffers_when_modern_buffers_are_absent() -> super::Result<()> {
        let source = tile_source(
            vec![tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None)],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_use_modern_buffers_only_when_both_are_present() -> super::Result<()> {
        let modern = BncCell::minimal().encode();
        let source = tile_source(
            vec![tile_row(
                0,
                1,
                vec![0xff],
                vec![0xff],
                Some(modern),
                Some(vec![0, 0]),
                None,
            )],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_fall_back_to_both_pre_bnc_buffers_for_partial_modern_storage() -> super::Result<()>
    {
        let source = tile_source(
            vec![tile_row(
                0,
                1,
                vec![0; 8],
                vec![0, 0],
                Some(vec![0xff]),
                None,
                None,
            )],
            1,
            1,
        );
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_rows_honor_wide_offset_units() -> super::Result<()> {
        let cell = BncCell::minimal().encode();
        let mut storage = cell.clone();
        storage.extend_from_slice(&cell);
        let source = tile_source(
            vec![tile_row(
                0,
                2,
                Vec::new(),
                vec![0xff, 0xff, 0xff, 0xff],
                Some(storage),
                Some(vec![0, 0, 3, 0]),
                Some(true),
            )],
            2,
            1,
        );
        let table = parse_projected_rows(&source, 2)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        assert_eq!(table.get_cell(0, 1), Some(&CellValue::Empty));
        Ok(())
    }

    #[test]
    fn tile_row_projection_is_atomic_when_a_later_row_is_malformed() -> super::Result<()> {
        let valid = tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        let mut malformed =
            tile_row(1, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        malformed.pop();
        let mut source = Vec::new();
        append_varint_field(&mut source, 1, 1)?;
        append_varint_field(&mut source, 2, 1)?;
        append_varint_field(&mut source, 3, 2)?;
        append_varint_field(&mut source, 4, 2)?;
        append_length_delimited_field(&mut source, 5, &valid)?;
        append_length_delimited_field(&mut source, 5, &malformed)?;

        let options = numbers_table_cell_storage_codec::DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            16,
            usize::MAX,
            usize::MAX,
        );
        assert!(
            numbers_table_cell_storage_codec::decode_tile_with_visitor(&source, options, &mut ())
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn later_malformed_wire_overrides_prior_semantic_tile_error() -> super::Result<()> {
        let mut invalid_cell = vec![0; 8];
        invalid_cell[2] = 99;
        let first = tile_row(0, 1, invalid_cell, vec![0, 0], None, None, None).encode_to_vec();
        let mut malformed =
            tile_row(1, 1, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        malformed.pop();
        let mut source = Vec::new();
        append_varint_field(&mut source, 1, 1)?;
        append_varint_field(&mut source, 2, 1)?;
        append_varint_field(&mut source, 3, 2)?;
        append_varint_field(&mut source, 4, 2)?;
        append_length_delimited_field(&mut source, 5, &first)?;
        append_length_delimited_field(&mut source, 5, &malformed)?;

        let (decode_result, table, _budget, semantic_error, materialized_cells) =
            run_tile_visitor(&source, 1, SemanticLimits::default())?;
        assert!(decode_result.is_err());
        assert!(
            matches!(semantic_error, Some(Error::ParseError(message)) if message.contains("Unsupported Numbers pre-BNC cell type 99"))
        );
        assert_eq!(materialized_cells, 1);
        assert_eq!(table.cell_count(), 0);
        Ok(())
    }

    #[test]
    fn aggregate_cell_limit_precedes_retained_semantic_error_without_row_growth()
    -> super::Result<()> {
        let mut invalid_cell = vec![0; 8];
        invalid_cell[2] = 99;
        let first = tile_row(0, 1, invalid_cell, vec![0, 0], None, None, None).encode_to_vec();
        let second = tile_row(1, 3, vec![0; 8], vec![0, 0], None, None, None).encode_to_vec();
        let source = tile_source(
            vec![
                tst::TileRowInfo::decode(first.as_slice()).map_err(Error::protobuf)?,
                tst::TileRowInfo::decode(second.as_slice()).map_err(Error::protobuf)?,
            ],
            4,
            2,
        );
        let limits = SemanticLimits::default()
            .with_projection_limits(1, SemanticLimits::MAX_OUTPUT_TEXT_BYTES)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let (decode_result, table, mut budget, semantic_error, materialized_cells) =
            run_tile_visitor(&source, 1, limits)?;
        let report = decode_result.map_err(|error| Error::InvalidFormat(format!("{error:?}")))?;
        assert!(matches!(semantic_error, Some(Error::ParseError(_))));
        assert_eq!(materialized_cells, 4);
        assert_eq!(table.cell_count(), 0);
        budget.charge_decode_report(report)?;
        let error = budget
            .charge_materialized_cells(materialized_cells)
            .err()
            .ok_or_else(|| {
                Error::InvalidFormat("aggregate cell limit was not enforced".to_owned())
            })?;
        assert!(matches!(
            error,
            Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 4,
                maximum: 1,
                path: SemanticPath::StructuredTables,
            }
        ));
        Ok(())
    }

    #[test]
    fn tile_rows_ignore_declared_counts_when_records_are_valid() -> super::Result<()> {
        let source = tile_source(
            vec![tile_row(0, 1, vec![0; 8], vec![0, 0], None, None, None)],
            99,
            77,
        );
        let projected = decode_tile_scalars(&source);
        assert_eq!(projected.num_cells(), 99);
        assert_eq!(projected.num_rows(), 77);
        let table = parse_projected_rows(&source, 1)?;
        assert_eq!(table.get_cell(0, 0), Some(&CellValue::Empty));
        Ok(())
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
            comments: Some(&comments),
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
    fn rooted_table_call_graph_enforces_one_cumulative_reference_budget() -> super::Result<()> {
        fn extract(max_references: usize) -> super::Result<usize> {
            let package = Package::open(native_fixture())?;
            let limits = SemanticLimits::new(
                SemanticLimits::MAX_OBJECTS,
                SemanticLimits::MAX_SHEETS,
                SemanticLimits::MAX_TABLES,
                max_references,
            )
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
            let extractor =
                TableDataExtractor::new(&package.state.components, &package.state.index, limits)
                    .without_comments();

            // The fixture has one rooted sheet, drawable, and TableInfo edge.
            // Production charges these before entering the table-model and
            // sidecar call graph; mirror that exact prefix here.
            extractor.charge_references(3)?;
            let entry = package
                .state
                .index
                .iter_entries_by_type(super::TABLE_MODEL_MESSAGE_TYPE)
                .next()
                .ok_or_else(|| Error::InvalidFormat("native table model is missing".to_owned()))?;
            let object = package
                .state
                .index
                .resolve_ref(&package.state.components, entry.id())?
                .ok_or_else(|| Error::InvalidFormat("native table model is missing".to_owned()))?;
            extractor.extract_reachable_table_from_object(
                &object,
                SemanticPath::Drawable { sheet: 0, index: 0 },
            )?;
            Ok(extractor.projection_budget.borrow().references)
        }

        let exact = extract(SemanticLimits::MAX_REFERENCES)?;
        assert!(
            exact > 3,
            "table/list/rich/tile projection must charge references"
        );
        assert_eq!(extract(exact)?, exact);
        let tight = extract(exact - 1);
        assert!(
            matches!(
                tight,
                Err(Error::SemanticLimit {
                    kind: SemanticLimitKind::References,
                    observed,
                    maximum,
                    ..
                }) if observed == exact && maximum == exact - 1
            ),
            "unexpected tight cumulative result: {tight:?}; exact={exact}"
        );
        Ok(())
    }

    #[test]
    fn strict_selected_list_rejects_conflicting_entry_payloads_in_both_modes() {
        let mut conflicting = string_list_entry(1, "selected");
        conflicting.formula = Some(formula(Vec::new()));
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            vec![conflicting.clone()],
                            Vec::new(),
                        )],
                    )
                    .expect("list object"),
                ],
                document_projection,
                |extractor| load_string_list(extractor, 90).map(|_| ()),
            );
            assert!(
                matches!(&result, Err(Error::InvalidFormat(message)) if message.contains("no selected payload")),
                "unexpected selected-list result: {result:?}"
            );
        }
    }

    #[test]
    fn wrong_list_candidates_charge_work_not_references_or_text_in_both_modes() {
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::Formula,
                            Vec::new(),
                            Vec::new(),
                        )],
                    )
                    .expect("list object"),
                ],
                document_projection,
                |extractor| load_string_list(extractor, 90),
            );
            let error = result.expect_err("wrong list must be rejected");
            assert!(matches!(error, Error::InvalidFormat(_)));
        }
    }

    #[test]
    fn wrong_list_candidate_does_not_consume_exhausted_admission_budgets() -> super::Result<()> {
        let mut wrong = list_message(
            tst::table_data_list::ListType::Formula,
            vec![tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                reference: Some(reference(99)),
                ..Default::default()
            }],
            Vec::new(),
        );
        let mut deep_prefix = nested_unknown_group_prefix(60);
        deep_prefix.extend_from_slice(&wrong.data);
        wrong.data = deep_prefix;
        let selected = list_message(
            tst::table_data_list::ListType::String,
            Vec::new(),
            Vec::new(),
        );

        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![archive_object(90, vec![wrong.clone(), selected.clone()])?],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    budget.references = budget.max_references;
                    budget.staging_text_bytes = DEFAULT_MAX_TEXT_BYTES;
                    let table = load_string_list_with_budget(extractor, 90, &mut budget)?;
                    assert!(table.is_empty());
                    assert_eq!(budget.references, budget.max_references);
                    assert_eq!(budget.staging_text_bytes, DEFAULT_MAX_TEXT_BYTES);
                    assert!(budget.payload_work > 0);
                    Ok(())
                },
            );
            result?;
        }
        Ok(())
    }

    #[test]
    fn duplicate_roots_and_segments_do_not_admit_reference_or_text_budget() -> super::Result<()> {
        let duplicate_entry = tst::table_data_list::ListEntry {
            key: 2,
            refcount: 1,
            string: Some("duplicate".to_owned()),
            reference: Some(reference(99)),
            ..Default::default()
        };
        for document_projection in [false, true] {
            let result = with_list_extractor(
                vec![archive_object(
                    90,
                    vec![
                        list_message(
                            tst::table_data_list::ListType::String,
                            vec![string_list_entry(1, "admitted")],
                            Vec::new(),
                        ),
                        list_message(
                            tst::table_data_list::ListType::String,
                            vec![duplicate_entry.clone()],
                            Vec::new(),
                        ),
                    ],
                )?],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    let result = load_string_list_with_budget(extractor, 90, &mut budget);
                    assert!(matches!(result, Err(Error::InvalidFormat(_))));
                    assert_eq!(budget.references, 0);
                    assert_eq!(budget.staging_text_bytes, "admitted".len());
                    Ok(())
                },
            );
            result?;

            let result = with_list_extractor(
                vec![
                    archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            Vec::new(),
                            vec![91],
                        )],
                    )?,
                    archive_object(
                        91,
                        vec![
                            segment_message(
                                tst::table_data_list::ListType::String,
                                1,
                                2,
                                vec![string_list_entry(1, "admitted")],
                            ),
                            segment_message(
                                tst::table_data_list::ListType::String,
                                1,
                                2,
                                vec![duplicate_entry.clone()],
                            ),
                        ],
                    )?,
                ],
                document_projection,
                |extractor| {
                    let mut budget = ProjectionBudget::new(SemanticLimits::default());
                    let result = load_string_list_with_budget(extractor, 90, &mut budget);
                    assert!(matches!(result, Err(Error::InvalidFormat(_))));
                    assert_eq!(budget.references, 1);
                    assert_eq!(budget.staging_text_bytes, "admitted".len());
                    Ok(())
                },
            );
            result?;
        }
        Ok(())
    }

    #[test]
    fn structural_entry_errors_win_after_a_semantic_conversion_error() -> super::Result<()> {
        for document_projection in [false, true] {
            let result =
                with_list_extractor(
                    vec![archive_object(
                        90,
                        vec![list_message(
                            tst::table_data_list::ListType::String,
                            vec![
                                string_list_entry(1, "first"),
                                string_list_entry(2, "later"),
                                string_list_entry(2, "later-duplicate"),
                            ],
                            Vec::new(),
                        )],
                    )?],
                    document_projection,
                    |extractor| {
                        let mut budget = ProjectionBudget::new(SemanticLimits::default());
                        let mut calls = 0usize;
                        let mut converter =
                        |entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
                         _budget: &mut ProjectionBudget| {
                            calls += 1;
                            if entry.key() == 1 {
                                return Err(Error::InvalidFormat(
                                    "synthetic semantic conversion failure".to_owned(),
                                ));
                            }
                            Ok(entry.string_value().unwrap_or_default().to_owned())
                        };
                        let result = extractor.load_table_data_list_entries(
                            90,
                            tst::table_data_list::ListType::String,
                            &mut budget,
                            &mut converter,
                        );
                        assert!(matches!(
                            result,
                            Err(Error::InvalidFormat(message))
                                if message.contains("duplicate keys")
                        ));
                        assert_eq!(calls, 1);
                        Ok(())
                    },
                );
            result?;
        }
        Ok(())
    }

    #[test]
    fn rejected_formula_candidate_keeps_wire_work_monotonic() {
        let mut budget = ProjectionBudget::new(SemanticLimits::default());
        let mut candidate = budget;
        candidate.formula_wire_bytes = 17;
        budget.commit_attempt(candidate, false);
        assert_eq!(budget.formula_wire_bytes, 17);
    }

    #[test]
    fn list_root_and_segment_order_duplicates_and_ranges_are_strict() -> super::Result<()> {
        let ordered = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        vec![string_list_entry(2, "root")],
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        1,
                        vec![string_list_entry(1, "segment")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|(table, _)| table),
        )?;
        assert_eq!(
            ordered.as_ref(),
            &[(1, "segment".to_owned()), (2, "root".to_owned())]
        );

        let duplicate_segment_key = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        2,
                        vec![string_list_entry(1, "a"), string_list_entry(1, "b")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(
            duplicate_segment_key,
            Err(Error::InvalidFormat(_))
        ));

        let outside_range = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        10,
                        1,
                        vec![string_list_entry(1, "outside")],
                    )],
                )?,
            ],
            false,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(outside_range, Err(Error::InvalidFormat(_))));

        let duplicate_segment_id = with_list_extractor(
            vec![
                archive_object(
                    90,
                    vec![list_message(
                        tst::table_data_list::ListType::String,
                        Vec::new(),
                        vec![91, 91],
                    )],
                )?,
                archive_object(
                    91,
                    vec![segment_message(
                        tst::table_data_list::ListType::String,
                        1,
                        1,
                        vec![string_list_entry(1, "segment")],
                    )],
                )?,
            ],
            true,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(duplicate_segment_id, Err(Error::InvalidFormat(_))));

        let duplicate_root = with_list_extractor(
            vec![archive_object(
                90,
                vec![list_message(
                    tst::table_data_list::ListType::String,
                    vec![string_list_entry(1, "a"), string_list_entry(1, "b")],
                    Vec::new(),
                )],
            )?],
            false,
            |extractor| load_string_list(extractor, 90).map(|_| ()),
        );
        assert!(matches!(duplicate_root, Err(Error::InvalidFormat(_))));

        for document_projection in [false, true] {
            let mut malformed = list_message(
                tst::table_data_list::ListType::String,
                vec![string_list_entry(1, "callback-before-wire-error")],
                Vec::new(),
            );
            malformed.data.push(0);
            let result = with_list_extractor(
                vec![archive_object(90, vec![malformed])?],
                document_projection,
                |extractor| load_string_list(extractor, 90).map(|_| ()),
            );
            assert!(matches!(result, Err(Error::InvalidFormat(_))));
        }
        Ok(())
    }

    #[test]
    fn wrong_list_candidate_rolls_back_retention_but_keeps_decode_work() -> super::Result<()> {
        let bytes = compatibility_package(vec![
            archive_object(
                90,
                vec![empty_list(tst::table_data_list::ListType::Formula)],
            )?,
            archive_object(
                91,
                vec![empty_list(tst::table_data_list::ListType::Formula)],
            )?,
            archive_object(
                92,
                vec![
                    empty_list(tst::table_data_list::ListType::String),
                    empty_list(tst::table_data_list::ListType::Formula),
                ],
            )?,
        ])?;
        let components = Components::from_bytes(&bytes, Limits::default())?;
        let index = Index::from_components(&components, SemanticLimits::MAX_OBJECTS)?;
        let extractor = TableDataExtractor::new(&components, &index, SemanticLimits::default())
            .without_comments();

        let invalid = tst::TableModelArchive {
            table_name: "rejected".to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(90),
                formula_table: reference(91),
                ..Default::default()
            },
            ..Default::default()
        };
        assert!(matches!(
            extractor.parse_table_model(invalid, false, None),
            Err(Error::InvalidFormat(_))
        ));
        let rejected = *extractor.projection_budget.borrow();
        assert_eq!(rejected.references, 0);
        assert_eq!(rejected.materialized_cells, 0);
        assert_eq!(rejected.output_text_bytes, 0);
        assert!(rejected.payload_work > 0);

        let valid = tst::TableModelArchive {
            table_name: "accepted".to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                string_table: reference(92),
                formula_table: reference(92),
                ..Default::default()
            },
            ..Default::default()
        };
        let table = extractor.parse_table_model(valid, false, None)?;
        assert_eq!(table.name(), "accepted");
        let published = *extractor.projection_budget.borrow();
        assert!(published.payload_work > rejected.payload_work);
        assert_eq!(published.output_text_bytes, "accepted".len());
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
        let mut depth_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut tight_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut false_positive_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut full_work_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut nested_work_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut malformed_budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut duplicate_budget = FormulaReferenceBudget::new(
            1,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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
        let mut budget = FormulaReferenceBudget::new(
            crate::MAX_REFERENCES,
            MAX_FORMULA_WORK,
            MAX_FORMULA_WIRE_BYTES,
            DEFAULT_MAX_TEXT_BYTES,
        );
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

    include!("extractor_rich_text_tests.rs");
}
