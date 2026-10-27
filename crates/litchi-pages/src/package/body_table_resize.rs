//! Resizing a body table to a new row and column count (record 0371).
//!
//! The port follows the migration host's `resize_attached_table_in_package`:
//! a shrink drops the stored rows past the new row count (refused when one
//! still holds cells, or when a kept row holds cells past the new column
//! count), every header bucket keeps only the headers inside the new
//! extent, the row/column UID map gains or drops identities at its end (a
//! new identity is one above the largest lower and upper halves, as the
//! host allocates it), the stroke sidecar records the new dimensions and the
//! model records the new counts, and the table's own formula owner moves its
//! table and body extents (record 0370e). Nothing else (header counts,
//! formulas, merges) is touched, as in the host.

use std::fmt;
use std::num::NonZeroU64;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::wire::{
    WireView, patch_varint_field, repeated_length_delimited_payloads, repeated_varint_values,
    rewrite_repeated_length_delimited_fields, rewrite_repeated_varint_fields,
};
use litchi_numbers_wire::table_topology::extents::{Extents, resize_extents};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, MessageRewrite, ObjectIndex,
};
use super::body_table_sort_execution::{located_message, reference, row_cells, varint};
use crate::selector::BodyTableSelector;

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const UID_MAP_MESSAGE_TYPES: [u32; 2] = [6_267, 6_200];
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const MODEL_DATA_STORE_FIELD: u32 = 4;
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
const MODEL_UID_MAP_FIELD: u32 = 46;
const MODEL_STROKE_SIDECAR_FIELD: u32 = 49;
const STORE_ROW_HEADERS_FIELD: u32 = 1;
const MODEL_HEADER_ROWS_FIELD: u32 = 9;
const MODEL_HEADER_COLUMNS_FIELD: u32 = 10;
const MODEL_FOOTER_ROWS_FIELD: u32 = 11;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `TSCE.FormulaOwnerDependenciesArchive.formula_owner`.
const OWNER_FORMULA_OWNER_FIELD: u32 = 11;
const STORE_COLUMN_HEADERS_FIELD: u32 = 2;
const STORE_TILES_FIELD: u32 = 3;
const HEADER_STORAGE_BUCKETS_FIELD: u32 = 2;
const TILE_STORAGE_TILES_FIELD: u32 = 1;
const TILE_STORAGE_SIZE_FIELD: u32 = 2;
const TILE_REFERENCE_ID_FIELD: u32 = 1;
const TILE_REFERENCE_TILE_FIELD: u32 = 2;
const TILE_NUMROWS_FIELD: u32 = 4;
const TILE_ROWS_FIELD: u32 = 5;
const ROW_INDEX_FIELD: u32 = 1;
const BUCKET_HEADERS_FIELD: u32 = 2;
const HEADER_INDEX_FIELD: u32 = 1;
const HEADER_CELLS_FIELD: u32 = 4;
const SIDECAR_COLUMN_COUNT_FIELD: u32 = 2;
const SIDECAR_ROW_COUNT_FIELD: u32 = 3;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;
const DEFAULT_TILE_SIZE: u64 = 256;
const MAX_TABLE_UIDS: usize = 1_100_000;

/// A finite resource charged by one resize.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payloads or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector resolved no body table.
    #[error("the Pages body has no such table")]
    TableNotFound,
    /// A table needs at least one row and one column, and at most
    /// 1,100,000 row and column identities together.
    #[error("the table size {rows}×{columns} is invalid")]
    InvalidSize {
        /// Requested rows.
        rows: usize,
        /// Requested columns.
        columns: usize,
    },
    /// A shrink would drop stored cells.
    #[error("the shrink would drop stored cells")]
    StoredCells,
    /// The table stores legacy (pre-BNC) rows.
    #[error("the table stores legacy rows")]
    LegacyRows,
    /// No size was staged.
    #[error("no table size is staged")]
    NoStagedOperation,
    /// A second size was staged on one edit.
    #[error("one table-size edit stages one size")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact table resizes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::InvalidSize { .. } => RefusalKind::InvalidValue,
            Self::StoredCells | Self::LegacyRows | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::NoStagedOperation | Self::OperationAlreadyStaged => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

/// Failure from a table resize.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The table graph is malformed.
    #[error("the selected Pages table is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages table resize {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages table resize")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the new size.
    #[error("the resized Pages table failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages table resize patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                },
                observed,
                maximum,
            },
        }
    }
}

fn wire(error: litchi_iwa_common::Error) -> Error {
    graph::map_wire_error(error).into()
}

/// One table size staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    table: usize,
    size: Option<(usize, usize)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("table", &self.table)
            .field("size", &self.size)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the table's new row and column count (header and footer rows
    /// included).
    ///
    /// # Errors
    ///
    /// Refuses an invalid size and a second size.
    pub fn set(&mut self, rows: usize, columns: usize) -> Result<&mut Self, Error> {
        if self.size.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        if rows == 0 || columns == 0 || rows.saturating_add(columns) > MAX_TABLE_UIDS {
            return Err(Refusal::InvalidSize { rows, columns }.into());
        }
        self.size = Some((rows, columns));
        Ok(self)
    }

    /// Validate and publish the staged size.
    pub fn commit(self) -> Result<Commit, Error> {
        let (rows, columns) = self.size.ok_or(Refusal::NoStagedOperation)?;
        let package = resize(self.source, self.table, rows, columns)?;
        let target = package.state.source.shared_source();
        let changed = !std::sync::Arc::ptr_eq(&target, &self.source.state.source.shared_source());
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics: Diagnostics { changed },
        })
    }
}

/// Exact-source reversible table-resize patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the exact artifact is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics of one resize.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the size changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// The validated immutable result of one resize.
#[must_use = "a resize commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Begin resizing a body table.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selector resolves no table.
    pub fn edit_body_table_size<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let catalog = self.body_tables().map_err(|_error| Error::InvalidSource)?;
        let table = match selector {
            BodyTableSelector::Position(position) => {
                (position.get() < catalog.len()).then_some(position.get())
            },
            BodyTableSelector::Name(name) => {
                let matches = (0..catalog.len())
                    .filter(|index| {
                        catalog
                            .get(*index)
                            .is_some_and(|entry| entry.name() == name)
                    })
                    .collect::<Vec<_>>();
                match matches.as_slice() {
                    [index] => Some(*index),
                    _ => None,
                }
            },
        }
        .ok_or(Refusal::TableNotFound)?;
        Ok(Edit {
            source: self,
            table,
            size: None,
        })
    }

    /// Replay an exact-source table-resize patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_table_size(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            graph::reopen(self, patch.artifacts.target())?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
            },
        })
    }
}

/// One axis of the UID map: sorted identities (raw `TSP.UUID` payloads),
/// index-for-uid and uid-for-index lists.
struct UidAxis {
    sorted_field: u32,
    index_field: u32,
    uid_field: u32,
}

const ROW_AXIS: UidAxis = UidAxis {
    sorted_field: 4,
    index_field: 5,
    uid_field: 6,
};
const COLUMN_AXIS: UidAxis = UidAxis {
    sorted_field: 1,
    index_field: 2,
    uid_field: 3,
};

/// The UID-map payload with one axis resized (the host's `resize_uid_axis`).
fn resize_axis(
    payload: &[u8],
    axis: &UidAxis,
    old_len: usize,
    new_len: usize,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let mut sorted = repeated_length_delimited_payloads(payload, axis.sorted_field)
        .map_err(wire)?
        .into_iter()
        .map(<[u8]>::to_vec)
        .collect::<Vec<_>>();
    let mut index_for_uid = repeated_varint_values(payload, axis.index_field).map_err(wire)?;
    let mut uid_for_index = repeated_varint_values(payload, axis.uid_field).map_err(wire)?;
    if sorted.len() != old_len || index_for_uid.len() != old_len || uid_for_index.len() != old_len {
        return Err(Error::InvalidSource);
    }
    if new_len < old_len {
        for logical in (new_len..old_len).rev() {
            let slot = uid_for_index.pop().ok_or(Error::InvalidSource)?;
            let slot = usize::try_from(slot).map_err(|_error| Error::InvalidSource)?;
            if slot >= sorted.len() || index_for_uid.get(slot).copied() != Some(logical as u64) {
                return Err(Error::InvalidSource);
            }
            sorted.remove(slot);
            index_for_uid.remove(slot);
            for value in &mut uid_for_index {
                if *value > slot as u64 {
                    *value -= 1;
                }
            }
        }
    } else {
        for logical in old_len..new_len {
            let (mut lower, mut upper) = (0_u64, 0_u64);
            for uuid in &sorted {
                let view = budget.parse(uuid, 3)?;
                lower = lower.max(varint(&view, UUID_LOWER_FIELD)?.unwrap_or(0));
                upper = upper.max(varint(&view, UUID_UPPER_FIELD)?.unwrap_or(0));
            }
            let lower = lower.checked_add(1).ok_or(Error::InvalidSource)?;
            let upper = upper.checked_add(1).ok_or(Error::InvalidSource)?;
            let mut uuid = Vec::new();
            litchi_iwa_common::wire::append_varint_field(&mut uuid, UUID_LOWER_FIELD, lower)
                .map_err(wire)?;
            litchi_iwa_common::wire::append_varint_field(&mut uuid, UUID_UPPER_FIELD, upper)
                .map_err(wire)?;
            sorted.push(uuid);
            index_for_uid.push(logical as u64);
            uid_for_index.push((sorted.len() - 1) as u64);
        }
    }
    let data = rewrite_repeated_length_delimited_fields(payload, axis.sorted_field, &sorted)
        .map_err(wire)?;
    let data =
        rewrite_repeated_varint_fields(&data, axis.index_field, &index_for_uid).map_err(wire)?;
    rewrite_repeated_varint_fields(&data, axis.uid_field, &uid_for_index).map_err(wire)
}

#[allow(clippy::too_many_lines)]
fn resize(source: &Package, table: usize, rows: usize, columns: usize) -> Result<Package, Error> {
    let selector = BodyTableSelector::index(table);
    let mut budget = Budget::new(source)?;
    let index = ObjectIndex::new(source, &mut budget)?;
    let (model_identifier, info_identifier) = {
        let mut lock_budget = super::table_lock::WireBudget::new(source.state.source.limits())
            .map_err(|_error| Error::InvalidSource)?;
        let target = source
            .resolve_body_table_with_budget(selector, &mut lock_budget)
            .map_err(|_error| Error::InvalidSource)?;
        (target.model_identifier, target.drawable_identifier)
    };
    let (model_component, model_message, model_type, model) =
        located_message(&index, model_identifier, &[TABLE_MODEL_MESSAGE_TYPE])?;
    let model_view = budget.parse(model, 1)?;
    let old_rows =
        usize::try_from(varint(&model_view, MODEL_ROWS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    let old_columns =
        usize::try_from(varint(&model_view, MODEL_COLUMNS_FIELD)?.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?;
    if (rows, columns) == (old_rows, old_columns) {
        return Ok(source.clone());
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let store = graph::unique_field(&model_view, MODEL_DATA_STORE_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let store_view = budget.parse(store, 2)?;
    let mut rewrites = Vec::new();

    // Tiles: trim stored rows past the new row count.
    let tiles = graph::unique_field(&store_view, STORE_TILES_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let tiles_view = budget.parse(tiles, 3)?;
    let tile_size =
        usize::try_from(varint(&tiles_view, TILE_STORAGE_SIZE_FIELD)?.unwrap_or(DEFAULT_TILE_SIZE))
            .ok()
            .filter(|size| *size != 0)
            .ok_or(Error::InvalidSource)?;
    for reference_payload in
        repeated_length_delimited_payloads(tiles, TILE_STORAGE_TILES_FIELD).map_err(wire)?
    {
        let view = budget.parse(reference_payload, 4)?;
        let key =
            usize::try_from(varint(&view, TILE_REFERENCE_ID_FIELD)?.ok_or(Error::InvalidSource)?)
                .map_err(|_error| Error::InvalidSource)?;
        let tile: NonZeroU64 = reference(&view, TILE_REFERENCE_TILE_FIELD, &mut budget)?
            .ok_or(Error::InvalidSource)?;
        let (component, message_index, message_type, payload) =
            located_message(&index, tile, &[TILE_MESSAGE_TYPE])?;
        let mut kept = Vec::new();
        let mut removed = false;
        for raw in repeated_length_delimited_payloads(payload, TILE_ROWS_FIELD).map_err(wire)? {
            let view = budget.parse(raw, 3)?;
            let local = varint(&view, ROW_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
            let global =
                key * tile_size + usize::try_from(local).map_err(|_error| Error::InvalidSource)?;
            let cells = row_cells(raw, &mut budget)?.ok_or(Refusal::LegacyRows)?;
            if global >= rows {
                if cells.iter().any(Option::is_some) {
                    return Err(Refusal::StoredCells.into());
                }
                removed = true;
                continue;
            }
            if cells.iter().skip(columns).any(Option::is_some) {
                return Err(Refusal::StoredCells.into());
            }
            kept.push((local, raw.to_vec()));
        }
        if removed {
            let numrows = kept.iter().map(|(local, _)| local + 1).max().unwrap_or(0);
            let data = patch_varint_field(payload, TILE_NUMROWS_FIELD, true, Some(numrows))
                .map_err(wire)?;
            let records = kept.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
            let data = rewrite_repeated_length_delimited_fields(&data, TILE_ROWS_FIELD, &records)
                .map_err(wire)?;
            rewrites.push(MessageRewrite {
                component,
                object: tile,
                message_index,
                message_type,
                data,
            });
        }
    }

    // Header buckets keep the headers inside the new extent.
    let row_headers = graph::unique_field(&store_view, STORE_ROW_HEADERS_FIELD, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let mut buckets = Vec::new();
    for bucket in repeated_length_delimited_payloads(row_headers, HEADER_STORAGE_BUCKETS_FIELD)
        .map_err(wire)?
    {
        buckets.push((graph::parse_reference(bucket, &mut budget)?, rows));
    }
    if let Some(column_bucket) = reference(&store_view, STORE_COLUMN_HEADERS_FIELD, &mut budget)? {
        buckets.push((column_bucket, columns));
    }
    for (bucket, limit) in buckets {
        let (component, message_index, message_type, payload) =
            located_message(&index, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let mut kept = Vec::new();
        for raw in
            repeated_length_delimited_payloads(payload, BUCKET_HEADERS_FIELD).map_err(wire)?
        {
            let view: WireView<'_> = budget.parse(raw, 3)?;
            let header = varint(&view, HEADER_INDEX_FIELD)?.ok_or(Error::InvalidSource)?;
            if usize::try_from(header).map_err(|_error| Error::InvalidSource)? < limit {
                // The host re-writes every kept header with an explicit
                // cell count (`rewrite_header_bucket_wire`).
                let cells = varint(&view, HEADER_CELLS_FIELD)?.unwrap_or(0);
                kept.push(
                    patch_varint_field(raw, HEADER_CELLS_FIELD, true, Some(cells)).map_err(wire)?,
                );
            }
        }
        let data = rewrite_repeated_length_delimited_fields(payload, BUCKET_HEADERS_FIELD, &kept)
            .map_err(wire)?;
        // The host rewrites every bucket, so its component is re-serialized
        // even when no header changes.
        rewrites.push(MessageRewrite {
            component,
            object: bucket,
            message_index,
            message_type,
            data,
        });
    }

    // The row/column UID map.
    if let Some(uid_map) = reference(&model_view, MODEL_UID_MAP_FIELD, &mut budget)? {
        let (component, message_index, message_type, payload) =
            located_message(&index, uid_map, &UID_MAP_MESSAGE_TYPES)?;
        let data = resize_axis(payload, &ROW_AXIS, old_rows, rows, &mut budget)?;
        let data = resize_axis(&data, &COLUMN_AXIS, old_columns, columns, &mut budget)?;
        rewrites.push(MessageRewrite {
            component,
            object: uid_map,
            message_index,
            message_type,
            data,
        });
    }

    // The stroke sidecar's dimensions.
    if let Some(sidecar) = reference(&model_view, MODEL_STROKE_SIDECAR_FIELD, &mut budget)? {
        let (component, message_index, message_type, payload) =
            located_message(&index, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
        let view = budget.parse(payload, 2)?;
        let data = patch_varint_field(
            payload,
            SIDECAR_COLUMN_COUNT_FIELD,
            varint(&view, SIDECAR_COLUMN_COUNT_FIELD)?.is_some(),
            Some(columns as u64),
        )
        .map_err(wire)?;
        let data = patch_varint_field(
            &data,
            SIDECAR_ROW_COUNT_FIELD,
            varint(&view, SIDECAR_ROW_COUNT_FIELD)?.is_some(),
            Some(rows as u64),
        )
        .map_err(wire)?;
        rewrites.push(MessageRewrite {
            component,
            object: sidecar,
            message_index,
            message_type,
            data,
        });
    }

    // The table's own formula owner: its table and body extents (record
    // 0370e fixed the host to move them).
    let count = |number: u32| -> Result<u32, Error> {
        u32::try_from(varint(&model_view, number)?.unwrap_or(0))
            .map_err(|_error| Error::InvalidSource)
    };
    let (header_rows, header_columns, footer_rows) = (
        count(MODEL_HEADER_ROWS_FIELD)?,
        count(MODEL_HEADER_COLUMNS_FIELD)?,
        count(MODEL_FOOTER_ROWS_FIELD)?,
    );
    let extents = |rows: usize, columns: usize| -> Result<Extents, Error> {
        Ok(Extents::of(
            u32::try_from(rows).map_err(|_error| Error::InvalidSource)?,
            u32::try_from(columns).map_err(|_error| Error::InvalidSource)?,
            header_rows,
            header_columns,
            footer_rows,
        ))
    };
    let (before, after) = (extents(old_rows, old_columns)?, extents(rows, columns)?);
    for (component_index, archive) in source.state.source.components().iter().enumerate() {
        for object in &archive.archive().objects {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
                    continue;
                }
                budget.work(message.data.len())?;
                let view = budget.parse(&message.data, 1)?;
                if reference(&view, OWNER_FORMULA_OWNER_FIELD, &mut budget)?
                    != Some(info_identifier)
                {
                    continue;
                }
                let Some(data) = resize_extents(&message.data, before, after)
                    .map_err(|_error| Error::InvalidSource)?
                else {
                    continue;
                };
                let object = object
                    .archive_info
                    .identifier
                    .and_then(NonZeroU64::new)
                    .ok_or(Error::InvalidSource)?;
                rewrites.push(MessageRewrite {
                    component: index.component_name(component_index)?,
                    object,
                    message_index,
                    message_type: message.type_,
                    data,
                });
            }
        }
    }

    // The model's counts.
    let data =
        patch_varint_field(model, MODEL_ROWS_FIELD, true, Some(rows as u64)).map_err(wire)?;
    let data =
        patch_varint_field(&data, MODEL_COLUMNS_FIELD, true, Some(columns as u64)).map_err(wire)?;
    rewrites.push(MessageRewrite {
        component: model_component,
        object: model_identifier,
        message_index: model_message,
        message_type: model_type,
        data,
    });
    let candidate = graph::publish(source, rewrites)?;
    let reopened = candidate
        .body_tables()
        .map_err(|_error| Error::Verification)?;
    let _ = reopened;
    Ok(candidate)
}
