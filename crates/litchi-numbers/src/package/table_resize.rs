//! Implementation of the public [`crate::table::resize`] transaction:
//! resizing a table at its trailing edges (record 0370e).
//!
//! Growth adds blank trailing rows and columns; shrinking is accepted only
//! when the removed trailing region holds no stored cell. The tiles lose
//! the removed rows' records, the header buckets the removed rows' and
//! columns' headers, the row/column UID map its trailing identities (new
//! ones are allocated above every existing one), the stroke sidecar takes
//! the new size, and the table's own formula owner moves its table and body
//! extents. This is the host editor's `resize_table`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use thiserror::Error as ThisError;

use super::Package;
use super::structure::{self, Fault, Published, Working, metadata::Transition};
use super::table_topology::wire::{
    rebuild, repeated_bytes, single_bytes, single_reference, single_varint, varint_field,
};
use super::table_topology::{self as topology, TILE_MESSAGE_TYPE, storage};
use crate::selector::{SheetSelector, TableSelector};

structure::structural_family!("table-resize");

const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const UID_MAP_MESSAGE_TYPES: [u32; 2] = [6_200, 6_267];
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
/// `TableModelArchive.number_of_rows` / `number_of_columns`.
const MODEL_ROWS_FIELD: u32 = 6;
const MODEL_COLUMNS_FIELD: u32 = 7;
/// `TableModelArchive.base_column_row_uids`.
const MODEL_UID_MAP_FIELD: u32 = 46;
/// `TableModelArchive.stroke_sidecar`.
const MODEL_STROKE_FIELD: u32 = 49;
/// The most row and column identities one table may need.
const MAX_TABLE_UIDS: u64 = 1_100_000;

/// A typed, content-redacted resize failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested size.
    #[error("table resize candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table resize patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// Commit was requested before staging a size.
    #[error("no table size is staged")]
    NoStagedOperation,
    /// The size is zero, leaves no body row or column, or needs more row
    /// and column identities than a table may hold.
    #[error("the table cannot take this size")]
    InvalidSize,
    /// The removed trailing region holds a stored cell.
    #[error("the removed rows or columns hold stored cells")]
    StoredCells,
    /// The table's storage uses a layout this route does not rewrite.
    #[error("source representation does not support resizing this table")]
    UnsupportedSource,
    /// A surviving object depends on what the resize would change.
    #[error("a surviving object depends on the table's size")]
    UnsupportedDependency,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::InvalidSize | Self::StoredCells => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        use topology::{Error as E, Refusal as R};
        match error {
            E::Refused(R::SheetNotFound) => Self::Refused(Refusal::SheetNotFound),
            E::Refused(R::TableNotFound) => Self::Refused(Refusal::TableNotFound),
            E::Refused(_) => Self::Refused(Refusal::UnsupportedSource),
            E::Allocation { amount } => Self::Allocation { amount },
            E::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

fn storage_fault(error: litchi_numbers_wire::table_topology::Error) -> Error {
    match Fault::from(error) {
        Fault::Unsupported => Error::Refused(Refusal::UnsupportedSource),
        fault => fault.into(),
    }
}

/// One immutable edit that resizes one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    size: Option<(u32, u32)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("size", &self.size)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `rows` × `columns` (header and footer rows and header columns
    /// included) as the table's size.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidSize`] for a zero size or one needing more row and
    /// column identities than a table may hold.
    pub fn resize(mut self, rows: u32, columns: u32) -> Result<Self, Error> {
        if rows == 0 || columns == 0 || u64::from(rows) + u64::from(columns) > MAX_TABLE_UIDS {
            return Err(Error::Refused(Refusal::InvalidSize));
        }
        self.size = Some((rows, columns));
        Ok(self)
    }

    /// Validate and atomically publish the staged size; the current size
    /// is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::StoredCells`],
    /// [`Refusal::InvalidSize`] for a size leaving no body row or column,
    /// source, resource and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let (rows, columns) = self
            .size
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let Some(published) = resize(self.source, self.sheet, self.table, rows, columns)? else {
            return Commit::unchanged(self.source);
        };
        let table = published.package.workbook_ref().sheets()[self.sheet]
            .tables()
            .nth(self.table)
            .ok_or(Error::Verification)?;
        if (table.row_count(), table.column_count()) != (rows, columns) {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

fn replace(
    working: &mut Working<'_>,
    object: u64,
    index: usize,
    data: Vec<u8>,
) -> Result<(), Error> {
    super::table_cell_text_style::replace_message(working, object, index, data)?;
    Ok(())
}

#[allow(
    clippy::too_many_lines,
    reason = "one resize: tiles, headers, UID map, strokes, owner extents and model"
)]
fn resize(
    source: &Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
) -> Result<Option<Published>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let target = topology::locate(&working, sheet, table)?;
    let (model_index, model) = topology::message(&working, target.model, &[6_001])?;
    let model = model.to_vec();
    let before = topology::geometry(&model)?;
    if (before.rows, before.columns) == (rows, columns) {
        return Ok(None);
    }
    if rows <= before.header_rows + before.footer_rows || columns <= before.header_columns {
        return Err(Error::Refused(Refusal::InvalidSize));
    }
    let store = single_bytes(&model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();

    // Tiles: removed rows go; no stored cell may lie outside the new size.
    let (tiles, size) = topology::tiles(&store)?;
    for (key, tile) in &tiles {
        let (index, data) = topology::message(&working, *tile, &[TILE_MESSAGE_TYPE])?;
        let base = u64::from(*key) * size;
        let mut removed = false;
        let mut maximum: Option<u64> = None;
        let mut fault = None;
        let rewritten = rebuild(data, |field| {
            if field.number != 5 {
                return Ok(None);
            }
            let local = single_varint(field.payload, 1)?.ok_or(Fault::Invalid)?;
            if single_bytes(field.payload, 6)?.is_none() {
                fault.get_or_insert(Refusal::UnsupportedSource);
                return Ok(None);
            }
            let cells = storage::bnc_cells(field.payload).map_err(Fault::from)?;
            if base + local >= u64::from(rows) {
                if !cells.is_empty() {
                    fault.get_or_insert(Refusal::StoredCells);
                    return Ok(None);
                }
                removed = true;
                return Ok(Some(Vec::new()));
            }
            if cells.iter().any(|(column, _)| *column >= columns) {
                fault.get_or_insert(Refusal::StoredCells);
            }
            maximum = maximum.max(Some(local));
            Ok(None)
        })?;
        if let Some(refusal) = fault {
            return Err(Error::Refused(refusal));
        }
        if removed {
            // The host editor counts the tile's rows by extent here.
            let numrows = maximum.map_or(0, |maximum| maximum + 1);
            let rewritten = rebuild(&rewritten, |field| {
                if field.number == 4 {
                    return varint_field(4, numrows).map(Some);
                }
                Ok(None)
            })?;
            replace(&mut working, *tile, index, rewritten)?;
        }
    }

    // Header buckets keep the headers inside the new size.
    let headers = single_bytes(&store, 1)?.ok_or(Error::InvalidSource)?;
    let mut buckets: Vec<(u64, u32)> = Vec::new();
    for reference in repeated_bytes(headers, 2)? {
        buckets.push((
            single_varint(reference, 1)?.ok_or(Error::InvalidSource)?,
            rows,
        ));
    }
    if let Some(column_bucket) = single_reference(&store, 2)? {
        buckets.push((column_bucket, columns));
    }
    for (bucket, limit) in buckets {
        if bucket == 0 {
            continue;
        }
        let (index, data) = topology::message(&working, bucket, &[HEADER_BUCKET_MESSAGE_TYPE])?;
        let mut changed = false;
        let rewritten = rebuild(data, |field| {
            if field.number != 2 || single_varint(field.payload, 1)?.unwrap_or(0) < u64::from(limit)
            {
                return Ok(None);
            }
            changed = true;
            Ok(Some(Vec::new()))
        })?;
        if changed {
            replace(&mut working, bucket, index, rewritten)?;
        }
    }

    // The row/column UID map.
    if let Some(uids) = single_reference(&model, MODEL_UID_MAP_FIELD)? {
        let (index, data) = topology::message(&working, uids, &UID_MAP_MESSAGE_TYPES)?;
        let map =
            storage::UidMap::read(data, before.rows, before.columns).map_err(storage_fault)?;
        let rewritten = map
            .resize(before.rows, rows, before.columns, columns)
            .map_err(storage_fault)?;
        replace(&mut working, uids, index, rewritten)?;
    }

    // The stroke sidecar's size.
    if let Some(sidecar) = single_reference(&model, MODEL_STROKE_FIELD)? {
        let (index, data) = topology::message(&working, sidecar, &[STROKE_SIDECAR_MESSAGE_TYPE])?;
        let mut rewritten = data.to_vec();
        for (field, value) in [(2, columns), (3, rows)] {
            rewritten = litchi_iwa_common::wire::patch_varint_field(
                &rewritten,
                field,
                single_varint(&rewritten, field)?.is_some(),
                Some(u64::from(value)),
            )
            .map_err(|error| structure::map_wire(&error))?;
        }
        replace(&mut working, sidecar, index, rewritten)?;
    }

    // The table's own formula owner: its table and body extents.
    let extents = |rows: u32, columns: u32| {
        topology::calc::Extents::of(
            rows,
            columns,
            before.header_rows,
            before.header_columns,
            before.footer_rows,
        )
    };
    let mut owner = None;
    for object in working.source().state.components.iter_objects() {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == FORMULA_OWNER_MESSAGE_TYPE
                && single_reference(&message.data, 11)? == Some(target.info)
            {
                owner = Some((
                    object.archive_info.identifier.ok_or(Error::InvalidSource)?,
                    index,
                    message.data.clone(),
                ));
            }
        }
    }
    if let Some((identifier, index, data)) = owner
        && let Some(rewritten) = topology::calc::resize_extents(
            &data,
            extents(before.rows, before.columns),
            extents(rows, columns),
        )
        .map_err(storage_fault)?
    {
        replace(&mut working, identifier, index, rewritten)?;
    }

    // The model's size.
    let mut data = model;
    for (field, value) in [(MODEL_ROWS_FIELD, rows), (MODEL_COLUMNS_FIELD, columns)] {
        data = rebuild(&data, |candidate| {
            if candidate.number == field {
                return varint_field(field, u64::from(value)).map(Some);
            }
            Ok(None)
        })?;
    }
    replace(&mut working, target.model, model_index, data)?;
    Ok(Some(working.publish(&Transition::default())?))
}

impl Package {
    /// Start an immutable resize of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package.edit_table_resize(0, 0)?.resize(20, 6)?.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_resize<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet.into())
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table.into())?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            size: None,
        })
    }

    /// Apply a table-resize patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_resize(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
