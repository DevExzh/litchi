//! Body-table cell values (record 0309): text, numbers, booleans and
//! clears in a Pages body table, through the shared TST cell writer
//! (`litchi_numbers_wire::table_write`).
//!
//! A write rewrites only the cell's storage tile, the table's string list
//! and its row and column header counts; the cell keeps its style and
//! format keys. A cell that holds a formula, or one a formula or chart
//! reads, is refused, so no calculation cache goes stale. The patch keeps
//! both exact artifacts, so its inverse restores the source byte for byte.

use std::collections::BTreeMap;
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_numbers_wire::{
    formula_cache::{self, refresh},
    table_write::{self, CellInput},
};
use thiserror::Error as ThisError;

use super::section_text_style::{component_name_of, parse_member};
use super::{Package, find_object, table_lock};
use crate::BodyTableSelector;

const TABLE_MODEL: u32 = 6_001;
const TILE: u32 = 6_002;
const DATA_LIST: [u32; 2] = [6_005, 6_201];
const HEADER_BUCKET: u32 = 6_006;
const FORMULA_OWNER: u32 = 4_008;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE: u32 = 4_000;
const HEADER_BUCKET_ROWS: u32 = 65_536;

/// A value a body-table cell takes.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Value {
    /// Plain text.
    Text(String),
    /// A finite number.
    Number(f64),
    /// A boolean.
    Boolean(bool),
    /// No value (the cell's style and format stay).
    Empty,
}

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Cells staged in one edit.
    Cells,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("cells")
    }
}

/// A typed, content-redacted body-table value failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted body-table source is invalid")]
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
    #[error("could not allocate {amount} body-table value items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged values.
    #[error("body-table value candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("body-table value patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body table matches the selector.
    #[error("no body table matches the selector")]
    TableNotFound,
    /// The cell is outside the table, or a number is not finite.
    #[error("cell outside the table or value not finite")]
    InvalidCell,
    /// The cell holds a formula, or a formula that reads it is outside the
    /// shared cache refresh (record 0376).
    #[error("the cell takes part in a formula graph")]
    Formulas,
    /// The cell storage is outside the shared writer (a legacy tile, a
    /// missing tile, a rich-text, formula or commented cell).
    #[error("the table's cell storage is outside the shared writer")]
    UnsupportedStorage,
    /// Commit was requested before staging a cell.
    #[error("no body-table cell value is staged")]
    NoStagedOperation,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support body-table values")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::InvalidCell { .. } => RefusalKind::InvalidValue,
            Self::Formulas { .. }
            | Self::UnsupportedStorage { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_core::Error> for Error {
    fn from(error: litchi_iwa_core::Error) -> Self {
        match error {
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Self::Allocation { amount: requested }
            },
            _ => Self::InvalidSource,
        }
    }
}

fn storage_error(_error: table_write::WriteError) -> Error {
    Error::Refused(Refusal::UnsupportedStorage)
}

fn style_error(error: super::section_text_style::Error) -> Error {
    use super::section_text_style::Error as E;
    match error {
        E::Refused(crate::section::text_style::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        E::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_target_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(table_lock::Refusal::TableNotFound)
        | table_lock::Error::Refused(table_lock::Refusal::AmbiguousTableName) => {
            Error::Refused(Refusal::TableNotFound)
        },
        table_lock::Error::Refused(table_lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// A reversible, exact-source body-table value patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact artifact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one body-table value publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    cells: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.cells != 0
    }

    /// Number of cells written.
    #[must_use]
    pub const fn cells(self) -> usize {
        self.cells
    }
}

/// One validated body-table value publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One immutable edit that writes cell values of one body table.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<table_lock::BodyTableTarget>,
    cells: Vec<(u32, u32, Value)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("cells", &self.cells.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn target(&mut self) -> Result<&table_lock::BodyTableTarget, Error> {
        if self.target.is_none() {
            self.target = Some(
                self.selector
                    .resolve(self.source)
                    .map_err(map_target_error)?,
            );
        }
        self.target.as_ref().ok_or(Error::InvalidSource)
    }

    /// Stage `value` in the cell at zero-based `row` and `column`; a later
    /// stage of the same cell replaces the earlier one.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidCell`] outside the table or for a non-finite number.
    pub fn set(&mut self, row: u32, column: u32, value: Value) -> Result<&mut Self, Error> {
        let target = self.target()?.clone();
        if row >= target.table_rows || column >= target.table_columns {
            return Err(Error::Refused(Refusal::InvalidCell));
        }
        if let Value::Number(number) = value
            && !number.is_finite()
        {
            return Err(Error::Refused(Refusal::InvalidCell));
        }
        self.cells.retain(|(r, c, _)| (*r, *c) != (row, column));
        self.cells
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.cells.push((row, column, value));
        Ok(self)
    }

    /// Validate and atomically publish the staged values.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; [`Refusal::Formulas`] for a formula
    /// cell or a dependent formula the cache refresh does not support; [`Refusal::UnsupportedStorage`]; source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let mut this = self;
        this.target()?;
        let edit = this;
        if edit.cells.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let candidate = publish(&edit)?;
        let read = candidate
            .body_table_cells(BodyTableSelector::index(
                edit.target
                    .as_ref()
                    .ok_or(Error::InvalidSource)?
                    .table_position,
            ))
            .map_err(|_error| Error::Verification)?;
        for (row, column, value) in &edit.cells {
            let stored = read
                .table()
                .get(litchi_iwa_common::table::coordinate::At::new(*row, *column));
            use litchi_iwa_common::table::cell::value::Value as Read;
            let ok = match value {
                Value::Text(text) => stored == Some(&Read::Text(text.clone())),
                Value::Number(number) => {
                    matches!(stored, Some(Read::Number(stored)) if stored.get() == *number)
                },
                Value::Boolean(flag) => stored == Some(&Read::Boolean(*flag)),
                Value::Empty => !matches!(
                    stored,
                    Some(Read::Text(_) | Read::Number(_) | Read::Boolean(_))
                ),
            };
            if !ok {
                return Err(Error::Verification);
            }
        }
        let source_owner = SharedBytes::from_shared_slice(edit.source.state.source.shared_source());
        let target_owner = SharedBytes::from_shared_slice(candidate.state.source.shared_source());
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            },
            diagnostics: Diagnostics {
                cells: edit.cells.len(),
            },
        })
    }
}

/// The staged payloads of one write: object -> (message type, payload).
struct Staged<'p> {
    package: &'p Package,
    payloads: BTreeMap<u64, (u32, Vec<u8>)>,
}

impl Staged<'_> {
    fn payload(&self, identifier: u64, types: &[u32]) -> Result<(u32, &[u8]), Error> {
        if let Some((message_type, payload)) = self.payloads.get(&identifier) {
            return Ok((*message_type, payload));
        }
        let object = find_object(self.package.state.source.components(), identifier)
            .ok_or(Error::InvalidSource)?;
        let mut found = None;
        for message in &object.messages {
            if types.contains(&message.type_)
                && found
                    .replace((message.type_, message.data.as_slice()))
                    .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
        found.ok_or(Error::InvalidSource)
    }
}

fn reference(payload: &[u8], number: u32) -> Result<Option<u64>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
    for field in view.fields() {
        if field.number() == number {
            let inner = litchi_iwa_common::wire::WireView::parse(
                field.payload(),
                litchi_iwa_common::WireLimits::DEFAULT,
            )
            .map_err(|_error| Error::InvalidSource)?;
            for value in inner.fields() {
                if value.number() == 1 {
                    let (identifier, _) =
                        litchi_iwa_common::varint::decode_varint_from_bytes(value.payload())
                            .map_err(|_error| Error::InvalidSource)?;
                    return Ok(Some(identifier));
                }
            }
        }
    }
    Ok(None)
}

fn varint(payload: &[u8], number: u32) -> Result<Option<u64>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
    for field in view.fields() {
        if field.number() == number && field.wire_type() == 0 {
            let (value, _) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                .map_err(|_error| Error::InvalidSource)?;
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn message(payload: &[u8], number: u32) -> Result<Vec<&[u8]>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
    Ok(view
        .fields()
        .filter(|field| field.number() == number && field.wire_type() == 2)
        .map(|field| field.payload())
        .collect())
}

fn cell_input(value: &Value) -> CellInput<'_> {
    match value {
        Value::Text(text) => CellInput::Text(text),
        Value::Number(number) => CellInput::Number(*number),
        Value::Boolean(flag) => CellInput::Boolean(*flag),
        Value::Empty => CellInput::Empty,
    }
}

#[allow(clippy::too_many_lines, reason = "one staged publication")]
fn publish(edit: &Edit<'_>) -> Result<Package, Error> {
    let target = edit.target.as_ref().ok_or(Error::InvalidSource)?;
    let package = edit.source;
    let limits = package
        .numbers_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    // The calculation graph: plan the refresh of the formula caches that
    // read the written cells (record 0376) before any payload changes.
    let mut engine = None;
    let mut owners = Vec::new();
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ == FORMULA_OWNER {
                    owners.push(formula_cache::DependencyPayload {
                        identifier: object.archive_info.identifier.ok_or(Error::InvalidSource)?,
                        bytes: &message.data,
                    });
                } else if message.type_ == CALCULATION_ENGINE
                    && engine.replace(message.data.as_slice()).is_some()
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let object = |identifier: u64| -> Option<&[u8]> {
        find_object(package.state.source.components(), identifier)
            .and_then(|object| object.messages.first())
            .map(|message| message.data.as_slice())
    };
    let refresh = match engine {
        Some(engine) if !owners.is_empty() => {
            let mut inputs = Vec::new();
            inputs
                .try_reserve_exact(edit.cells.len())
                .map_err(|_error| Error::Allocation {
                    amount: edit.cells.len(),
                })?;
            for (row, column, value) in &edit.cells {
                inputs.push((*row, *column, cell_input(value)));
            }
            inputs.sort_unstable_by_key(|(row, column, _)| (*row, *column));
            let writes: Vec<refresh::Write<'_>> = inputs
                .iter()
                .map(|(row, column, value)| refresh::Write {
                    row: *row,
                    column: *column,
                    value,
                })
                .collect();
            refresh::plan(
                refresh::Source {
                    engine,
                    owners: &owners,
                    table: target.drawable_identifier.get(),
                },
                &writes,
                &object,
                &limits,
            )
            .map_err(|error| match error {
                formula_cache::Error::UnsupportedDependency(_) => Error::Refused(Refusal::Formulas),
                _ => Error::Refused(Refusal::UnsupportedStorage),
            })?
        },
        _ => refresh::Plan::default(),
    };
    let mut staged = Staged {
        package,
        payloads: BTreeMap::new(),
    };
    let (_, model_payload) = staged.payload(target.model_identifier.get(), &[TABLE_MODEL])?;
    let model_payload = model_payload.to_vec();
    let store = message(&model_payload, 4)?
        .first()
        .copied()
        .ok_or(Error::InvalidSource)?
        .to_vec();
    let tiles = message(&store, 3)?
        .first()
        .copied()
        .ok_or(Error::InvalidSource)?
        .to_vec();
    let size = varint(&tiles, 2)?.unwrap_or(256).max(1);
    let mut tile_ids = BTreeMap::new();
    for entry in message(&tiles, 1)? {
        let key = varint(entry, 1)?.ok_or(Error::InvalidSource)?;
        let tile = reference(entry, 2)?.ok_or(Error::InvalidSource)?;
        tile_ids.insert(key, tile);
    }
    let strings = reference(&store, 4)?.ok_or(Error::InvalidSource)?;
    let buckets: Vec<u64> = message(&store, 1)?
        .first()
        .map(|headers| -> Result<Vec<u64>, Error> {
            message(headers, 2)?
                .into_iter()
                .map(|bucket| {
                    let view = litchi_iwa_common::wire::WireView::parse(
                        bucket,
                        litchi_iwa_common::WireLimits::DEFAULT,
                    )
                    .map_err(|_error| Error::InvalidSource)?;
                    let field = view.fields().next().ok_or(Error::InvalidSource)?;
                    litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                        .map(|(value, _)| value)
                        .map_err(|_error| Error::InvalidSource)
                })
                .collect()
        })
        .transpose()?
        .unwrap_or_default();
    let column_headers = reference(&store, 2)?;
    for (row, column, value) in &edit.cells {
        let tile = *tile_ids
            .get(&(u64::from(*row) / size))
            .ok_or(Error::Refused(Refusal::UnsupportedStorage))?;
        let local = u32::try_from(u64::from(*row) % size).map_err(|_error| Error::InvalidSource)?;
        let (tile_type, tile_payload) = staged.payload(tile, &[TILE])?;
        let tile_payload = tile_payload.to_vec();
        let previous = table_write::tile_cell(&tile_payload, local, *column, &limits)
            .map_err(storage_error)?;
        let (list_type, list_payload) = staged.payload(strings, &DATA_LIST)?;
        let mut list = list_payload.to_vec();
        let mut key = None;
        if let Value::Text(text) = value {
            let (interned, interned_key) =
                table_write::list_intern(&list, text, &limits).map_err(storage_error)?;
            list = interned;
            key = Some(interned_key);
        }
        let input = cell_input(value);
        let (cell, released) = table_write::cell_bytes(previous.as_deref(), &input, key, &limits)
            .map_err(storage_error)?;
        if let Some(released) = released {
            list = table_write::list_release(&list, released, &limits).map_err(storage_error)?;
        }
        let written = table_write::tile_set_cell(
            &tile_payload,
            local,
            *column,
            target.table_columns,
            cell.as_deref(),
            &limits,
        )
        .map_err(storage_error)?;
        let delta = i64::from(cell.is_some()) - i64::from(written.previous.is_some());
        staged.payloads.insert(tile, (tile_type, written.tile));
        staged.payloads.insert(strings, (list_type, list));
        if delta != 0 {
            let bucket = *buckets
                .get((*row / HEADER_BUCKET_ROWS) as usize)
                .ok_or(Error::Refused(Refusal::UnsupportedStorage))?;
            let (bucket_type, payload) = staged.payload(bucket, &[HEADER_BUCKET])?;
            let adjusted =
                table_write::header_adjust(payload, *row, delta, &limits).map_err(storage_error)?;
            staged.payloads.insert(bucket, (bucket_type, adjusted));
            if let Some(column_headers) = column_headers {
                let (header_type, payload) = staged.payload(column_headers, &[HEADER_BUCKET])?;
                let adjusted = table_write::header_adjust(payload, *column, delta, &limits)
                    .map_err(storage_error)?;
                staged
                    .payloads
                    .insert(column_headers, (header_type, adjusted));
            }
        }
    }
    for caches in &refresh.tiles {
        let (tile_type, payload) = staged.payload(caches.tile, &[TILE])?;
        let refreshed = caches
            .apply(payload, target.table_columns, &limits)
            .map_err(|_error| Error::Refused(Refusal::Formulas))?;
        staged.payloads.insert(caches.tile, (tile_type, refreshed));
    }
    // One rewrite per touched member.
    let catalog = &package.state.source;
    let limits = catalog.limits();
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
    let snappy_limits = limits
        .snappy_limits()
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
    let mut members: BTreeMap<String, Archive> = BTreeMap::new();
    for (identifier, (message_type, payload)) in staged.payloads {
        let member = component_name_of(package, identifier).map_err(style_error)?;
        if !members.contains_key(&member) {
            let archive = parse_member(package, &member).map_err(style_error)?;
            members.insert(member.clone(), archive);
        }
        let archive = members.get_mut(&member).ok_or(Error::InvalidSource)?;
        let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == message_type)
            .ok_or(Error::InvalidSource)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            archive_limits,
        )?;
    }
    let mut encoded: Vec<(String, Vec<u8>)> = Vec::new();
    for (name, archive) in &members {
        let bytes = archive.to_bytes(archive_limits)?;
        encoded.push((name.clone(), SnappyStream::compress(&bytes, snappy_limits)?));
    }
    let edits: Vec<EntryEdit<'_>> = encoded
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let bytes = catalog
        .package()
        .reassemble_with_changes_to_bytes(&[], &edits, &[], limits)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
    let candidate = Package::from_bytes_with_archive_limits(&bytes, limits)
        .map_err(|_error| Error::Verification)?;
    let after = candidate.state.source.package();
    for (before, now) in catalog.package().iter().zip(after.iter()) {
        let rewritten = encoded.iter().any(|(name, _)| name == before.name());
        if before.name() != now.name() || (!rewritten && before.data() != now.data()) {
            return Err(Error::Verification);
        }
    }
    Ok(candidate)
}

impl Package {
    /// Start an immutable cell-value edit of one body table.
    ///
    /// ```no_run
    /// use litchi_pages::table::values::Value;
    /// use litchi_pages::{BodyTableSelector, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("report.pages")?;
    /// let mut edit = package.edit_body_table_values(BodyTableSelector::index(0))?;
    /// edit.set(1, 1, Value::Text("Q3".into()))?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_body_table_values(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`]; source errors.
    pub fn edit_body_table_values<'table, S>(&self, selector: S) -> Result<Edit<'_>, Error>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_target_error)?,
            target: None,
            cells: Vec::new(),
        })
    }

    /// Apply a body-table value patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_body_table_values(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_owner = SharedBytes::from_shared_slice(self.state.source.shared_source());
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            let target = Arc::<[u8]>::from(patch.artifacts.target_owner().as_ref());
            Package::from_bytes_with_archive_limits(&target, self.state.source.limits())
                .map_err(|_error| Error::Verification)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics::default(),
        })
    }
}
