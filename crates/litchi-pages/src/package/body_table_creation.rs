//! Exact-source creation of a body table (record 0329).
//!
//! A new, empty table is cloned from the structure of the document's first
//! body table (the template): its drawable, model and emptied storage come
//! from `litchi_numbers_wire::table_create` (the payload half of
//! `litchi-numbers`' table creation, record 0301), with styles shared. Every
//! clone goes into its source's package member; a copy of the template's
//! drawable attachment names the new table; one object-replacement
//! character (`U+FFFC`) and an attachment-table entry anchor it in the body
//! story at the requested UTF-16 position, shifting later text attributes
//! as a native text insertion does. The table gets its own formula owner in
//! the calculation engine, the clones of registered objects get registry
//! UUIDs and every new cross-member reference is registered. The graph is
//! the host editor's `add_table` graph (records 0370f and 0371): the
//! same identifier order, model references, kept drawable frame and
//! registration order; only the table UUID differs in value. Seeded cell
//! values are then written by the body-table value transaction
//! ([`crate::table::values`]); the patch covers both. No clock or
//! randomness is consulted: the same insertion into the same source
//! publishes the same bytes.

use std::collections::{BTreeMap, HashMap};
use std::fmt;

use litchi_core::Position;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, repeated_length_delimited_payloads,
    repeated_varint_values,
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text_wire::{RewriteBehavior, StorageRewriteExecutionLimits};
use litchi_numbers_wire::table_create::{
    self, CALCULATION_ENGINE_MESSAGE_TYPE, FORMULA_OWNER_MESSAGE_TYPE, FORMULA_OWNER_VERSIONS,
    Frame, OwnerGeometry, Role, Shape, TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE,
    TableUuid,
};
use thiserror::Error as ThisError;

use super::structure::{self, Artifacts, External, Fault, FaultLimit, Transition, Working};
use super::{Package, body_table_values, table_lock};
use crate::selector::BodyTableSelector;

/// `TSWP.StorageArchive` (the body story).
const BODY_STORAGE_TYPES: [u32; 2] = [2_001, 2_022];
/// `TSWP.StorageArchive.table_attachment`.
const ATTACHMENT_TABLE_FIELD: u32 = 9;
/// The object-replacement character a drawable attachment anchors at.
const ANCHOR: &str = "\u{fffc}";

/// A new body table for [`Edit::insert`]: its size, an
/// optional name and seeded cell values.
///
/// ```
/// use litchi_pages::table::creation::NewTable;
/// use litchi_pages::table::values::Value;
///
/// let table = NewTable::new(3, 2)
///     .with_name("Budget")
///     .with_cell(0, 1, Value::Text("Q3".into()));
/// assert_eq!((table.rows(), table.columns()), (3, 2));
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct NewTable {
    rows: u32,
    columns: u32,
    name: Option<String>,
    cells: Vec<(u32, u32, body_table_values::Value)>,
}

impl NewTable {
    /// An empty table of `rows` × `columns` cells (header rows and columns
    /// included, as the template's counts allow).
    #[must_use]
    pub const fn new(rows: u32, columns: u32) -> Self {
        Self {
            rows,
            columns,
            name: None,
            cells: Vec::new(),
        }
    }

    /// This table with a name (the default is `Table N`, after the
    /// document's tables).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// This table with a value in one cell (zero-based); a later value of
    /// the same cell replaces an earlier one.
    #[must_use]
    pub fn with_cell(mut self, row: u32, column: u32, value: body_table_values::Value) -> Self {
        self.cells.retain(|(r, c, _)| (*r, *c) != (row, column));
        self.cells.push((row, column, value));
        self
    }

    /// The number of rows.
    #[must_use]
    pub const fn rows(&self) -> u32 {
        self.rows
    }

    /// The number of columns.
    #[must_use]
    pub const fn columns(&self) -> u32 {
        self.columns
    }

    /// The requested name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The seeded cells, in staging order.
    #[must_use]
    pub fn cells(&self) -> &[(u32, u32, body_table_values::Value)] {
        &self.cells
    }
}

/// A finite resource governed by a body-table creation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes of one decoded component or payload.
    PayloadBytes,
    /// Object identifiers.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "identifiers",
        })
    }
}

/// A body-table creation failure that never exposes native identifiers,
/// component names or authored content.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The package cannot be edited safely (malformed or inconsistent).
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error(
        "Pages body-table creation {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// An allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table creation")]
    Allocation {
        /// The requested amount.
        amount: usize,
    },
    /// Writing a seeded cell failed.
    #[error("a seeded cell of the new Pages table could not be written")]
    Values(#[source] body_table_values::Error),
    /// The published candidate did not verify.
    #[error("the Pages body-table creation candidate failed verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Pages body-table creation patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Pages source does not support structural edits")]
    UnsupportedSource,
    /// The document holds no body table whose structure a new table can
    /// copy.
    #[error("the Pages document has no body table to copy")]
    NoTemplate,
    /// The size is empty or past the native row and column identity
    /// ceiling.
    #[error("a new Pages table of {rows} x {columns} cells is not supported")]
    InvalidSize {
        /// The requested rows.
        rows: u32,
        /// The requested columns.
        columns: u32,
    },
    /// A seeded cell lies outside the new table or holds a non-finite
    /// number.
    #[error("a seeded cell lies outside the new Pages table")]
    InvalidCell,
    /// The anchor is past the end of the body story.
    #[error("the Pages table anchor is outside the body story")]
    PositionOutOfBounds,
    /// The template or the body uses a native shape this transaction does
    /// not rewrite.
    #[error("the Pages table template or body is outside the creation transaction")]
    Unsupported,
    /// The edit already stages its one insertion.
    #[error("the Pages body-table creation edit already stages an insertion")]
    OperationAlreadyStaged,
    /// The edit was committed without a staged insertion.
    #[error("the Pages body-table creation edit stages no insertion")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } | Self::NoTemplate { .. } | Self::Unsupported { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::InvalidSize { .. }
            | Self::InvalidCell { .. }
            | Self::PositionOutOfBounds { .. } => RefusalKind::InvalidValue,
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::Unsupported),
            Fault::Invalid | Fault::Dependency => Self::InvalidSource,
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

/// Exact-source-checked reversible body-table creation patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: Artifacts,
    table: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("table", &self.table)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_noop()
    }

    /// The exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            table: self.table,
        }
    }
}

/// Compact evidence for one verified body-table creation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    created_objects: usize,
    touched_components: usize,
    seeded_cells: usize,
}

impl Diagnostics {
    /// Whether the commit changed the package (false only for a no-op patch
    /// replay).
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many native objects the creation inserted.
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }

    /// How many package components the creation rewrote.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// How many seeded cells were written.
    #[must_use]
    pub const fn seeded_cells(self) -> usize {
        self.seeded_cells
    }
}

/// The verified result of one body-table creation.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the committed package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// The new table's position in the body-table catalog
    /// ([`Package::body_tables`]).
    #[must_use]
    pub const fn table(&self) -> Position {
        Position::new(self.patch.table)
    }

    /// Split the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// One immutable edit that stages exactly one body-table insertion.
///
/// Start it with [`Package::edit_body_table_creation`]; the source package is
/// never changed, and dropping the edit discards the staged insertion.
pub struct Edit<'a> {
    source: &'a Package,
    insertion: Option<(Position, NewTable)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "insertion",
                &self
                    .insertion
                    .as_ref()
                    .map(|(anchor, table)| (anchor, table.rows, table.columns)),
            )
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage a new, empty body table anchored at UTF-16 code unit `anchor`
    /// of the body story, cloned from the structure of the document's first
    /// body table (record 0329). The anchor character joins the paragraph
    /// holding `anchor`; seeded cells are written with the body-table value
    /// writer.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one insertion,
    /// [`Refusal::InvalidSize`] and [`Refusal::InvalidCell`] for a table this
    /// transaction cannot create.
    pub fn insert(mut self, anchor: Position, table: NewTable) -> Result<Self, Error> {
        if self.insertion.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        if !table_create::admits_size(table.rows, table.columns) {
            return Err(Error::Refused(Refusal::InvalidSize {
                rows: table.rows,
                columns: table.columns,
            }));
        }
        if table.cells.iter().any(|(row, column, value)| {
            *row >= table.rows
                || *column >= table.columns
                || matches!(value, body_table_values::Value::Number(number) if !number.is_finite())
        }) {
            return Err(Error::Refused(Refusal::InvalidCell));
        }
        self.insertion = Some((anchor, table));
        Ok(self)
    }

    /// Validate, publish and reopen the staged insertion.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit,
    /// [`Refusal::UnsupportedSource`] for a source that is not exact,
    /// [`Refusal::NoTemplate`] for a document without a body table,
    /// [`Refusal::PositionOutOfBounds`], [`Refusal::Unsupported`] for a
    /// template or body shape outside the transaction, and typed source,
    /// limit, allocation and verification errors. Failure never publishes a
    /// partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let (anchor, table) = self
            .insertion
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        if !self.source.state.source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        create(self.source, anchor.get(), &table)
    }
}

impl Package {
    /// Start an immutable single-insertion body-table creation edit.
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_pages::Package;
    /// use litchi_pages::table::creation::NewTable;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("report.pages")?;
    /// let commit = package
    ///     .edit_body_table_creation()
    ///     .insert(Position::new(0), NewTable::new(3, 2))?
    ///     .commit()?;
    /// assert_eq!(
    ///     commit.package().body_tables()?.len(),
    ///     package.body_tables()?.len() + 1
    /// );
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_body_table_creation(&self) -> Edit<'_> {
        Edit {
            source: self,
            insertion: None,
        }
    }

    /// Apply a body-table creation patch (or its inverse) only to its exact
    /// immutable source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source, and typed source errors when the target cannot be reopened.
    pub fn apply_body_table_creation(&self, patch: &Patch) -> Result<Commit, Error> {
        let package = patch.artifacts.apply(self)?.ok_or(Error::PatchConflict)?;
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                ..Diagnostics::default()
            },
        })
    }
}

/// The single payload of `message_type` of `object`.
fn single(object: &ArchiveObject, message_type: u32) -> Result<&[u8], Error> {
    match object.messages.as_slice() {
        [message] if message.type_ == message_type => Ok(&message.data),
        _ => Err(Error::Refused(Refusal::Unsupported)),
    }
}

fn storage(_error: litchi_numbers_wire::table_write::WriteError) -> Error {
    Error::Refused(Refusal::Unsupported)
}

/// A new single-message object whose header copies `template`'s message
/// versions and merge flag and declares `references`.
fn fresh(
    working: &Working<'_>,
    identifier: u64,
    template: Option<&ArchiveObject>,
    message_type: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        working.archive_limits(),
    )
    .map_err(|error| Error::from(structure::map_core(&error)))?;
    if let Some(template) = template {
        object.archive_info.should_merge = template.archive_info.should_merge;
        if let Some(info) = template.archive_info.message_infos.first() {
            object.archive_info.message_infos[0].versions = info.versions.clone();
        }
    }
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

/// The references a clone declares: its template's declared references,
/// mapped through `remap`, kept only where the clone's payload holds them.
fn declared(template: &ArchiveObject, data: &[u8], remap: &dyn Fn(u64) -> Option<u64>) -> Vec<u64> {
    let mut output = Vec::new();
    for identifier in template
        .archive_info
        .message_infos
        .first()
        .map(|info| info.object_references.as_slice())
        .unwrap_or_default()
    {
        let mapped = remap(*identifier).unwrap_or(*identifier);
        if table_create::references_identifier(data, mapped) && !output.contains(&mapped) {
            output.push(mapped);
        }
    }
    output
}

/// The objects a model payload references, sorted, as the host editor's
/// `table_model_references` lists them (record 0370f): its style and owner
/// references, and its data store's headers, tiles and lists.
fn model_references(model: &[u8]) -> Result<Vec<u64>, Error> {
    const MODEL_FIELDS: [u32; 43] = [
        3, 24, 25, 26, 27, 18, 19, 20, 21, 30, 36, 48, 5, 34, 35, 38, 46, 49, 60, 61, 62, 63, 64,
        65, 66, 67, 68, 69, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 85, 86, 87, 88, 89,
    ];
    const STORE_FIELDS: [u32; 15] = [2, 4, 5, 6, 11, 12, 16, 13, 15, 17, 18, 19, 20, 21, 22];
    fn at(data: &[u8], path: &[u32], output: &mut Vec<u64>) -> Result<(), Error> {
        match path {
            [] => Ok(()),
            [last] => {
                output.extend(table_create::references(data, *last).map_err(storage)?);
                Ok(())
            },
            [first, rest @ ..] => {
                for payload in repeated_length_delimited_payloads(data, *first)
                    .map_err(|_error| Error::InvalidSource)?
                {
                    at(payload, rest, output)?;
                }
                Ok(())
            },
        }
    }
    let mut output = Vec::new();
    for field in MODEL_FIELDS {
        at(model, &[field], &mut output)?;
    }
    at(model, &[4, 1, 2], &mut output)?;
    at(model, &[4, 3, 1, 2], &mut output)?;
    for field in STORE_FIELDS {
        at(model, &[4, field], &mut output)?;
    }
    output.retain(|identifier| *identifier != 0);
    output.sort_unstable();
    output.dedup();
    Ok(output)
}

/// The calculation engine (identifier, message index), when the document
/// has one.
fn engine(source: &Package) -> Result<Option<(u64, usize)>, Error> {
    let mut found = None;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                    let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
                    if found.replace((identifier, index)).is_some() {
                        return Err(Error::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(found)
}

/// The decoded double at field `number` of `data`, if present.
fn double(data: &[u8], number: u32) -> Option<f64> {
    let view = litchi_iwa_common::wire::parse_wire_view(data, WireLimits::DEFAULT).ok()?;
    view.fields()
        .filter(|field| field.number() == number && field.wire_type() == 1)
        .last()
        .and_then(|field| field.payload().try_into().ok())
        .map(f64::from_le_bytes)
}

/// The template drawable's position, when it records one.
/// The template drawable's geometry pair `member` (1 position, 2 size).
fn template_geometry(info: &[u8], member: u32) -> Option<(f32, f32)> {
    let drawable = repeated_length_delimited_payloads(info, 1).ok()?;
    let geometry = repeated_length_delimited_payloads(drawable.first()?, 1).ok()?;
    let position = repeated_length_delimited_payloads(geometry.first()?, member).ok()?;
    let view =
        litchi_iwa_common::wire::parse_wire_view(position.first()?, WireLimits::DEFAULT).ok()?;
    let mut point = (0.0_f32, 0.0_f32);
    for field in view.fields() {
        let value = field
            .payload()
            .try_into()
            .ok()
            .map(f32::from_le_bytes)
            .filter(|value| value.is_finite());
        match (field.number(), field.wire_type(), value) {
            (1, 5, Some(value)) => point.0 = value,
            (2, 5, Some(value)) => point.1 = value,
            _ => {},
        }
    }
    Some(point)
}

/// The cross-member reference registrations a transaction adds.
#[derive(Default)]
struct References {
    additions: Vec<External>,
}

impl References {
    /// Register `target` as referenced from the component of `from`, unless
    /// both live in one member or the registry already records it.
    fn add(
        &mut self,
        working: &Working<'_>,
        registry: &structure::Registry,
        from: u64,
        target: u64,
    ) -> Result<(), Error> {
        let from_member = working.locate(from)?.ok_or(Error::InvalidSource)?;
        let target_member = working.locate(target)?.ok_or(Error::InvalidSource)?;
        if from_member == target_member {
            return Ok(());
        }
        let source = registry.component_for_member(&from_member)?.identifier;
        let destination = registry.component_for_member(&target_member)?.identifier;
        let covered = registry.covers(source, destination, target)
            || self.additions.iter().any(|external| {
                external.source == source
                    && external.target == destination
                    && external.object == target
            });
        if !covered {
            self.additions.push(External {
                source,
                target: destination,
                object: target,
                weak: None,
            });
        }
        Ok(())
    }
}

/// Insert the anchor character and its attachment-table entry into the
/// body storage payload.
fn anchored_body(
    source: &Package,
    payload: &[u8],
    anchor: usize,
    attachment: u64,
) -> Result<Vec<u8>, Error> {
    let limits = super::storage_rewrite_limits(source.state.source.limits())
        .map_err(|_error| Error::InvalidSource)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        anchor..anchor,
        ANCHOR,
        RewriteBehavior::PreserveOnEqualText,
        limits,
    )
    .map_err(|_error| Error::Refused(Refusal::Unsupported))?;
    let requirements = prepared.execution_requirements();
    let rewritten = prepared
        .execute(StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(|_error| Error::Refused(Refusal::Unsupported))?
        .into_bytes();
    // The attachment table, with the new entry in character order.
    let wire = WireLimits::default()
        .with_input_bytes(rewritten.len().max(1))
        .map_err(|_error| Error::InvalidSource)?;
    let view = WireView::parse(&rewritten, wire).map_err(|_error| Error::InvalidSource)?;
    let tables: Vec<&[u8]> = view
        .fields()
        .filter(|field| field.number() == ATTACHMENT_TABLE_FIELD)
        .map(|field| field.payload())
        .collect();
    if tables.len() > 1 {
        return Err(Error::Refused(Refusal::Unsupported));
    }
    let mut entry = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut entry, 1, anchor as u64)
        .map_err(|_error| Error::InvalidSource)?;
    append_length_delimited_field(
        &mut entry,
        2,
        &table_create::reference_payload(attachment),
        WireLimits::DEFAULT,
    )
    .map_err(|_error| Error::InvalidSource)?;
    let mut table = Vec::new();
    let mut placed = false;
    if let Some(existing) = tables.first() {
        let view = litchi_iwa_common::wire::parse_wire_view(existing, WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
        for field in view.fields() {
            if field.number() != 1 || field.wire_type() != 2 {
                return Err(Error::Refused(Refusal::Unsupported));
            }
            let index = repeated_varint_values(field.payload(), 1)
                .map_err(|_error| Error::InvalidSource)?;
            let [index] = index.as_slice() else {
                return Err(Error::InvalidSource);
            };
            if *index == anchor as u64 {
                return Err(Error::InvalidSource);
            }
            if !placed && *index > anchor as u64 {
                append_length_delimited_field(&mut table, 1, &entry, WireLimits::DEFAULT)
                    .map_err(|_error| Error::InvalidSource)?;
                placed = true;
            }
            table.extend_from_slice(field.raw());
        }
    }
    if !placed {
        append_length_delimited_field(&mut table, 1, &entry, WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
    }
    let present = !tables.is_empty();
    litchi_iwa_common::wire::patch_length_delimited_field(
        &rewritten,
        ATTACHMENT_TABLE_FIELD,
        present,
        Some(&table),
    )
    .map_err(|_error| Error::InvalidSource)
}

/// The UTF-16 length of a storage payload's text.
fn prepared_len(payload: &[u8]) -> Result<usize, Error> {
    let texts =
        repeated_length_delimited_payloads(payload, 3).map_err(|_error| Error::InvalidSource)?;
    let mut length = 0_usize;
    for text in texts {
        let text = std::str::from_utf8(text).map_err(|_error| Error::InvalidSource)?;
        length += text.encode_utf16().count();
    }
    Ok(length)
}

fn create(source: &Package, anchor: usize, table: &NewTable) -> Result<Commit, Error> {
    if source
        .body_tables()
        .map_err(|_error| Error::Refused(Refusal::Unsupported))?
        .is_empty()
    {
        return Err(Error::Refused(Refusal::NoTemplate));
    }
    let mut budget = table_lock::WireBudget::new(source.state.source.limits())
        .map_err(|_| Error::InvalidSource)?;
    let targets = table_lock::native_body_table_targets_with_budget(source, &mut budget)
        .map_err(|_error| Error::Refused(Refusal::Unsupported))?;
    let template = targets.first().ok_or(Error::Refused(Refusal::NoTemplate))?;
    let template_info = template.drawable_identifier.get();
    let template_model = template.model_identifier.get();
    let template_attachment = template.attachment_identifier.get();
    let source_tables = targets.len();
    let name = match table.name.as_deref() {
        Some(name) => name.to_owned(),
        None => format!("Table {}", source_tables + 1),
    };
    let mut working = Working::new(source)?;
    let registry = working
        .registry()?
        .ok_or(Error::Refused(Refusal::Unsupported))?;

    let info_object = working.object(template_info)?.clone();
    let info_template = single(&info_object, TABLE_INFO_MESSAGE_TYPE)?.to_vec();
    if table_create::single_reference(&info_template, 2).map_err(storage)? != Some(template_model) {
        return Err(Error::InvalidSource);
    }
    let body = table_create::drawable_parent(&info_template)
        .map_err(storage)?
        .ok_or(Error::Refused(Refusal::Unsupported))?;
    let model_object = working.object(template_model)?.clone();
    let model_template = single(&model_object, TABLE_MODEL_MESSAGE_TYPE)?.to_vec();
    let owned: BTreeMap<u64, Role> =
        table_create::owned_storage(&model_template).map_err(storage)?;
    let attachment_object = working.object(template_attachment)?.clone();
    let [attachment_message] = attachment_object.messages.as_slice() else {
        return Err(Error::Refused(Refusal::Unsupported));
    };
    if table_create::single_reference(&attachment_message.data, 1).map_err(storage)?
        != Some(template_info)
    {
        return Err(Error::Refused(Refusal::Unsupported));
    }

    // Fresh identities, in the host editor's order (record 0370f): the
    // drawable, the model, the owned storage in source-identifier order,
    // the formula owner (when the document has an engine), then the
    // attachment.
    let engine_root = engine(source)?;
    let storage_count = owned.len();
    let fresh_identifiers =
        working.allocate(storage_count + 3 + usize::from(engine_root.is_some()))?;
    let (info, new_model) = (fresh_identifiers[0], fresh_identifiers[1]);
    let owner = fresh_identifiers[2 + storage_count];
    let attachment = fresh_identifiers[fresh_identifiers.len() - 1];
    let remap: HashMap<u64, u64> = owned
        .keys()
        .copied()
        .zip(fresh_identifiers[2..2 + storage_count].iter().copied())
        .collect();
    let seed = working.identity_seed(info)?;
    let uuid = TableUuid::derive(seed);
    let mut inserted: Vec<(u64, u64)> = Vec::with_capacity(owned.len() + 3);

    // The storage clones, each in its source's member.
    for (identifier, role) in &owned {
        let object = working.object(*identifier)?.clone();
        let [message] = object.messages.as_slice() else {
            return Err(Error::Refused(Refusal::Unsupported));
        };
        let member = working.locate(*identifier)?.ok_or(Error::InvalidSource)?;
        let data =
            table_create::empty_storage(*role, &message.data, table.rows, table.columns, new_model)
                .map_err(storage)?;
        let clone = fresh(
            &working,
            remap[identifier],
            Some(&object),
            message.type_,
            data,
            Vec::new(),
        )?;
        working.insert(&member, clone)?;
        inserted.push((remap[identifier], *identifier));
    }

    // The model, with its styles shared and its storage remapped.
    let shape = Shape {
        name: &name,
        rows: table.rows,
        columns: table.columns,
        row_height: None,
        column_width: None,
    };
    let model_data = table_create::model_payload(
        &model_template,
        &|identifier| remap.get(&identifier).copied(),
        &uuid,
        &shape,
    )
    .map_err(storage)?;
    let geometry = OwnerGeometry::of(&model_data, table.rows, table.columns).map_err(storage)?;
    let row_height = double(&model_data, 16)
        .filter(|value| *value > 0.0)
        .ok_or(Error::Refused(Refusal::Unsupported))?;
    let column_width = double(&model_data, 17)
        .filter(|value| *value > 0.0)
        .ok_or(Error::Refused(Refusal::Unsupported))?;
    let model_references = model_references(&model_data)?;
    let model_member = working
        .locate(template_model)?
        .ok_or(Error::InvalidSource)?;
    let clone = fresh(
        &working,
        new_model,
        Some(&model_object),
        TABLE_MODEL_MESSAGE_TYPE,
        model_data,
        model_references,
    )?;
    working.insert(&model_member, clone)?;
    inserted.push((new_model, template_model));

    // The drawable, parented to the body story.
    // The host editor keeps the template drawable's geometry; a template
    // without a size gets the default row and column extents.
    #[expect(
        clippy::cast_possible_truncation,
        reason = "native drawable geometry is single precision"
    )]
    let frame = {
        let (x, y) = template_geometry(&info_template, 1).unwrap_or((0.0, 0.0));
        let (width, height) = template_geometry(&info_template, 2).unwrap_or((
            (column_width * f64::from(table.columns)) as f32,
            (row_height * f64::from(table.rows)) as f32,
        ));
        Frame {
            x,
            y,
            width,
            height,
        }
    };
    let info_data =
        table_create::info_payload(&info_template, body, new_model, frame).map_err(storage)?;
    // The parent and the model, as the host editor declares them.
    let info_references = vec![body, new_model];
    let info_member = working.locate(template_info)?.ok_or(Error::InvalidSource)?;
    let clone = fresh(
        &working,
        info,
        Some(&info_object),
        TABLE_INFO_MESSAGE_TYPE,
        info_data,
        info_references,
    )?;
    working.insert(&info_member, clone)?;
    inserted.push((info, template_info));

    // The attachment, a copy of the template's naming the new drawable.
    let attachment_data = litchi_iwa_common::wire::patch_length_delimited_field(
        &attachment_message.data,
        1,
        true,
        Some(&table_create::reference_payload(info)),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let attachment_references = declared(&attachment_object, &attachment_data, &|identifier| {
        (identifier == template_info).then_some(info)
    });
    let attachment_member = working
        .locate(template_attachment)?
        .ok_or(Error::InvalidSource)?;
    let clone = fresh(
        &working,
        attachment,
        Some(&attachment_object),
        attachment_message.type_,
        attachment_data,
        attachment_references,
    )?;
    working.insert(&attachment_member, clone)?;
    inserted.push((attachment, template_attachment));

    // The body story: the anchor character and its attachment entry.
    let limits = working.archive_limits();
    let body_object = working.object(body)?;
    let body_index = structure::unique_message(body_object, &BODY_STORAGE_TYPES)?;
    let body_payload = body_object.messages[body_index].data.clone();
    if anchor > prepared_len(&body_payload)? {
        return Err(Error::Refused(Refusal::PositionOutOfBounds));
    }
    let body_data = anchored_body(source, &body_payload, anchor, attachment)?;
    structure::replace(
        working.object_mut(body)?,
        body_index,
        body_data,
        &[],
        &[attachment],
        limits,
    )?;

    // The table's own formula owner, when the document has an engine.
    let mut references = References::default();
    if let Some((engine_id, engine_index)) = engine_root {
        let engine_member = working.locate(engine_id)?.ok_or(Error::InvalidSource)?;
        let engine_data = working.object(engine_id)?.messages[engine_index]
            .data
            .clone();
        let owners: Vec<&[u8]> = source
            .state
            .source
            .components()
            .get(&engine_member)
            .map(|component| {
                component
                    .archive()
                    .objects
                    .iter()
                    .flat_map(|object| object.messages.iter())
                    .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
                    .map(|message| message.data.as_slice())
                    .collect()
            })
            .unwrap_or_default();
        let internal = table_create::next_internal_owner(&owners, &engine_data).map_err(storage)?;
        let uid = uuid.owner_uid();
        let owner_data =
            table_create::owner_payload(uid, internal, info, geometry).map_err(storage)?;
        let mut owner_object = fresh(
            &working,
            owner,
            None,
            FORMULA_OWNER_MESSAGE_TYPE,
            owner_data,
            Vec::new(),
        )?;
        owner_object.archive_info.message_infos[0].versions =
            FORMULA_OWNER_VERSIONS.to_vec().into();
        working.insert(&engine_member, owner_object)?;
        let engine_data =
            table_create::engine_payload(&engine_data, owner, internal, uid).map_err(storage)?;
        structure::replace(
            working.object_mut(engine_id)?,
            engine_index,
            engine_data,
            &[],
            &[owner],
            limits,
        )?;
        references.add(&working, &registry, owner, info)?;
        // The host registers the new owner's UUID; it has no template, so
        // it is listed as its own.
        inserted.push((owner, owner));
    }

    // Registry: UUIDs for the clones of registered objects, and every new
    // cross-member reference.
    let mut transition = Transition::default();
    // In identifier order, as the host editor registers them.
    inserted.sort_by_key(|(clone, _)| *clone);
    for (clone, template) in &inserted {
        let member = working.locate(*clone)?.ok_or(Error::InvalidSource)?;
        let component = registry.component_for_member(&member)?.identifier;
        if clone == template
            || registry
                .uuids
                .iter()
                .any(|binding| binding.object == *template)
        {
            transition.register_uuid(&registry, component, *clone, seed)?;
        }
        let targets = working.object(*clone)?.archive_info.message_infos[0]
            .object_references
            .clone();
        for target in targets {
            references.add(&working, &registry, *clone, target)?;
        }
    }
    references.add(&working, &registry, body, attachment)?;
    transition.external_additions = references.additions;
    let created_objects = inserted.len();
    let (candidate, touched_components) = working.publish(&transition)?;

    // Read the new table back through the catalog and the cell reader.
    let catalog = candidate
        .body_tables()
        .map_err(|_error| Error::Verification)?;
    if catalog.len() != source_tables + 1 {
        return Err(Error::Verification);
    }
    let mut budget = table_lock::WireBudget::new(candidate.state.source.limits())
        .map_err(|_error| Error::Verification)?;
    let position = table_lock::native_body_table_targets_with_budget(&candidate, &mut budget)
        .map_err(|_error| Error::Verification)?
        .iter()
        .position(|target| target.drawable_identifier.get() == info)
        .ok_or(Error::Verification)?;
    let snapshot = catalog.get(position).ok_or(Error::Verification)?;
    if snapshot.name() != name
        || (snapshot.rows(), snapshot.columns()) != (table.rows, table.columns)
    {
        return Err(Error::Verification);
    }
    candidate
        .body_table_cells(BodyTableSelector::index(position))
        .map_err(|_error| Error::Verification)?;

    // Seeded values, through the body-table value writer.
    let mut package = candidate;
    if !table.cells.is_empty() {
        let mut edit = package
            .edit_body_table_values(BodyTableSelector::index(position))
            .map_err(Error::Values)?;
        for (row, column, value) in &table.cells {
            edit.set(*row, *column, value.clone())
                .map_err(Error::Values)?;
        }
        package = edit.commit().map_err(Error::Values)?.into_package();
    }
    Ok(Commit {
        patch: Patch {
            artifacts: Artifacts::new(source, &package),
            table: position,
        },
        package,
        diagnostics: Diagnostics {
            changed: true,
            created_objects,
            touched_components,
            seeded_cells: table.cells.len(),
        },
    })
}
