//! Implementation of the public [`crate::cell::list`] transactions:
//! paragraph lists of the text in one cell (record 0370b).
//!
//! A cell's text carries lists only as rich text: a `TST.RichTextPayload`
//! in its table's rich-text list names a `TSWP.StorageArchive` whose
//! list-style table gives each paragraph its list. A plain cell (empty, or a
//! string key) reads as unlisted paragraphs; the first change that needs a
//! list promotes it the way the host editor's `cell_paragraph_list` does:
//!
//! - the table gains a rich-text list when it has none;
//! - a cell storage holding the cell's text is stored beside the table
//!   model, its paragraphs styled by the cell's text style and the
//!   stylesheet's `None` list preset, with a rich-text payload naming it;
//! - a rich-text entry names the payload, the cell becomes a rich-text cell
//!   pointing at it, and the cell's string entry is released.
//!
//! The list change itself runs W1-D's engine
//! ([`litchi_iwa_text_wire::list`], record 0374) and is published by
//! `litchi_iwa_text_wire::list::publish`. Host `set_x`/`reset_x` pairs collapse into one
//! property verb ([`Edit::set`]/[`Edit::reset`]); `reset_list` and
//! `reset_level` are [`Edit::set_list`] with [`ParagraphList::None`] and
//! [`Edit::set_level`] with [`ParagraphListLevel::ZERO`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
    ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Change, Property, PropertyKind};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{
    self, Fault, Published, Working,
    metadata::{self, Transition},
};
use super::table_cell_text_style as cells;
use super::table_topology::wire::{
    bytes_field, push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_bytes,
    single_reference, single_varint, varint_field,
};
use super::table_topology::{
    self as topology, LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE, TILE_MESSAGE_TYPE, storage,
};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;
use litchi_iwa_table::structure::list_workspace::{
    ListWorkspace, ReadStyles as ListStyles, storage_payload,
};
use litchi_iwa_text_wire::list::publish::{self as storage_list, Failure, Publisher};

structure::structural_family!("cell-list", Wire(litchi_iwa_common::LimitKind));

/// `TST.RichTextPayloadArchive`.
const RICH_TEXT_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
/// The message version of the objects a promotion stores.
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
/// `DataStore.stringTable`.
const STRING_LIST_FIELD: u32 = 4;
/// `DataStore.rich_text_table`.
const RICH_TEXT_LIST_FIELD: u32 = 17;
/// `TableDataList.ListType.RICH_TEXT_PAYLOAD`.
const RICH_TEXT_LIST_TYPE: u64 = 8;

/// A typed, content-redacted cell list failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell list source is invalid")]
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
    #[error("could not allocate {amount} cell list items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell list candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell list patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// Commit was requested before staging a change.
    #[error("no cell list change is staged")]
    NoStagedOperation,
    /// The position is not the start of a paragraph of the cell's text.
    #[error("position is not a paragraph start")]
    NotAParagraphStart,
    /// The paragraph's list kind has no such property (a bullet of a
    /// numbered paragraph, any list property of an unlisted paragraph).
    #[error("the paragraph's list has no such property")]
    NotApplicable,
    /// Placements must start at zero and strictly increase.
    #[error("list placements must start at zero and strictly increase")]
    InvalidPlacements,
    /// The cell's paragraphs do not share one list.
    #[error("the cell's paragraphs do not share one list")]
    NotUniform,
    /// The cell holds a value other than text (a number, a formula, …).
    #[error("the cell does not hold text")]
    NotText,
    /// The cell's rich text or styles live outside what this transaction
    /// rewrites (a shared rich-text payload, a segmented list, no `None`
    /// list preset).
    #[error("a surviving object depends on the cell's rich-text graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell lists")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::NotAParagraphStart | Self::InvalidPlacements => RefusalKind::InvalidValue,
            Self::NotApplicable | Self::NotUniform | Self::NotText => {
                RefusalKind::UnsupportedOperation
            },
            Self::UnsupportedDependency | Self::UnsupportedSource => {
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
        from_cells(cells::from_topology(error))
    }
}

fn from_cells(error: cells::Error) -> Error {
    use cells::{Error as E, LimitKind as K, Refusal as R};
    match error {
        E::Refused(R::SheetNotFound) => Error::Refused(Refusal::SheetNotFound),
        E::Refused(R::TableNotFound) => Error::Refused(Refusal::TableNotFound),
        E::Refused(R::CellNotFound) => Error::Refused(Refusal::CellNotFound),
        E::Refused(R::UnsupportedDependency) => Error::Refused(Refusal::UnsupportedDependency),
        E::Refused(R::UnsupportedSource) => Error::Refused(Refusal::UnsupportedSource),
        E::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                K::PackageBytes => LimitKind::PackageBytes,
                K::PayloadBytes => LimitKind::PayloadBytes,
                K::Identifiers => LimitKind::Identifiers,
                K::ScratchBytes => LimitKind::ScratchBytes,
                K::AllocationEvents => LimitKind::AllocationEvents,
                K::Wire(kind) => LimitKind::Wire(kind),
            },
            observed,
            maximum,
        },
        E::Allocation { amount } => Error::Allocation { amount },
        E::Verification => Error::Verification,
        E::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn from_list(error: list::Error) -> Error {
    match error {
        list::Error::Refused(list::Refusal::NotAParagraphStart) => {
            Error::Refused(Refusal::NotAParagraphStart)
        },
        list::Error::Refused(list::Refusal::NotApplicable) => {
            Error::Refused(Refusal::NotApplicable)
        },
        list::Error::Refused(list::Refusal::InvalidPlacements) => {
            Error::Refused(Refusal::InvalidPlacements)
        },
        list::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        list::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl From<Failure<Fault>> for Error {
    fn from(failure: Failure<Fault>) -> Self {
        match failure {
            Failure::Workspace(fault) => fault.into(),
            Failure::List(error) => from_list(error),
        }
    }
}

/// One staged change of one cell.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    List(ParagraphList),
    Lists(Vec<ParagraphListPlacement>),
    Level(TextPosition, ParagraphListLevel),
    Numbering(TextPosition, ParagraphListNumbering),
    Set(TextPosition, Property),
    Reset(TextPosition, PropertyKind),
}

/// One immutable edit that changes paragraph lists of cells of one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, position: At, change: Staged) -> Result<Self, Error> {
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((position, change));
        Ok(self)
    }

    /// Stage giving every paragraph of the cell at `position` the list
    /// preset `list` ([`ParagraphList::None`] removes the lists).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_list(self, position: At, list: ParagraphList) -> Result<Self, Error> {
        self.stage(position, Staged::List(list))
    }

    /// Stage giving the cell's paragraphs the presets of `placements`: each
    /// placement's list runs from its paragraph to the next placement.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_lists(
        self,
        position: At,
        placements: impl IntoIterator<Item = ParagraphListPlacement>,
    ) -> Result<Self, Error> {
        let mut staged = Vec::new();
        for placement in placements {
            staged
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            staged.push(placement);
        }
        self.stage(position, Staged::Lists(staged))
    }

    /// Stage the list level of the paragraph starting at `paragraph`
    /// ([`ParagraphListLevel::ZERO`] resets it).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_level(
        self,
        position: At,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Level(paragraph, level))
    }

    /// Stage whether the paragraph starting at `paragraph` continues or
    /// restarts its list's numbering.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_numbering(
        self,
        position: At,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Numbering(paragraph, numbering))
    }

    /// Stage `property` as the effective list property of the paragraph
    /// starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(
        self,
        position: At,
        paragraph: TextPosition,
        property: Property,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Set(paragraph, property))
    }

    /// Stage restoring the list's own value of `kind` for the paragraph
    /// starting at `paragraph` (its variation's override goes).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(
        self,
        position: At,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Reset(paragraph, kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the cell as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; the list refusals; source, resource
    /// and verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.table, *position, change)? {
                verify(
                    &published.package,
                    self.sheet,
                    self.table,
                    *position,
                    change,
                )?;
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The published change reads back.
fn verify(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
    change: &Staged,
) -> Result<(), Error> {
    let reader = Reader::new(package, sheet, table, position)?;
    let holds = match change {
        Staged::List(preset) => reader.list()? == *preset,
        Staged::Lists(_) | Staged::Reset(..) => true,
        Staged::Level(paragraph, level) => reader.level(*paragraph)? == *level,
        Staged::Numbering(paragraph, numbering) => reader.numbering(*paragraph)? == *numbering,
        Staged::Set(paragraph, property) => {
            reader.property(*paragraph, property.kind())? == *property
        },
    };
    if holds {
        Ok(())
    } else {
        Err(Error::Verification)
    }
}

/// The cell's table: its model, data store and geometry.
struct Table {
    model: u64,
    store: Vec<u8>,
    geometry: topology::Geometry,
}

fn table(working: &Working<'_>, sheet: usize, table: usize) -> Result<Table, Error> {
    let target = topology::locate(working, sheet, table).map_err(Error::from)?;
    let (_, model) = topology::message(working, target.model, &[6_001]).map_err(Error::from)?;
    let geometry = topology::geometry(model).map_err(Error::from)?;
    let store = single_bytes(model, 4)?
        .ok_or(Error::InvalidSource)?
        .to_vec();
    Ok(Table {
        model: target.model,
        store,
        geometry,
    })
}

/// What the cell's text is stored as.
enum Text {
    /// Plain text (an empty cell reads as the empty string) and its string
    /// key.
    Plain(String, Option<u32>),
    /// A rich-text storage exclusive to the cell.
    Rich(u64),
}

/// The entry of `key` in a table data list (root or segments).
fn list_entry(working: &Working<'_>, list: u64, key: u32) -> Result<Vec<u8>, Error> {
    let (_, data) = topology::message(working, list, &[LIST_MESSAGE_TYPE]).map_err(Error::from)?;
    let find = |data: &[u8]| -> Result<Option<Vec<u8>>, Error> {
        for entry in repeated_bytes(data, 3)? {
            if single_varint(entry, 1)? == Some(u64::from(key)) {
                return Ok(Some(entry.to_vec()));
            }
        }
        Ok(None)
    };
    if let Some(found) = find(data)? {
        return Ok(found);
    }
    for reference in repeated_bytes(data, 4)? {
        let segment = single_varint(reference, 1)?.ok_or(Error::InvalidSource)?;
        let (_, segment_data) = topology::message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])
            .map_err(Error::from)?;
        if let Some(found) = find(segment_data)? {
            return Ok(found);
        }
    }
    Err(Error::InvalidSource)
}

/// How the cell at `position` stores its text.
fn cell_text(working: &Working<'_>, table: &Table, position: At) -> Result<Text, Error> {
    let Some(located) = cells::locate_cell(working, &table.store, position).map_err(from_cells)?
    else {
        return Ok(Text::Plain(String::new(), None));
    };
    let cell = litchi_numbers_wire::BncCell::parse(
        &located.bytes,
        &crate::cell::wire::limits_for_source(&located.bytes),
    )
    .map_err(|_error| Error::InvalidSource)?;
    match cell.stored_value() {
        litchi_numbers_wire::StoredValue::Empty => Ok(Text::Plain(String::new(), None)),
        litchi_numbers_wire::StoredValue::Text(key) => {
            let list =
                single_reference(&table.store, STRING_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
            let entry = list_entry(working, list, key)?;
            let text = single_bytes(&entry, 3)?.ok_or(Error::InvalidSource)?;
            let text = std::str::from_utf8(text).map_err(|_error| Error::InvalidSource)?;
            Ok(Text::Plain(text.to_owned(), Some(key)))
        },
        litchi_numbers_wire::StoredValue::RichText(key) => {
            let list = single_reference(&table.store, RICH_TEXT_LIST_FIELD)?
                .ok_or(Error::InvalidSource)?;
            let entry = list_entry(working, list, key)?;
            let payload = single_reference(&entry, 9)?.ok_or(Error::InvalidSource)?;
            // Only a payload this cell alone names is rewritten in place;
            // the host copies a shared one first.
            if single_varint(&entry, 2)?.unwrap_or(0) != 1
                || referenced_elsewhere(working, payload, list)?
            {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let (_, data) = topology::message(working, payload, &[RICH_TEXT_PAYLOAD_MESSAGE_TYPE])
                .map_err(Error::from)?;
            let storage = single_reference(data, 1)?.ok_or(Error::InvalidSource)?;
            Ok(Text::Rich(storage))
        },
        _ => Err(Error::Refused(Refusal::NotText)),
    }
}

/// Whether an object other than `owner` and `identifier` itself references
/// `identifier`.
fn referenced_elsewhere(working: &Working<'_>, identifier: u64, owner: u64) -> Result<bool, Error> {
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            let id = object.archive_info.identifier.unwrap_or_default();
            if id == owner || id == identifier {
                continue;
            }
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&identifier))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// The UTF-16 paragraph starts of plain text.
fn paragraph_starts(text: &str) -> Result<Vec<u32>, Error> {
    let mut starts = vec![0];
    let mut index = 0_u32;
    for character in text.chars() {
        index = index
            .checked_add(character.len_utf16() as u32)
            .ok_or(Error::InvalidSource)?;
        if character == '\n' {
            starts.push(index);
        }
    }
    Ok(starts)
}

fn require_start(text: &str, paragraph: TextPosition) -> Result<(), Error> {
    if paragraph_starts(text)?
        .binary_search(&paragraph.utf16_index())
        .is_ok()
    {
        Ok(())
    } else {
        Err(Error::Refused(Refusal::NotAParagraphStart))
    }
}

/// Reads of one cell's lists.
struct Reader<'p> {
    working: Working<'p>,
    text: Text,
    limits: WireLimits,
}

impl<'p> Reader<'p> {
    fn new(
        package: &'p Package,
        sheet: usize,
        table_index: usize,
        position: At,
    ) -> Result<Self, Error> {
        let working = Working::new(package)?;
        let table = table(&working, sheet, table_index)?;
        if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        let text = cell_text(&working, &table, position)?;
        let limits = cells::style_wire_limits(&working).map_err(from_cells)?;
        Ok(Self {
            working,
            text,
            limits,
        })
    }

    fn storage(&self) -> Result<Vec<u8>, Error> {
        let Text::Rich(storage) = self.text else {
            return Err(Error::InvalidSource);
        };
        Ok(storage_payload(&self.working, storage)?)
    }

    fn list(&self) -> Result<ParagraphList, Error> {
        if matches!(self.text, Text::Plain(..)) {
            return Ok(ParagraphList::None);
        }
        let storage = self.storage()?;
        let style = list::uniform_style(&storage, self.limits).map_err(|error| match error {
            list::Error::InvalidSource { .. } => Error::Refused(Refusal::NotUniform),
            other => from_list(other),
        })?;
        list::preset(&ListStyles(&self.working), style, self.limits).map_err(from_list)
    }

    fn lists(&self) -> Result<Vec<ParagraphListPlacement>, Error> {
        if matches!(self.text, Text::Plain(..)) {
            return Ok(vec![ParagraphListPlacement::new(
                TextPosition::ZERO,
                ParagraphList::None,
            )]);
        }
        list::placements(&self.storage()?, &ListStyles(&self.working), self.limits)
            .map_err(from_list)
    }

    fn levels(&self) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        if matches!(self.text, Text::Plain(..)) {
            return Ok(vec![ParagraphListLevelPlacement::new(
                TextPosition::ZERO,
                ParagraphListLevel::ZERO,
            )]);
        }
        list::levels(&self.storage()?, self.limits).map_err(from_list)
    }

    fn level(&self, paragraph: TextPosition) -> Result<ParagraphListLevel, Error> {
        if let Text::Plain(text, _) = &self.text {
            require_start(text, paragraph)?;
            return Ok(ParagraphListLevel::ZERO);
        }
        list::level(&self.storage()?, paragraph, self.limits).map_err(from_list)
    }

    fn numbering(&self, paragraph: TextPosition) -> Result<ParagraphListNumbering, Error> {
        if let Text::Plain(text, _) = &self.text {
            require_start(text, paragraph)?;
            return Ok(ParagraphListNumbering::Continue);
        }
        list::numbering(&self.storage()?, paragraph, self.limits).map_err(from_list)
    }

    fn property(&self, paragraph: TextPosition, kind: PropertyKind) -> Result<Property, Error> {
        if let Text::Plain(text, _) = &self.text {
            require_start(text, paragraph)?;
            return Err(Error::Refused(Refusal::NotApplicable));
        }
        list::property(
            &self.storage()?,
            &ListStyles(&self.working),
            paragraph,
            kind,
            self.limits,
        )
        .map_err(from_list)
    }
}

/// A new single-message object at the promotion's message version.
fn new_object(
    working: &Working<'_>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> Result<litchi_iwa_core::ArchiveObject, Error> {
    let mut object = structure::new_object(
        identifier,
        None,
        message_type,
        data,
        references,
        working.archive_limits(),
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec().into();
    Ok(object)
}

/// Append `reference` to the first message's references of `identifier`.
fn add_reference(
    working: &mut Working<'_>,
    identifier: u64,
    index: usize,
    reference: u64,
) -> Result<(), Error> {
    let info = working
        .object_mut(identifier)?
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    if !info.object_references.contains(&reference) {
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(reference);
    }
    Ok(())
}

/// The next key of a table data list: past `next_list_id` and every key
/// it holds (the host's `next_table_data_list_key`).
fn next_key(working: &Working<'_>, list: &[u8]) -> Result<u32, Error> {
    let mut next = single_varint(list, 2)?.unwrap_or(0).max(1);
    let mut visit = |data: &[u8]| -> Result<(), Error> {
        for entry in repeated_bytes(data, 3)? {
            if let Some(key) = single_varint(entry, 1)? {
                next = next.max(key.checked_add(1).ok_or(Error::InvalidSource)?);
            }
        }
        Ok(())
    };
    visit(list)?;
    for reference in repeated_bytes(list, 4)? {
        let segment = single_varint(reference, 1)?.ok_or(Error::InvalidSource)?;
        let (_, data) = topology::message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])
            .map_err(Error::from)?;
        visit(data)?;
    }
    u32::try_from(next).map_err(|_error| Error::Refused(Refusal::UnsupportedSource))
}

/// Point the cell at `position` at rich-text entry `key`, storing it when
/// it is empty.
fn write_rich_text(
    working: &mut Working<'_>,
    table: &Table,
    position: At,
    key: u32,
) -> Result<(), Error> {
    let Some(located) = cells::locate_cell(working, &table.store, position).map_err(from_cells)?
    else {
        let mut cell = litchi_numbers_wire::BncCell::minimal();
        cell.set_rich_text(key);
        return cells::materialize_cell(
            working,
            &table.store,
            table.geometry.columns,
            position,
            &cell,
        )
        .map_err(from_cells);
    };
    let limits = crate::cell::wire::limits_for_source(&located.bytes);
    let mut cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &limits)
        .map_err(|_error| Error::InvalidSource)?;
    cell.set_rich_text(key);
    let encoded = crate::cell::wire::encode(&cell, &limits).map_err(|_| Error::InvalidSource)?;
    let (_, tile_data) =
        topology::message(working, located.tile, &[TILE_MESSAGE_TYPE]).map_err(Error::from)?;
    let mut field_index = 0_usize;
    let rewritten = rebuild(tile_data, |field| {
        let this = field_index;
        field_index += 1;
        if this != located.row_field {
            return Ok(None);
        }
        let row = storage::replace_cell(field.payload, position.column(), &encoded)
            .map_err(Fault::from)?
            .ok_or(Fault::Invalid)?;
        bytes_field(5, &row).map(Some)
    })?;
    cells::replace_message(working, located.tile, located.tile_index, rewritten)?;
    Ok(())
}

/// A `TSP.Reference` to `identifier`.
fn reference(identifier: u64) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    push_varint_field(&mut output, 1, identifier)?;
    Ok(output)
}

/// The host's new cell storage (`cell_storage`), encoded in field-number
/// order: `kind` CELL, the stylesheet, the text, `in_document`, one
/// paragraph style and one list style at zero, a paragraph-data,
/// paragraph-start and bidi entry per paragraph, and an empty drop-cap
/// style at zero.
fn cell_storage(
    text: &str,
    stylesheet: u64,
    paragraph_style: u64,
    list_style: u64,
) -> Result<Vec<u8>, Error> {
    let object_table = |identifier: Option<u64>| -> Result<Vec<u8>, Error> {
        let mut attribute = Vec::new();
        push_varint_field(&mut attribute, 1, 0)?;
        if let Some(identifier) = identifier {
            push_bytes_field(&mut attribute, 2, &reference(identifier)?)?;
        }
        let mut table = Vec::new();
        push_bytes_field(&mut table, 1, &attribute)?;
        Ok(table)
    };
    let mut para_data = Vec::new();
    for start in paragraph_starts(text)? {
        let mut attribute = Vec::new();
        push_varint_field(&mut attribute, 1, u64::from(start))?;
        push_varint_field(&mut attribute, 2, 0)?;
        push_varint_field(&mut attribute, 3, 0)?;
        push_bytes_field(&mut para_data, 1, &attribute)?;
    }
    let mut storage = Vec::new();
    push_varint_field(&mut storage, 1, 5)?;
    push_bytes_field(&mut storage, 2, &reference(stylesheet)?)?;
    push_bytes_field(&mut storage, 3, text.as_bytes())?;
    push_bytes_field(&mut storage, 5, &object_table(Some(paragraph_style))?)?;
    push_bytes_field(&mut storage, 6, &para_data)?;
    push_bytes_field(&mut storage, 7, &object_table(Some(list_style))?)?;
    push_varint_field(&mut storage, 10, 1)?;
    push_bytes_field(&mut storage, 14, &para_data)?;
    push_bytes_field(&mut storage, 24, &para_data)?;
    push_bytes_field(&mut storage, 28, &object_table(None)?)?;
    Ok(storage)
}

/// The new storage's references, sorted and deduplicated as the host's
/// `storage_object_references` lists them.
fn storage_references(stylesheet: u64, paragraph_style: u64, list_style: u64) -> Vec<u64> {
    let mut references = vec![stylesheet, paragraph_style, list_style];
    references.sort_unstable();
    references.dedup();
    references
}

/// Promote the plain cell at `position` to a rich-text storage holding
/// `text` (the host editor's `promote_plain_cell`); returns the storage.
#[allow(
    clippy::too_many_lines,
    clippy::too_many_arguments,
    reason = "one promotion: rich-text list, storage, payload, entry, cell, string release"
)]
fn promote(
    working: &mut Working<'_>,
    registry: &metadata::Registry,
    transition: &mut Transition,
    sheet: usize,
    table_index: usize,
    table: &Table,
    position: At,
    text: &str,
    string_key: Option<u32>,
    limits: WireLimits,
) -> Result<u64, Error> {
    let paragraph_style =
        cells::current_style(working, sheet, table_index, position).map_err(from_cells)?;
    let stylesheet = cells::stylesheet_of(working, paragraph_style).map_err(from_cells)?;
    let sheet_entry = working.locate(stylesheet)?.ok_or(Error::InvalidSource)?;
    let none_preset = {
        let archive = working.archive(&sheet_entry)?;
        let candidates = archive.objects.iter().filter_map(|object| {
            let identifier = object.archive_info.identifier?;
            object
                .messages
                .iter()
                .find(|message| message.type_ == list::MESSAGE_TYPE)
                .map(|message| (identifier, message.data.as_slice()))
        });
        list::find_preset(candidates, stylesheet, ParagraphList::None, limits)
            .map_err(from_list)?
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?
    };
    let model_entry = working.locate(table.model)?.ok_or(Error::InvalidSource)?;
    let existing = single_reference(&table.store, RICH_TEXT_LIST_FIELD)?;
    let identifiers = working.allocate(if existing.is_some() { 2 } else { 3 })?;
    let (rich_list, storage_id, payload_id) = match (existing, identifiers.as_slice()) {
        (Some(list), [storage, payload]) => (list, *storage, *payload),
        (None, [list, storage, payload]) => (*list, *storage, *payload),
        _ => return Err(Error::InvalidSource),
    };
    if existing.is_none() {
        // `TST.TableDataList { list_type: RICH_TEXT_PAYLOAD, next_list_id: 1,
        // is_new_for_bnc: true }`.
        let mut list = Vec::new();
        push_varint_field(&mut list, 1, RICH_TEXT_LIST_TYPE)?;
        push_varint_field(&mut list, 2, 1)?;
        push_varint_field(&mut list, 5, 1)?;
        let object = new_object(working, rich_list, LIST_MESSAGE_TYPE, list, Vec::new())?;
        working.insert(&model_entry, object)?;
        let model_object = working.object(table.model)?;
        let index = structure::unique_message(model_object, &[6_001])?;
        let reference = reference(rich_list)?;
        let data = litchi_iwa_common::wire::patch_nested_length_delimited_field(
            &model_object.messages[index].data,
            &[4, RICH_TEXT_LIST_FIELD],
            false,
            Some(&reference),
        )
        .map_err(|error| structure::map_wire(&error))?;
        cells::replace_message(working, table.model, index, data)?;
        add_reference(working, table.model, index, rich_list)?;
    }

    // The storage: the cell's text, one paragraph style and the `None`
    // list for every paragraph.
    let storage_archive = cell_storage(text, stylesheet, paragraph_style, none_preset)?;
    let storage_object = new_object(
        working,
        storage_id,
        storage_list::STORAGE_MESSAGE_TYPE,
        storage_archive,
        storage_references(stylesheet, paragraph_style, none_preset),
    )?;
    let row = position.row();
    let column = position.column();
    let packed_data = if row <= u32::from(u16::MAX) && column <= u32::from(u16::MAX) {
        column << 16 | row
    } else {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    };
    // `TST.RichTextPayloadArchive { storage, cellid { packedData } }`.
    let mut payload = Vec::new();
    push_bytes_field(&mut payload, 1, &reference(storage_id)?)?;
    let mut cell_id = vec![0x0d];
    cell_id.extend_from_slice(&packed_data.to_le_bytes());
    push_bytes_field(&mut payload, 3, &cell_id)?;
    let payload_object = new_object(
        working,
        payload_id,
        RICH_TEXT_PAYLOAD_MESSAGE_TYPE,
        payload,
        vec![storage_id],
    )?;
    // The payload and its storage live beside the rich-text list (native
    // Numbers keeps it in its own `DataList` component); their styles are
    // referenced across components.
    let list_entry_name = working.locate(rich_list)?.ok_or(Error::InvalidSource)?;
    working.insert(&list_entry_name, storage_object)?;
    working.insert(&list_entry_name, payload_object)?;
    if let Ok(owner) = registry.component_for_entry(&list_entry_name) {
        let owner = owner.identifier;
        for style in [stylesheet, paragraph_style, none_preset] {
            let entry = working.locate(style)?.ok_or(Error::InvalidSource)?;
            let target = registry.component_for_entry(&entry)?.identifier;
            transition.reference(registry, owner, target, style)?;
        }
    }

    // The entry naming the payload.
    let (list_index, list_data) =
        topology::message(working, rich_list, &[LIST_MESSAGE_TYPE]).map_err(Error::from)?;
    let list_data = list_data.to_vec();
    let key = next_key(working, &list_data)?;
    let mut rebuilt = rebuild(&list_data, |field| {
        if field.number == 2 {
            return varint_field(2, u64::from(key) + 1).map(Some);
        }
        Ok(None)
    })?;
    if single_varint(&rebuilt, 2)?.is_none() {
        return Err(Error::InvalidSource);
    }
    let mut entry = Vec::new();
    push_varint_field(&mut entry, 1, u64::from(key))?;
    push_varint_field(&mut entry, 2, 1)?;
    let mut reference = Vec::new();
    push_varint_field(&mut reference, 1, payload_id)?;
    push_bytes_field(&mut entry, 9, &reference)?;
    push_bytes_field(&mut rebuilt, 3, &entry)?;
    cells::replace_message(working, rich_list, list_index, rebuilt)?;
    add_reference(working, rich_list, list_index, payload_id)?;

    write_rich_text(working, table, position, key)?;
    if let Some(old) = string_key {
        let list =
            single_reference(&table.store, STRING_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
        let (index, data) =
            topology::message(working, list, &[LIST_MESSAGE_TYPE]).map_err(Error::from)?;
        let released = litchi_numbers_wire::table_write::list_release(
            data,
            old,
            &litchi_numbers_wire::Limits::DEFAULT,
        )
        .map_err(|error| match error {
            litchi_numbers_wire::table_write::WriteError::Unsupported => {
                Error::Refused(Refusal::UnsupportedDependency)
            },
            _ => Error::InvalidSource,
        })?;
        cells::replace_message(working, list, index, released)?;
    }
    Ok(storage_id)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    change: &Staged,
) -> Result<Option<Published>, Error> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    let limits = cells::style_wire_limits(&working).map_err(from_cells)?;
    let table = table(&working, sheet, table_index)?;
    if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let text = cell_text(&working, &table, position)?;
    // A plain cell keeps unlisted paragraphs without promotion.
    if let Text::Plain(plain, _) = &text {
        match change {
            Staged::List(ParagraphList::None) => return Ok(None),
            Staged::Level(paragraph, ParagraphListLevel::ZERO)
            | Staged::Numbering(paragraph, ParagraphListNumbering::Continue)
            | Staged::Reset(paragraph, _) => {
                require_start(plain, *paragraph)?;
                return Ok(None);
            },
            _ => {},
        }
    }
    // One preset for the whole cell needs paragraphs that share one list
    // (the host reads the cell's list first).
    if let (Text::Rich(storage), Staged::List(_)) = (&text, change) {
        let payload = storage_payload(&working, *storage)?;
        list::uniform_style(&payload, limits).map_err(|error| match error {
            list::Error::InvalidSource { .. } => Error::Refused(Refusal::NotUniform),
            other => from_list(other),
        })?;
    }
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let mut transition = Transition::default();
    let (storage, promoted) = match text {
        Text::Rich(storage) => (storage, false),
        Text::Plain(plain, key) => (
            promote(
                &mut working,
                &registry,
                &mut transition,
                sheet,
                table_index,
                &table,
                position,
                &plain,
                key,
                limits,
            )?,
            true,
        ),
    };
    let mut workspace = ListWorkspace {
        working: &mut working,
        registry: &registry,
        transition: &mut transition,
    };
    let mut publisher = Publisher::new(&mut workspace, limits);
    let changed = match change {
        Staged::List(preset) => publisher.set_list(storage, *preset)?,
        Staged::Lists(placements) => publisher.set_lists(storage, placements)?,
        Staged::Level(paragraph, level) => publisher.set_level(storage, *paragraph, *level)?,
        Staged::Numbering(paragraph, numbering) => {
            publisher.set_numbering(storage, *paragraph, *numbering)?
        },
        Staged::Set(paragraph, property) => {
            publisher.property(storage, *paragraph, &Change::Set(property.clone()))?
        },
        Staged::Reset(paragraph, kind) => {
            publisher.property(storage, *paragraph, &Change::Reset(*kind))?
        },
    };
    if !changed && !promoted {
        return Ok(None);
    }
    Ok(Some(working.publish(&transition)?))
}

impl Package {
    /// The list preset every paragraph of one cell's text shares
    /// ([`ParagraphList::None`] for plain text).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], [`Refusal::NotUniform`]
    /// when the paragraphs differ, [`Refusal::NotText`], source errors.
    pub fn table_cell_list<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<ParagraphList, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.list()
    }

    /// The list preset of each run of paragraphs of one cell's text.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list`], without [`Refusal::NotUniform`].
    pub fn table_cell_lists<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Vec<ParagraphListPlacement>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.lists()
    }

    /// The list level of each run of paragraphs of one cell's text.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_lists`].
    pub fn table_cell_list_levels<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.levels()
    }

    /// The list level of the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_lists`], [`Refusal::NotAParagraphStart`].
    pub fn table_cell_list_level<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.level(paragraph)
    }

    /// Whether the paragraph starting at `paragraph` continues or restarts
    /// its list's numbering.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list_level`].
    pub fn table_cell_list_numbering<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.numbering(paragraph)
    }

    /// The effective list property `kind` of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list_level`], [`Refusal::NotApplicable`] when
    /// the paragraph's list has no such property.
    pub fn table_cell_list_property<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Reader::new(self, sheet, table, position)?.property(paragraph, kind)
    }

    /// Start an immutable cell-list edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::list::ParagraphList;
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package
    ///     .edit_table_cell_lists(0, 0)?
    ///     .set_list(At::new(1, 1), ParagraphList::Bullet)?
    ///     .commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_lists<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        let leaf = self.workbook_ref().sheets()[sheet]
            .tables()
            .nth(table)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            rows: leaf.row_count(),
            columns: leaf.column_count(),
            changes: Vec::new(),
        })
    }

    /// Apply a cell-list patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_lists(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_cell_list_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize), Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table).map_err(Error::from)?;
        Ok((sheet, table))
    }
}
