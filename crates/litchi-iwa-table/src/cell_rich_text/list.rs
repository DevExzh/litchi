//! Paragraph lists of the text in one cell (lifted from `litchi-numbers`,
//! record 0370b).
//!
//! A cell's text carries lists only as rich text: a `TST.RichTextPayload`
//! in its table's rich-text list names a `TSWP.StorageArchive` whose
//! list-style table gives each paragraph its list. A plain cell (empty, or a
//! string key) reads as unlisted paragraphs; the first change that needs a
//! list promotes it the way the host editor's `cell_paragraph_list` does:
//!
//! - the table gains a rich-text list when it has none;
//! - a cell storage holding the cell's text is stored beside the rich-text
//!   list, its paragraphs styled by the cell's text style and the
//!   stylesheet's `None` list preset, with a rich-text payload naming it;
//! - a rich-text entry names the payload, the cell becomes a rich-text cell
//!   pointing at it, and the cell's string entry is released.
//!
//! The list change itself runs W1-D's engine
//! ([`litchi_iwa_text_wire::list`], record 0374) and is published by
//! [`super::storage_list`].

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::table::coordinate::At;
pub use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
    ParagraphListPlacement,
};
pub use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Change};
pub use litchi_iwa_text_wire::list::{Property, PropertyKind};

use super::storage_list::{Failure, ListStyles, Publisher};
use super::{
    Error, LIST_MESSAGE_TYPE, LIST_SEGMENT_MESSAGE_TYPE, Refusal, TABLE_MODEL_MESSAGE_TYPE, Table,
    cell_limits, current_style, locate_cell, materialize_cell, message, replace_message,
    rewrite_cell, style_wire_limits, stylesheet_of, table,
};
use crate::structure::metadata::{self, Transition};
use crate::structure::wire::{
    push_bytes_field, push_varint_field, rebuild, repeated_bytes, single_bytes, single_reference,
    single_varint, varint_field,
};
use crate::structure::{self, Published, TableLeaf, Working};

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

/// The list engine's error as an engine error.
#[must_use]
pub fn from_list(error: list::Error) -> Error {
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
            kind: super::LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        list::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl From<Failure> for Error {
    fn from(failure: Failure) -> Self {
        match failure {
            Failure::Fault(fault) => fault.into(),
            Failure::List(error) => from_list(error),
        }
    }
}

/// One staged list operation of one cell.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Operation {
    /// Every paragraph takes one preset ([`ParagraphList::None`] removes
    /// the lists).
    List(ParagraphList),
    /// Each placement's preset runs from its paragraph to the next one.
    Lists(Vec<ParagraphListPlacement>),
    /// The list level of the paragraph starting at the position.
    Level(TextPosition, ParagraphListLevel),
    /// Whether the paragraph starting at the position continues or
    /// restarts its numbering.
    Numbering(TextPosition, ParagraphListNumbering),
    /// One list property of the paragraph starting at the position.
    Set(TextPosition, Property),
    /// Restore the list's own value of one property of the paragraph.
    Reset(TextPosition, PropertyKind),
}

/// Whether the change `change` reads back on `package`.
/// The published change reads back.
///
/// # Errors
///
/// [`Error::Verification`] when it does not; the reader's errors.
pub fn verify<P: TableLeaf>(
    package: &P,
    model: u64,
    position: At,
    change: &Operation,
) -> Result<(), Error> {
    let reader = Reader::new(Working::new(package)?, model, position)?;
    let holds = match change {
        Operation::List(preset) => reader.list()? == *preset,
        Operation::Lists(_) | Operation::Reset(..) => true,
        Operation::Level(paragraph, level) => reader.level(*paragraph)? == *level,
        Operation::Numbering(paragraph, numbering) => reader.numbering(*paragraph)? == *numbering,
        Operation::Set(paragraph, property) => {
            reader.property(*paragraph, property.kind())? == *property
        },
    };
    if holds {
        Ok(())
    } else {
        Err(Error::Verification)
    }
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
fn list_entry<P: TableLeaf>(
    working: &Working<'_, P>,
    list: u64,
    key: u32,
) -> Result<Vec<u8>, Error> {
    let (_, data) = message(working, list, &[LIST_MESSAGE_TYPE])?;
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
        let (_, segment_data) = message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])?;
        if let Some(found) = find(segment_data)? {
            return Ok(found);
        }
    }
    Err(Error::InvalidSource)
}

/// How the cell at `position` stores its text.
fn cell_text<P: TableLeaf>(
    working: &Working<'_, P>,
    table: &Table,
    position: At,
) -> Result<Text, Error> {
    let Some(located) = locate_cell(working, &table.store, position)? else {
        return Ok(Text::Plain(String::new(), None));
    };
    let cell = litchi_numbers_wire::BncCell::parse(&located.bytes, &cell_limits(&located.bytes))
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
            let (_, data) = message(working, payload, &[RICH_TEXT_PAYLOAD_MESSAGE_TYPE])?;
            let storage = single_reference(data, 1)?.ok_or(Error::InvalidSource)?;
            Ok(Text::Rich(storage))
        },
        _ => Err(Error::Refused(Refusal::NotText)),
    }
}

/// Whether an object other than `owner` and `identifier` itself references
/// `identifier`.
fn referenced_elsewhere<P: TableLeaf>(
    working: &Working<'_, P>,
    identifier: u64,
    owner: u64,
) -> Result<bool, Error> {
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
pub struct Reader<'p, P> {
    working: Working<'p, P>,
    text: Text,
    limits: WireLimits,
}

impl<'p, P: TableLeaf> Reader<'p, P> {
    /// The reader of the cell at `position` of the table whose model is
    /// `model`, over a fresh working set.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table; [`Refusal::NotText`]
    /// for a cell holding another value; [`Refusal::UnsupportedDependency`]
    /// for a shared rich-text payload; source errors.
    pub fn new(working: Working<'p, P>, model: u64, position: At) -> Result<Self, Error> {
        let table = table(&working, model)?;
        if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        let text = cell_text(&working, &table, position)?;
        let limits = style_wire_limits(&working)?;
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
        Ok(super::storage_list::storage(&self.working, storage)?.payload)
    }

    /// The list preset every paragraph shares ([`ParagraphList::None`] for
    /// plain text).
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn list(&self) -> Result<ParagraphList, Error> {
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

    /// The list preset of each run of paragraphs.
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn lists(&self) -> Result<Vec<ParagraphListPlacement>, Error> {
        if matches!(self.text, Text::Plain(..)) {
            return Ok(vec![ParagraphListPlacement::new(
                TextPosition::ZERO,
                ParagraphList::None,
            )]);
        }
        list::placements(&self.storage()?, &ListStyles(&self.working), self.limits)
            .map_err(from_list)
    }

    /// The list level of each run of paragraphs.
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn levels(&self) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        if matches!(self.text, Text::Plain(..)) {
            return Ok(vec![ParagraphListLevelPlacement::new(
                TextPosition::ZERO,
                ParagraphListLevel::ZERO,
            )]);
        }
        list::levels(&self.storage()?, self.limits).map_err(from_list)
    }

    /// The list level of the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn level(&self, paragraph: TextPosition) -> Result<ParagraphListLevel, Error> {
        if let Text::Plain(text, _) = &self.text {
            require_start(text, paragraph)?;
            return Ok(ParagraphListLevel::ZERO);
        }
        list::level(&self.storage()?, paragraph, self.limits).map_err(from_list)
    }

    /// Whether the paragraph starting at `paragraph` continues or restarts
    /// its list's numbering.
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn numbering(&self, paragraph: TextPosition) -> Result<ParagraphListNumbering, Error> {
        if let Text::Plain(text, _) = &self.text {
            require_start(text, paragraph)?;
            return Ok(ParagraphListNumbering::Continue);
        }
        list::numbering(&self.storage()?, paragraph, self.limits).map_err(from_list)
    }

    /// The effective list property `kind` of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// The list refusals, source and wire errors.
    pub fn property(&self, paragraph: TextPosition, kind: PropertyKind) -> Result<Property, Error> {
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
fn new_object<P: TableLeaf>(
    working: &Working<'_, P>,
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
fn add_reference<P: TableLeaf>(
    working: &mut Working<'_, P>,
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
fn next_key<P: TableLeaf>(working: &Working<'_, P>, list: &[u8]) -> Result<u32, Error> {
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
        let (_, data) = message(working, segment, &[LIST_SEGMENT_MESSAGE_TYPE])?;
        visit(data)?;
    }
    u32::try_from(next).map_err(|_error| Error::Refused(Refusal::UnsupportedSource))
}

/// Point the cell at `position` at rich-text entry `key`, storing it when
/// it is empty.
fn write_rich_text<P: TableLeaf>(
    working: &mut Working<'_, P>,
    table: &Table,
    position: At,
    key: u32,
) -> Result<(), Error> {
    let Some(located) = locate_cell(working, &table.store, position)? else {
        let mut cell = litchi_numbers_wire::BncCell::minimal();
        cell.set_rich_text(key);
        return materialize_cell(
            working,
            &table.store,
            table.geometry.columns,
            position,
            &cell,
        );
    };
    rewrite_cell(working, &located, position, |cell| cell.set_rich_text(key))
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
fn promote<P: TableLeaf>(
    working: &mut Working<'_, P>,
    registry: &metadata::Registry,
    transition: &mut Transition,
    table: &Table,
    position: At,
    text: &str,
    string_key: Option<u32>,
    limits: WireLimits,
) -> Result<u64, Error> {
    let paragraph_style = current_style(working, table.model, position)?;
    let stylesheet = stylesheet_of(working, paragraph_style)?;
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
        let index = structure::unique_message(model_object, &[TABLE_MODEL_MESSAGE_TYPE])?;
        let reference = reference(rich_list)?;
        let data = litchi_iwa_common::wire::patch_nested_length_delimited_field(
            &model_object.messages[index].data,
            &[4, RICH_TEXT_LIST_FIELD],
            false,
            Some(&reference),
        )
        .map_err(|error| structure::map_wire(&error))?;
        replace_message(working, table.model, index, data)?;
        add_reference(working, table.model, index, rich_list)?;
    }

    // The storage: the cell's text, one paragraph style and the `None`
    // list for every paragraph.
    let storage_archive = cell_storage(text, stylesheet, paragraph_style, none_preset)?;
    let storage_object = new_object(
        working,
        storage_id,
        super::storage_list::STORAGE_MESSAGE_TYPE,
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
    let (list_index, list_data) = message(working, rich_list, &[LIST_MESSAGE_TYPE])?;
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
    replace_message(working, rich_list, list_index, rebuilt)?;
    add_reference(working, rich_list, list_index, payload_id)?;

    write_rich_text(working, table, position, key)?;
    if let Some(old) = string_key {
        let list =
            single_reference(&table.store, STRING_LIST_FIELD)?.ok_or(Error::InvalidSource)?;
        let (index, data) = message(working, list, &[LIST_MESSAGE_TYPE])?;
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
        replace_message(working, list, index, released)?;
    }
    Ok(storage_id)
}

/// Apply one staged operation to the cell at `position` of the table whose
/// model is `model`, publishing the working set. `None` is an exact no-op.
///
/// # Errors
///
/// The list refusals; [`Refusal::UnsupportedSource`] without a package
/// metadata registry; source, resource and candidate errors.
pub fn apply<P: TableLeaf>(
    mut working: Working<'_, P>,
    model: u64,
    position: At,
    change: &Operation,
) -> Result<Option<Published<P>>, Error> {
    working.keep_previews();
    let limits = style_wire_limits(&working)?;
    let table = table(&working, model)?;
    if position.row() >= table.geometry.rows || position.column() >= table.geometry.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let text = cell_text(&working, &table, position)?;
    // A plain cell keeps unlisted paragraphs without promotion.
    if let Text::Plain(plain, _) = &text {
        match change {
            Operation::List(ParagraphList::None) => return Ok(None),
            Operation::Level(paragraph, ParagraphListLevel::ZERO)
            | Operation::Numbering(paragraph, ParagraphListNumbering::Continue)
            | Operation::Reset(paragraph, _) => {
                require_start(plain, *paragraph)?;
                return Ok(None);
            },
            _ => {},
        }
    }
    // One preset for the whole cell needs paragraphs that share one list
    // (the host reads the cell's list first).
    if let (Text::Rich(storage), Operation::List(_)) = (&text, change) {
        let payload = super::storage_list::storage(&working, *storage)?.payload;
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
                &table,
                position,
                &plain,
                key,
                limits,
            )?,
            true,
        ),
    };
    let mut publisher = Publisher {
        working: &mut working,
        registry: &registry,
        transition: &mut transition,
        limits,
    };
    let changed = match change {
        Operation::List(preset) => publisher.set_list(storage, *preset)?,
        Operation::Lists(placements) => publisher.set_lists(storage, placements)?,
        Operation::Level(paragraph, level) => publisher.set_level(storage, *paragraph, *level)?,
        Operation::Numbering(paragraph, numbering) => {
            publisher.set_numbering(storage, *paragraph, *numbering)?
        },
        Operation::Set(paragraph, property) => {
            publisher.property(storage, *paragraph, &Change::Set(property.clone()))?
        },
        Operation::Reset(paragraph, kind) => {
            publisher.property(storage, *paragraph, &Change::Reset(*kind))?
        },
    };
    if !changed && !promoted {
        return Ok(None);
    }
    Ok(Some(working.publish(&transition)?))
}
