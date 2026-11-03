//! Storage-scoped ranged-object table engine (record 0374).
//!
//! A `TSWP.StorageArchive` marks hyperlinks and date-time fields (the
//! smart-field table, field 11), bookmarks (field 15) and highlights and
//! comments (field 23) with an `ObjectAttributeTable`: sorted boundaries at
//! UTF-16 indexes, each naming the object that covers the text up to the
//! next boundary (or none). This engine reads those ranges and adds, moves
//! or removes one object's range, exactly as the retired host's (commit `27b4840684`)
//! hyperlink, bookmark and annotation families do: boundaries keep their
//! bytes unless their object changes, neighbouring boundaries naming the
//! same object coalesce, and the table is dropped when no object is left.
//!
//! The engine rewrites the storage payload only. The caller owns the
//! object the range names (see the payload codecs in [`crate::smart_field`]), allocates its identifier and keeps the storage
//! message's object references with [`Rewrite::added`] and
//! [`Rewrite::removed`].

use litchi_iwa_common::wire::{
    patch_length_delimited_field, rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::position::{TextPosition, TextRange};
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, NativeResult, Writer};

const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const TEXT_FIELD: u32 = 3;
const OVERLAPPING_HIGHLIGHT_TABLE_FIELD: u32 = 25;

/// A storage's ranged-object table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Table {
    /// `table_smartfield` (field 11): hyperlinks and date-time fields.
    SmartField,
    /// `table_bookmark` (field 15): ranged bookmarks.
    Bookmark,
    /// `table_highlight` (field 23): highlights and text comments.
    Highlight,
}

impl Table {
    /// The storage field holding the table.
    #[must_use]
    pub const fn field(self) -> u32 {
        match self {
            Self::SmartField => 11,
            Self::Bookmark => 15,
            Self::Highlight => 23,
        }
    }
}

/// One object-bearing range of a table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Entry {
    object: u64,
    range: TextRange,
}

impl Entry {
    /// The object covering the range.
    #[must_use]
    pub const fn object(self) -> u64 {
        self.object
    }

    /// The covered UTF-16 range.
    #[must_use]
    pub const fn range(self) -> TextRange {
        self.range
    }
}

/// One change of a table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// Cover a currently unoccupied range with a new object.
    Add {
        /// The new object (allocated by the caller).
        object: u64,
        /// The range it covers.
        range: TextRange,
    },
    /// Move an object the table already names to another unoccupied range.
    Move {
        /// The object.
        object: u64,
        /// Its new range.
        range: TextRange,
    },
    /// Drop an object's range.
    Remove {
        /// The object.
        object: u64,
    },
}

/// A rewritten storage payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    payload: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
}

impl Rewrite {
    /// The storage payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Take the storage payload.
    #[must_use]
    pub fn into_payload(self) -> Vec<u8> {
        self.payload
    }

    /// The object the storage message's references gain.
    #[must_use]
    pub const fn added(&self) -> Option<u64> {
        self.added
    }

    /// The object the storage message's references lose.
    #[must_use]
    pub const fn removed(&self) -> Option<u64> {
        self.removed
    }
}

/// Why a well-formed table change is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// A range bound is past the text or splits a surrogate pair.
    NotABoundary,
    /// The range overlaps another object's range.
    Occupied,
    /// The table does not name the object exactly once.
    NotOwned,
    /// The storage carries overlapping highlights (field 25), which this
    /// engine does not rewrite.
    OverlappingHighlights,
}

/// Why a table operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The storage payload or its table is malformed.
    #[error("invalid ranged-object source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A well-formed change this engine refuses.
    #[error("ranged-object change refused: {0:?}")]
    Refused(Refusal),
    /// A configured wire resource was exhausted.
    #[error("ranged-object {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} ranged-object items or bytes")]
    Allocation {
        /// The attempted item or byte capacity.
        amount: usize,
    },
}

impl From<NativeError> for Error {
    fn from(error: NativeError) -> Self {
        match error {
            NativeError::Invalid(reason) => Self::InvalidSource { reason },
            NativeError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind,
                observed,
                maximum,
            },
            NativeError::Allocation { amount } => Self::Allocation { amount },
        }
    }
}

/// The result of a table operation.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
struct Boundary {
    index: u32,
    object: Option<u64>,
    raw: Vec<u8>,
}

/// The storage's text as UTF-16 scalar boundaries: the indexes where a
/// boundary may fall, and the total length.
struct Text {
    starts: Vec<u32>,
    len: u32,
}

impl Text {
    fn decode(storage: &[u8], budget: &mut Budget) -> NativeResult<Self> {
        let mut starts = Vec::new();
        let mut len = 0u32;
        native::fields(storage, 0, budget, |field, _| {
            if field.number == TEXT_FIELD {
                let fragment = std::str::from_utf8(field.bytes()?)
                    .map_err(|_| NativeError::Invalid("storage text is not UTF-8"))?;
                native::reserve(&mut starts, fragment.len())?;
                for character in fragment.chars() {
                    starts.push(len);
                    let width = u32::try_from(character.len_utf16()).unwrap_or(2);
                    len = len
                        .checked_add(width)
                        .ok_or(NativeError::Invalid("storage text UTF-16 length overflow"))?;
                }
            }
            Ok(())
        })?;
        Ok(Self { starts, len })
    }

    fn is_boundary(&self, index: u32) -> bool {
        index == self.len || self.starts.binary_search(&index).is_ok()
    }
}

fn table_payload<'a>(
    storage: &'a [u8],
    table: Table,
    budget: &mut Budget,
) -> NativeResult<(Option<&'a [u8]>, bool)> {
    let mut found = Vec::new();
    let mut overlapping = false;
    native::fields(storage, 0, budget, |field, _| {
        if field.number == table.field() {
            native::reserve(&mut found, 1)?;
            found.push(field.bytes()?);
        } else if field.number == OVERLAPPING_HIGHLIGHT_TABLE_FIELD {
            field.bytes()?;
            overlapping = true;
        }
        Ok(())
    })?;
    match found.as_slice() {
        [] => Ok((None, overlapping)),
        [table] => Ok((Some(table), overlapping)),
        _ => Err(NativeError::Invalid(
            "storage repeats a ranged-object table",
        )),
    }
}

fn boundaries(table: &[u8], text: &Text, budget: &mut Budget) -> NativeResult<Vec<Boundary>> {
    let mut boundaries = Vec::new();
    native::fields(table, 1, budget, |field, budget| {
        if field.number != TABLE_ENTRIES_FIELD {
            return Ok(());
        }
        let raw = field.bytes()?;
        let mut index = 0u32;
        let mut object: Option<native::Reference> = None;
        native::fields(raw, 2, budget, |inner, budget| {
            match inner.number {
                ENTRY_INDEX_FIELD => index = inner.uint32()?,
                ENTRY_OBJECT_FIELD => object
                    .get_or_insert_with(native::Reference::default)
                    .merge(inner.bytes()?, 3, budget)?,
                _ => {},
            }
            Ok(())
        })?;
        native::reserve(&mut boundaries, 1)?;
        boundaries.push(Boundary {
            index,
            object: object.map(|reference| reference.identifier),
            raw: raw.to_vec(),
        });
        Ok(())
    })?;
    if boundaries.is_empty() {
        return Err(NativeError::Invalid(
            "storage has an empty ranged-object table",
        ));
    }
    if boundaries[0].index != 0 {
        return Err(NativeError::Invalid(
            "ranged-object table must begin at UTF-16 index zero",
        ));
    }
    let mut previous = None;
    for boundary in &boundaries {
        if previous.is_some_and(|index| index >= boundary.index) {
            return Err(NativeError::Invalid(
                "ranged-object boundaries are not strictly increasing",
            ));
        }
        if boundary.object == Some(0) {
            return Err(NativeError::Invalid(
                "ranged-object table names object zero",
            ));
        }
        if boundary.object.is_some() && boundary.index == text.len {
            return Err(NativeError::Invalid(
                "ranged-object table has an empty range at the end",
            ));
        }
        if !text.is_boundary(boundary.index) {
            return Err(NativeError::Invalid(
                "ranged-object boundary is not a scalar boundary",
            ));
        }
        previous = Some(boundary.index);
    }
    Ok(boundaries)
}

/// The object-bearing ranges of `table`, in text order.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed storage or table; resource
/// errors as configured.
pub fn entries(storage: &[u8], table: Table, limits: WireLimits) -> Result<Vec<Entry>> {
    let mut budget = Budget::new(limits);
    let text = Text::decode(storage, &mut budget)?;
    let (payload, _) = table_payload(storage, table, &mut budget)?;
    let Some(payload) = payload else {
        return Ok(Vec::new());
    };
    let boundaries = boundaries(payload, &text, &mut budget)?;
    let mut entries = Vec::new();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(object) = boundary.object else {
            continue;
        };
        let end = boundaries
            .get(position + 1)
            .map_or(text.len, |next| next.index);
        let range = TextRange::new(
            TextPosition::from_utf16_code_units(boundary.index),
            TextPosition::from_utf16_code_units(end),
        )
        .map_err(|_| Error::InvalidSource {
            reason: "ranged-object table has an empty range",
        })?;
        native::reserve(&mut entries, 1)?;
        entries.push(Entry { object, range });
    }
    Ok(entries)
}

fn ensure_available(
    range: TextRange,
    boundaries: &[Boundary],
    ignored: Option<u64>,
    text: &Text,
) -> Result<()> {
    let (start, end) = (range.start().utf16_index(), range.end().utf16_index());
    if !text.is_boundary(start) || !text.is_boundary(end) || end > text.len {
        return Err(Error::Refused(Refusal::NotABoundary));
    }
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(object) = boundary.object else {
            continue;
        };
        if Some(object) == ignored {
            continue;
        }
        let occupied_end = boundaries
            .get(position + 1)
            .map_or(text.len, |next| next.index);
        if start < occupied_end && boundary.index < end {
            return Err(Error::Refused(Refusal::Occupied));
        }
    }
    Ok(())
}

fn reference_payload(identifier: u64, limits: WireLimits) -> NativeResult<Vec<u8>> {
    let mut writer = Writer::new(limits);
    native::Reference { identifier }.encode(&mut writer)?;
    Ok(writer.into_bytes())
}

fn new_boundary(index: u32, object: Option<u64>, limits: WireLimits) -> NativeResult<Boundary> {
    let mut writer = Writer::new(limits);
    writer.uint32(ENTRY_INDEX_FIELD, index)?;
    if let Some(identifier) = object {
        writer.message(ENTRY_OBJECT_FIELD, |writer| {
            native::Reference { identifier }.encode(writer)
        })?;
    }
    Ok(Boundary {
        index,
        object,
        raw: writer.into_bytes(),
    })
}

fn patch_object(
    raw: &[u8],
    current: Option<u64>,
    replacement: Option<u64>,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    let reference = replacement
        .map(|identifier| reference_payload(identifier, limits))
        .transpose()?;
    patch_length_delimited_field(
        raw,
        ENTRY_OBJECT_FIELD,
        current.is_some(),
        reference.as_deref(),
    )
    .map_err(native::common)
}

fn set_boundary(
    boundaries: &mut Vec<Boundary>,
    index: u32,
    object: Option<u64>,
    limits: WireLimits,
) -> NativeResult<()> {
    if let Some(boundary) = boundaries
        .iter_mut()
        .find(|boundary| boundary.index == index)
    {
        if boundary.object != object {
            boundary.raw = patch_object(&boundary.raw, boundary.object, object, limits)?;
            boundary.object = object;
        }
        return Ok(());
    }
    native::reserve(boundaries, 1)?;
    boundaries.push(new_boundary(index, object, limits)?);
    boundaries.sort_by_key(|boundary| boundary.index);
    Ok(())
}

fn add_range(
    boundaries: &mut Vec<Boundary>,
    range: TextRange,
    object: u64,
    limits: WireLimits,
) -> NativeResult<()> {
    if boundaries.is_empty() {
        boundaries.push(new_boundary(0, None, limits)?);
    }
    set_boundary(
        boundaries,
        range.start().utf16_index(),
        Some(object),
        limits,
    )?;
    set_boundary(boundaries, range.end().utf16_index(), None, limits)?;
    boundaries.dedup_by(|current, previous| current.object == previous.object);
    Ok(())
}

fn remove_range(boundaries: &mut Vec<Boundary>, object: u64, limits: WireLimits) -> Result<()> {
    let matches: Vec<usize> = boundaries
        .iter()
        .enumerate()
        .filter(|(_, boundary)| boundary.object == Some(object))
        .map(|(position, _)| position)
        .collect();
    let [position] = matches.as_slice() else {
        return Err(Error::Refused(Refusal::NotOwned));
    };
    let boundary = &mut boundaries[*position];
    boundary.raw = patch_object(&boundary.raw, boundary.object, None, limits)?;
    boundary.object = None;
    boundaries.dedup_by(|current, previous| current.object == previous.object);
    Ok(())
}

fn encode_table(
    original: Option<&[u8]>,
    boundaries: Vec<Boundary>,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    let entries: Vec<Vec<u8>> = boundaries
        .into_iter()
        .map(|boundary| boundary.raw)
        .collect();
    match original {
        Some(table) => {
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &entries)
                .map_err(native::common)
        },
        None => {
            let mut writer = Writer::new(limits);
            for entry in &entries {
                writer.bytes(TABLE_ENTRIES_FIELD, entry)?;
            }
            Ok(writer.into_bytes())
        },
    }
}

/// Apply one [`Change`] to `table` of the storage payload.
///
/// # Errors
///
/// [`Error::Refused`] for a range that is not on scalar boundaries or that
/// overlaps another object, an object the table does not name once, or a
/// highlight change on a storage with overlapping highlights;
/// [`Error::InvalidSource`] for a malformed storage; resource errors as
/// configured.
pub fn rewrite(
    storage: &[u8],
    table: Table,
    change: &Change,
    limits: WireLimits,
) -> Result<Rewrite> {
    let mut budget = Budget::new(limits);
    let text = Text::decode(storage, &mut budget)?;
    let (original, overlapping) = table_payload(storage, table, &mut budget)?;
    if table == Table::Highlight && overlapping && !matches!(change, Change::Remove { .. }) {
        return Err(Error::Refused(Refusal::OverlappingHighlights));
    }
    let mut current = match original {
        Some(payload) => boundaries(payload, &text, &mut budget)?,
        None => Vec::new(),
    };
    let (added, removed) = match *change {
        Change::Add { object, range } => {
            ensure_available(range, &current, None, &text)?;
            add_range(&mut current, range, object, limits)?;
            (Some(object), None)
        },
        Change::Move { object, range } => {
            ensure_available(range, &current, Some(object), &text)?;
            remove_range(&mut current, object, limits)?;
            ensure_available(range, &current, None, &text)?;
            add_range(&mut current, range, object, limits)?;
            (None, None)
        },
        Change::Remove { object } => {
            remove_range(&mut current, object, limits)?;
            (None, Some(object))
        },
    };
    let replacement = if matches!(change, Change::Remove { .. })
        && !current.iter().any(|boundary| boundary.object.is_some())
    {
        None
    } else {
        Some(encode_table(original, current, limits)?)
    };
    budget.admit_bytes(storage.len())?;
    let payload = patch_length_delimited_field(
        storage,
        table.field(),
        original.is_some(),
        replacement.as_deref(),
    )
    .map_err(native::common)?;
    native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?;
    Ok(Rewrite {
        payload,
        added,
        removed,
    })
}

#[cfg(test)]
mod tests {
    use litchi_iwa_protos::tswp;
    use prost::Message;

    use super::*;

    fn storage(text: &str) -> Vec<u8> {
        tswp::StorageArchive {
            text: vec![text.to_owned()],
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn range(start: u32, end: u32) -> TextRange {
        TextRange::new(
            TextPosition::from_utf16_code_units(start),
            TextPosition::from_utf16_code_units(end),
        )
        .unwrap()
    }

    #[test]
    fn add_move_and_remove_round_trip_through_the_generated_schema() {
        let limits = WireLimits::DEFAULT;
        let source = storage("hello world");
        let added = rewrite(
            &source,
            Table::SmartField,
            &Change::Add {
                object: 7,
                range: range(0, 5),
            },
            limits,
        )
        .unwrap();
        assert_eq!(added.added(), Some(7));
        let decoded = tswp::StorageArchive::decode(added.payload()).unwrap();
        let entries = &decoded.table_smartfield.unwrap().entries;
        assert_eq!(entries.len(), 2);
        assert_eq!(
            super::entries(added.payload(), Table::SmartField, limits).unwrap(),
            [Entry {
                object: 7,
                range: range(0, 5)
            }]
        );
        let moved = rewrite(
            added.payload(),
            Table::SmartField,
            &Change::Move {
                object: 7,
                range: range(6, 11),
            },
            limits,
        )
        .unwrap();
        assert_eq!(
            super::entries(moved.payload(), Table::SmartField, limits).unwrap()[0].range(),
            range(6, 11)
        );
        let removed = rewrite(
            moved.payload(),
            Table::SmartField,
            &Change::Remove { object: 7 },
            limits,
        )
        .unwrap();
        assert_eq!(removed.removed(), Some(7));
        assert_eq!(removed.payload(), source.as_slice());
    }

    #[test]
    fn occupied_ranges_and_split_surrogates_are_refused() {
        let limits = WireLimits::DEFAULT;
        let source = storage("a😀bc");
        assert_eq!(
            rewrite(
                &source,
                Table::Bookmark,
                &Change::Add {
                    object: 3,
                    range: range(0, 2)
                },
                limits
            ),
            Err(Error::Refused(Refusal::NotABoundary))
        );
        let added = rewrite(
            &source,
            Table::Bookmark,
            &Change::Add {
                object: 3,
                range: range(0, 3),
            },
            limits,
        )
        .unwrap();
        assert_eq!(
            rewrite(
                added.payload(),
                Table::Bookmark,
                &Change::Add {
                    object: 4,
                    range: range(1, 4)
                },
                limits
            ),
            Err(Error::Refused(Refusal::Occupied))
        );
        assert_eq!(
            rewrite(
                added.payload(),
                Table::Bookmark,
                &Change::Remove { object: 9 },
                limits
            ),
            Err(Error::Refused(Refusal::NotOwned))
        );
    }
}
