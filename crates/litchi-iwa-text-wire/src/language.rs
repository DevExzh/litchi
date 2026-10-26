//! Storage-scoped language-run engine (record 0374).
//!
//! A storage's `table_language` (field 19) is a `StringAttributeTable`:
//! sorted boundaries at UTF-16 indexes, each naming the BCP 47 tag of the
//! text up to the next boundary (no tag: automatic). The engine reads the
//! runs and sets, removes or resets boundaries as the `litchi-iwa` host's
//! `text::Draft` language family does: touched entries keep their other
//! bytes, neighbouring entries with the same tag coalesce, and every other
//! storage byte is kept.

use litchi_iwa_common::wire::{
    patch_length_delimited_field, rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::language::{TextLanguage, TextLanguageRun};
use litchi_iwa_text::position::TextPosition;
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, NativeResult, Writer};

const LANGUAGE_TABLE_FIELD: u32 = 19;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_LANGUAGE_FIELD: u32 = 2;
const TEXT_FIELD: u32 = 3;

/// One language change of a storage.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    /// Make `language` effective from `position` to the next boundary.
    Set {
        /// A scalar boundary of the text.
        position: TextPosition,
        /// The language.
        language: TextLanguage,
    },
    /// Remove the boundary at `position` (not zero), so the text inherits
    /// the preceding run.
    RemoveBoundary(TextPosition),
    /// Remove every explicit language.
    Reset,
}

/// Why a well-formed language change is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The position is past the text or splits a surrogate pair.
    NotABoundary,
    /// The index-zero boundary goes only with [`Change::Reset`].
    FirstBoundary,
}

/// Why a language operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The storage payload or its language table is malformed.
    #[error("invalid language-table source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A well-formed change this engine refuses.
    #[error("language change refused: {0:?}")]
    Refused(Refusal),
    /// A configured wire resource was exhausted.
    #[error("language {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} language items or bytes")]
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

/// The result of a language operation.
pub type Result<T> = std::result::Result<T, Error>;

struct Entry {
    index: u32,
    tag: Option<String>,
    raw: Vec<u8>,
}

struct Source<'a> {
    starts: Vec<u32>,
    len: u32,
    table: Option<&'a [u8]>,
    entries: Vec<Entry>,
}

impl Source<'_> {
    fn is_boundary(&self, index: u32) -> bool {
        index == self.len || self.starts.binary_search(&index).is_ok()
    }
}

fn language(tag: Option<&str>) -> NativeResult<TextLanguage> {
    tag.map(TextLanguage::tag)
        .transpose()
        .map(Option::unwrap_or_default)
        .map_err(|_| NativeError::Invalid("language table has an invalid language tag"))
}

fn decode<'a>(storage: &'a [u8], budget: &mut Budget) -> NativeResult<Source<'a>> {
    let mut starts = Vec::new();
    let mut len = 0u32;
    let mut tables = Vec::new();
    native::fields(storage, 0, budget, |field, _| {
        if field.number == TEXT_FIELD {
            let fragment = std::str::from_utf8(field.bytes()?)
                .map_err(|_| NativeError::Invalid("storage text is not UTF-8"))?;
            native::reserve(&mut starts, fragment.len())?;
            for character in fragment.chars() {
                starts.push(len);
                len = len
                    .checked_add(u32::try_from(character.len_utf16()).unwrap_or(2))
                    .ok_or(NativeError::Invalid("storage text UTF-16 length overflow"))?;
            }
        } else if field.number == LANGUAGE_TABLE_FIELD {
            native::reserve(&mut tables, 1)?;
            tables.push(field.bytes()?);
        }
        Ok(())
    })?;
    let table = match tables.as_slice() {
        [] => None,
        [table] => Some(*table),
        _ => return Err(NativeError::Invalid("storage repeats its language table")),
    };
    let mut entries = Vec::new();
    if let Some(table) = table {
        native::fields(table, 1, budget, |field, budget| {
            if field.number != TABLE_ENTRIES_FIELD {
                return Ok(());
            }
            let raw = field.bytes()?;
            let mut index = 0u32;
            let mut tag = None;
            native::fields(raw, 2, budget, |inner, budget| {
                match inner.number {
                    ENTRY_INDEX_FIELD => index = inner.uint32()?,
                    ENTRY_LANGUAGE_FIELD => tag = Some(inner.string(budget)?),
                    _ => {},
                }
                Ok(())
            })?;
            native::reserve(&mut entries, 1)?;
            entries.push(Entry {
                index,
                tag,
                raw: raw.to_vec(),
            });
            Ok(())
        })?;
        if entries.is_empty() {
            return Err(NativeError::Invalid("storage has an empty language table"));
        }
    }
    let source = Source {
        starts,
        len,
        table,
        entries,
    };
    if let Some(first) = source.entries.first()
        && first.index != 0
    {
        return Err(NativeError::Invalid(
            "language table must begin at UTF-16 index zero",
        ));
    }
    let mut previous = None;
    for entry in &source.entries {
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(NativeError::Invalid(
                "language boundaries are not strictly increasing",
            ));
        }
        language(entry.tag.as_deref())?;
        if !source.is_boundary(entry.index) {
            return Err(NativeError::Invalid(
                "language boundary is not a scalar boundary",
            ));
        }
        previous = Some(entry.index);
    }
    Ok(source)
}

/// Every explicit language boundary, in text order.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed storage or table; resource
/// errors as configured.
pub fn runs(storage: &[u8], limits: WireLimits) -> Result<Vec<TextLanguageRun>> {
    let source = decode(storage, &mut Budget::new(limits))?;
    let mut runs = Vec::new();
    native::reserve(&mut runs, source.entries.len())?;
    for entry in &source.entries {
        runs.push(TextLanguageRun::new(
            TextPosition::from_utf16_code_units(entry.index),
            language(entry.tag.as_deref())?,
        ));
    }
    Ok(runs)
}

fn effective(source: &Source<'_>, position: TextPosition) -> NativeResult<TextLanguage> {
    let Some(entry) = source
        .entries
        .iter()
        .take_while(|entry| entry.index <= position.utf16_index())
        .last()
    else {
        return Ok(TextLanguage::Automatic);
    };
    language(entry.tag.as_deref())
}

/// The effective language at one scalar boundary.
///
/// # Errors
///
/// [`Error::Refused`] when `position` is not a scalar boundary; otherwise
/// as [`runs`].
pub fn at(storage: &[u8], position: TextPosition, limits: WireLimits) -> Result<TextLanguage> {
    let source = decode(storage, &mut Budget::new(limits))?;
    if !source.is_boundary(position.utf16_index()) {
        return Err(Error::Refused(Refusal::NotABoundary));
    }
    Ok(effective(&source, position)?)
}

fn entry_payload(index: u32, tag: Option<&str>, limits: WireLimits) -> NativeResult<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.uint32(ENTRY_INDEX_FIELD, index)?;
    if let Some(tag) = tag {
        writer.bytes(ENTRY_LANGUAGE_FIELD, tag.as_bytes())?;
    }
    Ok(writer.into_bytes())
}

/// Apply one [`Change`]; `None` when it changes nothing.
///
/// # Errors
///
/// [`Error::Refused`] for a position that is not a scalar boundary or for
/// removing the index-zero boundary; otherwise as [`runs`].
pub fn rewrite(storage: &[u8], change: &Change, limits: WireLimits) -> Result<Option<Vec<u8>>> {
    let mut budget = Budget::new(limits);
    let source = decode(storage, &mut budget)?;
    let table = match change {
        Change::Set { position, language } => {
            if !source.is_boundary(position.utf16_index()) {
                return Err(Error::Refused(Refusal::NotABoundary));
            }
            if effective(&source, *position)? == *language {
                return Ok(None);
            }
            let tag = language.as_tag().map(|tag| tag.as_str());
            Some(match source.table {
                Some(table) => {
                    set_existing(table, source.entries, position.utf16_index(), tag, limits)?
                },
                None => {
                    let mut writer = Writer::new(limits);
                    if position.utf16_index() != 0 {
                        writer.bytes(TABLE_ENTRIES_FIELD, &entry_payload(0, None, limits)?)?;
                    }
                    writer.bytes(
                        TABLE_ENTRIES_FIELD,
                        &entry_payload(position.utf16_index(), tag, limits)?,
                    )?;
                    writer.into_bytes()
                },
            })
        },
        Change::RemoveBoundary(position) => {
            if *position == TextPosition::ZERO {
                return Err(Error::Refused(Refusal::FirstBoundary));
            }
            if !source.is_boundary(position.utf16_index()) {
                return Err(Error::Refused(Refusal::NotABoundary));
            }
            let Some(table) = source.table else {
                return Ok(None);
            };
            if !source
                .entries
                .iter()
                .any(|entry| entry.index == position.utf16_index())
            {
                return Ok(None);
            }
            let mut retained: Vec<Entry> = source
                .entries
                .into_iter()
                .filter(|entry| entry.index != position.utf16_index())
                .collect();
            retained.dedup_by(|current, previous| current.tag == previous.tag);
            let retained: Vec<Vec<u8>> = retained.into_iter().map(|entry| entry.raw).collect();
            Some(
                rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &retained)
                    .map_err(native::common)?,
            )
        },
        Change::Reset => {
            if source.table.is_none() {
                return Ok(None);
            }
            None
        },
    };
    budget.admit_bytes(storage.len())?;
    let payload = patch_length_delimited_field(
        storage,
        LANGUAGE_TABLE_FIELD,
        source.table.is_some(),
        table.as_deref(),
    )
    .map_err(native::common)?;
    native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?;
    Ok(Some(payload))
}

fn set_existing(
    table: &[u8],
    mut entries: Vec<Entry>,
    index: u32,
    tag: Option<&str>,
    limits: WireLimits,
) -> NativeResult<Vec<u8>> {
    if let Some(entry) = entries.iter_mut().find(|entry| entry.index == index) {
        entry.raw = patch_length_delimited_field(
            &entry.raw,
            ENTRY_LANGUAGE_FIELD,
            entry.tag.is_some(),
            tag.map(str::as_bytes),
        )
        .map_err(native::common)?;
        entry.tag = tag.map(str::to_owned);
    } else {
        native::reserve(&mut entries, 1)?;
        entries.push(Entry {
            index,
            tag: tag.map(str::to_owned),
            raw: entry_payload(index, tag, limits)?,
        });
        entries.sort_by_key(|entry| entry.index);
    }
    entries.dedup_by(|current, previous| current.tag == previous.tag);
    let encoded: Vec<Vec<u8>> = entries.into_iter().map(|entry| entry.raw).collect();
    rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
        .map_err(native::common)
}
