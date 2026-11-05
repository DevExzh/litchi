//! Storage-scoped number-attachment engine (record 0374).
//!
//! A number attachment (page number, page count, …) is a
//! `TSWP.NumberAttachmentArchive` (2043) anchored to a U+FFFC character by
//! the storage's attachment table (field 9). The host inserts the U+FFFC
//! with the storage text rewrite, then names the new object at that index;
//! removal deletes the character, which drops the table entry. This module
//! adds the table entry ([`anchor`]), reads the anchors ([`anchors`]) and
//! reads, encodes and patches the attachment payload.

use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_nested_length_delimited_field, patch_nested_varint_field,
    patch_varint_field, rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::number_attachment_codec;
use litchi_iwa_text::number_attachment::{
    TextNumberAttachmentFormat, TextNumberAttachmentKind, TextNumberAttachmentSettings,
    TextNumberAttachmentText,
};
use litchi_iwa_text::position::TextPosition;
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, NativeResult, Writer};

/// `TSWP.NumberAttachmentArchive`.
pub const MESSAGE_TYPE: u32 = 2_043;

const ATTACHMENT_TABLE_FIELD: u32 = 9;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const TEXT_FIELD: u32 = 3;
const OBJECT_REPLACEMENT_UNIT: u16 = 0xfffc;
const SUPER_FIELD: u32 = 1;
const STRING_EQUIVALENT_FIELD: u32 = 1;
const KIND_FIELD: u32 = 2;
const NUMBER_FORMAT_FIELD: u32 = 2;
const STRING_VALUE_FIELD: u32 = 3;
const NUMBER_FORMAT_NAME_FIELD: u32 = 4;

/// Why a well-formed attachment change is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The position does not hold a U+FFFC on a scalar boundary.
    NotAnchorCharacter,
    /// The position already anchors an attachment.
    Occupied,
}

/// Why a number-attachment operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The storage, its attachment table or the payload is malformed.
    #[error("invalid number-attachment source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A well-formed change this engine refuses.
    #[error("number-attachment change refused: {0:?}")]
    Refused(Refusal),
    /// A configured wire resource was exhausted.
    #[error("number-attachment {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} number-attachment items or bytes")]
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

/// The result of a number-attachment operation.
pub type Result<T> = std::result::Result<T, Error>;

/// One attachment anchor: the U+FFFC position and the object it names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Anchor {
    position: TextPosition,
    object: u64,
}

impl Anchor {
    /// The U+FFFC position.
    #[must_use]
    pub const fn position(self) -> TextPosition {
        self.position
    }

    /// The attachment object.
    #[must_use]
    pub const fn object(self) -> u64 {
        self.object
    }
}

struct Entry {
    index: u32,
    object: u64,
    raw: Vec<u8>,
}

struct Source<'a> {
    units: Vec<u16>,
    starts: Vec<u32>,
    table: Option<&'a [u8]>,
    entries: Vec<Entry>,
}

impl Source<'_> {
    fn anchors_character(&self, index: u32) -> bool {
        self.starts.binary_search(&index).is_ok()
            && usize::try_from(index)
                .ok()
                .and_then(|index| self.units.get(index))
                == Some(&OBJECT_REPLACEMENT_UNIT)
    }
}

fn decode<'a>(storage: &'a [u8], budget: &mut Budget) -> NativeResult<Source<'a>> {
    let mut units = Vec::new();
    let mut starts = Vec::new();
    let mut tables = Vec::new();
    native::fields(storage, 0, budget, |field, _| {
        if field.number == TEXT_FIELD {
            let fragment = std::str::from_utf8(field.bytes()?)
                .map_err(|_| NativeError::Invalid("storage text is not UTF-8"))?;
            native::reserve(&mut units, fragment.len())?;
            native::reserve(&mut starts, fragment.len())?;
            for character in fragment.chars() {
                starts
                    .push(u32::try_from(units.len()).map_err(|_| {
                        NativeError::Invalid("storage text UTF-16 length overflow")
                    })?);
                let mut buffer = [0u16; 2];
                units.extend_from_slice(character.encode_utf16(&mut buffer));
            }
        } else if field.number == ATTACHMENT_TABLE_FIELD {
            native::reserve(&mut tables, 1)?;
            tables.push(field.bytes()?);
        }
        Ok(())
    })?;
    let table = match tables.as_slice() {
        [] => None,
        [table] => Some(*table),
        _ => return Err(NativeError::Invalid("storage repeats its attachment table")),
    };
    let mut entries = Vec::new();
    if let Some(table) = table {
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
            native::reserve(&mut entries, 1)?;
            entries.push(Entry {
                index,
                object: object.map_or(0, |reference| reference.identifier),
                raw: raw.to_vec(),
            });
            Ok(())
        })?;
        if entries.is_empty() {
            return Err(NativeError::Invalid(
                "storage has an empty attachment table",
            ));
        }
    }
    let source = Source {
        units,
        starts,
        table,
        entries,
    };
    let mut previous = None;
    for entry in &source.entries {
        if entry.object == 0 {
            return Err(NativeError::Invalid("attachment table names object zero"));
        }
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(NativeError::Invalid(
                "attachment positions are not strictly increasing",
            ));
        }
        if !source.anchors_character(entry.index) {
            return Err(NativeError::Invalid(
                "attachment is not anchored to a U+FFFC character",
            ));
        }
        previous = Some(entry.index);
    }
    Ok(source)
}

/// Every attachment anchor of the storage, in text order (any attachment
/// kind; the caller resolves which objects are number attachments).
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed storage or table; resource
/// errors as configured.
pub fn anchors(storage: &[u8], limits: WireLimits) -> Result<Vec<Anchor>> {
    let source = decode(storage, &mut Budget::new(limits))?;
    Ok(source
        .entries
        .iter()
        .map(|entry| Anchor {
            position: TextPosition::from_utf16_code_units(entry.index),
            object: entry.object,
        })
        .collect())
}

/// Name `object` at `position`, which must already hold a U+FFFC (insert
/// it with [`crate::rewrite_storage_text_with_behavior`] first). The
/// storage message's references gain `object`.
///
/// # Errors
///
/// [`Error::Refused`] when the position is not a U+FFFC or already anchors
/// an attachment; otherwise as [`anchors`].
pub fn anchor(
    storage: &[u8],
    position: TextPosition,
    object: u64,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let mut source = decode(storage, &mut budget)?;
    let index = position.utf16_index();
    if source.entries.iter().any(|entry| entry.index == index) {
        return Err(Error::Refused(Refusal::Occupied));
    }
    if !source.anchors_character(index) {
        return Err(Error::Refused(Refusal::NotAnchorCharacter));
    }
    let mut writer = Writer::new(limits);
    writer.uint32(ENTRY_INDEX_FIELD, index)?;
    writer.message(ENTRY_OBJECT_FIELD, |writer| {
        native::Reference { identifier: object }.encode(writer)
    })?;
    native::reserve(&mut source.entries, 1)?;
    source.entries.push(Entry {
        index,
        object,
        raw: writer.into_bytes(),
    });
    source.entries.sort_by_key(|entry| entry.index);
    let encoded: Vec<Vec<u8>> = source.entries.into_iter().map(|entry| entry.raw).collect();
    let table = match source.table {
        Some(table) => {
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
                .map_err(native::common)?
        },
        None => {
            let mut writer = Writer::new(limits);
            for entry in &encoded {
                writer.bytes(TABLE_ENTRIES_FIELD, entry)?;
            }
            writer.into_bytes()
        },
    };
    budget.admit_bytes(storage.len())?;
    let payload = patch_length_delimited_field(
        storage,
        ATTACHMENT_TABLE_FIELD,
        source.table.is_some(),
        Some(&table),
    )
    .map_err(native::common)?;
    native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?;
    Ok(payload)
}

fn text(value: Option<&str>) -> Result<Option<TextNumberAttachmentText>> {
    value
        .map(|value| TextNumberAttachmentText::from_boxed(value.into()))
        .transpose()
        .map_err(|_| Error::InvalidSource {
            reason: "number attachment holds an invalid text value",
        })
}

/// The settings of a `TSWP.NumberAttachmentArchive`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload, a missing kind or an
/// invalid text value; resource errors as configured.
pub fn settings(payload: &[u8], limits: WireLimits) -> Result<TextNumberAttachmentSettings> {
    native::admit(
        LimitKind::InputBytes,
        payload.len(),
        limits.max_input_bytes(),
    )?;
    let snapshot = number_attachment_codec::decode_number_attachment(
        payload,
        number_attachment_codec::DecodeOptions::for_source(payload),
    )
    .map_err(|_| Error::InvalidSource {
        reason: "number attachment failed strict validation",
    })?;
    let kind = snapshot.kind().ok_or(Error::InvalidSource {
        reason: "number attachment is missing its kind",
    })?;
    Ok(TextNumberAttachmentSettings {
        kind: TextNumberAttachmentKind::from_raw(kind),
        string_equivalent: text(snapshot.string_equivalent())?,
        number_format: snapshot
            .number_format()
            .map(TextNumberAttachmentFormat::from_native_value),
        string_value: text(snapshot.string_value())?,
        number_format_name: text(snapshot.number_format_name())?,
    })
}

/// A canonical `TSWP.NumberAttachmentArchive`.
///
/// # Errors
///
/// Resource errors as configured.
pub fn encode(settings: &TextNumberAttachmentSettings, limits: WireLimits) -> Result<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.message(SUPER_FIELD, |writer| {
        if let Some(value) = &settings.string_equivalent {
            writer.bytes(STRING_EQUIVALENT_FIELD, value.as_str().as_bytes())?;
        }
        writer.int32(KIND_FIELD, settings.kind.as_raw())
    })?;
    if let Some(format) = settings.number_format {
        writer.uint32(NUMBER_FORMAT_FIELD, format.native_value())?;
    }
    if let Some(value) = &settings.string_value {
        writer.bytes(STRING_VALUE_FIELD, value.as_str().as_bytes())?;
    }
    if let Some(value) = &settings.number_format_name {
        writer.bytes(NUMBER_FORMAT_NAME_FIELD, value.as_str().as_bytes())?;
    }
    Ok(writer.into_bytes())
}

#[allow(
    clippy::cast_sign_loss,
    reason = "protobuf int32 encodes the sign-extended 64-bit value"
)]
const fn int32_varint(value: i32) -> u64 {
    value as i64 as u64
}

/// `payload` with its settings replaced field by field; every other byte
/// is kept.
///
/// # Errors
///
/// As [`settings`].
pub fn with_settings(
    payload: &[u8],
    settings: &TextNumberAttachmentSettings,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let current = self::settings(payload, limits)?;
    let wire = native::common;
    let mut data = patch_nested_length_delimited_field(
        payload,
        &[SUPER_FIELD, STRING_EQUIVALENT_FIELD],
        current.string_equivalent.is_some(),
        settings
            .string_equivalent
            .as_ref()
            .map(|value| value.as_str().as_bytes()),
    )
    .map_err(wire)?;
    data = patch_nested_varint_field(
        &data,
        &[SUPER_FIELD, KIND_FIELD],
        true,
        Some(int32_varint(settings.kind.as_raw())),
    )
    .map_err(wire)?;
    data = patch_varint_field(
        &data,
        NUMBER_FORMAT_FIELD,
        current.number_format.is_some(),
        settings
            .number_format
            .map(|value| u64::from(value.native_value())),
    )
    .map_err(wire)?;
    for (field, present, value) in [
        (
            STRING_VALUE_FIELD,
            current.string_value.is_some(),
            settings.string_value.as_ref(),
        ),
        (
            NUMBER_FORMAT_NAME_FIELD,
            current.number_format_name.is_some(),
            settings.number_format_name.as_ref(),
        ),
    ] {
        data = patch_length_delimited_field(
            &data,
            field,
            present,
            value.map(|value| value.as_str().as_bytes()),
        )
        .map_err(wire)?;
    }
    native::admit(
        LimitKind::OutputBytes,
        data.len(),
        limits.max_output_bytes(),
    )?;
    Ok(data)
}
