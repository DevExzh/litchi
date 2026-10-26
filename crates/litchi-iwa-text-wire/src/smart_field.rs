//! Payload codecs of the objects a storage's ranged tables name: hyperlink
//! and bookmark smart fields (record 0374).
//!
//! Each object is one message whose `TSWP.SmartFieldArchive` carries a
//! text-attribute UUID. The host draws that UUID at random; the codecs take
//! it from the caller ([`TextAttributeUuid::new`]), so a format package
//! decides how it is generated. Writes encode canonically; setting changes
//! keep every other byte of the payload.

use litchi_iwa_common::wire::{patch_length_delimited_field, patch_varint_field};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::hyperlink_codec;
use litchi_iwa_text::bookmark::{Name as BookmarkName, Settings as BookmarkSettings, Visibility};
use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, Writer};

/// `TSWP.HyperlinkFieldArchive`.
pub const HYPERLINK_MESSAGE_TYPE: u32 = 2_032;
/// `TSWP.BookmarkFieldArchive`.
pub const BOOKMARK_MESSAGE_TYPE: u32 = 2_035;

const SUPER_FIELD: u32 = 1;
const UUID_FIELD: u32 = 1;
const HYPERLINK_TARGET_FIELD: u32 = 2;
const BOOKMARK_NAME_FIELD: u32 = 2;
const BOOKMARK_RANGED_FIELD: u32 = 3;
const BOOKMARK_HIDDEN_FIELD: u32 = 4;
const RANGED_BOOKMARK: u32 = 1;

/// Why a smart-field payload could not be read or written.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The payload is malformed or holds a value the semantic types reject.
    #[error("invalid smart-field source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A caller-supplied text-attribute UUID is not a hyphenated UUID.
    #[error("text-attribute UUID must be 36 hyphenated hexadecimal characters")]
    InvalidUuid,
    /// A configured wire resource was exhausted.
    #[error("smart-field {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} smart-field items or bytes")]
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

/// The result of a smart-field codec.
pub type Result<T> = std::result::Result<T, Error>;

/// A smart field's hyphenated text-attribute UUID
/// (`xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextAttributeUuid(Box<str>);

impl TextAttributeUuid {
    /// Validate a hyphenated UUID.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidUuid`] unless `value` is 36 characters with hyphens
    /// at 8, 13, 18 and 23 and hexadecimal digits elsewhere.
    pub fn new(value: &str) -> Result<Self> {
        let valid = value.len() == 36
            && value.bytes().enumerate().all(|(index, byte)| match index {
                8 | 13 | 18 | 23 => byte == b'-',
                _ => byte.is_ascii_hexdigit(),
            });
        if valid {
            Ok(Self(value.into()))
        } else {
            Err(Error::InvalidUuid)
        }
    }

    /// The UUID text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

fn admit_input(payload: &[u8], limits: WireLimits) -> Result<()> {
    native::admit(
        LimitKind::InputBytes,
        payload.len(),
        limits.max_input_bytes(),
    )?;
    Ok(())
}

fn admit_output(payload: &[u8], limits: WireLimits) -> Result<()> {
    native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?;
    Ok(())
}

fn decode_hyperlink(
    payload: &[u8],
    limits: WireLimits,
) -> Result<hyperlink_codec::HyperlinkSnapshot<'_>> {
    admit_input(payload, limits)?;
    hyperlink_codec::decode_hyperlink(payload, hyperlink_codec::DecodeOptions::for_source(payload))
        .map_err(|_| Error::InvalidSource {
            reason: "hyperlink payload failed strict validation",
        })
}

/// The text-attribute UUID and target of a `TSWP.HyperlinkFieldArchive`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload, a missing or invalid
/// UUID or a missing or invalid target; resource errors as configured.
pub fn hyperlink(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(TextAttributeUuid, TextHyperlinkTarget)> {
    let snapshot = decode_hyperlink(payload, limits)?;
    let uuid = snapshot.text_attribute_uuid().ok_or(Error::InvalidSource {
        reason: "hyperlink is missing its text-attribute UUID",
    })?;
    let uuid = TextAttributeUuid::new(uuid).map_err(|_| Error::InvalidSource {
        reason: "hyperlink has an invalid text-attribute UUID",
    })?;
    let target = snapshot.url_ref().ok_or(Error::InvalidSource {
        reason: "hyperlink is missing its target",
    })?;
    let target = TextHyperlinkTarget::new(target).map_err(|_| Error::InvalidSource {
        reason: "hyperlink has an invalid target",
    })?;
    Ok((uuid, target))
}

/// A canonical `TSWP.HyperlinkFieldArchive`.
///
/// # Errors
///
/// Resource errors as configured.
pub fn encode_hyperlink(
    uuid: &TextAttributeUuid,
    target: &TextHyperlinkTarget,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let payload =
        hyperlink_codec::encode_hyperlink(uuid.as_str(), target.as_str()).map_err(|_| {
            Error::Allocation {
                amount: target.as_str().len(),
            }
        })?;
    admit_output(&payload, limits)?;
    Ok(payload)
}

/// `payload` with its target replaced; every other byte is kept.
///
/// # Errors
///
/// As [`hyperlink`].
pub fn hyperlink_with_target(
    payload: &[u8],
    target: &TextHyperlinkTarget,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let present = decode_hyperlink(payload, limits)?.url_ref().is_some();
    let rewritten = patch_length_delimited_field(
        payload,
        HYPERLINK_TARGET_FIELD,
        present,
        Some(target.as_str().as_bytes()),
    )
    .map_err(native::common)?;
    admit_output(&rewritten, limits)?;
    Ok(rewritten)
}

#[derive(Default)]
struct BookmarkFields {
    uuid: Option<String>,
    name: Option<String>,
    ranged: Option<u32>,
    hidden: Option<u32>,
}

fn bookmark_fields(payload: &[u8], limits: WireLimits) -> Result<BookmarkFields> {
    let mut budget = Budget::new(limits);
    let mut fields = BookmarkFields::default();
    native::fields(payload, 0, &mut budget, |field, budget| {
        match field.number {
            SUPER_FIELD => native::fields(field.bytes()?, 1, budget, |inner, budget| {
                if inner.number == UUID_FIELD {
                    fields.uuid = Some(inner.string(budget)?);
                }
                Ok(())
            })?,
            BOOKMARK_NAME_FIELD => fields.name = Some(field.string(budget)?),
            BOOKMARK_RANGED_FIELD => fields.ranged = Some(field.uint32()?),
            BOOKMARK_HIDDEN_FIELD => fields.hidden = Some(field.uint32()?),
            _ => {},
        }
        Ok(())
    })?;
    Ok(fields)
}

const fn visibility_from_native(raw: u32) -> Visibility {
    match raw {
        0 => Visibility::Visible,
        1 => Visibility::Hidden,
        unknown => Visibility::Unknown(unknown),
    }
}

const fn visibility_to_native(visibility: Visibility) -> u32 {
    match visibility {
        Visibility::Visible => 0,
        Visibility::Hidden => 1,
        Visibility::Unknown(raw) => raw,
    }
}

/// The text-attribute UUID and settings of a ranged
/// `TSWP.BookmarkFieldArchive`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload, a missing or invalid
/// UUID, an unranged bookmark or an invalid name; resource errors as
/// configured.
pub fn bookmark(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(TextAttributeUuid, BookmarkSettings)> {
    let fields = bookmark_fields(payload, limits)?;
    let uuid = fields.uuid.ok_or(Error::InvalidSource {
        reason: "bookmark is missing its text-attribute UUID",
    })?;
    let uuid = TextAttributeUuid::new(&uuid).map_err(|_| Error::InvalidSource {
        reason: "bookmark has an invalid text-attribute UUID",
    })?;
    if fields.ranged != Some(RANGED_BOOKMARK) {
        return Err(Error::InvalidSource {
            reason: "bookmark is not a ranged bookmark",
        });
    }
    let settings =
        BookmarkSettings::new().with_visibility(visibility_from_native(fields.hidden.unwrap_or(0)));
    let settings =
        match fields.name {
            Some(name) => settings.with_name(BookmarkName::try_from(name).map_err(|_| {
                Error::InvalidSource {
                    reason: "bookmark has an invalid name",
                }
            })?),
            None => settings,
        };
    Ok((uuid, settings))
}

/// A canonical ranged `TSWP.BookmarkFieldArchive`.
///
/// # Errors
///
/// Resource errors as configured.
pub fn encode_bookmark(
    uuid: &TextAttributeUuid,
    settings: &BookmarkSettings,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.message(SUPER_FIELD, |writer| {
        writer.bytes(UUID_FIELD, uuid.as_str().as_bytes())
    })?;
    if let Some(name) = settings.name() {
        writer.bytes(BOOKMARK_NAME_FIELD, name.as_str().as_bytes())?;
    }
    writer.uint32(BOOKMARK_RANGED_FIELD, RANGED_BOOKMARK)?;
    writer.uint32(
        BOOKMARK_HIDDEN_FIELD,
        visibility_to_native(settings.visibility()),
    )?;
    Ok(writer.into_bytes())
}

/// `payload` with its name and visibility replaced; every other byte is
/// kept.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload; resource errors as
/// configured.
pub fn bookmark_with_settings(
    payload: &[u8],
    settings: &BookmarkSettings,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let fields = bookmark_fields(payload, limits)?;
    let named = patch_length_delimited_field(
        payload,
        BOOKMARK_NAME_FIELD,
        fields.name.is_some(),
        settings.name().map(|name| name.as_str().as_bytes()),
    )
    .map_err(native::common)?;
    let rewritten = patch_varint_field(
        &named,
        BOOKMARK_HIDDEN_FIELD,
        fields.hidden.is_some(),
        Some(u64::from(visibility_to_native(settings.visibility()))),
    )
    .map_err(native::common)?;
    admit_output(&rewritten, limits)?;
    Ok(rewritten)
}
