//! Payload codecs of the objects a storage's ranged tables name: hyperlink
//! and bookmark smart fields (record 0374).
//!
//! Each object is one message whose `TSWP.SmartFieldArchive` carries a
//! text-attribute UUID. The host draws that UUID at random; the codecs take
//! it from the caller ([`TextAttributeUuid::new`]), so a format package
//! decides how it is generated. Writes encode canonically; setting changes
//! keep every other byte of the payload.

use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_nested_fixed64_field, patch_varint_field,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::hyperlink_codec;
use litchi_iwa_text::bookmark::{Name as BookmarkName, Settings as BookmarkSettings, Visibility};
use litchi_iwa_text::date_time::{
    Format as DateTimeFormat, FormatterStyle, Instant, LocaleIdentifier,
    Settings as DateTimeSettings, UpdatePlan,
};
use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, Writer};

/// `TSWP.HyperlinkFieldArchive`.
pub const HYPERLINK_MESSAGE_TYPE: u32 = 2_032;
/// `TSWP.BookmarkFieldArchive`.
pub const BOOKMARK_MESSAGE_TYPE: u32 = 2_035;
/// `TSWP.DateTimeSmartFieldArchive`.
pub const DATE_TIME_MESSAGE_TYPE: u32 = 2_034;

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

const DATE_FORMAT_FIELD: u32 = 2;
const DATE_LOCALE_FIELD: u32 = 3;
const DATE_STYLE_FIELD: u32 = 4;
const TIME_STYLE_FIELD: u32 = 5;
const UPDATE_PLAN_FIELD: u32 = 6;
const NEEDS_UPDATE_FIELD: u32 = 7;
const DATE_FIELD: u32 = 8;
const DATE_SECONDS_FIELD: u32 = 1;

#[derive(Default)]
struct DateTimeFields {
    uuid: Option<String>,
    format: Option<String>,
    locale: Option<String>,
    date_style: Option<i32>,
    time_style: Option<i32>,
    update_plan: Option<i32>,
    needs_update: Option<bool>,
    date: Option<f64>,
}

fn date_time_fields(payload: &[u8], limits: WireLimits) -> Result<DateTimeFields> {
    let mut budget = Budget::new(limits);
    let mut fields = DateTimeFields::default();
    native::fields(payload, 0, &mut budget, |field, budget| {
        match field.number {
            SUPER_FIELD => native::fields(field.bytes()?, 1, budget, |inner, budget| {
                if inner.number == UUID_FIELD {
                    fields.uuid = Some(inner.string(budget)?);
                }
                Ok(())
            })?,
            DATE_FORMAT_FIELD => fields.format = Some(field.string(budget)?),
            DATE_LOCALE_FIELD => fields.locale = Some(field.string(budget)?),
            DATE_STYLE_FIELD => fields.date_style = Some(field.int32()?),
            TIME_STYLE_FIELD => fields.time_style = Some(field.int32()?),
            UPDATE_PLAN_FIELD => fields.update_plan = Some(field.int32()?),
            NEEDS_UPDATE_FIELD => fields.needs_update = Some(field.bool()?),
            DATE_FIELD => {
                let seconds = fields.date.get_or_insert(0.0);
                native::fields(field.bytes()?, 1, budget, |inner, _| {
                    if inner.number == DATE_SECONDS_FIELD {
                        *seconds = inner.double()?;
                    }
                    Ok(())
                })?;
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok(fields)
}

fn invalid_date_time(_: litchi_iwa_text::date_time::Error) -> Error {
    Error::InvalidSource {
        reason: "Date & Time field holds an invalid setting",
    }
}

/// The text-attribute UUID and settings of a `TSWP.DateTimeSmartFieldArchive`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload, a missing or invalid
/// UUID or an invalid setting; resource errors as configured.
pub fn date_time(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(TextAttributeUuid, DateTimeSettings)> {
    let fields = date_time_fields(payload, limits)?;
    let uuid = fields.uuid.ok_or(Error::InvalidSource {
        reason: "Date & Time field is missing its text-attribute UUID",
    })?;
    let uuid = TextAttributeUuid::new(&uuid).map_err(|_| Error::InvalidSource {
        reason: "Date & Time field has an invalid text-attribute UUID",
    })?;
    let settings = DateTimeSettings::new(
        fields
            .format
            .map(DateTimeFormat::try_from)
            .transpose()
            .map_err(invalid_date_time)?,
        fields
            .locale
            .map(LocaleIdentifier::try_from)
            .transpose()
            .map_err(invalid_date_time)?,
        fields.date_style.map(FormatterStyle::from_raw),
        fields.time_style.map(FormatterStyle::from_raw),
        fields.update_plan.map(UpdatePlan::from_raw),
        fields.needs_update,
        fields
            .date
            .map(Instant::from_reference_date_seconds)
            .transpose()
            .map_err(invalid_date_time)?,
    )
    .map_err(invalid_date_time)?;
    Ok((uuid, settings))
}

/// A canonical `TSWP.DateTimeSmartFieldArchive`.
///
/// # Errors
///
/// Resource errors as configured.
pub fn encode_date_time(
    uuid: &TextAttributeUuid,
    settings: &DateTimeSettings,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.message(SUPER_FIELD, |writer| {
        writer.bytes(UUID_FIELD, uuid.as_str().as_bytes())
    })?;
    if let Some(format) = settings.format() {
        writer.bytes(DATE_FORMAT_FIELD, format.as_str().as_bytes())?;
    }
    if let Some(locale) = settings.locale_identifier() {
        writer.bytes(DATE_LOCALE_FIELD, locale.as_str().as_bytes())?;
    }
    if let Some(style) = settings.date_style() {
        writer.int32(DATE_STYLE_FIELD, style.as_raw())?;
    }
    if let Some(style) = settings.time_style() {
        writer.int32(TIME_STYLE_FIELD, style.as_raw())?;
    }
    if let Some(plan) = settings.update_plan() {
        writer.int32(UPDATE_PLAN_FIELD, plan.as_raw())?;
    }
    if let Some(needs_update) = settings.needs_update() {
        writer.bool(NEEDS_UPDATE_FIELD, needs_update)?;
    }
    if let Some(instant) = settings.instant() {
        writer.message(DATE_FIELD, |writer| {
            writer.double(DATE_SECONDS_FIELD, instant.reference_date_seconds())
        })?;
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
/// [`Error::InvalidSource`] for a malformed payload; resource errors as
/// configured.
pub fn date_time_with_settings(
    payload: &[u8],
    settings: &DateTimeSettings,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let current = date_time_fields(payload, limits)?;
    let wire = native::common;
    let mut data = patch_length_delimited_field(
        payload,
        DATE_FORMAT_FIELD,
        current.format.is_some(),
        settings.format().map(|value| value.as_str().as_bytes()),
    )
    .map_err(wire)?;
    data = patch_length_delimited_field(
        &data,
        DATE_LOCALE_FIELD,
        current.locale.is_some(),
        settings
            .locale_identifier()
            .map(|value| value.as_str().as_bytes()),
    )
    .map_err(wire)?;
    for (number, present, replacement) in [
        (
            DATE_STYLE_FIELD,
            current.date_style.is_some(),
            settings.date_style().map(FormatterStyle::as_raw),
        ),
        (
            TIME_STYLE_FIELD,
            current.time_style.is_some(),
            settings.time_style().map(FormatterStyle::as_raw),
        ),
        (
            UPDATE_PLAN_FIELD,
            current.update_plan.is_some(),
            settings.update_plan().map(UpdatePlan::as_raw),
        ),
    ] {
        data = patch_varint_field(&data, number, present, replacement.map(int32_varint))
            .map_err(wire)?;
    }
    data = patch_varint_field(
        &data,
        NEEDS_UPDATE_FIELD,
        current.needs_update.is_some(),
        settings.needs_update().map(u64::from),
    )
    .map_err(wire)?;
    data = match (current.date, settings.instant()) {
        (Some(_), Some(instant)) => patch_nested_fixed64_field(
            &data,
            &[DATE_FIELD, DATE_SECONDS_FIELD],
            true,
            Some(instant.reference_date_seconds().to_bits()),
        )
        .map_err(wire)?,
        (None, Some(instant)) => {
            let mut writer = Writer::new(limits);
            writer.double(DATE_SECONDS_FIELD, instant.reference_date_seconds())?;
            patch_length_delimited_field(&data, DATE_FIELD, false, Some(&writer.into_bytes()))
                .map_err(wire)?
        },
        (Some(_), None) => {
            patch_length_delimited_field(&data, DATE_FIELD, true, None).map_err(wire)?
        },
        (None, None) => data,
    };
    admit_output(&data, limits)?;
    Ok(data)
}
