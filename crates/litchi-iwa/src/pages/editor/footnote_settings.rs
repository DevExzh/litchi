//! Typed, lossless Pages footnote and endnote settings CRUD.

use super::document_options::{locate_settings, settings_message_index};
use super::*;

const FOOTNOTE_KIND_FIELD: u32 = 30;
const FOOTNOTE_FORMAT_FIELD: u32 = 31;
const FOOTNOTE_NUMBERING_FIELD: u32 = 32;
const FOOTNOTE_GAP_FIELD: u32 = 33;

impl PagesEditor {
    /// Read the lossless settings shown by Pages' Footnotes formatter.
    pub fn footnote_settings(&self) -> Result<FootnoteSettings> {
        let location = locate_settings(self.text.package())?;
        decode_footnote_settings(&location.data, &location.settings)
    }

    /// Replace Pages' footnote and endnote settings transactionally.
    ///
    /// Only the four note fields are patched. Unknown fields, field order, and
    /// all unrelated document settings remain byte-for-byte unchanged.
    pub fn set_footnote_settings(&mut self, settings: FootnoteSettings) -> Result<()> {
        validate_footnote_settings(settings)?;
        let current = self.footnote_settings()?;
        if current == settings {
            return Ok(());
        }

        let location = locate_settings(self.text.package())?;
        let mut staged = self.text.package().clone();
        staged.update_archive(&location.archive_name, |archive| {
            let object = archive.object_mut(location.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages settings object {} is missing",
                    location.identifier
                ))
            })?;
            let message_index = settings_message_index(object, location.identifier)?;
            let original = object.messages[message_index].data.as_slice();
            let native = tp::SettingsArchive::decode(original)?;
            if decode_footnote_settings(original, &native)? != current {
                return Err(Error::InvalidFormat(
                    "Pages footnote settings changed during mutation".to_owned(),
                ));
            }
            let data = encode_footnote_settings(original, &native, settings)?;
            let message_type = object.messages[message_index].type_;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;

        let verified = Self::from_package(staged)?;
        if verified.footnote_settings()? != settings {
            return Err(Error::InvalidFormat(
                "Pages footnote settings failed round-trip validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn decode_footnote_settings(
    original: &[u8],
    native: &tp::SettingsArchive,
) -> Result<FootnoteSettings> {
    for (field, present, raw) in [
        (
            FOOTNOTE_KIND_FIELD,
            native.footnote_kind.is_some(),
            native.footnote_kind.map(i32_varint),
        ),
        (
            FOOTNOTE_FORMAT_FIELD,
            native.footnote_format.is_some(),
            native.footnote_format.map(i32_varint),
        ),
        (
            FOOTNOTE_NUMBERING_FIELD,
            native.footnote_numbering.is_some(),
            native.footnote_numbering.map(i32_varint),
        ),
        (
            FOOTNOTE_GAP_FIELD,
            native.footnote_gap.is_some(),
            native.footnote_gap.map(i32_varint),
        ),
    ] {
        patch_varint_field(original, field, present, raw)?;
    }

    let gap = native
        .footnote_gap
        .map(|points| {
            u32::try_from(points)
                .map_err(|_| {
                    Error::InvalidFormat(format!(
                        "Pages footnote gap {points} must be non-negative"
                    ))
                })
                .and_then(|points| FootnoteGap::new(points).map_err(semantic_error))
        })
        .transpose()?;
    Ok(FootnoteSettings {
        kind: native.footnote_kind.map(FootnoteKind::from_raw),
        format: native.footnote_format.map(FootnoteFormat::from_raw),
        numbering: native.footnote_numbering.map(FootnoteNumbering::from_raw),
        gap,
    })
}

fn encode_footnote_settings(
    original: &[u8],
    native: &tp::SettingsArchive,
    settings: FootnoteSettings,
) -> Result<Vec<u8>> {
    decode_footnote_settings(original, native)?;
    let mut data = original.to_vec();
    for (field, present, replacement) in [
        (
            FOOTNOTE_KIND_FIELD,
            native.footnote_kind.is_some(),
            settings.kind.map(FootnoteKind::as_raw),
        ),
        (
            FOOTNOTE_FORMAT_FIELD,
            native.footnote_format.is_some(),
            settings.format.map(FootnoteFormat::as_raw),
        ),
        (
            FOOTNOTE_NUMBERING_FIELD,
            native.footnote_numbering.is_some(),
            settings.numbering.map(FootnoteNumbering::as_raw),
        ),
        (
            FOOTNOTE_GAP_FIELD,
            native.footnote_gap.is_some(),
            settings.gap.map(|gap| gap.points() as i32),
        ),
    ] {
        data = patch_varint_field(&data, field, present, replacement.map(i32_varint))?;
    }
    let verified = tp::SettingsArchive::decode(data.as_slice())?;
    if decode_footnote_settings(&data, &verified)? != settings {
        return Err(Error::InvalidFormat(
            "Pages footnote-settings wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn validate_footnote_settings(settings: FootnoteSettings) -> Result<()> {
    settings.validate().map_err(semantic_error)
}

fn semantic_error(error: litchi_pages::footnote::Error) -> Error {
    Error::ParseError(format!("invalid Pages footnote settings: {error}"))
}

const fn i32_varint(value: i32) -> u64 {
    value as i64 as u64
}
