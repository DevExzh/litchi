//! Lossless protobuf wire handling for Pages document formatter options.

use super::*;

const BODY_FIELD: u32 = 1;
const HEADERS_FIELD: u32 = 2;
const FOOTERS_FIELD: u32 = 3;
const HYPHENATION_FIELD: u32 = 9;
const LIGATURES_FIELD: u32 = 10;
const FACING_PAGES_FIELD: u32 = 34;

pub(super) fn read_document_options_wire(
    original: &[u8],
    settings: &SettingsArchive,
) -> Result<DocumentOptions> {
    for (field, label, decoded) in [
        (BODY_FIELD, "document body", settings.body),
        (HEADERS_FIELD, "headers", settings.headers),
        (FOOTERS_FIELD, "footers", settings.footers),
        (HYPHENATION_FIELD, "hyphenation", settings.hyphenation),
        (LIGATURES_FIELD, "ligatures", settings.use_ligatures),
        (FACING_PAGES_FIELD, "facing pages", settings.facing_pages),
    ] {
        require_optional_bool(original, field, label, decoded)?;
    }
    Ok(options_from_settings(settings))
}

pub(super) fn write_document_options_wire(
    original: &[u8],
    settings: &SettingsArchive,
    options: DocumentOptions,
) -> Result<Vec<u8>> {
    read_document_options_wire(original, settings)?;
    let mut data = original.to_vec();
    for (field, present, replacement) in [
        (BODY_FIELD, settings.body.is_some(), options.body_enabled()),
        (
            HEADERS_FIELD,
            settings.headers.is_some(),
            options.headers_enabled(),
        ),
        (
            FOOTERS_FIELD,
            settings.footers.is_some(),
            options.footers_enabled(),
        ),
        (
            HYPHENATION_FIELD,
            settings.hyphenation.is_some(),
            options.automatic_hyphenation(),
        ),
        (
            LIGATURES_FIELD,
            settings.use_ligatures.is_some(),
            options.ligatures_enabled(),
        ),
        (
            FACING_PAGES_FIELD,
            settings.facing_pages.is_some(),
            options.facing_pages(),
        ),
    ] {
        data = patch_varint_field(&data, field, present, replacement.map(u64::from))?;
    }
    let verified = SettingsArchive::decode(data.as_slice())?;
    if read_document_options_wire(&data, &verified)? != options {
        return Err(Error::InvalidFormat(
            "Pages document-options wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn require_optional_bool(
    original: &[u8],
    field: u32,
    label: &str,
    decoded: Option<bool>,
) -> Result<()> {
    let values = crate::wire::repeated_varint_values(original, field)?;
    let expected = decoded.map(u64::from);
    if values.as_slice() != expected.as_slice() {
        return Err(Error::InvalidFormat(format!(
            "Pages {label} wire value is missing, duplicated, non-Boolean, or inconsistent"
        )));
    }
    Ok(())
}
