//! Safety checks for ODS protection snapshots and candidate transactions.

use crate::model::style_protection;
use litchi_core::{Error, Result};

use super::{
    codec::Location,
    model::{Document, Sheet, Styles},
};

const MAX_PROTECTION_TEXT_BYTES: usize = 64 * 1024;

pub(crate) fn validate_snapshot(
    source: &str,
    location: &Location,
    document: &Document,
    sheets: &[Sheet],
    styles: &Styles,
) -> Result<()> {
    if source.len() > super::codec::MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "ODS protection source exceeds the snapshot limit".to_string(),
        ));
    }
    if location.sheets().len() != sheets.len() {
        return Err(Error::InvalidFormat(
            "ODS protection sheet catalog is inconsistent with content.xml".to_string(),
        ));
    }
    validate_document(document)?;
    validate_sheets(sheets, location)?;
    validate_styles(location, styles, location.automatic_xml(), None)
}

pub(crate) fn validate_candidate(
    location: &Location,
    original_sheets: &[Sheet],
    original_styles: &Styles,
    candidate_document: &Document,
    candidate_sheets: &[Sheet],
    candidate_styles: &Styles,
) -> Result<()> {
    validate_document(candidate_document)?;
    validate_sheets(candidate_sheets, location)?;
    if candidate_sheets.len() != original_sheets.len() {
        return Err(Error::InvalidFormat(
            "ODS protection edits cannot add or remove worksheets".to_string(),
        ));
    }
    for (original, candidate) in original_sheets.iter().zip(candidate_sheets) {
        if original.name != candidate.name {
            return Err(Error::InvalidFormat(
                "ODS protection edits cannot rename or reorder worksheets".to_string(),
            ));
        }
    }

    let automatic_source = location.automatic_xml();
    validate_styles(
        location,
        candidate_styles,
        automatic_source,
        Some(original_styles),
    )
}

fn validate_document(document: &Document) -> Result<()> {
    validate_key(&document.key)
}

fn validate_sheets(sheets: &[Sheet], location: &Location) -> Result<()> {
    if sheets.len() > location.sheet_limit() {
        return Err(Error::InvalidFormat(
            "ODS protection sheet catalog exceeds the source limit".to_string(),
        ));
    }
    let mut names = std::collections::HashSet::with_capacity(sheets.len());
    for sheet in sheets {
        if sheet.name.len() > MAX_PROTECTION_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "ODS protection sheet name exceeds the size limit".to_string(),
            ));
        }
        if !names.insert(sheet.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODS protection sheet name '{}'",
                sheet.name
            )));
        }
        validate_key(&sheet.key)?;
    }
    Ok(())
}

fn validate_key(key: &super::model::Key) -> Result<()> {
    for (label, value) in [
        ("protection key", key.value.as_deref()),
        (
            "protection digest algorithm",
            key.digest_algorithm.as_deref(),
        ),
        (
            "secondary protection digest algorithm",
            key.secondary_digest_algorithm.as_deref(),
        ),
    ] {
        if value.is_some_and(|value| value.len() > MAX_PROTECTION_TEXT_BYTES) {
            return Err(Error::InvalidFormat(format!(
                "ODS {label} exceeds the size limit"
            )));
        }
    }
    Ok(())
}

fn validate_styles(
    location: &Location,
    styles: &Styles,
    automatic_source: Option<&str>,
    original: Option<&Styles>,
) -> Result<()> {
    let automatic_changed = original.is_none_or(|original| {
        original.automatic() != styles.automatic() || original.conditional() != styles.conditional()
    });
    if !automatic_changed {
        return Ok(());
    }

    let common = style_protection::common_table_cell_style_names(location.styles_xml())?;
    style_protection::validate_conditional_style_collection(styles.conditional(), &common)?;
    let automatic_xml = automatic_source.unwrap_or("");
    style_protection::validate_protection_style_document(
        location.styles_xml(),
        automatic_xml,
        &styles
            .automatic()
            .iter()
            .map(|style| {
                let mut value =
                    style_protection::TableStyle::new(style.name.clone(), style.protection);
                if let Some(parent) = &style.parent_name {
                    value = value.with_parent_style_name(parent.clone());
                }
                value
            })
            .collect::<Vec<_>>(),
    )?;

    if styles.automatic().len() > 65_536 {
        return Err(Error::InvalidFormat(
            "ODS protection style catalog exceeds the size limit".to_string(),
        ));
    }
    if styles.conditional().len() > 65_536 {
        return Err(Error::InvalidFormat(
            "ODS conditional protection-style catalog exceeds the size limit".to_string(),
        ));
    }
    for style in styles.automatic() {
        if style.name.len() > MAX_PROTECTION_TEXT_BYTES
            || style
                .parent_name
                .as_deref()
                .is_some_and(|name| name.len() > MAX_PROTECTION_TEXT_BYTES)
        {
            return Err(Error::InvalidFormat(
                "ODS protection style name exceeds the size limit".to_string(),
            ));
        }
    }

    // The source is used by the codec for a second, exact safety gate. Keep
    // this check explicit so a future owner cannot accidentally treat a
    // missing automatic-styles part as an editable named-style part.
    if automatic_source.is_none() && location.has_automatic_owner() {
        return Err(Error::InvalidFormat(
            "ODS protection automatic-styles owner could not be located".to_string(),
        ));
    }
    Ok(())
}
