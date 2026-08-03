//! Transactional plain-text Drop Cap CRUD shared by every iWork suite.

mod native;
mod storage;
mod types;

#[cfg(test)]
mod tests;

pub use types::{
    DropCapCharacterCount, DropCapCharacterScale, DropCapCornerRadius, DropCapLineCount,
    DropCapOutdent, DropCapPadding, DropCapRaisedLines, DropCapWrap, ParagraphDropCap,
    ParagraphDropCapPlacement, ParagraphStart,
};

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::{Error, IWorkPackage, Result};

use crate::text::style_registry::{
    register_private_style, unregister_owner_reference_if_unused, unregister_private_style,
};

pub(super) fn paragraph_drop_caps(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<ParagraphDropCapPlacement>> {
    let storage = storage::locate(package, storage_id)?;
    if storage.entries.iter().all(|entry| entry.style_id.is_none()) {
        return Ok(Vec::new());
    }
    let base = native::base_style(package, storage.stylesheet_id)?;
    storage
        .entries
        .iter()
        .filter_map(|entry| {
            entry
                .style_id
                .filter(|style_id| *style_id != base.identifier)
                .map(|style_id| {
                    storage::validate_paragraph_start(&storage.text, entry.paragraph_start)?;
                    let location = native::locate_style(package, style_id)?;
                    validate_style_ownership(
                        &location,
                        style_id,
                        storage.stylesheet_id,
                        &base.archive_name,
                    )?;
                    Ok(ParagraphDropCapPlacement {
                        paragraph_start: entry.paragraph_start,
                        drop_cap: native::plain_text_model(style_id, &location)?,
                    })
                })
        })
        .collect()
}

pub(super) fn paragraph_drop_cap(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph_start: ParagraphStart,
) -> Result<Option<ParagraphDropCap>> {
    Ok(paragraph_drop_caps(package, storage_id)?
        .into_iter()
        .find(|placement| placement.paragraph_start == paragraph_start)
        .map(|placement| placement.drop_cap))
}

pub(super) fn set_paragraph_drop_cap(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph_start: ParagraphStart,
    drop_cap: ParagraphDropCap,
) -> Result<()> {
    let storage = storage::locate(package, storage_id)?;
    storage::validate_paragraph_start(&storage.text, paragraph_start)?;
    let old_reference_style_id = storage
        .entries
        .iter()
        .find(|entry| entry.paragraph_start == paragraph_start)
        .and_then(|entry| entry.style_id);
    let base = native::base_style(package, storage.stylesheet_id)?;
    let old_style_id = old_reference_style_id.filter(|style_id| *style_id != base.identifier);
    if let Some(style_id) = old_style_id {
        let location = native::locate_style(package, style_id)?;
        validate_style_ownership(
            &location,
            style_id,
            storage.stylesheet_id,
            &base.archive_name,
        )?;
        if native::plain_text_model(style_id, &location)? == drop_cap {
            return Ok(());
        }
        if native::is_exclusive(package, style_id)? {
            let parent_style_id = native::parent_style_id(&location.style, style_id)?;
            let replacement = native::variation_object(
                style_id,
                parent_style_id,
                storage.stylesheet_id,
                drop_cap,
            )?;
            let mut staged = package.clone();
            native::replace_variation(&mut staged, &location, replacement)?;
            validate_drop_cap(&staged, storage_id, paragraph_start, Some(drop_cap))?;
            *package = staged;
            return Ok(());
        }
    }

    let new_style_id = next_object_identifier(package)?;
    let new_style = native::variation_object(
        new_style_id,
        base.identifier,
        storage.stylesheet_id,
        drop_cap,
    )?;
    let mut staged = package.clone();
    insert_style_variation(
        &mut staged,
        &base.archive_name,
        storage.stylesheet_id,
        base.identifier,
        new_style_id,
        new_style,
    )?;
    storage::patch_entry(
        &mut staged,
        &storage,
        paragraph_start,
        old_reference_style_id,
        Some(new_style_id),
    )?;
    register_private_style(
        &mut staged,
        &storage.archive_name,
        &base.archive_name,
        new_style_id,
    )?;
    if let Some(old_style_id) = old_reference_style_id {
        unregister_owner_reference_if_unused(
            &mut staged,
            &storage.archive_name,
            &base.archive_name,
            old_style_id,
        )?;
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    validate_drop_cap(&staged, storage_id, paragraph_start, Some(drop_cap))?;
    *package = staged;
    Ok(())
}

pub(super) fn remove_paragraph_drop_cap(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph_start: ParagraphStart,
) -> Result<bool> {
    let storage = storage::locate(package, storage_id)?;
    storage::validate_paragraph_start(&storage.text, paragraph_start)?;
    let referenced_style_id = storage
        .entries
        .iter()
        .find(|entry| entry.paragraph_start == paragraph_start)
        .and_then(|entry| entry.style_id);
    let base = native::base_style(package, storage.stylesheet_id)?;
    let Some(style_id) = referenced_style_id.filter(|style_id| *style_id != base.identifier) else {
        return Ok(false);
    };
    let location = native::locate_style(package, style_id)?;
    validate_style_ownership(
        &location,
        style_id,
        storage.stylesheet_id,
        &base.archive_name,
    )?;
    native::plain_text_model(style_id, &location)?;
    let exclusive = native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    storage::patch_entry(&mut staged, &storage, paragraph_start, Some(style_id), None)?;
    if exclusive {
        let parent_style_id = native::parent_style_id(&location.style, style_id)?;
        remove_style_variation(
            &mut staged,
            &location.archive_name,
            storage.stylesheet_id,
            parent_style_id,
            style_id,
        )?;
        unregister_private_style(
            &mut staged,
            &storage.archive_name,
            &location.archive_name,
            style_id,
            None,
        )?;
        release_package_identifier_suffix(&mut staged, &[style_id])?;
    } else {
        unregister_owner_reference_if_unused(
            &mut staged,
            &storage.archive_name,
            &location.archive_name,
            style_id,
        )?;
    }
    validate_drop_cap(&staged, storage_id, paragraph_start, None)?;
    *package = staged;
    Ok(true)
}

fn validate_style_ownership(
    location: &native::DropCapStyleLocation,
    style_id: u64,
    expected_stylesheet_id: u64,
    expected_archive_name: &str,
) -> Result<()> {
    if location.archive_name != expected_archive_name
        || native::stylesheet_id(&location.style, style_id)? != expected_stylesheet_id
    {
        return Err(Error::InvalidFormat(format!(
            "iWork Drop Cap style {style_id} is not owned by its text storage's stylesheet"
        )));
    }
    Ok(())
}

fn validate_drop_cap(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph_start: ParagraphStart,
    expected: Option<ParagraphDropCap>,
) -> Result<()> {
    if paragraph_drop_cap(package, storage_id, paragraph_start)? != expected {
        return Err(Error::InvalidFormat(
            "iWork Drop Cap mutation failed validation".to_owned(),
        ));
    }
    Ok(())
}
