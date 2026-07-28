//! Transactional paragraph-list CRUD shared by Pages, Numbers, and Keynote.

mod levels;
mod native;
mod storage;
mod types;

use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::text::style_registry::{
    insert_private_style, object_archive_name, register_private_style, register_style_reference,
    unregister_owner_reference_if_unused,
};
use crate::{Error, IWorkPackage, Result};

pub use types::{ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement};

pub(super) use levels::{
    paragraph_list_level, paragraph_list_levels, reset_paragraph_list_level,
    set_paragraph_list_level,
};

/// Build one canonical stylesheet object for a source-created document theme.
///
/// Scratch documents and editor-created private styles must share the same
/// native nine-level definitions so the apps present consistent presets.
pub(crate) fn preset_style_object(
    identifier: u64,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<crate::archive::ArchiveObject> {
    native::style_object(identifier, stylesheet_id, preset)
}

/// Locate a canonical list preset owned by one stylesheet.
pub(crate) fn preset_style_id(
    package: &IWorkPackage,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<Option<u64>> {
    let archive_name = object_archive_name(package, stylesheet_id)?;
    native::find_preset_style(package, &archive_name, stylesheet_id, preset)
}

pub(crate) fn paragraph_list(package: &IWorkPackage, storage_id: u64) -> Result<ParagraphList> {
    let storage = storage::locate(package, storage_id)?;
    let style = native::locate_style(package, storage.style_id)?;
    native::paragraph_list(&style.style)
}

pub(super) fn set_paragraph_list(
    package: &mut IWorkPackage,
    storage_id: u64,
    list: ParagraphList,
) -> Result<()> {
    let storage = storage::locate(package, storage_id)?;
    let current = native::locate_style(package, storage.style_id)?;
    if native::paragraph_list(&current.style).ok() == Some(list) {
        return Ok(());
    }
    let stylesheet_id = native::stylesheet_id(package, &current.style, storage.style_id)?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != current.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {} is not stored with stylesheet {stylesheet_id}",
            storage.style_id
        )));
    }

    let existing = native::find_preset_style(package, &current.archive_name, stylesheet_id, list)?;
    let target_style_id = match existing {
        Some(identifier) => identifier,
        None => next_object_identifier(package)?,
    };
    let mut staged = package.clone();
    if existing.is_some() {
        register_style_reference(
            &mut staged,
            &storage.archive_name,
            &current.archive_name,
            target_style_id,
        )?;
    } else {
        let object = native::style_object(target_style_id, stylesheet_id, list)?;
        insert_private_style(
            &mut staged,
            &current.archive_name,
            stylesheet_id,
            target_style_id,
            object,
        )?;
        register_private_style(
            &mut staged,
            &storage.archive_name,
            &current.archive_name,
            target_style_id,
        )?;
        set_package_last_object_identifier(&mut staged, target_style_id)?;
    }
    storage::patch_style_reference(
        &mut staged,
        &storage.archive_name,
        storage_id,
        storage.style_id,
        target_style_id,
    )?;
    unregister_owner_reference_if_unused(
        &mut staged,
        &storage.archive_name,
        &current.archive_name,
        storage.style_id,
    )?;
    if paragraph_list(&staged, storage_id)? != list {
        return Err(Error::InvalidFormat(
            "iWork paragraph-list update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    let storage = storage::locate(package, storage_id)?;
    let current = native::locate_style(package, storage.style_id)?;
    if native::paragraph_list(&current.style).ok() == Some(ParagraphList::None) {
        return Ok(false);
    }
    set_paragraph_list(package, storage_id, ParagraphList::None)?;
    Ok(true)
}

#[cfg(test)]
mod level_tests;
#[cfg(test)]
mod tests;
