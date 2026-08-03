//! Transactional paragraph-list CRUD shared by Pages, Numbers, and Keynote.

mod bullet;
mod geometry;
mod indentation;
mod label_color;
mod levels;
mod native;
mod number_format;
mod number_scale;
mod number_tiering;
mod numbering;
mod storage;
mod types;
mod variation;

use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::text::style_registry::{
    insert_private_style, object_archive, object_archive_name, register_private_style,
    register_style_reference, unregister_owner_reference_if_unused,
};
use crate::{Error, IWorkPackage, Result};

pub use types::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
    ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
    ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberPunctuation,
    ParagraphListNumberScale, ParagraphListNumberSequence, ParagraphListNumberTiering,
    ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
};

pub(crate) use bullet::paragraph_list_bullet;
pub(super) use bullet::{reset_paragraph_list_bullet, set_paragraph_list_bullet};
pub(crate) use geometry::paragraph_list_bullet_geometry;
pub(super) use geometry::{
    reset_paragraph_list_bullet_geometry, set_paragraph_list_bullet_geometry,
};
pub(crate) use indentation::paragraph_list_indentation;
pub(super) use indentation::{reset_paragraph_list_indentation, set_paragraph_list_indentation};
pub(crate) use label_color::paragraph_list_label_color;
pub(super) use label_color::{reset_paragraph_list_label_color, set_paragraph_list_label_color};
pub(crate) use levels::{paragraph_list_level, paragraph_list_levels};
pub(super) use levels::{reset_paragraph_list_level, set_paragraph_list_level};
pub(crate) use number_format::paragraph_list_number_format;
pub(super) use number_format::{
    reset_paragraph_list_number_format, set_paragraph_list_number_format,
};
pub(crate) use number_scale::paragraph_list_number_scale;
pub(super) use number_scale::{reset_paragraph_list_number_scale, set_paragraph_list_number_scale};
pub(crate) use number_tiering::paragraph_list_number_tiering;
pub(super) use number_tiering::{
    reset_paragraph_list_number_tiering, set_paragraph_list_number_tiering,
};
pub(crate) use numbering::paragraph_list_numbering;
pub(super) use numbering::set_paragraph_list_numbering;

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
    let (_, archive) = object_archive(package, stylesheet_id)?;
    native::find_preset_style_in_archive(&archive, stylesheet_id, preset)
}

pub(crate) fn paragraph_list(package: &IWorkPackage, storage_id: u64) -> Result<ParagraphList> {
    let storage = storage::locate(package, storage_id)?;
    native::resolved_paragraph_list(package, storage.style_id)
}

pub(crate) fn paragraph_lists(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<ParagraphListPlacement>> {
    let storage = storage::locate_boundaries(package, storage_id)?;
    storage
        .boundaries
        .into_iter()
        .map(|(index, style_id)| {
            Ok(ParagraphListPlacement::new(
                crate::text::ParagraphStart::from_utf16_index(index as usize)?,
                native::resolved_paragraph_list(package, style_id)?,
            ))
        })
        .collect()
}

pub(super) fn set_paragraph_lists(
    package: &mut IWorkPackage,
    storage_id: u64,
    placements: &[ParagraphListPlacement],
) -> Result<()> {
    let storage = storage::locate_boundaries(package, storage_id)?;
    let normalized = normalize_placements(storage_id, placements, &storage.paragraph_starts)?;
    if paragraph_lists(package, storage_id)? == normalized {
        return Ok(());
    }

    let first_style_id = storage
        .boundaries
        .first()
        .map(|entry| entry.1)
        .ok_or_else(|| Error::InvalidFormat("iWork list-style table is empty".to_owned()))?;
    let current = native::locate_style(package, first_style_id)?;
    let stylesheet_id = native::stylesheet_id(package, &current.style, first_style_id)?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != current.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {first_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    for &(_, style_id) in &storage.boundaries {
        let style = native::locate_style(package, style_id)?;
        if style.archive_name != current.archive_name
            || native::stylesheet_id(package, &style.style, style_id)? != stylesheet_id
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} mixes unrelated list styles"
            )));
        }
        native::resolved_paragraph_list(package, style_id)?;
    }

    let mut staged = package.clone();
    let mut style_ids = [None; 3];
    for placement in &normalized {
        let slot = placement.list.preset_index();
        if style_ids[slot].is_some() {
            continue;
        }
        let existing = native::find_preset_style(
            &staged,
            &current.archive_name,
            stylesheet_id,
            placement.list,
        )?;
        let target_style_id = match existing {
            Some(identifier) => identifier,
            None => next_object_identifier(&staged)?,
        };
        if existing.is_some() {
            register_style_reference(
                &mut staged,
                &storage.archive_name,
                &current.archive_name,
                target_style_id,
            )?;
        } else {
            let object = native::style_object(target_style_id, stylesheet_id, placement.list)?;
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
        style_ids[slot] = Some(target_style_id);
    }
    let boundaries = normalized
        .iter()
        .map(|placement| {
            Ok((
                placement.paragraph.utf16_index(),
                style_ids[placement.list.preset_index()].ok_or_else(|| {
                    Error::InvalidFormat("paragraph list preset was not registered".to_owned())
                })?,
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    let mut old_style_ids = storage
        .boundaries
        .iter()
        .map(|entry| entry.1)
        .collect::<Vec<_>>();
    storage::replace_boundaries(
        &mut staged,
        &storage,
        storage_id,
        &old_style_ids,
        &boundaries,
    )?;
    old_style_ids.sort_unstable();
    old_style_ids.dedup();
    for old_style_id in old_style_ids {
        unregister_owner_reference_if_unused(
            &mut staged,
            &storage.archive_name,
            &current.archive_name,
            old_style_id,
        )?;
    }
    if paragraph_lists(&staged, storage_id)? != normalized {
        return Err(Error::InvalidFormat(
            "iWork paragraph-list placement update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

fn normalize_placements(
    storage_id: u64,
    placements: &[ParagraphListPlacement],
    paragraph_starts: &[u32],
) -> Result<Vec<ParagraphListPlacement>> {
    if placements
        .first()
        .map(|placement| placement.paragraph.utf16_index())
        != Some(0)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph-list placements must begin at UTF-16 index zero"
        )));
    }
    let mut normalized = Vec::with_capacity(placements.len());
    let mut previous_index = None;
    for &placement in placements {
        let index = placement.paragraph.utf16_index();
        if previous_index.is_some_and(|previous| previous >= index) {
            return Err(Error::InvalidFormat(
                "paragraph-list placements must be strictly increasing".to_owned(),
            ));
        }
        if paragraph_starts.binary_search(&index).is_err() {
            return Err(Error::InvalidFormat(format!(
                "UTF-16 index {index} is not a paragraph start in iWork text storage {storage_id}"
            )));
        }
        if normalized
            .last()
            .is_some_and(|previous: &ParagraphListPlacement| previous.list == placement.list)
        {
            previous_index = Some(index);
            continue;
        }
        normalized.push(placement);
        previous_index = Some(index);
    }
    Ok(normalized)
}

pub(super) fn set_paragraph_list(
    package: &mut IWorkPackage,
    storage_id: u64,
    list: ParagraphList,
) -> Result<()> {
    let storage = storage::locate(package, storage_id)?;
    let current = native::locate_style(package, storage.style_id)?;
    if native::resolved_paragraph_list(package, storage.style_id).ok() == Some(list) {
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
        &storage,
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
    if native::resolved_paragraph_list(package, storage.style_id).ok() == Some(ParagraphList::None)
    {
        return Ok(false);
    }
    set_paragraph_list(package, storage_id, ParagraphList::None)?;
    Ok(true)
}

#[cfg(test)]
mod level_tests;
#[cfg(test)]
mod numbering_tests;
#[cfg(test)]
mod tests;
