//! Per-paragraph text-bullet glyph CRUD.

use crate::package_metadata::release_package_identifier_suffix;
use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::types::{ParagraphList, ParagraphListBullet};
use super::variation::{
    effective_style_id, paragraph_boundaries_with_style, style_isolated_to_paragraph,
};
use super::{levels, native, storage};

pub(crate) fn paragraph_list_bullet(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListBullet> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if native::resolved_paragraph_list(package, style_id)? != ParagraphList::Bullet {
        return Err(Error::InvalidFormat(format!(
            "paragraph at UTF-16 index {} in iWork text storage {storage_id} is not a text-bullet list",
            paragraph.utf16_index()
        )));
    }
    let strings = native::effective_bullet_strings(package, style_id)?;
    ParagraphListBullet::new(strings[usize::from(level.get())].clone())
}

pub(in crate::text) fn set_paragraph_list_bullet(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    bullet: &ParagraphListBullet,
) -> Result<()> {
    if paragraph_list_bullet(package, storage_id, paragraph)?.as_str() == bullet.as_str() {
        return Ok(());
    }
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    let style = native::locate_style(package, style_id)?;
    let stylesheet_id = native::stylesheet_id(package, &style.style, style_id)?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let mut strings = native::effective_bullet_strings(package, style_id)?;
    strings[usize::from(level.get())] = bullet.as_str().to_owned();

    let can_update_in_place = style.style.super_.parent.is_some()
        && style_isolated_to_paragraph(&boundaries, paragraph)?
        && native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    if can_update_in_place {
        native::replace_direct_bullet_strings(
            &mut staged,
            &style.archive_name,
            style_id,
            &strings,
        )?;
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let variation = native::variation_object(new_style_id, style_id, stylesheet_id, strings)?;
        insert_style_variation(
            &mut staged,
            &style.archive_name,
            stylesheet_id,
            style_id,
            new_style_id,
            variation,
        )?;
        register_private_style(
            &mut staged,
            &boundaries.archive_name,
            &style.archive_name,
            new_style_id,
        )?;
        let replacements = paragraph_boundaries_with_style(&boundaries, paragraph, new_style_id)?;
        let old_style_ids = boundaries
            .boundaries
            .iter()
            .map(|entry| entry.1)
            .collect::<Vec<_>>();
        storage::replace_boundaries(
            &mut staged,
            &boundaries.archive_name,
            storage_id,
            &old_style_ids,
            &replacements,
        )?;
        set_package_last_object_identifier(&mut staged, new_style_id)?;
    }
    if paragraph_list_bullet(&staged, storage_id, paragraph)? != *bullet {
        return Err(Error::InvalidFormat(
            "iWork paragraph text-bullet update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_bullet(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<bool> {
    if paragraph_list_bullet(package, storage_id, paragraph)?.as_str()
        == ParagraphListBullet::STANDARD
    {
        return Ok(false);
    }
    set_paragraph_list_bullet(
        package,
        storage_id,
        paragraph,
        &ParagraphListBullet::default(),
    )?;
    collapse_redundant_variation(package, storage_id, paragraph)?;
    Ok(true)
}

fn collapse_redundant_variation(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<()> {
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if !style_isolated_to_paragraph(&boundaries, paragraph)?
        || !native::is_exclusive(package, style_id)?
    {
        return Ok(());
    }
    let style = native::locate_style(package, style_id)?;
    let Some(parent_style_id) = style
        .style
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        .filter(|identifier| *identifier != 0)
    else {
        return Ok(());
    };
    if style.style.override_count != Some(1) || style.style.strings.is_empty() {
        return Ok(());
    }
    if native::effective_bullet_strings(package, style_id)?
        != native::effective_bullet_strings(package, parent_style_id)?
    {
        return Ok(());
    }
    let stylesheet_id = native::stylesheet_id(package, &style.style, style_id)?;
    let replacements = paragraph_boundaries_with_style(&boundaries, paragraph, parent_style_id)?;
    let old_style_ids = boundaries
        .boundaries
        .iter()
        .map(|entry| entry.1)
        .collect::<Vec<_>>();
    let mut staged = package.clone();
    storage::replace_boundaries(
        &mut staged,
        &boundaries.archive_name,
        storage_id,
        &old_style_ids,
        &replacements,
    )?;
    remove_style_variation(
        &mut staged,
        &style.archive_name,
        stylesheet_id,
        parent_style_id,
        style_id,
    )?;
    unregister_private_style(
        &mut staged,
        &boundaries.archive_name,
        &style.archive_name,
        style_id,
        Some(parent_style_id),
    )?;
    release_package_identifier_suffix(&mut staged, &[style_id])?;
    if paragraph_list_bullet(&staged, storage_id, paragraph)?.as_str()
        != ParagraphListBullet::STANDARD
    {
        return Err(Error::InvalidFormat(
            "iWork paragraph text-bullet reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}
