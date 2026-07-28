//! Per-paragraph text-bullet glyph CRUD.

use crate::package_metadata::release_package_identifier_suffix;
use crate::package_metadata::{next_object_identifier, set_package_last_object_identifier};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::storage::ListBoundaryStorage;
use super::types::{ParagraphList, ParagraphListBullet};
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

fn effective_style_id(storage: &ListBoundaryStorage, paragraph: ParagraphStart) -> Result<u64> {
    let target = paragraph.utf16_index();
    if storage.paragraph_starts.binary_search(&target).is_err() {
        return Err(Error::InvalidFormat(format!(
            "UTF-16 index {target} is not a paragraph start"
        )));
    }
    storage
        .boundaries
        .iter()
        .take_while(|entry| entry.0 <= target)
        .last()
        .map(|entry| entry.1)
        .ok_or_else(|| {
            Error::InvalidFormat("iWork text storage has no list-style boundary at zero".to_owned())
        })
}

fn style_isolated_to_paragraph(
    storage: &ListBoundaryStorage,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let target = paragraph.utf16_index();
    let target_boundary = storage
        .boundaries
        .binary_search_by_key(&target, |entry| entry.0)
        .is_ok();
    if !target_boundary {
        return Ok(false);
    }
    let Some(next_paragraph) = storage
        .paragraph_starts
        .iter()
        .copied()
        .find(|start| *start > target)
    else {
        return Ok(true);
    };
    Ok(storage
        .boundaries
        .binary_search_by_key(&next_paragraph, |entry| entry.0)
        .is_ok())
}

fn paragraph_boundaries_with_style(
    storage: &ListBoundaryStorage,
    paragraph: ParagraphStart,
    replacement_style_id: u64,
) -> Result<Vec<(u32, u64)>> {
    let target = paragraph.utf16_index();
    if storage.paragraph_starts.binary_search(&target).is_err() {
        return Err(Error::InvalidFormat(format!(
            "UTF-16 index {target} is not a paragraph start"
        )));
    }
    let mut result = Vec::with_capacity(storage.boundaries.len().saturating_add(2));
    for &start in &storage.paragraph_starts {
        let style_id = if start == target {
            replacement_style_id
        } else {
            storage
                .boundaries
                .iter()
                .take_while(|entry| entry.0 <= start)
                .last()
                .map(|entry| entry.1)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "iWork text storage has no list-style boundary at zero".to_owned(),
                    )
                })?
        };
        if result
            .last()
            .is_none_or(|previous: &(u32, u64)| previous.1 != style_id)
        {
            result.push((start, style_id));
        }
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn storage() -> ListBoundaryStorage {
        ListBoundaryStorage {
            archive_name: "Index/Test.iwa".to_owned(),
            boundaries: vec![(0, 10), (20, 11)],
            paragraph_starts: vec![0, 10, 20],
        }
    }

    #[test]
    fn changing_one_paragraph_splits_and_restores_boundaries() {
        assert_eq!(
            paragraph_boundaries_with_style(
                &storage(),
                ParagraphStart::from_utf16_index(10).unwrap(),
                99,
            )
            .unwrap(),
            [(0, 10), (10, 99), (20, 11)]
        );
    }

    #[test]
    fn only_single_paragraph_style_spans_are_isolated() {
        let storage = storage();
        assert!(!style_isolated_to_paragraph(&storage, ParagraphStart::ZERO).unwrap());
        assert!(
            !style_isolated_to_paragraph(&storage, ParagraphStart::from_utf16_index(10).unwrap())
                .unwrap()
        );
        assert!(
            style_isolated_to_paragraph(&storage, ParagraphStart::from_utf16_index(20).unwrap())
                .unwrap()
        );
    }
}
