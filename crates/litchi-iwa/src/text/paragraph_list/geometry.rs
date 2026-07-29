//! Per-paragraph text-bullet size and baseline CRUD.

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::unregister_private_style;
use crate::text::style_registry::{object_archive_name, register_private_style};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::bullet::{
    effective_style_id, paragraph_boundaries_with_style, style_isolated_to_paragraph,
};
use super::types::{
    ParagraphList, ParagraphListBulletBaselineOffset, ParagraphListBulletGeometry,
    ParagraphListBulletScale,
};
use super::{levels, native, storage};

pub(crate) fn paragraph_list_bullet_geometry(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListBulletGeometry> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if native::resolved_paragraph_list(package, style_id)? != ParagraphList::Bullet {
        return Err(Error::InvalidFormat(format!(
            "paragraph at UTF-16 index {} in iWork text storage {storage_id} is not a text-bullet list",
            paragraph.utf16_index()
        )));
    }
    let geometries = native::effective_bullet_geometries(package, style_id)?;
    let geometry = geometries[usize::from(level.get())];
    Ok(ParagraphListBulletGeometry::new(
        ParagraphListBulletScale::from_ratio(
            geometry
                .scale
                .unwrap_or(ParagraphListBulletScale::ONE.ratio()),
        )?,
        ParagraphListBulletBaselineOffset::from_points(
            geometry
                .baseline_offset
                .unwrap_or(ParagraphListBulletBaselineOffset::ZERO.points()),
        )?,
    ))
}

pub(in crate::text) fn set_paragraph_list_bullet_geometry(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    geometry: ParagraphListBulletGeometry,
) -> Result<()> {
    if paragraph_list_bullet_geometry(package, storage_id, paragraph)? == geometry {
        return Ok(());
    }
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    let style = native::locate_style(package, style_id)?;
    let stylesheet_id = native::stylesheet_id(package, &style.style, style_id)?;
    if object_archive_name(package, stylesheet_id)? != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let mut geometries = native::effective_bullet_geometries(package, style_id)?;
    let target = &mut geometries[usize::from(level.get())];
    target.scale = Some(geometry.scale.ratio());
    target.baseline_offset = Some(geometry.baseline_offset.points());
    target.scale_with_text = Some(true);

    let can_update_in_place = style.style.super_.parent.is_some()
        && style_isolated_to_paragraph(&boundaries, paragraph)?
        && native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    if can_update_in_place {
        native::replace_direct_bullet_geometries(
            &mut staged,
            &style.archive_name,
            style_id,
            &geometries,
        )?;
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let variation =
            native::geometry_variation_object(new_style_id, style_id, stylesheet_id, geometries)?;
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
    if paragraph_list_bullet_geometry(&staged, storage_id, paragraph)? != geometry {
        return Err(Error::InvalidFormat(
            "iWork paragraph text-bullet geometry update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_bullet_geometry(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let standard = ParagraphListBulletGeometry::standard(level);
    if paragraph_list_bullet_geometry(package, storage_id, paragraph)? == standard {
        return Ok(false);
    }
    set_paragraph_list_bullet_geometry(package, storage_id, paragraph, standard)?;
    collapse_or_clear_redundant_geometry(package, storage_id, paragraph)?;
    Ok(true)
}

fn collapse_or_clear_redundant_geometry(
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
    let parent_style_id = native::parent_style_id(&style.style, style_id)?;
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let parent_geometries = native::effective_bullet_geometries(package, parent_style_id)?;
    let parent = parent_geometries[usize::from(level.get())];
    let standard = ParagraphListBulletGeometry::standard(level);
    if parent
        .scale
        .unwrap_or(ParagraphListBulletScale::ONE.ratio())
        != standard.scale.ratio()
        || parent
            .baseline_offset
            .unwrap_or(ParagraphListBulletBaselineOffset::ZERO.points())
            != standard.baseline_offset.points()
    {
        return Ok(());
    }
    let Some(override_count) = style.style.override_count else {
        return Ok(());
    };
    if style.style.geometries.is_empty() {
        return Ok(());
    }
    let mut staged = package.clone();
    if override_count == 1 {
        let stylesheet_id = native::stylesheet_id(&staged, &style.style, style_id)?;
        let replacements =
            paragraph_boundaries_with_style(&boundaries, paragraph, parent_style_id)?;
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
    } else {
        native::remove_direct_bullet_geometries(&mut staged, &style.archive_name, style_id)?;
    }
    if paragraph_list_bullet_geometry(&staged, storage_id, paragraph)? != standard {
        return Err(Error::InvalidFormat(
            "iWork paragraph text-bullet geometry reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}
