//! Per-paragraph bullet and number label-color CRUD.

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::types::{ParagraphList, ParagraphListLabelColor};
use super::variation::{
    effective_style_id, paragraph_boundaries_with_style, style_isolated_to_paragraph,
};
use super::{native, storage};

pub(crate) fn paragraph_list_label_color(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListLabelColor> {
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if native::resolved_paragraph_list(package, style_id)? == ParagraphList::None {
        return Err(Error::InvalidFormat(format!(
            "paragraph at UTF-16 index {} in iWork text storage {storage_id} is not a list",
            paragraph.utf16_index()
        )));
    }
    native::effective_label_color(package, style_id)
}

pub(in crate::text) fn set_paragraph_list_label_color(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    color: ParagraphListLabelColor,
) -> Result<()> {
    if paragraph_list_label_color(package, storage_id, paragraph)? == color {
        return Ok(());
    }
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    let style = native::locate_style(package, style_id)?;
    let stylesheet_id = native::stylesheet_id(package, &style.style, style_id)?;
    if object_archive_name(package, stylesheet_id)? != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let can_update_in_place = style.style.super_.parent.is_some()
        && style_isolated_to_paragraph(&boundaries, paragraph)?
        && native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    if can_update_in_place {
        native::replace_direct_label_color(&mut staged, &style, color)?;
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let variation =
            native::label_color_variation_object(new_style_id, style_id, stylesheet_id, color)?;
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
            &boundaries,
            storage_id,
            &old_style_ids,
            &replacements,
        )?;
        set_package_last_object_identifier(&mut staged, new_style_id)?;
    }
    if paragraph_list_label_color(&staged, storage_id, paragraph)? != color {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-label color update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_label_color(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<bool> {
    if paragraph_list_label_color(package, storage_id, paragraph)?
        == ParagraphListLabelColor::Automatic
    {
        return Ok(false);
    }
    set_paragraph_list_label_color(
        package,
        storage_id,
        paragraph,
        ParagraphListLabelColor::Automatic,
    )?;
    collapse_or_clear_redundant_color(package, storage_id, paragraph)?;
    Ok(true)
}

fn collapse_or_clear_redundant_color(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<()> {
    let located = storage::locate_boundaries_with_archive(package, storage_id)?;
    let boundaries = &located.location;
    let storage_archive_name = boundaries.archive_name.clone();
    let style_id = effective_style_id(boundaries, paragraph)?;
    if !style_isolated_to_paragraph(boundaries, paragraph)?
        || !native::is_exclusive(package, style_id)?
    {
        return Ok(());
    }
    let style = native::locate_style(package, style_id)?;
    let parent_style_id = native::parent_style_id(&style.style, style_id)?;
    if native::effective_label_color(package, parent_style_id)?
        != ParagraphListLabelColor::Automatic
    {
        return Ok(());
    }
    let Some(override_count) = style.style.override_count else {
        return Ok(());
    };
    if style.style.font_color_null.is_none() && style.style.font_color.is_none() {
        return Ok(());
    }
    let mut staged = package.clone();
    if override_count == 1 {
        let stylesheet_id = native::stylesheet_id(&staged, &style.style, style_id)?;
        let replacements = paragraph_boundaries_with_style(boundaries, paragraph, parent_style_id)?;
        let old_style_ids = boundaries
            .boundaries
            .iter()
            .map(|entry| entry.1)
            .collect::<Vec<_>>();
        storage::replace_boundaries_with_archive(
            &mut staged,
            located,
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
            &storage_archive_name,
            &style.archive_name,
            style_id,
            Some(parent_style_id),
        )?;
        release_package_identifier_suffix(&mut staged, &[style_id])?;
    } else {
        native::remove_direct_label_color(&mut staged, &style)?;
    }
    if paragraph_list_label_color(&staged, storage_id, paragraph)?
        != ParagraphListLabelColor::Automatic
    {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-label color reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}
