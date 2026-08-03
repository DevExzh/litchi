//! Per-paragraph numbered-list label-format CRUD.

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::types::{ParagraphList, ParagraphListNumberFormat};
use super::variation::{
    effective_style_id, paragraph_boundaries_with_style, style_isolated_to_paragraph,
};
use super::{levels, native, storage};

pub(crate) fn paragraph_list_number_format(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListNumberFormat> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if native::resolved_paragraph_list(package, style_id)? != ParagraphList::Numbered {
        return Err(Error::InvalidFormat(format!(
            "paragraph at UTF-16 index {} in iWork text storage {storage_id} is not a numbered list",
            paragraph.utf16_index()
        )));
    }
    let formats = native::effective_number_types(package, style_id)?;
    native::number_format_from_native(formats[usize::from(level.get())])
}

pub(in crate::text) fn set_paragraph_list_number_format(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    format: ParagraphListNumberFormat,
) -> Result<()> {
    if paragraph_list_number_format(package, storage_id, paragraph)? == format {
        return Ok(());
    }
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    let located_style = native::locate_style_with_archive(package, style_id)?;
    let style = &located_style.location;
    let stylesheet_id = native::stylesheet_id(package, &style.style, style_id)?;
    if object_archive_name(package, stylesheet_id)? != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let mut number_types = native::effective_number_types(package, style_id)?;
    number_types[usize::from(level.get())] = native::number_format_to_native(format);

    let can_update_in_place = style.style.super_.parent.is_some()
        && style_isolated_to_paragraph(&boundaries, paragraph)?
        && native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    if can_update_in_place {
        native::replace_direct_number_types_with_archive(
            &mut staged,
            located_style,
            &number_types,
        )?;
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let variation = native::number_format_variation_object(
            new_style_id,
            style_id,
            stylesheet_id,
            number_types,
        )?;
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
    if paragraph_list_number_format(&staged, storage_id, paragraph)? != format {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-number format update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_number_format(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<bool> {
    if paragraph_list_number_format(package, storage_id, paragraph)?
        == ParagraphListNumberFormat::DECIMAL
    {
        return Ok(false);
    }
    set_paragraph_list_number_format(
        package,
        storage_id,
        paragraph,
        ParagraphListNumberFormat::DECIMAL,
    )?;
    collapse_or_clear_redundant_format(package, storage_id, paragraph)?;
    Ok(true)
}

fn collapse_or_clear_redundant_format(
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
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let parent_formats = native::effective_number_types(package, parent_style_id)?;
    if native::number_format_from_native(parent_formats[usize::from(level.get())])?
        != ParagraphListNumberFormat::DECIMAL
    {
        return Ok(());
    }
    let Some(override_count) = style.style.override_count else {
        return Ok(());
    };
    if style.style.number_types.is_empty() {
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
        native::remove_direct_number_types(&mut staged, &style)?;
    }
    if paragraph_list_number_format(&staged, storage_id, paragraph)?
        != ParagraphListNumberFormat::DECIMAL
    {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-number format reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}
