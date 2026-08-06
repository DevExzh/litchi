//! Per-paragraph list-label and text-gap indentation CRUD.

use crate::package_metadata::{
    next_object_identifier, release_package_identifier_suffix, set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::text::style_registry::{
    object_archive_name, register_private_style, unregister_private_style,
};
use crate::{Error, IWorkPackage, Result};

use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListIndentation, ParagraphListLabelIndent, ParagraphListTextGap,
};
use litchi_iwa_text::position::TextPosition;
use super::variation::{
    effective_style_id, paragraph_boundaries_with_style, style_isolated_to_paragraph,
};
use super::{levels, native, storage};

pub(crate) fn paragraph_list_indentation(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
) -> Result<ParagraphListIndentation> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    if native::resolved_paragraph_list(package, style_id)? == ParagraphList::None {
        return Err(Error::InvalidFormat(format!(
            "paragraph at UTF-16 index {} in iWork text storage {storage_id} is not a list",
            paragraph.utf16_index()
        )));
    }
    let indents = native::effective_list_indents(package, style_id)?;
    let text_indents = native::effective_list_text_indents(package, style_id)?;
    Ok(ParagraphListIndentation::new(
        ParagraphListLabelIndent::from_points(indents[usize::from(level.get())])?,
        ParagraphListTextGap::from_em(text_indents[usize::from(level.get())])?,
    ))
}

pub(in crate::text) fn set_paragraph_list_indentation(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
    indentation: ParagraphListIndentation,
) -> Result<()> {
    if paragraph_list_indentation(package, storage_id, paragraph)? == indentation {
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
    let mut indents = native::effective_list_indents(package, style_id)?;
    let mut text_indents = native::effective_list_text_indents(package, style_id)?;
    let level_index = usize::from(level.get());
    indents[level_index] = indentation.label_from_margin.points();
    text_indents[level_index] = indentation.text_from_label.em();

    let can_update_in_place = style.style.super_.parent.is_some()
        && style_isolated_to_paragraph(&boundaries, paragraph)?
        && native::is_exclusive(package, style_id)?;
    let mut staged = package.clone();
    if can_update_in_place {
        native::replace_direct_list_indentation_with_archive(
            &mut staged,
            located_style,
            &indents,
            &text_indents,
        )?;
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let variation = native::indentation_variation_object(
            new_style_id,
            style_id,
            stylesheet_id,
            indents,
            text_indents,
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
    if paragraph_list_indentation(&staged, storage_id, paragraph)? != indentation {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-indentation update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_indentation(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
) -> Result<bool> {
    let level = levels::paragraph_list_level(package, storage_id, paragraph)?;
    let boundaries = storage::locate_boundaries(package, storage_id)?;
    let style_id = effective_style_id(&boundaries, paragraph)?;
    let list = native::resolved_paragraph_list(package, style_id)?;
    let standard = ParagraphListIndentation::standard(list, level)?;
    if paragraph_list_indentation(package, storage_id, paragraph)? == standard {
        return Ok(false);
    }
    set_paragraph_list_indentation(package, storage_id, paragraph, standard)?;
    collapse_or_clear_redundant_indentation(package, storage_id, paragraph)?;
    Ok(true)
}

fn collapse_or_clear_redundant_indentation(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
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
    let list = native::resolved_paragraph_list(package, parent_style_id)?;
    let standard = ParagraphListIndentation::standard(list, level)?;
    let parent_indents = native::effective_list_indents(package, parent_style_id)?;
    let parent_text_indents = native::effective_list_text_indents(package, parent_style_id)?;
    let level_index = usize::from(level.get());
    if parent_indents[level_index] != standard.label_from_margin.points()
        || parent_text_indents[level_index] != standard.text_from_label.em()
    {
        return Ok(());
    }
    let Some(override_count) = style.style.override_count else {
        return Ok(());
    };
    if style.style.indents.is_empty() && style.style.text_indents.is_empty() {
        return Ok(());
    }
    let direct_override_count = u32::from(!style.style.indents.is_empty())
        + u32::from(!style.style.text_indents.is_empty());
    let mut staged = package.clone();
    if override_count == direct_override_count {
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
        native::remove_direct_list_indentation(&mut staged, &style)?;
    }
    if paragraph_list_indentation(&staged, storage_id, paragraph)? != standard {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-indentation reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}
