//! Composable native paragraph-style CRUD shared by Pages, Numbers, and Keynote.

mod native;
mod storage;

use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::{Error, IWorkPackage, Result};

use self::native::ParagraphStyleOverrides;
use super::style::{ParagraphLineSpacing, ParagraphSpacing, TextAlignment};

#[derive(Debug, Clone, Copy)]
enum ParagraphProperty {
    Alignment(TextAlignment),
    LineSpacing(ParagraphLineSpacing),
    Spacing(ParagraphSpacing),
}

#[derive(Debug, Clone, Copy)]
enum ParagraphPropertyKind {
    Alignment,
    LineSpacing,
    Spacing,
}

pub(super) fn paragraph_alignment(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextAlignment> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_alignment(package, storage.style_id)
}

pub(super) fn set_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
    alignment: TextAlignment,
) -> Result<()> {
    if paragraph_alignment(package, storage_id)? == alignment {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::Alignment(alignment))
}

pub(super) fn reset_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::Alignment)
}

pub(super) fn paragraph_line_spacing(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphLineSpacing> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_line_spacing(package, storage.style_id)
}

pub(super) fn set_paragraph_line_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
    spacing: ParagraphLineSpacing,
) -> Result<()> {
    if paragraph_line_spacing(package, storage_id)? == spacing {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::LineSpacing(spacing))
}

pub(super) fn reset_paragraph_line_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::LineSpacing)
}

pub(super) fn paragraph_spacing(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphSpacing> {
    let storage = storage::locate(package, storage_id)?;
    native::inherited_spacing(package, storage.style_id)
}

pub(super) fn set_paragraph_spacing(
    package: &mut IWorkPackage,
    storage_id: u64,
    spacing: ParagraphSpacing,
) -> Result<()> {
    if paragraph_spacing(package, storage_id)? == spacing {
        return Ok(());
    }
    set_property(package, storage_id, ParagraphProperty::Spacing(spacing))
}

pub(super) fn reset_paragraph_spacing(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    reset_property(package, storage_id, ParagraphPropertyKind::Spacing)
}

fn set_property(
    package: &mut IWorkPackage,
    storage_id: u64,
    property: ParagraphProperty,
) -> Result<()> {
    let storage = storage::locate(package, storage_id)?;
    let style = native::locate_style(package, storage.style_id)?;
    let stylesheet_id = native::stylesheet_id(&style.style, storage.style_id)?;
    let stylesheet_archive_name = native::object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != style.archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {} is not stored with stylesheet {stylesheet_id}",
            storage.style_id
        )));
    }

    if let Some(mut overrides) = native::direct_overrides(&style.style, &style.message.data)?
        && native::is_exclusive(package, storage.style_id)?
    {
        apply_property(&mut overrides, property);
        let parent_style_id = native::parent_style_id(&style.style, storage.style_id)?;
        let replacement =
            native::variation_object(storage.style_id, parent_style_id, stylesheet_id, overrides)?;
        let mut staged = package.clone();
        native::replace_variation(
            &mut staged,
            &style.archive_name,
            storage.style_id,
            replacement,
        )?;
        validate_property(&staged, storage_id, property)?;
        *package = staged;
        return Ok(());
    }

    let new_style_id = next_object_identifier(package)?;
    let mut overrides = ParagraphStyleOverrides::default();
    apply_property(&mut overrides, property);
    let new_style =
        native::variation_object(new_style_id, storage.style_id, stylesheet_id, overrides)?;
    let mut staged = package.clone();
    storage::patch_style_reference(
        &mut staged,
        &storage.archive_name,
        storage_id,
        storage.style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        &style.archive_name,
        stylesheet_id,
        storage.style_id,
        new_style_id,
        new_style,
    )?;
    register_new_style(
        &mut staged,
        &storage.archive_name,
        &style.archive_name,
        new_style_id,
    )?;
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    validate_property(&staged, storage_id, property)?;
    *package = staged;
    Ok(())
}

fn reset_property(
    package: &mut IWorkPackage,
    storage_id: u64,
    kind: ParagraphPropertyKind,
) -> Result<bool> {
    let storage = storage::locate(package, storage_id)?;
    let style = native::locate_style(package, storage.style_id)?;
    let Some(mut overrides) = native::direct_overrides(&style.style, &style.message.data)? else {
        return Ok(false);
    };
    if !has_property(overrides, kind) || !native::is_exclusive(package, storage.style_id)? {
        return Ok(false);
    }
    clear_property(&mut overrides, kind);
    let parent_style_id = native::parent_style_id(&style.style, storage.style_id)?;
    let stylesheet_id = native::stylesheet_id(&style.style, storage.style_id)?;
    let expected = inherited_property(package, parent_style_id, kind)?;
    let mut staged = package.clone();
    if overrides.is_empty() {
        storage::patch_style_reference(
            &mut staged,
            &storage.archive_name,
            storage_id,
            storage.style_id,
            parent_style_id,
        )?;
        remove_style_variation(
            &mut staged,
            &style.archive_name,
            stylesheet_id,
            parent_style_id,
            storage.style_id,
        )?;
        unregister_removed_style(
            &mut staged,
            &storage.archive_name,
            &style.archive_name,
            storage.style_id,
            parent_style_id,
        )?;
        release_package_identifier_suffix(&mut staged, &[storage.style_id])?;
    } else {
        let replacement =
            native::variation_object(storage.style_id, parent_style_id, stylesheet_id, overrides)?;
        native::replace_variation(
            &mut staged,
            &style.archive_name,
            storage.style_id,
            replacement,
        )?;
    }
    validate_expected_property(&staged, storage_id, expected)?;
    *package = staged;
    Ok(true)
}

fn apply_property(overrides: &mut ParagraphStyleOverrides, property: ParagraphProperty) {
    match property {
        ParagraphProperty::Alignment(alignment) => overrides.alignment = Some(alignment),
        ParagraphProperty::LineSpacing(spacing) => overrides.line_spacing = Some(spacing),
        ParagraphProperty::Spacing(spacing) => {
            overrides.space_before = Some(spacing.before);
            overrides.space_after = Some(spacing.after);
        },
    }
}

fn has_property(overrides: ParagraphStyleOverrides, kind: ParagraphPropertyKind) -> bool {
    match kind {
        ParagraphPropertyKind::Alignment => overrides.alignment.is_some(),
        ParagraphPropertyKind::LineSpacing => overrides.line_spacing.is_some(),
        ParagraphPropertyKind::Spacing => {
            overrides.space_before.is_some() || overrides.space_after.is_some()
        },
    }
}

fn clear_property(overrides: &mut ParagraphStyleOverrides, kind: ParagraphPropertyKind) {
    match kind {
        ParagraphPropertyKind::Alignment => overrides.alignment = None,
        ParagraphPropertyKind::LineSpacing => overrides.line_spacing = None,
        ParagraphPropertyKind::Spacing => {
            overrides.space_before = None;
            overrides.space_after = None;
        },
    }
}

fn inherited_property(
    package: &IWorkPackage,
    style_id: u64,
    kind: ParagraphPropertyKind,
) -> Result<ParagraphProperty> {
    match kind {
        ParagraphPropertyKind::Alignment => Ok(ParagraphProperty::Alignment(
            native::inherited_alignment(package, style_id)?,
        )),
        ParagraphPropertyKind::LineSpacing => Ok(ParagraphProperty::LineSpacing(
            native::inherited_line_spacing(package, style_id)?,
        )),
        ParagraphPropertyKind::Spacing => Ok(ParagraphProperty::Spacing(
            native::inherited_spacing(package, style_id)?,
        )),
    }
}

fn validate_property(
    package: &IWorkPackage,
    storage_id: u64,
    expected: ParagraphProperty,
) -> Result<()> {
    validate_expected_property(package, storage_id, expected).map_err(|_| {
        Error::InvalidFormat("iWork paragraph-style update failed validation".to_owned())
    })
}

fn validate_expected_property(
    package: &IWorkPackage,
    storage_id: u64,
    expected: ParagraphProperty,
) -> Result<()> {
    let matches = match expected {
        ParagraphProperty::Alignment(alignment) => {
            paragraph_alignment(package, storage_id)? == alignment
        },
        ParagraphProperty::LineSpacing(spacing) => {
            paragraph_line_spacing(package, storage_id)? == spacing
        },
        ParagraphProperty::Spacing(spacing) => paragraph_spacing(package, storage_id)? == spacing,
    };
    if matches {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "iWork paragraph-style property does not match its expected value".to_owned(),
        ))
    }
}

fn register_new_style(
    package: &mut IWorkPackage,
    storage_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    add_component_object_uuids(package, style_component, &[style_id])?;
    if let Some(storage_component) = component_identifier_for_entry(package, storage_archive_name)?
        && storage_component != style_component
    {
        add_component_external_reference(package, storage_component, style_component, style_id)?;
    }
    Ok(())
}

fn unregister_removed_style(
    package: &mut IWorkPackage,
    storage_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
    parent_style_id: u64,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    remove_component_object_uuids(package, style_component, &[style_id])?;
    remove_component_external_references_to_object(package, style_component, style_id)?;
    if let Some(storage_component) = component_identifier_for_entry(package, storage_archive_name)?
        && storage_component != style_component
    {
        add_component_external_reference(
            package,
            storage_component,
            style_component,
            parent_style_id,
        )?;
    }
    Ok(())
}

#[cfg(test)]
mod tests;
