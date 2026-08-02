//! Selection and application of named paragraph styles.

use crate::package_metadata::release_package_identifier_suffix;
use crate::shapes::remove_style_variation;
use crate::{Error, IWorkPackage, Result};

use super::paragraph_alignment::{native, storage};
use super::paragraph_following_style::{NamedParagraphStyle, ParagraphStyleId};
use super::style_registry::{
    object_archive_name, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};

/// The named style selected for one uniform text storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AppliedParagraphStyle {
    style: NamedParagraphStyle,
    has_overrides: bool,
}

impl AppliedParagraphStyle {
    fn new(style: NamedParagraphStyle, has_overrides: bool) -> Self {
        Self {
            style,
            has_overrides,
        }
    }

    /// Return the selected named style.
    pub const fn style(&self) -> &NamedParagraphStyle {
        &self.style
    }

    /// Whether the text storage has direct overrides on the selected style.
    pub const fn has_overrides(&self) -> bool {
        self.has_overrides
    }
}

pub(super) fn applied_named_paragraph_style(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<AppliedParagraphStyle> {
    let storage = storage::locate(package, storage_id)?;
    resolve_applied_style(package, storage.style_id)
}

pub(super) fn apply_named_paragraph_style(
    package: &mut IWorkPackage,
    storage_id: u64,
    target: ParagraphStyleId,
) -> Result<NamedParagraphStyle> {
    let storage = storage::locate(package, storage_id)?;
    let target_style = selectable_style(package, storage.style_id, target)?;
    let current = resolve_applied_style(package, storage.style_id)?;
    if current.style.id() == target && !current.has_overrides {
        return Ok(target_style);
    }

    let current_location = native::locate_style(package, storage.style_id)?;
    let current_is_named = current.style.id().get() == storage.style_id;
    let target_location = native::locate_style(package, target.get())?;
    let mut staged = package.clone();
    storage::patch_style_reference(
        &mut staged,
        &storage.archive_name,
        storage_id,
        storage.style_id,
        target.get(),
    )?;

    if !current_is_named && native::is_exclusive(package, storage.style_id)? {
        remove_exclusive_variation(
            &mut staged,
            &storage.archive_name,
            storage.style_id,
            current.style.id(),
            target,
        )?;
    } else {
        register_style_reference(
            &mut staged,
            &storage.archive_name,
            &target_location.archive_name,
            target.get(),
        )?;
        unregister_owner_reference_if_unused(
            &mut staged,
            &storage.archive_name,
            &current_location.archive_name,
            storage.style_id,
        )?;
    }

    let applied = applied_named_paragraph_style(&staged, storage_id)?;
    if applied.style.id() != target || applied.has_overrides {
        return Err(Error::InvalidFormat(
            "named iWork paragraph-style application failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(target_style)
}

fn resolve_applied_style(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<AppliedParagraphStyle> {
    let selectable = native::named_paragraph_styles(package, first_style_id)?;
    let mut current = first_style_id;
    let mut visited = Vec::new();
    loop {
        if let Some(style) = selectable.iter().find(|style| style.id().get() == current) {
            return Ok(AppliedParagraphStyle::new(
                style.clone(),
                current != first_style_id,
            ));
        }
        if visited.contains(&current) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph-style selection contains an inheritance cycle at {current}"
            )));
        }
        visited.push(current);
        let location = native::locate_style(package, current)?;
        if native::direct_overrides(&location.style, &location.message.data)?.is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {current} is neither selectable nor a supported variation"
            )));
        }
        current = native::parent_style_id(&location.style, current)?;
    }
}

fn selectable_style(
    package: &IWorkPackage,
    first_style_id: u64,
    target: ParagraphStyleId,
) -> Result<NamedParagraphStyle> {
    native::named_paragraph_styles(package, first_style_id)?
        .into_iter()
        .find(|style| style.id() == target)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {} is not a named style in this text stylesheet",
                target.get()
            ))
        })
}

fn remove_exclusive_variation(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    style_id: u64,
    base_style_id: ParagraphStyleId,
    target: ParagraphStyleId,
) -> Result<()> {
    let mut current_id = style_id;
    let mut current_location = native::locate_style(package, current_id)?;
    loop {
        let parent_style_id = native::parent_style_id(&current_location.style, current_id)?;
        let stylesheet_id = native::stylesheet_id(&current_location.style, current_id)?;
        let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
        if stylesheet_archive_name != current_location.archive_name {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph-style variation {current_id} is not stored with stylesheet {stylesheet_id}"
            )));
        }
        remove_style_variation(
            package,
            &current_location.archive_name,
            stylesheet_id,
            parent_style_id,
            current_id,
        )?;
        unregister_private_style(
            package,
            owner_archive_name,
            &current_location.archive_name,
            current_id,
            Some(target.get()),
        )?;
        release_package_identifier_suffix(package, &[current_id])?;
        if parent_style_id == base_style_id.get()
            || !native::is_unreferenced(package, parent_style_id)?
        {
            return Ok(());
        }
        current_id = parent_style_id;
        current_location = native::locate_style(package, current_id)?;
        if native::direct_overrides(&current_location.style, &current_location.message.data)?
            .is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph-style variation chain contains unsupported style {current_id}"
            )));
        }
    }
}
