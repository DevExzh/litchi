//! Transactional named paragraph-style redefinition from one applied override chain.

use crate::{Error, IWorkPackage, Result};

use super::paragraph_alignment::{native, storage};
use super::paragraph_following_style::NamedParagraphStyle;
use super::paragraph_style_apply::{applied_named_paragraph_style, apply_named_paragraph_style};

pub(super) fn redefine_applied_named_paragraph_style(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<NamedParagraphStyle> {
    let storage = storage::locate(package, storage_id)?;
    let applied = applied_named_paragraph_style(package, storage_id)?;
    if !applied.has_overrides() {
        return Err(Error::InvalidFormat(
            "applied iWork named paragraph style has no overrides to redefine".to_owned(),
        ));
    }

    let mut current = storage.style_id;
    let target = applied.style().id();
    let mut variations = Vec::new();
    while current != target.get() {
        if variations.contains(&current) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph-style override chain contains a cycle at {current}"
            )));
        }
        let location = native::locate_style(package, current)?;
        if native::direct_overrides(&location.style, &location.message.data)?.is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {current} is not a supported variation"
            )));
        }
        variations.push(current);
        current = native::parent_style_id(&location.style, current)?;
    }

    let mut staged = package.clone();
    native::redefine_named_style(&mut staged, target.get(), &variations)?;
    let redefined = apply_named_paragraph_style(&mut staged, storage_id, target)?;
    let verified = applied_named_paragraph_style(&staged, storage_id)?;
    if verified.style() != &redefined || verified.has_overrides() {
        return Err(Error::InvalidFormat(
            "named iWork paragraph-style redefinition failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(redefined)
}
