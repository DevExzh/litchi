//! Package-metadata registration for slide styles.

use super::*;

pub(super) fn update_package_metadata(
    package: &mut IWorkPackage,
    slide_archive: &str,
    stylesheet_archive: &str,
    remove_style_id: Option<u64>,
    new_style_id: Option<u64>,
) -> Result<()> {
    let slide_component = component_identifier_for_entry(package, slide_archive)?;
    let stylesheet_component = component_identifier_for_entry(package, stylesheet_archive)?;
    if let Some(component) = stylesheet_component {
        if let Some(old) = remove_style_id {
            remove_component_object_uuids(package, component, &[old])?;
            remove_component_external_references_to_object(package, component, old)?;
        }
        if let Some(new_style_id) = new_style_id {
            add_component_object_uuids(package, component, &[new_style_id])?;
        }
    }
    if let (Some(source), Some(target), Some(new_style_id)) =
        (slide_component, stylesheet_component, new_style_id)
        && source != target
    {
        add_component_external_reference(package, source, target, new_style_id)?;
    }
    Ok(())
}

pub(super) fn ensure_slide_style_external_reference(
    package: &mut IWorkPackage,
    slide_archive: &str,
    stylesheet_archive: &str,
    style_id: u64,
) -> Result<()> {
    if let (Some(source), Some(target)) = (
        component_identifier_for_entry(package, slide_archive)?,
        component_identifier_for_entry(package, stylesheet_archive)?,
    ) && source != target
    {
        add_component_external_reference(package, source, target, style_id)?;
    }
    Ok(())
}
