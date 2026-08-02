//! PackageMetadata registration for newly created slide components.

use super::graph::NoteSource;
use super::layout::ResolvedLayout;
use super::*;

#[allow(clippy::too_many_arguments)]
pub(super) fn register_created_slide(
    package: &mut IWorkPackage,
    layout: &ResolvedLayout,
    new_archive_name: &str,
    node_archive_name: &str,
    new_node_id: u64,
    new_slide_id: u64,
    remap: &HashMap<u64, u64>,
    note_source: &NoteSource,
) -> Result<()> {
    let Some(template_component) = component_identifier_for_entry(package, &layout.archive_name)?
    else {
        return Ok(());
    };
    let locator = new_archive_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Invalid component name {new_archive_name}"))
        })?;
    clone_component_registration(package, template_component, new_slide_id, locator, remap)?;
    let note_targets = note_source
        .object_ids
        .iter()
        .filter_map(|identifier| remap.get(identifier).copied())
        .collect::<Vec<_>>();
    add_component_object_uuids(package, new_slide_id, &note_targets)?;
    add_component_link(package, new_slide_id, template_component)?;

    let archive = package.archive(new_archive_name)?;
    let internal = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .collect::<HashSet<_>>();
    let external = archive
        .objects
        .iter()
        .flat_map(|object| &object.archive_info.message_infos)
        .flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        })
        .copied()
        .filter(|identifier| !internal.contains(identifier) && *identifier != layout.slide_id)
        .collect::<HashSet<_>>();
    for identifier in external {
        if let Some(component) = component_identifier_for_object_uuid(package, identifier)?
            && component != new_slide_id
        {
            add_component_external_reference(package, new_slide_id, component, identifier)?;
        }
    }

    if let Some(document_component) = component_identifier_for_entry(package, node_archive_name)? {
        add_component_object_uuids(package, document_component, &[new_node_id])?;
        add_component_link(package, document_component, new_slide_id)?;
    }
    Ok(())
}
