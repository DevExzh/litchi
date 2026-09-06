//! PackageMetadata dependency reconciliation for Keynote layout reassignment.

use super::*;

pub(super) struct SlideLayoutDependencies {
    old_external_object_references: HashSet<u64>,
    old_template_component: Option<u64>,
}

pub(super) fn capture_slide_layout_dependencies(
    package: &IWorkPackage,
    graph: &ObjectGraph,
    slide_id: u64,
    current: &kn::SlideArchive,
) -> Result<SlideLayoutDependencies> {
    let archive_name = graph.archive_name(slide_id)?;
    let old_external_object_references = package.with_parsed_archive(archive_name, |archive| {
        Ok(external_object_references(archive))
    })?;
    let old_template_component = if let Some(template) = current.template_slide.as_ref() {
        let archive_name = graph.archive_name(template.identifier)?;
        component_identifier_for_entry(package, archive_name)?
    } else {
        None
    };
    Ok(SlideLayoutDependencies {
        old_external_object_references,
        old_template_component,
    })
}

pub(super) fn reconcile_slide_layout_dependencies(
    package: &mut IWorkPackage,
    slide_archive_name: &str,
    target_archive_name: &str,
    target_slide_id: u64,
    previous: &SlideLayoutDependencies,
) -> Result<()> {
    let Some(slide_component) = component_identifier_for_entry(package, slide_archive_name)? else {
        return Ok(());
    };
    let Some(target_component) = component_identifier_for_entry(package, target_archive_name)?
    else {
        return Ok(());
    };
    let final_external_object_references = package
        .with_parsed_archive(slide_archive_name, |archive| {
            Ok(external_object_references(archive))
        })?;

    let mut final_external_components = HashSet::new();
    let mut final_object_references = final_external_object_references
        .iter()
        .copied()
        .collect::<Vec<_>>();
    final_object_references.sort_unstable();
    for identifier in final_object_references {
        // The template relationship is represented by a component-only edge;
        // retaining an object-bearing edge to that same template object would
        // differ from Keynote's native PackageMetadata shape.
        if identifier == target_slide_id {
            continue;
        }
        let Some(target_component) = component_identifier_for_object_uuid(package, identifier)?
        else {
            continue;
        };
        if target_component == slide_component {
            continue;
        }
        final_external_components.insert(target_component);
        add_component_external_reference(package, slide_component, target_component, identifier)?;
    }

    let mut stale_object_references = previous
        .old_external_object_references
        .iter()
        .filter(|identifier| !final_external_object_references.contains(identifier))
        .copied()
        .collect::<Vec<_>>();
    stale_object_references.sort_unstable();
    for identifier in stale_object_references {
        let Some(old_component) = component_identifier_for_object_uuid(package, identifier)? else {
            continue;
        };
        if old_component != slide_component {
            crate::package_metadata::remove_component_external_reference(
                package,
                slide_component,
                old_component,
                identifier,
            )?;
        }
    }

    if target_component != slide_component {
        add_component_link(package, slide_component, target_component)?;
    }
    if let Some(old_template_component) = previous.old_template_component
        && old_template_component != slide_component
        && old_template_component != target_component
        && !final_external_components.contains(&old_template_component)
    {
        crate::package_metadata::remove_component_link(
            package,
            slide_component,
            old_template_component,
        )?;
    }
    Ok(())
}

fn external_object_references(archive: &Archive) -> HashSet<u64> {
    let internal = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .collect::<HashSet<_>>();
    archive
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
        .filter(|identifier| !internal.contains(identifier))
        .collect()
}
