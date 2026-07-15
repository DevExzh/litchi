//! Materialization lifecycle for layout-owned Keynote images.

use super::*;
use graph::{
    current_layout_image_roots, layout_image_roots, match_live_image_roots, private_graph_union,
    reject_layout_movies,
};
use metadata::{
    object_data_references, registered_subset, remapped_registered_subset,
    update_component_metadata,
};
use slide_create::graph::take_identifier;
use wire::{materialize_image_object, rewrite_slide_media_roots};

mod graph;
mod metadata;
mod wire;

pub(super) use graph::template_image_roots;
pub(super) use wire::materialize_cloned_images;

pub(super) fn materialize(
    package: &mut IWorkPackage,
    graph: &ObjectGraph,
    slide_id: u64,
    current: &kn::SlideArchive,
    target: &slide_create::layout::ResolvedLayout,
) -> Result<()> {
    let slide_archive_name = graph.archive_name(slide_id)?.to_owned();
    let slide_archive = package.archive(&slide_archive_name)?;
    let target_archive = package.archive(&target.archive_name)?;
    let target_roots = layout_image_roots(graph, &target.slide);
    let old_roots = current_layout_image_roots(graph, current)?;
    let live_roots = match_live_image_roots(graph, current, &old_roots)?;

    if target_roots.is_empty() && live_roots.is_empty() {
        reject_layout_movies(graph, &target.slide, "target layout")?;
        return Ok(());
    }
    reject_layout_movies(graph, &target.slide, "target layout")?;

    let old_graph_ids =
        private_graph_union(&slide_archive, &live_roots, "materialized layout image")?;
    if old_graph_ids.contains(&slide_id) {
        return Err(Error::InvalidFormat(
            "Keynote layout image graph reaches its owning slide".to_owned(),
        ));
    }
    let source_graph_ids =
        private_graph_union(&target_archive, &target_roots, "layout image template")?;

    let mut next_identifier = next_object_identifier(package)?;
    let mut remap = HashMap::with_capacity(source_graph_ids.len() + 1);
    remap.insert(target.slide_id, slide_id);
    for identifier in &source_graph_ids {
        remap.insert(*identifier, take_identifier(&mut next_identifier)?);
    }
    let new_roots = target_roots
        .iter()
        .map(|identifier| remap[identifier])
        .collect::<Vec<_>>();
    let mut cloned = source_graph_ids
        .iter()
        .map(|identifier| {
            let source = target_archive.object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote layout image object {identifier} disappeared"
                ))
            })?;
            let mut object = clone_slide_object(source, &remap)?;
            if target_roots.contains(identifier) {
                materialize_image_object(&mut object, slide_id)?;
            }
            Ok(object)
        })
        .collect::<Result<Vec<_>>>()?;

    let old_data_references = object_data_references(
        old_graph_ids
            .iter()
            .filter_map(|identifier| slide_archive.object(*identifier)),
    )?;
    let new_data_references = object_data_references(cloned.iter())?;
    let old_uuid_ids = registered_subset(package, &slide_archive_name, &old_graph_ids)?;
    let new_uuid_ids =
        remapped_registered_subset(package, &target.archive_name, &source_graph_ids, &remap)?;

    package.update_archive(&slide_archive_name, |archive| {
        for object in cloned.drain(..) {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;
    rewrite_slide_media_roots(
        package,
        &slide_archive_name,
        slide_id,
        &live_roots,
        &target_roots,
        &new_roots,
        current,
        &target.slide,
    )?;
    package.update_archive(&slide_archive_name, |archive| {
        for identifier in &old_graph_ids {
            archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote layout image object {identifier} is missing"
                ))
            })?;
        }
        Ok(())
    })?;

    update_component_metadata(
        package,
        &slide_archive_name,
        &target.archive_name,
        target.slide_id,
        &old_uuid_ids,
        &new_uuid_ids,
        &old_data_references,
        &new_data_references,
    )?;
    set_package_last_object_identifier(package, next_identifier - 1)?;
    for identifier in old_graph_ids {
        if package_references_object(package, identifier)? {
            return Err(Error::InvalidFormat(format!(
                "Removed Keynote layout image object {identifier} is still referenced"
            )));
        }
    }
    Ok(())
}
