//! Discovery and unambiguous matching of layout-owned image graphs.

use super::*;
use slide_create::graph::private_clone_object_ids;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;

pub(in crate::keynote::editor) fn template_image_roots(
    archive: &Archive,
    slide: &kn::SlideArchive,
) -> Vec<u64> {
    slide
        .owned_drawables
        .iter()
        .filter(|reference| {
            archive.object(reference.identifier).is_some_and(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == IMAGE_MESSAGE_TYPE)
            })
        })
        .map(|reference| reference.identifier)
        .collect()
}

pub(super) fn layout_image_roots(graph: &ObjectGraph, slide: &kn::SlideArchive) -> Vec<u64> {
    slide
        .owned_drawables
        .iter()
        .filter(|reference| {
            graph
                .objects
                .get(&reference.identifier)
                .is_some_and(|messages| {
                    messages
                        .iter()
                        .any(|message| message.type_ == IMAGE_MESSAGE_TYPE)
                })
        })
        .map(|reference| reference.identifier)
        .collect()
}

pub(super) fn reject_layout_movies(
    graph: &ObjectGraph,
    slide: &kn::SlideArchive,
    context: &str,
) -> Result<()> {
    if slide.owned_drawables.iter().any(|reference| {
        graph
            .objects
            .get(&reference.identifier)
            .is_some_and(|messages| {
                messages
                    .iter()
                    .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            })
    }) {
        return Err(Error::InvalidFormat(format!(
            "Keynote {context} contains a movie that cannot yet be materialized safely"
        )));
    }
    Ok(())
}

pub(super) fn current_layout_image_roots(
    graph: &ObjectGraph,
    current: &kn::SlideArchive,
) -> Result<Vec<u64>> {
    let Some(template) = current.template_slide.as_ref() else {
        return Ok(Vec::new());
    };
    let slide: kn::SlideArchive =
        graph.decode_type(template.identifier, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    reject_layout_movies(graph, &slide, "current layout")?;
    Ok(layout_image_roots(graph, &slide))
}

pub(super) fn match_live_image_roots(
    graph: &ObjectGraph,
    current: &kn::SlideArchive,
    templates: &[u64],
) -> Result<Vec<u64>> {
    let candidates = layout_image_roots(graph, current);
    let mut used = HashSet::new();
    templates
        .iter()
        .map(|template| {
            let signature = image_signature(graph, *template)?;
            let mut matches = Vec::new();
            for candidate in candidates.iter().copied() {
                if !used.contains(&candidate)
                    && image_signature(graph, candidate)? == signature
                {
                    matches.push(candidate);
                }
            }
            let [identifier] = matches.as_slice() else {
                return Err(Error::InvalidFormat(format!(
                    "Keynote layout image {template} matched {} live slide images; expected exactly one",
                    matches.len()
                )));
            };
            used.insert(*identifier);
            Ok(*identifier)
        })
        .collect()
}

fn image_signature(graph: &ObjectGraph, identifier: u64) -> Result<tsd::ImageArchive> {
    let mut image: tsd::ImageArchive =
        graph.decode_type(identifier, IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")?;
    image.super_.parent = None;
    image.super_.title = None;
    image.super_.caption = None;
    image.mask = None;
    image.flags = None;
    Ok(image)
}

pub(super) fn private_graph_union(
    archive: &Archive,
    roots: &[u64],
    context: &str,
) -> Result<Vec<u64>> {
    let selected = roots
        .iter()
        .map(|root| private_clone_object_ids(archive, [*root], context))
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .flatten()
        .collect::<HashSet<_>>();
    Ok(archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| selected.contains(identifier))
        .collect())
}
