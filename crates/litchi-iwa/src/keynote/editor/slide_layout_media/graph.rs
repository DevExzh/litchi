//! Discovery and unambiguous matching of layout-owned media graphs.

use super::*;
use slide_create::graph::private_clone_object_ids;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const LIVE_VIDEO_INFO_FIELD: u32 = 100;
const MOVIE_MEDIA_PLACEHOLDER_FLAG: u32 = 1;

#[derive(Debug, Default)]
pub(in crate::keynote::editor) struct LayoutMediaRoots {
    pub(super) images: Vec<u64>,
    pub(super) live_videos: Vec<u64>,
    pub(super) movie_placeholders: Vec<u64>,
    template_movies: Vec<u64>,
}

impl LayoutMediaRoots {
    pub(super) fn is_empty(&self) -> bool {
        self.images.is_empty() && self.live_videos.is_empty() && self.movie_placeholders.is_empty()
    }

    pub(super) fn identifiers(&self) -> impl Iterator<Item = u64> + '_ {
        self.images
            .iter()
            .chain(&self.live_videos)
            .chain(&self.movie_placeholders)
            .copied()
    }

    pub(in crate::keynote::editor) fn template_movies(&self) -> &[u64] {
        &self.template_movies
    }
}

pub(in crate::keynote::editor) fn layout_media_roots(
    graph: &ObjectGraph,
    slide: &kn::SlideArchive,
    context: &str,
) -> Result<LayoutMediaRoots> {
    media_roots(graph, slide, MoviePolicy::CollectTemplateMovies, context)
}

fn live_slide_media_candidates(
    graph: &ObjectGraph,
    slide: &kn::SlideArchive,
) -> Result<LayoutMediaRoots> {
    media_roots(
        graph,
        slide,
        MoviePolicy::IgnoreTemplateMovies,
        "live slide",
    )
}

#[derive(Clone, Copy)]
enum MoviePolicy {
    CollectTemplateMovies,
    IgnoreTemplateMovies,
}

fn media_roots(
    graph: &ObjectGraph,
    slide: &kn::SlideArchive,
    movie_policy: MoviePolicy,
    context: &str,
) -> Result<LayoutMediaRoots> {
    let mut roots = LayoutMediaRoots::default();
    for reference in &slide.owned_drawables {
        let Some(messages) = graph.objects.get(&reference.identifier) else {
            return Err(Error::InvalidFormat(format!(
                "Keynote {context} drawable {} is missing",
                reference.identifier
            )));
        };
        let is_image = messages
            .iter()
            .any(|message| message.type_ == IMAGE_MESSAGE_TYPE);
        let is_movie = messages
            .iter()
            .any(|message| message.type_ == MOVIE_MESSAGE_TYPE);
        if is_image && is_movie {
            return Err(Error::InvalidFormat(format!(
                "Keynote {context} drawable {} is both an image and a movie",
                reference.identifier
            )));
        }
        if is_image {
            roots.images.push(reference.identifier);
        } else if is_movie {
            let movie: tsd::MovieArchive =
                graph.decode_type(reference.identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
            if movie.is_live_video == Some(true) {
                roots.live_videos.push(reference.identifier);
            } else if is_movie_placeholder(&movie) {
                roots.movie_placeholders.push(reference.identifier);
            } else if matches!(movie_policy, MoviePolicy::CollectTemplateMovies) {
                roots.template_movies.push(reference.identifier);
            }
        }
    }
    Ok(roots)
}

pub(super) fn current_layout_media_roots(
    graph: &ObjectGraph,
    current: &kn::SlideArchive,
) -> Result<LayoutMediaRoots> {
    let Some(template) = current.template_slide.as_ref() else {
        return Ok(LayoutMediaRoots::default());
    };
    let slide: kn::SlideArchive =
        graph.decode_type(template.identifier, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    layout_media_roots(graph, &slide, "current layout")
}

pub(super) fn match_live_media_roots(
    graph: &ObjectGraph,
    current: &kn::SlideArchive,
    templates: &LayoutMediaRoots,
) -> Result<LayoutMediaRoots> {
    let candidates = live_slide_media_candidates(graph, current)?;
    Ok(LayoutMediaRoots {
        images: match_roots(
            &templates.images,
            &candidates.images,
            |identifier| image_signature(graph, identifier),
            "image",
        )?,
        live_videos: match_roots(
            &templates.live_videos,
            &candidates.live_videos,
            |identifier| live_video_signature(graph, identifier),
            "live video",
        )?,
        movie_placeholders: match_roots(
            &templates.movie_placeholders,
            &candidates.movie_placeholders,
            |identifier| movie_placeholder_signature(graph, identifier),
            "movie placeholder",
        )?,
        template_movies: Vec::new(),
    })
}

fn match_roots<T: PartialEq>(
    templates: &[u64],
    candidates: &[u64],
    signature: impl Fn(u64) -> Result<T>,
    kind: &str,
) -> Result<Vec<u64>> {
    let mut available = candidates
        .iter()
        .copied()
        .map(|identifier| Ok((identifier, signature(identifier)?)))
        .collect::<Result<Vec<_>>>()?;
    let mut matched = Vec::with_capacity(templates.len());
    for template in templates {
        let expected = signature(*template)?;
        let mut position = None;
        let mut match_count = 0usize;
        for (index, (_, candidate)) in available.iter().enumerate() {
            if candidate == &expected {
                match_count += 1;
                position = Some(index);
            }
        }
        if match_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote layout {kind} {template} matched {} live slide objects; expected exactly one",
                match_count
            )));
        }
        let position = position.ok_or_else(|| {
            Error::InvalidFormat("Keynote media match count lost its position".to_owned())
        })?;
        matched.push(available.remove(position).0);
    }
    Ok(matched)
}

fn image_signature(graph: &ObjectGraph, identifier: u64) -> Result<tsd::ImageArchive> {
    let mut image: tsd::ImageArchive =
        graph.decode_type(identifier, IMAGE_MESSAGE_TYPE, "TSD.ImageArchive")?;
    clear_instance_references(&mut image.super_);
    image.mask = None;
    image.flags = None;
    Ok(image)
}

#[derive(PartialEq)]
struct MovieSignature {
    movie: tsd::MovieArchive,
    extension_payloads: Vec<Vec<u8>>,
}

fn live_video_signature(graph: &ObjectGraph, identifier: u64) -> Result<MovieSignature> {
    let data = graph.message_data_type(identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let mut movie = tsd::MovieArchive::decode(data)?;
    if movie.is_live_video != Some(true) {
        return Err(Error::InvalidFormat(format!(
            "Keynote layout movie {identifier} is not a live video"
        )));
    }
    clear_instance_references(&mut movie.super_);
    Ok(MovieSignature {
        movie,
        extension_payloads: repeated_length_delimited_payloads(data, LIVE_VIDEO_INFO_FIELD)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect(),
    })
}

fn movie_placeholder_signature(graph: &ObjectGraph, identifier: u64) -> Result<MovieSignature> {
    let data = graph.message_data_type(identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let mut movie = tsd::MovieArchive::decode(data)?;
    if movie.is_live_video == Some(true) || !is_movie_placeholder(&movie) {
        return Err(Error::InvalidFormat(format!(
            "Keynote layout movie {identifier} is not a file-backed media placeholder"
        )));
    }
    clear_instance_references(&mut movie.super_);
    Ok(MovieSignature {
        movie,
        extension_payloads: repeated_length_delimited_payloads(data, LIVE_VIDEO_INFO_FIELD)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect(),
    })
}

pub(super) fn is_movie_placeholder(movie: &tsd::MovieArchive) -> bool {
    movie
        .flags
        .is_some_and(|flags| flags & MOVIE_MEDIA_PLACEHOLDER_FLAG != 0)
}

fn clear_instance_references(drawable: &mut tsd::DrawableArchive) {
    drawable.parent = None;
    drawable.title = None;
    drawable.caption = None;
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
