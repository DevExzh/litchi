//! Standalone movie-object CRUD for Keynote slides.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::shapes::{DrawableGeometry, DrawableSize, geometry_from_drawable};

mod builds;
mod geometry;

use builds::*;
use geometry::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const BUILD_MESSAGE_TYPE: u32 = 8;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_BUILD_CHUNKS_FIELD: u32 = 43;
const MOVIE_DUPLICATE_OFFSET: f32 = 10.0;
const MOVIE_MEDIA_PLACEHOLDER_FLAG: u32 = 1;

/// Semantic role of a movie drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSlideMovieKind {
    /// Ordinary file-backed movie inserted by the user.
    File,
    /// File-backed replacement target materialized from a slide layout.
    MediaPlaceholder,
    /// Camera-backed live-video drawable.
    LiveVideo,
}

/// One movie drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideMovieInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub kind: KeynoteSlideMovieKind,
    pub movie_data_identifier: Option<u64>,
    pub poster_image_data_identifier: Option<u64>,
    pub geometry: DrawableGeometry,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

/// Result of removing one slide-owned movie and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideMovie {
    pub movie: KeynoteSlideMovieInfo,
    /// Assets culled because the removed movie held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

struct SlideMovieGraph {
    slide_id: u64,
    node_id: u64,
    archive_name: String,
    info: KeynoteSlideMovieInfo,
    object_ids: Vec<u64>,
    build_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
    data_references: Vec<(u64, u64)>,
}

impl KeynoteEditor {
    /// List movie drawables directly owned by one slide in drawable order.
    pub fn slide_movies(&self, slide_index: usize) -> Result<Vec<KeynoteSlideMovieInfo>> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let native: kn::SlideArchive =
            graph.decode_type(slide.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        native
            .owned_drawables
            .iter()
            .filter(|reference| {
                graph
                    .objects
                    .get(&reference.identifier)
                    .is_some_and(|messages| {
                        messages
                            .iter()
                            .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                    })
            })
            .map(|reference| movie_info(&graph, slide_index, reference.identifier))
            .collect()
    }

    /// Read typed geometry for one slide-owned movie.
    pub fn slide_movie_geometry(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(self
            .slide_movie_graph(slide_index, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update geometry on an ordinary file-backed movie while preserving unknown wire fields.
    pub fn set_slide_movie_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = self.require_file_movie(slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_movie_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_movie_geometry(slide_index, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Keynote movie geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace the video bytes referenced by one ordinary slide movie.
    ///
    /// Keynote duplicates share their media identifiers, so every movie sharing
    /// this identifier observes the replacement, matching native Keynote behavior.
    pub fn replace_slide_movie_data(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = self.require_file_movie(slide_index, drawable_object_id)?;
        let identifier = source.info.movie_data_identifier.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote movie {drawable_object_id} has no materialized video data"
            ))
        })?;
        self.replace_media(identifier, replacement)
    }

    /// Replace the poster image referenced by one ordinary slide movie.
    pub fn replace_slide_movie_poster(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = self.require_file_movie(slide_index, drawable_object_id)?;
        let identifier = source.info.poster_image_data_identifier.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote movie {drawable_object_id} has no materialized poster image"
            ))
        })?;
        self.replace_media(identifier, replacement)
    }

    /// Duplicate an ordinary file-backed movie using native shared-asset semantics.
    ///
    /// The movie, stand-in title/caption objects, and chunkless automatic movie
    /// builds receive fresh object identifiers and UUIDs. Embedded video and
    /// poster data remain shared, exactly as with Keynote's Duplicate command.
    pub fn duplicate_slide_movie(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
    ) -> Result<KeynoteSlideMovieInfo> {
        let source = self.require_file_movie(slide_index, source_drawable_object_id)?;
        let builds = self
            .slide_builds(slide_index)?
            .into_iter()
            .filter(|build| build.drawable_object_id == source_drawable_object_id)
            .collect::<Vec<_>>();
        let chunk_ids = builds
            .iter()
            .flat_map(|build| build.chunks.iter().map(|chunk| chunk.object_id))
            .collect::<Vec<_>>();

        let mut next_identifier = next_object_identifier(self.package())?;
        let mut remap = HashMap::with_capacity(
            source.object_ids.len() + source.build_ids.len() + chunk_ids.len(),
        );
        for identifier in source
            .object_ids
            .iter()
            .chain(&source.build_ids)
            .chain(&chunk_ids)
        {
            remap.insert(*identifier, take_movie_identifier(&mut next_identifier)?);
        }
        let new_drawable_id = remap[&source_drawable_object_id];
        let mut staged = self.package().clone();
        for identifier in &source.object_ids {
            let cloned = {
                let archive = self.package().archive(&source.archive_name)?;
                let object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote movie object {identifier} is missing during duplication"
                    ))
                })?;
                clone_slide_object(object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
        }
        let mut build_uuids = HashMap::with_capacity(builds.len());
        for build in &builds {
            let cloned = {
                let archive = self.package().archive(&source.archive_name)?;
                let object = archive.object(build.object_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote movie build {} is missing during duplication",
                        build.object_id
                    ))
                })?;
                clone_movie_build(
                    object,
                    remap[&build.object_id],
                    source_drawable_object_id,
                    new_drawable_id,
                )?
            };
            staged.update_archive(&source.archive_name, |archive| {
                archive.insert_object(cloned)
            })?;
            build_uuids.insert(build.object_id, new_build_uuid_and_seed().0);
        }
        for build in &builds {
            for chunk in &build.chunks {
                let cloned = {
                    let archive = self.package().archive(&source.archive_name)?;
                    let object = archive.object(chunk.object_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Keynote movie build chunk {} is missing during duplication",
                            chunk.object_id
                        ))
                    })?;
                    clone_movie_build_chunk(
                        object,
                        remap[&chunk.object_id],
                        build.object_id,
                        remap[&build.object_id],
                        build_uuids[&build.object_id],
                    )?
                };
                staged.update_archive(&source.archive_name, |archive| {
                    archive.insert_object(cloned)
                })?;
            }
        }
        offset_movie(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            MOVIE_DUPLICATE_OFFSET,
        )?;
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            None,
            Some(new_drawable_id),
        )?;
        let new_build_ids = source
            .build_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        let new_chunk_ids = chunk_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        append_slide_build_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            &new_build_ids,
            &new_chunk_ids,
        )?;
        let graph = ObjectGraph::read(self.package())?;
        let existing_chunk_count = self
            .slide_builds(slide_index)?
            .iter()
            .map(|build| build.chunks.len())
            .sum::<usize>();
        let new_event_count = existing_chunk_count
            .checked_add(new_chunk_ids.len())
            .ok_or_else(|| Error::InvalidFormat("Keynote build-event count overflow".to_owned()))?;
        patch_slide_build_cache(&mut staged, &graph, source.node_id, new_event_count)?;

        let new_uuid_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        let component =
            component_identifier_for_entry(&staged, &source.archive_name)?.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {} is not registered",
                    source.archive_name
                ))
            })?;
        add_component_object_uuids(&mut staged, component, &new_uuid_ids)?;
        for &(data_identifier, object_identifier) in &source.data_references {
            add_component_data_reference(
                &mut staged,
                component,
                data_identifier,
                remap[&object_identifier],
            )?;
        }
        set_package_last_object_identifier(&mut staged, next_identifier - 1)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slide_movies(slide_index)?
            .into_iter()
            .find(|movie| movie.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote movie duplication failed validation".to_owned())
            })?;
        let expected_position =
            source
                .info
                .geometry
                .position
                .map(|position| crate::shapes::DrawablePoint {
                    x: position.x + MOVIE_DUPLICATE_OFFSET,
                    y: position.y + MOVIE_DUPLICATE_OFFSET,
                });
        let cloned_build_count = verified
            .slide_builds(slide_index)?
            .iter()
            .filter(|build| build.drawable_object_id == new_drawable_id)
            .count();
        let cloned_chunk_count = verified
            .slide_builds(slide_index)?
            .iter()
            .filter(|build| build.drawable_object_id == new_drawable_id)
            .map(|build| build.chunks.len())
            .sum::<usize>();
        if created.kind != KeynoteSlideMovieKind::File
            || created.movie_data_identifier != source.info.movie_data_identifier
            || created.poster_image_data_identifier != source.info.poster_image_data_identifier
            || created.geometry.position != expected_position
            || cloned_build_count != builds.len()
            || cloned_chunk_count != chunk_ids.len()
        {
            return Err(Error::InvalidFormat(
                "Keynote movie duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Remove an ordinary file-backed movie, its private graph, and associated builds.
    pub fn remove_slide_movie(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideMovie> {
        let source = self.require_file_movie(slide_index, drawable_object_id)?;
        let mut working = self.clone();
        for build_id in &source.build_ids {
            working.remove_slide_build(slide_index, *build_id)?;
        }
        let source = working.require_file_movie(slide_index, drawable_object_id)?;

        let mut comments = IWorkDrawableCommentEditor::from_package(working.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
        let mut staged = comments.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        let component =
            component_identifier_for_entry(&staged, &source.archive_name)?.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {} is not registered",
                    source.archive_name
                ))
            })?;
        for &(data_identifier, object_identifier) in &source.data_references {
            remove_component_data_reference(
                &mut staged,
                component,
                data_identifier,
                object_identifier,
            )?;
        }
        for identifier in &source.object_ids {
            remove_object(&mut staged, &source.archive_name, *identifier)?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote movie object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, component, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let mut media = IWorkMediaEditor::from_package(staged)?;
        let mut removed_data_identifiers = Vec::new();
        let data_identifiers = source
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<HashSet<_>>();
        for identifier in data_identifiers {
            if media
                .asset(identifier)
                .is_some_and(|asset| !asset.is_referenced())
            {
                media.remove_unreferenced(identifier)?;
                removed_data_identifiers.push(identifier);
            }
        }
        removed_data_identifiers.sort_unstable();
        let staged = media.into_package();
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let remaining_media = verified.media_assets()?;
        if verified
            .slide_movies(slide_index)?
            .iter()
            .any(|movie| movie.drawable_object_id == drawable_object_id)
            || verified
                .slide_builds(slide_index)?
                .iter()
                .any(|build| build.drawable_object_id == drawable_object_id)
            || removed_data_identifiers.iter().any(|identifier| {
                remaining_media
                    .iter()
                    .any(|asset| asset.data_identifier == *identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "Keynote movie deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteSlideMovie {
            movie: source.info,
            removed_data_identifiers,
        })
    }

    fn require_file_movie(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<SlideMovieGraph> {
        let source = self.slide_movie_graph(slide_index, drawable_object_id)?;
        if source.info.kind != KeynoteSlideMovieKind::File {
            return Err(Error::ParseError(format!(
                "Keynote movie {drawable_object_id} is {:?}, not an ordinary file-backed movie",
                source.info.kind
            )));
        }
        Ok(source)
    }

    fn slide_movie_graph(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<SlideMovieGraph> {
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let native: kn::SlideArchive =
            graph.decode_type(slide.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        if !native
            .owned_drawables
            .iter()
            .any(|reference| reference.identifier == drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "Keynote movie {drawable_object_id} is not owned by slide {slide_index}"
            )));
        }
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        if graph.archive_name(drawable_object_id)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote movie {drawable_object_id} is outside slide component {archive_name}"
            )));
        }
        let archive = self.package().archive(&archive_name)?;
        let object_ids = slide_create::graph::private_clone_object_ids(
            &archive,
            [drawable_object_id],
            "slide movie",
        )?;
        if object_ids.contains(&slide.slide_id) {
            return Err(Error::InvalidFormat(
                "Keynote movie private graph reaches its owning slide".to_owned(),
            ));
        }
        let info = movie_info(&graph, slide_index, drawable_object_id)?;
        let build_ids = self
            .slide_builds(slide_index)?
            .into_iter()
            .filter(|build| build.drawable_object_id == drawable_object_id)
            .map(|build| build.object_id)
            .collect::<Vec<_>>();
        let component =
            component_identifier_for_entry(self.package(), &archive_name)?.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {archive_name} is not registered"
                ))
            })?;
        let registered = component_uuid_identifiers(self.package(), component)?.unwrap_or_default();
        let uuid_object_ids = object_ids
            .iter()
            .chain(&build_ids)
            .copied()
            .filter(|identifier| registered.contains(identifier))
            .collect::<Vec<_>>();
        let mut data_references = Vec::new();
        for identifier in &object_ids {
            let object = archive.object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote movie object {identifier} is missing"))
            })?;
            data_references.extend(
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .flat_map(|message| {
                        message
                            .data_references
                            .iter()
                            .chain(
                                message
                                    .field_infos
                                    .iter()
                                    .flat_map(|field| &field.data_references),
                            )
                            .map(|data| (*data, *identifier))
                    }),
            );
        }
        Ok(SlideMovieGraph {
            slide_id: slide.slide_id,
            node_id: slide.node_id,
            archive_name,
            info,
            object_ids,
            build_ids,
            uuid_object_ids,
            data_references,
        })
    }
}

fn movie_info(
    graph: &ObjectGraph,
    slide_index: usize,
    identifier: u64,
) -> Result<KeynoteSlideMovieInfo> {
    let movie: tsd::MovieArchive =
        graph.decode_type(identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let kind = if movie.is_live_video == Some(true) {
        KeynoteSlideMovieKind::LiveVideo
    } else if movie
        .flags
        .is_some_and(|flags| flags & MOVIE_MEDIA_PLACEHOLDER_FLAG != 0)
    {
        KeynoteSlideMovieKind::MediaPlaceholder
    } else {
        KeynoteSlideMovieKind::File
    };
    Ok(KeynoteSlideMovieInfo {
        slide_index,
        drawable_object_id: identifier,
        kind,
        movie_data_identifier: movie.movie_data.map(|reference| reference.identifier),
        poster_image_data_identifier: movie
            .poster_image_data
            .map(|reference| reference.identifier),
        geometry: geometry_from_drawable(&movie.super_)?,
        original_size: movie.original_size.map(drawable_size),
        natural_size: movie.natural_size.map(drawable_size),
    })
}

fn drawable_size(size: tsp::Size) -> DrawableSize {
    DrawableSize {
        width: size.width,
        height: size.height,
    }
}

fn take_movie_identifier(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}
