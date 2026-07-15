//! Standalone movie-object CRUD for Keynote slides.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::shapes::{DrawableGeometry, DrawablePoint, DrawableSize, geometry_from_drawable};
use std::time::Duration;

mod builds;
pub(in crate::keynote::editor) mod geometry;
pub(in crate::keynote::editor) mod graph;

use builds::*;
use geometry::*;
use graph::*;

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
    /// Independently positioned audio clip stored in a movie archive.
    Audio,
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

/// Typed layout and playback metadata for a newly created Keynote movie.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteSlideMovieOptions {
    /// Top-left position of the movie on the slide, in points.
    pub position: DrawablePoint,
    /// Displayed movie size on the slide, in points.
    pub size: DrawableSize,
    /// Untransformed media dimensions reported to Keynote, in points.
    pub natural_size: DrawableSize,
    /// Playable duration of the source movie.
    pub duration: Duration,
}

impl KeynoteSlideMovieOptions {
    /// Create options whose displayed and natural dimensions are identical.
    pub const fn new(position: DrawablePoint, size: DrawableSize, duration: Duration) -> Self {
        Self {
            position,
            size,
            natural_size: size,
            duration,
        }
    }

    /// Set dimensions independent of the displayed size.
    #[must_use]
    pub const fn with_natural_size(mut self, natural_size: DrawableSize) -> Self {
        self.natural_size = natural_size;
        self
    }
}

/// Result of removing one slide-owned movie and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideMovie {
    pub movie: KeynoteSlideMovieInfo,
    /// Assets culled because the removed movie held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

pub(in crate::keynote::editor) struct SlideMovieGraph {
    pub(in crate::keynote::editor) slide_id: u64,
    node_id: u64,
    pub(in crate::keynote::editor) archive_name: String,
    pub(in crate::keynote::editor) info: KeynoteSlideMovieInfo,
    pub(in crate::keynote::editor) object_ids: Vec<u64>,
    build_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
    data_references: Vec<(u64, u64)>,
}

impl KeynoteEditor {
    /// List movie drawables directly owned by one slide in drawable order.
    pub fn slide_movies(&self, slide_index: usize) -> Result<Vec<KeynoteSlideMovieInfo>> {
        Ok(self
            .slide_media_infos(slide_index)?
            .into_iter()
            .filter(|movie| movie.kind != KeynoteSlideMovieKind::Audio)
            .collect())
    }

    pub(in crate::keynote::editor) fn slide_media_infos(
        &self,
        slide_index: usize,
    ) -> Result<Vec<KeynoteSlideMovieInfo>> {
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

    /// Add an independently editable, file-backed movie to a slide.
    ///
    /// The movie, poster, title/caption stand-ins, automatic playback build,
    /// component registrations, UUIDs, and package media records are built from
    /// typed values. No source drawable or package template is copied.
    pub fn add_slide_movie(
        &mut self,
        slide_index: usize,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: KeynoteSlideMovieOptions,
    ) -> Result<KeynoteSlideMovieInfo> {
        let (geometry, duration_seconds) = movie_creation_values(options)?;
        let context = movie_creation_context(self, slide_index)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(self.package())?)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let movie_asset = media.insert_unreferenced(preferred_movie_filename, movie_data)?;
        if movie_asset.media_type != crate::MediaType::Video {
            return Err(Error::ParseError(format!(
                "Keynote slide movies require video data, not {}",
                movie_asset.media_type.name()
            )));
        }
        let poster_asset = media.insert_unreferenced(preferred_poster_filename, poster_data)?;
        if poster_asset.media_type != crate::MediaType::Image {
            return Err(Error::ParseError(format!(
                "Keynote movie posters require image data, not {}",
                poster_asset.media_type.name()
            )));
        }

        let mut staged = media.into_package();
        let objects = movie_objects(
            ids,
            context.slide_id,
            context.style_id,
            movie_asset.data_identifier,
            poster_asset.data_identifier,
            geometry,
            options.natural_size,
            duration_seconds,
        )?;
        staged.update_archive(&context.archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_slide_drawable_references(
            &mut staged,
            &context.archive_name,
            context.slide_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, context.component_id, &ids.all())?;
        for data_identifier in [movie_asset.data_identifier, poster_asset.data_identifier] {
            add_component_data_reference(
                &mut staged,
                context.component_id,
                data_identifier,
                ids.drawable,
            )?;
        }
        add_component_external_reference(
            &mut staged,
            context.component_id,
            context.stylesheet_component_id,
            context.style_id,
        )?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let mut verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slide_movies(slide_index)?
            .into_iter()
            .find(|movie| movie.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote movie creation failed validation".to_owned())
            })?;
        let created_graph = verified.slide_movie_graph(slide_index, ids.drawable)?;
        if created.kind != KeynoteSlideMovieKind::File
            || created.movie_data_identifier != Some(movie_asset.data_identifier)
            || created.poster_image_data_identifier != Some(poster_asset.data_identifier)
            || created.geometry != geometry
            || created.original_size != Some(options.natural_size)
            || created.natural_size != Some(options.natural_size)
            || created_graph.object_ids != ids.all()
            || verified.extract_media(movie_asset.data_identifier)? != movie_data
            || verified.extract_media(poster_asset.data_identifier)? != poster_data
        {
            return Err(Error::InvalidFormat(
                "Keynote movie creation produced an inconsistent graph".to_owned(),
            ));
        }

        let build = verified.add_slide_build(
            slide_index,
            ids.drawable,
            KeynoteBuildSettings::movie_start(),
        )?;
        if build.drawable_object_id != ids.drawable || build.chunks.len() != 1 {
            return Err(Error::InvalidFormat(
                "Keynote movie creation produced an inconsistent playback build".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
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
        self.remove_slide_media(slide_index, drawable_object_id, KeynoteSlideMovieKind::File)
    }

    pub(in crate::keynote::editor) fn remove_slide_media(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        expected_kind: KeynoteSlideMovieKind,
    ) -> Result<RemovedKeynoteSlideMovie> {
        let source =
            self.require_slide_media_kind(slide_index, drawable_object_id, expected_kind)?;
        let mut working = self.clone();
        for build_id in &source.build_ids {
            working.remove_slide_build(slide_index, *build_id)?;
        }
        let source =
            working.require_slide_media_kind(slide_index, drawable_object_id, expected_kind)?;

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
            .slide_media_infos(slide_index)?
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

    fn require_slide_media_kind(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
        expected_kind: KeynoteSlideMovieKind,
    ) -> Result<SlideMovieGraph> {
        let source = self.slide_movie_graph(slide_index, drawable_object_id)?;
        if source.info.kind != expected_kind {
            return Err(Error::ParseError(format!(
                "Keynote media {drawable_object_id} is {:?}, not {expected_kind:?}",
                source.info.kind
            )));
        }
        Ok(source)
    }

    pub(in crate::keynote::editor) fn slide_movie_graph(
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
    } else if movie.audio_only == Some(true) {
        KeynoteSlideMovieKind::Audio
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;

    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-movie";
    const REPLACEMENT_MOVIE: &[u8] = b"\0\0\0\x18ftypqt  replacement-movie";
    const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-poster";
    const REPLACEMENT_POSTER: &[u8] = b"GIF89areplacement-poster";
    const POSITION: DrawablePoint = DrawablePoint { x: 100.0, y: 120.0 };
    const DISPLAY_SIZE: DrawableSize = DrawableSize {
        width: 640.0,
        height: 360.0,
    };
    const NATURAL_SIZE: DrawableSize = DrawableSize {
        width: 1_280.0,
        height: 720.0,
    };

    fn options() -> KeynoteSlideMovieOptions {
        KeynoteSlideMovieOptions::new(POSITION, DISPLAY_SIZE, Duration::from_millis(1_250))
            .with_natural_size(NATURAL_SIZE)
    }

    #[test]
    fn scratch_presentation_supports_movie_crud_without_a_source_drawable() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch movie")
            .subtitle("No embedded package")
            .build()
            .unwrap();

        assert!(editor.slide_movies(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        let created = editor
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        assert_eq!(created.kind, KeynoteSlideMovieKind::File);
        assert_eq!(created.original_size, Some(NATURAL_SIZE));
        assert_eq!(created.natural_size, Some(NATURAL_SIZE));
        assert_eq!(created.geometry.position, Some(POSITION));
        assert_eq!(created.geometry.size, Some(DISPLAY_SIZE));
        let movie_data_identifier = created.movie_data_identifier.unwrap();
        let poster_data_identifier = created.poster_image_data_identifier.unwrap();
        assert_eq!(editor.extract_media(movie_data_identifier).unwrap(), MOVIE);
        assert_eq!(
            editor.extract_media(poster_data_identifier).unwrap(),
            POSTER
        );
        let builds = editor.slide_builds(0).unwrap();
        assert_eq!(builds.len(), 1);
        assert_eq!(builds[0].drawable_object_id, created.drawable_object_id);
        assert_eq!(builds[0].settings, KeynoteBuildSettings::movie_start());
        assert_eq!(builds[0].chunks.len(), 1);

        let roundtripped = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.slide_movies(0).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 48.0, y: 72.0 }),
            size: Some(DrawableSize {
                width: 320.0,
                height: 180.0,
            }),
            flags: Some(3),
            angle: Some(8.0),
        };
        editor
            .set_slide_movie_geometry(0, created.drawable_object_id, changed_geometry)
            .unwrap();
        assert_eq!(
            editor
                .slide_movie_geometry(0, created.drawable_object_id)
                .unwrap(),
            changed_geometry
        );
        assert_eq!(
            editor
                .replace_slide_movie_data(0, created.drawable_object_id, REPLACEMENT_MOVIE)
                .unwrap(),
            MOVIE
        );
        assert_eq!(
            editor
                .replace_slide_movie_poster(0, created.drawable_object_id, REPLACEMENT_POSTER)
                .unwrap(),
            POSTER
        );

        let duplicate = editor
            .duplicate_slide_movie(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(
            duplicate.movie_data_identifier,
            created.movie_data_identifier
        );
        assert_eq!(
            duplicate.poster_image_data_identifier,
            created.poster_image_data_identifier
        );
        let removed_original = editor
            .remove_slide_movie(0, created.drawable_object_id)
            .unwrap();
        assert!(removed_original.removed_data_identifiers.is_empty());
        let removed_duplicate = editor
            .remove_slide_movie(0, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [movie_data_identifier, poster_data_identifier]
        );
        assert!(editor.slide_movies(0).unwrap().is_empty());
        assert!(editor.slide_builds(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_movie_creation_is_transactional() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();

        for result in [
            editor.add_slide_movie(
                0,
                "payload.bin",
                b"not video",
                "poster.png",
                POSTER,
                options(),
            ),
            editor.add_slide_movie(
                0,
                "movie.mov",
                MOVIE,
                "payload.bin",
                b"not image",
                options(),
            ),
            editor.add_slide_movie(1, "movie.mov", MOVIE, "poster.png", POSTER, options()),
            editor.add_slide_movie(
                0,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                KeynoteSlideMovieOptions::new(POSITION, DISPLAY_SIZE, Duration::ZERO),
            ),
            editor.add_slide_movie(
                0,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                options().with_natural_size(DrawableSize {
                    width: f32::NAN,
                    height: 720.0,
                }),
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }
    }
}
