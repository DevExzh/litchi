//! Standalone movie-object creation and editing for Keynote slides.

use litchi_iwa_common::{WireLimits, media::Type as MediaType};
use litchi_iwa_protos::keynote_media_codec;
use litchi_keynote::slide::media::MovieKind;
use litchi_keynote::slide::movie::Options as SlideMovieOptions;

use super::*;
use crate::data_reference_registry::add_component_data_reference;
use crate::media::MediaAssetId;
use crate::shapes::{DrawableGeometry, DrawableProperties, DrawableSize, geometry_from_drawable};
use litchi_iwa_common::media::playback::MediaPlaybackSettings;

pub(in crate::keynote::editor) mod geometry;
pub(in crate::keynote::editor) mod graph;

use graph::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const MOVIE_DATA_FIELD: u32 = 14;
const POSTER_IMAGE_DATA_FIELD: u32 = 15;
const MOVIE_MEDIA_PLACEHOLDER_FLAG: u32 = 1;

pub(in crate::keynote::editor) fn movie_playback_wire_limits(
    package: &IWorkPackage,
) -> Result<WireLimits> {
    let archive_limits = package.limits().archive_limits();
    let source_bytes = archive_limits
        .max_message_bytes()
        .min(archive_limits.max_archive_bytes())
        .min(package.limits().max_iwa_stream_bytes())
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(source_bytes)
        .and_then(|limits| {
            limits.with_fields(
                source_bytes
                    .saturating_mul(4)
                    .clamp(1, WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| limits.with_output_bytes(source_bytes))
        .and_then(|limits| {
            limits.with_rewrite_work(
                source_bytes
                    .saturating_mul(8)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| {
            Error::InvalidFormat(format!("invalid Keynote movie playback limits: {error}"))
        })
}

/// One movie drawable owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideMovieInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub kind: MovieKind,
    /// Embedded video/audio data referenced by this movie, when materialized.
    pub movie_data_identifier: Option<MediaAssetId>,
    /// Embedded poster image data referenced by this movie, when materialized.
    pub poster_image_data_identifier: Option<MediaAssetId>,
    pub geometry: DrawableGeometry,
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Trim, poster, repeat, and volume settings when the media has a valid native playback
    /// range. Existing documents, live video, and unresolved media placeholders can omit it.
    pub playback: Option<MediaPlaybackSettings>,
    pub original_size: Option<DrawableSize>,
    pub natural_size: Option<DrawableSize>,
}

pub(in crate::keynote::editor) struct SlideMovieGraph {
    pub(in crate::keynote::editor) object_ids: Vec<u64>,
}

impl KeynoteEditor {
    /// List movie drawables directly owned by one slide in drawable order.
    pub fn slide_movies(&self, slide_index: usize) -> Result<Vec<KeynoteSlideMovieInfo>> {
        Ok(self
            .slide_media_infos(slide_index)?
            .into_iter()
            .filter(|movie| !movie.kind.is_audio())
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
        let slide_id = slide.native_ids()?.slide.get();
        let graph = ObjectGraph::read(self.package())?;
        let native: kn::SlideArchive =
            graph.decode_type(slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
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
            .map(|reference| movie_info(&graph, self.package(), slide_index, reference.identifier))
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
        options: SlideMovieOptions,
    ) -> Result<KeynoteSlideMovieInfo> {
        let (geometry, duration_seconds) = movie_creation_values(options)?;
        let context = movie_creation_context(self, slide_index)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(self.package())?)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let movie_asset = media.insert_unreferenced(preferred_movie_filename, movie_data)?;
        if movie_asset.media_type != MediaType::Video {
            return Err(Error::ParseError(format!(
                "Keynote slide movies require video data, not {}",
                movie_asset.media_type.name()
            )));
        }
        let poster_asset = media.insert_unreferenced(preferred_poster_filename, poster_data)?;
        if poster_asset.media_type != MediaType::Image {
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
            movie_asset.data_identifier.get(),
            poster_asset.data_identifier.get(),
            geometry,
            options.natural_size(),
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
                data_identifier.get(),
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
        if created.kind != MovieKind::File
            || created.movie_data_identifier != Some(movie_asset.data_identifier)
            || created.poster_image_data_identifier != Some(poster_asset.data_identifier)
            || created.geometry != geometry
            || created.original_size != Some(options.natural_size())
            || created.natural_size != Some(options.natural_size())
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
        let slide_ids = slide.native_ids()?;
        let slide_id = slide_ids.slide.get();
        let graph = ObjectGraph::read(self.package())?;
        let native: kn::SlideArchive =
            graph.decode_type(slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        if !native
            .owned_drawables
            .iter()
            .any(|reference| reference.identifier == drawable_object_id)
        {
            return Err(Error::ParseError(format!(
                "Keynote movie {drawable_object_id} is not owned by slide {slide_index}"
            )));
        }
        let archive_name = graph.archive_name(slide_id)?.to_owned();
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
        if object_ids.contains(&slide_id) {
            return Err(Error::InvalidFormat(
                "Keynote movie private graph reaches its owning slide".to_owned(),
            ));
        }
        // Retain the listing decoder's validation without storing a duplicate
        // property/playback snapshot in the private graph result.
        movie_info(&graph, self.package(), slide_index, drawable_object_id)?;
        Ok(SlideMovieGraph { object_ids })
    }
}

fn movie_info(
    graph: &ObjectGraph,
    package: &IWorkPackage,
    slide_index: usize,
    identifier: u64,
) -> Result<KeynoteSlideMovieInfo> {
    let raw = graph.message_data_type(identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let movie_data_identifier = movie_data_reference_identifier(raw, MOVIE_DATA_FIELD, identifier)?;
    let poster_image_data_identifier =
        movie_data_reference_identifier(raw, POSTER_IMAGE_DATA_FIELD, identifier)?;
    let movie: tsd::MovieArchive =
        graph.decode_type(identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    let movie_data = movie
        .movie_data
        .as_ref()
        .map(|reference| MediaAssetId::try_from(reference.identifier))
        .transpose()
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "Keynote movie {identifier} has an invalid video data identifier: {error}"
            ))
        })?;
    let poster_image_data = movie
        .poster_image_data
        .as_ref()
        .map(|reference| MediaAssetId::try_from(reference.identifier))
        .transpose()
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "Keynote movie {identifier} has an invalid poster data identifier: {error}"
            ))
        })?;
    if movie_data != movie_data_identifier || poster_image_data != poster_image_data_identifier {
        return Err(Error::InvalidFormat(format!(
            "Keynote movie {identifier} data-reference projection disagrees with MovieArchive"
        )));
    }
    let kind = if movie.is_live_video == Some(true) {
        MovieKind::LiveVideo
    } else if movie.audio_only == Some(true) {
        MovieKind::Audio
    } else if movie
        .flags
        .is_some_and(|flags| flags & MOVIE_MEDIA_PLACEHOLDER_FLAG != 0)
    {
        MovieKind::Placeholder
    } else {
        MovieKind::File
    };
    let playback = if movie.end_time.is_some() {
        let limits = movie_playback_wire_limits(package)?;
        Some(
            litchi_keynote::__decode_movie_playback_payload(raw, limits).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Keynote media {identifier} has invalid playback settings: {error}"
                ))
            })?,
        )
    } else {
        None
    };
    Ok(KeynoteSlideMovieInfo {
        slide_index,
        drawable_object_id: identifier,
        kind,
        movie_data_identifier,
        poster_image_data_identifier,
        geometry: geometry_from_drawable(&movie.super_)?,
        properties: crate::shapes::drawable_properties(&movie.super_),
        playback,
        original_size: movie.original_size.map(drawable_size),
        natural_size: movie.natural_size.map(drawable_size),
    })
}

fn movie_data_reference_identifier(
    source: &[u8],
    field_number: u32,
    movie_identifier: u64,
) -> Result<Option<MediaAssetId>> {
    let payloads = repeated_length_delimited_payloads(source, field_number)?;
    if payloads.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "Keynote movie {movie_identifier} field {field_number} repeats its data reference"
        )));
    }
    let Some(payload) = payloads.first().copied() else {
        return Ok(None);
    };
    let reference = keynote_media_codec::decode_data_reference(
        payload,
        keynote_media_codec::DecodeOptions::for_source(payload),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "Keynote movie {movie_identifier} field {field_number} has malformed data reference: {error}"
        ))
    })?;
    MediaAssetId::try_from(reference.identifier())
        .map(Some)
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "Keynote movie {movie_identifier} field {field_number} has an invalid data identifier: {error}"
            ))
        })
}

fn drawable_size(size: tsp::Size) -> DrawableSize {
    DrawableSize {
        width: size.width,
        height: size.height,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::RawMessage;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawableFlipAxis, DrawablePoint};
    use crate::wire::remove_repeated_length_delimited_field_where;
    use litchi_core::Position;
    use litchi_keynote::slide::audio::Options as SlideAudioOptions;
    use litchi_keynote::slide::media::geometry::MovieGeometry;
    use litchi_keynote::slide::media::{
        MediaLoopMode as KeynoteMediaLoopMode, MediaPlaybackSettings as KeynotePlaybackSettings,
        MediaProperties as KeynoteMediaProperties, MediaVolume as KeynoteMediaVolume,
    };
    use litchi_keynote::slide::media::{Point as KeynotePoint, Size as KeynoteSize};
    use litchi_keynote::{MovieSelector, Package as KeynotePackage, SlideSelector};
    use std::time::Duration;

    const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-movie";
    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio";
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
    const DUPLICATE_OFFSET: f32 = 10.0;
    const NATIVE_MEDIA_BASELINE: &[u8] =
        include_bytes!("../../../../../test-data/iwork/keynote/media-comments-baseline-native.key");

    fn options() -> SlideMovieOptions {
        SlideMovieOptions::new(POSITION, DISPLAY_SIZE, Duration::from_millis(1_250))
            .unwrap()
            .with_natural_size(NATURAL_SIZE)
            .unwrap()
    }

    // Keep editor setup/teardown local to these host tests while exercising
    // the focused package's selector-first movie title/caption APIs directly.
    fn focused_movie_package(editor: &KeynoteEditor) -> KeynotePackage {
        KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap()
    }

    fn replace_with_focused_movie_package(editor: &mut KeynoteEditor, package: &KeynotePackage) {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).unwrap();
        *editor = KeynoteEditor::from_bytes(&bytes).unwrap();
    }

    fn add_audio(
        editor: &mut KeynoteEditor,
        preferred_filename: &str,
        data: &[u8],
        options: SlideAudioOptions,
    ) -> KeynoteSlideAudioInfo {
        let package = focused_movie_package(editor);
        let commit = package
            .add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
        editor.slide_audio(0).unwrap().into_iter().last().unwrap()
    }

    fn remove_movie_poster(
        editor: &mut KeynoteEditor,
        drawable_object_id: u64,
        poster_data_identifier: MediaAssetId,
    ) {
        let archive_name = ObjectGraph::read(editor.package())
            .unwrap()
            .archive_name(drawable_object_id)
            .unwrap()
            .to_owned();
        let mut package = editor.package().clone();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(drawable_object_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                    .unwrap();
                let message = &object.messages[message_index];
                let data = remove_repeated_length_delimited_field_where(
                    &message.data,
                    POSTER_IMAGE_DATA_FIELD,
                    |_| Ok(true),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: MOVIE_MESSAGE_TYPE,
                        data,
                    },
                )?;
                object.archive_info.message_infos[message_index]
                    .data_references
                    .retain(|identifier| *identifier != poster_data_identifier.get());
                Ok(())
            })
            .unwrap();
        let bytes = package.to_bytes().unwrap();
        *editor = KeynoteEditor::from_bytes(&bytes).unwrap();
    }

    fn remove_movie_data(
        editor: &mut KeynoteEditor,
        drawable_object_id: u64,
        movie_data_identifier: MediaAssetId,
    ) {
        let archive_name = ObjectGraph::read(editor.package())
            .unwrap()
            .archive_name(drawable_object_id)
            .unwrap()
            .to_owned();
        let mut package = editor.package().clone();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(drawable_object_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                    .unwrap();
                let message = &object.messages[message_index];
                let data = remove_repeated_length_delimited_field_where(
                    &message.data,
                    MOVIE_DATA_FIELD,
                    |_| Ok(true),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: MOVIE_MESSAGE_TYPE,
                        data,
                    },
                )?;
                object.archive_info.message_infos[message_index]
                    .data_references
                    .retain(|identifier| *identifier != movie_data_identifier.get());
                Ok(())
            })
            .unwrap();
        let bytes = package.to_bytes().unwrap();
        *editor = KeynoteEditor::from_bytes(&bytes).unwrap();
    }

    fn duplicate_movie(editor: &mut KeynoteEditor, movie: MovieSelector) -> KeynoteSlideMovieInfo {
        let package = focused_movie_package(editor);
        let commit = package
            .duplicate_slide_movie(SlideSelector::index(0), movie)
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
        editor.slide_movies(0).unwrap().into_iter().last().unwrap()
    }

    fn remove_movie(editor: &mut KeynoteEditor, movie: MovieSelector) -> Vec<MediaAssetId> {
        let before = editor
            .media_assets()
            .unwrap()
            .into_iter()
            .map(|asset| asset.data_identifier)
            .collect::<Vec<_>>();
        let package = focused_movie_package(editor);
        let commit = package
            .remove_slide_movie(SlideSelector::index(0), movie)
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
        let after = editor
            .media_assets()
            .unwrap()
            .into_iter()
            .map(|asset| asset.data_identifier)
            .collect::<Vec<_>>();
        before
            .into_iter()
            .filter(|identifier| !after.contains(identifier))
            .collect()
    }

    fn movie_title(editor: &KeynoteEditor, movie: MovieSelector) -> Option<String> {
        focused_movie_package(editor)
            .slide_movie_title(SlideSelector::index(0), movie)
            .unwrap()
    }

    fn set_movie_title(editor: &mut KeynoteEditor, movie: MovieSelector, title: &str) {
        let package = focused_movie_package(editor);
        let commit = package
            .edit_slide_movie_title(SlideSelector::index(0), movie)
            .unwrap()
            .set(title)
            .unwrap()
            .commit()
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
    }

    fn remove_movie_title(editor: &mut KeynoteEditor, movie: MovieSelector) -> bool {
        let package = focused_movie_package(editor);
        let edit = package
            .edit_slide_movie_title(SlideSelector::index(0), movie)
            .unwrap();
        let had_title = edit.before().is_some();
        let commit = edit.clear().unwrap().commit().unwrap();
        if !commit.patch().is_noop() {
            replace_with_focused_movie_package(editor, commit.package());
        }
        had_title
    }

    fn movie_caption(editor: &KeynoteEditor, movie: MovieSelector) -> Option<String> {
        focused_movie_package(editor)
            .slide_movie_caption(SlideSelector::index(0), movie)
            .unwrap()
    }

    fn set_movie_caption(editor: &mut KeynoteEditor, movie: MovieSelector, caption: &str) {
        let package = focused_movie_package(editor);
        let commit = package
            .edit_slide_movie_caption(SlideSelector::index(0), movie)
            .unwrap()
            .set(caption)
            .unwrap()
            .commit()
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
    }

    fn remove_movie_caption(editor: &mut KeynoteEditor, movie: MovieSelector) -> bool {
        let package = focused_movie_package(editor);
        let edit = package
            .edit_slide_movie_caption(SlideSelector::index(0), movie)
            .unwrap();
        let had_caption = edit.before().is_some();
        let commit = edit.clear().unwrap().commit().unwrap();
        if !commit.patch().is_noop() {
            replace_with_focused_movie_package(editor, commit.package());
        }
        had_caption
    }

    fn properties(description: &str) -> KeynoteMediaProperties {
        KeynoteMediaProperties::new()
            .with_hyperlink_url(Some("https://example.test/keynote-movie".to_owned()))
            .with_locked(Some(true))
            .with_aspect_ratio_locked(Some(false))
            .with_accessibility_description(Some(description.to_owned()))
    }

    fn raw_properties(properties: &KeynoteMediaProperties) -> DrawableProperties {
        DrawableProperties {
            hyperlink_url: properties.hyperlink_url().map(str::to_owned),
            locked: properties.locked(),
            aspect_ratio_locked: properties.aspect_ratio_locked(),
            accessibility_description: properties.accessibility_description().map(str::to_owned),
        }
    }

    fn set_movie_properties(editor: &mut KeynoteEditor, properties: KeynoteMediaProperties) {
        let package = focused_movie_package(editor);
        let commit = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap()
            .set(properties)
            .unwrap()
            .commit()
            .unwrap();
        replace_with_focused_movie_package(editor, commit.package());
    }

    #[test]
    fn source_built_movie_properties_project_through_focused_updates() {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties projection")
            .build()
            .unwrap();
        seed.add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        let baseline = seed.slide_movies(0).unwrap().remove(0);
        let mut editor = KeynoteEditor::from_bytes(&seed.to_bytes().unwrap()).unwrap();
        let focused_before = focused_movie_package(&editor)
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap();
        assert_eq!(baseline.properties, raw_properties(&focused_before));

        for expected in [
            KeynoteMediaProperties::new()
                .with_hyperlink_url(Some(String::new()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("媒体 🎬".to_owned())),
            KeynoteMediaProperties::new()
                .with_hyperlink_url(Some("opaque target 日本語".to_owned()))
                .with_locked(Some(true))
                .with_aspect_ratio_locked(Some(true))
                .with_accessibility_description(Some("locked 🔒".to_owned())),
            KeynoteMediaProperties::new()
                .with_hyperlink_url(Some("second target".to_owned()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("unlocked again".to_owned())),
            KeynoteMediaProperties::default(),
        ] {
            set_movie_properties(&mut editor, expected.clone());
            let actual = editor.slide_movies(0).unwrap().remove(0);
            assert_eq!(actual.properties, raw_properties(&expected));
            assert_eq!(actual.kind, baseline.kind);
            assert_eq!(actual.movie_data_identifier, baseline.movie_data_identifier);
            assert_eq!(
                actual.poster_image_data_identifier,
                baseline.poster_image_data_identifier
            );
            assert_eq!(actual.geometry, baseline.geometry);
            assert_eq!(actual.playback, baseline.playback);
            assert_eq!(actual.original_size, baseline.original_size);
            assert_eq!(actual.natural_size, baseline.natural_size);
        }

        assert_eq!(
            editor
                .extract_media(baseline.movie_data_identifier.unwrap())
                .unwrap(),
            MOVIE
        );
    }

    #[test]
    fn native_file_movie_properties_match_focused_source_projection() {
        let editor = KeynoteEditor::from_bytes(NATIVE_MEDIA_BASELINE).unwrap();
        let focused = focused_movie_package(&editor);
        let infos = editor.slide_media_infos(0).unwrap();
        let mut file_count = 0;
        for (movie_position, info) in infos.into_iter().enumerate() {
            let focused_properties = focused
                .slide_media_properties(
                    SlideSelector::index(0),
                    MovieSelector::index(movie_position),
                )
                .unwrap();
            if info.kind != MovieKind::File {
                continue;
            }
            assert_eq!(info.properties, raw_properties(&focused_properties));
            file_count += 1;
        }
        assert!(file_count > 0);
    }

    #[test]
    fn source_built_movie_properties_support_an_absent_poster() {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties without poster")
            .build()
            .unwrap();
        let created = seed
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        let poster_data_identifier = created.poster_image_data_identifier.unwrap();
        remove_movie_poster(
            &mut seed,
            created.drawable_object_id,
            poster_data_identifier,
        );

        let baseline_bytes = seed.to_bytes().unwrap();
        let baseline_editor = KeynoteEditor::from_bytes(&baseline_bytes).unwrap();
        let baseline = baseline_editor.slide_movies(0).unwrap().remove(0);
        assert_eq!(baseline.kind, MovieKind::File);
        assert_eq!(baseline.poster_image_data_identifier, None);
        assert!(baseline.movie_data_identifier.is_some());

        let mut editor = KeynoteEditor::from_bytes(&baseline_bytes).unwrap();
        let expected = KeynoteMediaProperties::new()
            .with_hyperlink_url(Some("opaque target 日本語".to_owned()))
            .with_locked(Some(false))
            .with_aspect_ratio_locked(Some(true))
            .with_accessibility_description(Some("No poster 🎬".to_owned()));
        set_movie_properties(&mut editor, expected.clone());

        let actual = editor.slide_movies(0).unwrap().remove(0);
        assert_eq!(actual.poster_image_data_identifier, None);
        assert_eq!(actual.movie_data_identifier, baseline.movie_data_identifier);
        assert_eq!(actual.playback, baseline.playback);
        assert_eq!(actual.geometry, baseline.geometry);
        assert_eq!(actual.properties, raw_properties(&expected));
    }

    #[test]
    fn source_built_movie_properties_support_missing_content_data() {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties without content")
            .build()
            .unwrap();
        let created = seed
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        remove_movie_data(
            &mut seed,
            created.drawable_object_id,
            created.movie_data_identifier.unwrap(),
        );

        let editor = KeynoteEditor::from_bytes(&seed.to_bytes().unwrap()).unwrap();
        let movie = editor.slide_movies(0).unwrap().remove(0);
        assert_eq!(movie.kind, MovieKind::File);
        assert_eq!(movie.movie_data_identifier, None);
        assert_eq!(
            movie.properties,
            DrawableProperties {
                hyperlink_url: None,
                locked: Some(false),
                aspect_ratio_locked: Some(true),
                accessibility_description: None,
            }
        );

        let focused = focused_movie_package(&editor);
        let focused_properties = focused
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap();
        assert_eq!(movie.properties, raw_properties(&focused_properties));

        let mut before = Vec::new();
        focused.write_to(&mut before).unwrap();
        let rejected = focused
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .and_then(|edit| {
                edit.set(
                    KeynoteMediaProperties::new()
                        .with_accessibility_description(Some("missing content".to_owned())),
                )
            })
            .and_then(|edit| edit.commit());
        assert!(rejected.is_err());
        let mut after = Vec::new();
        focused.write_to(&mut after).unwrap();
        assert_eq!(after, before);
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
        assert_eq!(created.kind, MovieKind::File);
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

        let initial_playback = KeynotePackage::from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .slide_movie_playback_settings(
                SlideSelector::position(Position::new(0)),
                MovieSelector::index(0),
            )
            .unwrap()
            .unwrap();
        let changed_playback = KeynotePlaybackSettings {
            loop_mode: Some(KeynoteMediaLoopMode::BackAndForth),
            volume: Some(KeynoteMediaVolume::new(0.75).unwrap()),
            ..initial_playback
        };
        let package = KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = package
            .edit_slide_movie_playback_settings(
                SlideSelector::position(Position::new(0)),
                MovieSelector::index(0),
            )
            .unwrap()
            .set(changed_playback)
            .unwrap()
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        editor = KeynoteEditor::from_bytes(&bytes).unwrap();
        assert_eq!(
            KeynotePackage::from_bytes(&editor.to_bytes().unwrap())
                .unwrap()
                .slide_movie_playback_settings(
                    SlideSelector::position(Position::new(0)),
                    MovieSelector::index(0),
                )
                .unwrap()
                .unwrap(),
            changed_playback
        );
        let package = KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = package
            .edit_slide_movie_playback_settings(
                SlideSelector::position(Position::new(0)),
                MovieSelector::index(0),
            )
            .unwrap()
            .set(initial_playback)
            .unwrap()
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        editor = KeynoteEditor::from_bytes(&bytes).unwrap();

        let changed_properties = properties("Accessible Keynote movie");
        set_movie_properties(&mut editor, changed_properties.clone());
        assert_eq!(
            editor.slide_movies(0).unwrap().first().unwrap().properties,
            raw_properties(&changed_properties)
        );
        let cleared_properties = KeynoteMediaProperties::default();
        set_movie_properties(&mut editor, cleared_properties.clone());
        assert_eq!(
            editor.slide_movies(0).unwrap().first().unwrap().properties,
            raw_properties(&cleared_properties)
        );

        // Builder snapshots are not admitted by the focused movie-geometry
        // owner. Keep the host fail-closed and verify that rejection is
        // atomic; the focused package integration suite covers successful
        // geometry and transform edits on admitted sources.
        let changed_geometry = MovieGeometry::new(
            KeynotePoint { x: 48.0, y: 72.0 },
            KeynoteSize {
                width: 320.0,
                height: 180.0,
            },
        )
        .unwrap();
        let before_geometry = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_movie_geometry_by_selector(
                    Position::new(0),
                    MovieSelector::index(0),
                    changed_geometry,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_geometry);
        let before_flip = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_slide_movie_by_selector(
                    Position::new(0),
                    MovieSelector::index(0),
                    DrawableFlipAxis::Horizontal,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);

        assert_eq!(
            editor.slide_movies(0).unwrap()[0].geometry,
            created.geometry
        );
        assert_eq!(editor.extract_media(movie_data_identifier).unwrap(), MOVIE);
        assert_eq!(
            editor.extract_media(poster_data_identifier).unwrap(),
            POSTER
        );
        let source_geometry = editor.slide_movies(0).unwrap()[0].geometry;
        assert_eq!(
            editor
                .replace_media(movie_data_identifier, REPLACEMENT_MOVIE)
                .unwrap(),
            MOVIE
        );
        assert_eq!(
            editor
                .replace_media(poster_data_identifier, REPLACEMENT_POSTER)
                .unwrap(),
            POSTER
        );

        let duplicate_properties = properties("Duplicated Keynote movie");
        set_movie_properties(&mut editor, duplicate_properties.clone());

        let duplicate = duplicate_movie(&mut editor, MovieSelector::index(0));
        assert_eq!(
            duplicate.movie_data_identifier,
            created.movie_data_identifier
        );
        assert_eq!(
            duplicate.poster_image_data_identifier,
            created.poster_image_data_identifier
        );
        assert_eq!(
            duplicate.geometry.position,
            source_geometry.position.map(|position| DrawablePoint {
                x: position.x + DUPLICATE_OFFSET,
                y: position.y + DUPLICATE_OFFSET,
            })
        );
        assert_eq!(duplicate.geometry.size, source_geometry.size);
        assert_eq!(duplicate.geometry.flags, source_geometry.flags);
        assert_eq!(duplicate.geometry.angle, source_geometry.angle);
        assert_eq!(duplicate.properties, raw_properties(&duplicate_properties));
        let removed_original = remove_movie(&mut editor, MovieSelector::index(0));
        assert!(removed_original.is_empty());
        let removed_duplicate = remove_movie(&mut editor, MovieSelector::index(0));
        assert_eq!(
            removed_duplicate,
            [movie_data_identifier, poster_data_identifier]
        );
        assert!(editor.slide_movies(0).unwrap().is_empty());
        assert!(editor.slide_builds(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn scratch_presentation_supports_native_movie_title_caption_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Movie labels")
            .build()
            .unwrap();
        editor
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();

        assert_eq!(movie_title(&editor, MovieSelector::index(0)), None);
        assert_eq!(movie_caption(&editor, MovieSelector::index(0)), None);
        set_movie_title(&mut editor, MovieSelector::index(0), "Quarterly highlight");
        set_movie_caption(&mut editor, MovieSelector::index(0), "Revenue overview");
        assert_eq!(
            movie_title(&editor, MovieSelector::index(0)),
            Some("Quarterly highlight".to_owned())
        );
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(0)),
            Some("Revenue overview".to_owned())
        );

        set_movie_caption(
            &mut editor,
            MovieSelector::index(0),
            "Updated revenue overview",
        );
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(0)),
            Some("Updated revenue overview".to_owned())
        );

        let duplicate = duplicate_movie(&mut editor, MovieSelector::index(0));
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(1)),
            Some("Updated revenue overview".to_owned())
        );

        set_movie_title(&mut editor, MovieSelector::index(0), "Updated highlight");
        assert!(remove_movie_caption(&mut editor, MovieSelector::index(0)));
        assert!(!remove_movie_caption(&mut editor, MovieSelector::index(0)));
        assert!(remove_movie_title(&mut editor, MovieSelector::index(0)));
        assert_eq!(movie_title(&editor, MovieSelector::index(0)), None);
        assert_eq!(movie_caption(&editor, MovieSelector::index(0)), None);

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            movie_caption(&reopened, MovieSelector::index(1)),
            Some("Updated revenue overview".to_owned())
        );
        editor = reopened;
        remove_movie(&mut editor, MovieSelector::index(1));
        assert!(
            editor
                .slide_movies(0)
                .unwrap()
                .iter()
                .all(|item| item.drawable_object_id != duplicate.drawable_object_id)
        );
    }

    #[test]
    fn movie_caption_package_preserves_movie_selector_order_after_audio() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Movie caption selector order")
            .build()
            .unwrap();
        let audio = add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375)).unwrap(),
        );
        let movie = editor
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();

        assert_eq!(
            editor.slide_audio(0).unwrap()[0].drawable_object_id,
            audio.drawable_object_id
        );
        assert_eq!(
            editor.slide_movies(0).unwrap()[0].drawable_object_id,
            movie.drawable_object_id
        );

        set_movie_caption(&mut editor, MovieSelector::index(1), "Caption after audio");
        set_movie_title(&mut editor, MovieSelector::index(1), "Title after audio");
        assert_eq!(
            movie_title(&editor, MovieSelector::index(1)),
            Some("Title after audio".to_owned())
        );
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(1)),
            Some("Caption after audio".to_owned())
        );
        assert!(remove_movie_caption(&mut editor, MovieSelector::index(1)));
        assert_eq!(movie_caption(&editor, MovieSelector::index(1)), None);
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
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        assert_eq!(
            SlideMovieOptions::new(POSITION, DISPLAY_SIZE, Duration::ZERO),
            Err(litchi_keynote::Error::InvalidMovieDuration)
        );
        assert_eq!(
            options().with_natural_size(DrawableSize {
                width: f32::NAN,
                height: 720.0,
            }),
            Err(litchi_keynote::Error::InvalidMovieSize)
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        editor
            .add_slide_movie(0, "movie.mov", MOVIE, "poster.png", POSTER, options())
            .unwrap();
        let before_flip = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_slide_movie_by_selector(
                    Position::new(0),
                    MovieSelector::index(9),
                    DrawableFlipAxis::Horizontal,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);
        assert!(
            editor
                .restore_slide_movie_original_size_by_selector(
                    Position::new(0),
                    MovieSelector::index(9),
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);
        assert!(
            editor
                .flip_slide_movie_by_selector(
                    Position::new(0),
                    MovieSelector::index(0),
                    DrawableFlipAxis::Horizontal,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_flip);
    }
}
