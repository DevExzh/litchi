//! Independently positioned audio-object CRUD for Keynote slides.

use std::time::Duration;

use super::slide_movies::geometry::{set_movie_geometry, set_movie_properties};
use super::slide_movies::graph::{
    MovieObjectIds, audio_creation_values, audio_objects, movie_creation_context,
};
use super::*;
use crate::MediaPlaybackSettings;
use crate::data_reference_registry::add_component_data_reference;
use crate::media_playback::media_playback_settings;
use crate::media_playback::replace_movie_playback_settings;
use crate::shapes::{
    DrawablePoint, DrawableProperties, drawable_properties, geometry_from_drawable,
};

const AUDIO_ARCHIVE_MESSAGE_TYPE: u32 = 3_007;

/// One independently positioned audio clip owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideAudioInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub audio_data_identifier: u64,
    pub position: DrawablePoint,
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Trim, poster, repeat, and volume settings.
    pub playback: MediaPlaybackSettings,
    pub duration: Duration,
}

/// Typed placement and playback metadata for a newly created audio clip.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteSlideAudioOptions {
    /// Center point of Keynote's zero-size audio control, in slide points.
    pub position: DrawablePoint,
    /// Playable duration of the source audio.
    pub duration: Duration,
}

impl KeynoteSlideAudioOptions {
    pub const fn new(position: DrawablePoint, duration: Duration) -> Self {
        Self { position, duration }
    }
}

/// Result of removing one slide-owned audio clip and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideAudio {
    pub audio: KeynoteSlideAudioInfo,
    /// Assets culled because the removed clip held their final package reference.
    pub removed_data_identifiers: Vec<u64>,
}

impl KeynoteEditor {
    /// List independently positioned audio clips owned by one slide.
    pub fn slide_audio(&self, slide_index: usize) -> Result<Vec<KeynoteSlideAudioInfo>> {
        self.slide_media_infos(slide_index)?
            .into_iter()
            .filter(|media| media.kind == KeynoteSlideMovieKind::Audio)
            .map(|media| audio_info(self, slide_index, media.drawable_object_id))
            .collect()
    }

    /// Add an independently editable audio clip to a slide.
    ///
    /// The audio archive, title/caption stand-ins, automatic Start Audio build,
    /// component registrations, UUIDs, and package media record are constructed
    /// from typed values. No source drawable or package template is copied.
    pub fn add_slide_audio(
        &mut self,
        slide_index: usize,
        preferred_filename: &str,
        data: &[u8],
        options: KeynoteSlideAudioOptions,
    ) -> Result<KeynoteSlideAudioInfo> {
        let (geometry, duration_seconds) =
            audio_creation_values(options.position, options.duration)?;
        let context = movie_creation_context(self, slide_index)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(self.package())?)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != crate::MediaType::Audio {
            return Err(Error::ParseError(format!(
                "Keynote slide audio requires audio data, not {}",
                asset.media_type.name()
            )));
        }

        let mut staged = media.into_package();
        let objects = audio_objects(
            ids,
            context.slide_id,
            context.style_id,
            asset.data_identifier,
            geometry,
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
        add_component_data_reference(
            &mut staged,
            context.component_id,
            asset.data_identifier,
            ids.drawable,
        )?;
        add_component_external_reference(
            &mut staged,
            context.component_id,
            context.stylesheet_component_id,
            context.style_id,
        )?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let mut verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = audio_info(&verified, slide_index, ids.drawable)?;
        let created_graph = verified.slide_movie_graph(slide_index, ids.drawable)?;
        let expected_duration = Duration::try_from_secs_f64(f64::from(duration_seconds))
            .map_err(|error| Error::ParseError(error.to_string()))?;
        if created.audio_data_identifier != asset.data_identifier
            || created.position != options.position
            || created.duration != expected_duration
            || created_graph.info.kind != KeynoteSlideMovieKind::Audio
            || created_graph.object_ids != ids.all()
            || verified.extract_media(asset.data_identifier)? != data
        {
            return Err(Error::InvalidFormat(
                "Keynote audio creation produced an inconsistent graph".to_owned(),
            ));
        }

        let build = verified.add_slide_build(
            slide_index,
            ids.drawable,
            KeynoteBuildSettings::audio_start(),
        )?;
        if build.drawable_object_id != ids.drawable || build.chunks.len() != 1 {
            return Err(Error::InvalidFormat(
                "Keynote audio creation produced an inconsistent playback build".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the center position of one slide-owned audio control.
    pub fn slide_audio_position(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawablePoint> {
        Ok(require_audio(self, slide_index, drawable_object_id)?.position)
    }

    /// Move one slide-owned audio control while preserving its opaque media fields.
    pub fn set_slide_audio_position(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        position: DrawablePoint,
    ) -> Result<()> {
        require_audio(self, slide_index, drawable_object_id)?;
        let graph = self.slide_movie_graph(slide_index, drawable_object_id)?;
        let geometry = DrawableGeometry {
            position: Some(position),
            ..graph.info.geometry
        }
        .validate()?;
        let mut staged = self.package().clone();
        set_movie_geometry(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_audio_position(slide_index, drawable_object_id)? != position {
            return Err(Error::InvalidFormat(
                "Keynote audio position update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one slide-owned audio control.
    pub fn slide_audio_properties(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(require_audio(self, slide_index, drawable_object_id)?.properties)
    }

    /// Update audio accessibility, hyperlink, and lock properties.
    ///
    /// The typed update retains unknown native media fields and supports both
    /// clearing a property with `None` and encoding explicit boolean defaults.
    pub fn set_slide_audio_properties(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = self.slide_movie_graph(slide_index, drawable_object_id)?;
        if source.info.kind != KeynoteSlideMovieKind::Audio {
            return Err(Error::ParseError(format!(
                "Keynote media {drawable_object_id} is {:?}, not slide audio",
                source.info.kind
            )));
        }
        let mut staged = self.package().clone();
        set_movie_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_audio_properties(slide_index, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Keynote audio properties update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read trim, poster, repeat, and volume settings for one slide audio clip.
    pub fn slide_audio_playback_settings(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<MediaPlaybackSettings> {
        Ok(require_audio(self, slide_index, drawable_object_id)?.playback)
    }

    /// Update playback settings while retaining unrelated and unknown audio fields.
    pub fn set_slide_audio_playback_settings(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        settings: MediaPlaybackSettings,
    ) -> Result<()> {
        let source = self.slide_movie_graph(slide_index, drawable_object_id)?;
        if source.info.kind != KeynoteSlideMovieKind::Audio {
            return Err(Error::ParseError(format!(
                "Keynote media {drawable_object_id} is {:?}, not slide audio",
                source.info.kind
            )));
        }
        let mut staged = self.package().clone();
        let expected = replace_movie_playback_settings(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            "Keynote audio",
            settings,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_audio_playback_settings(slide_index, drawable_object_id)? != expected {
            return Err(Error::InvalidFormat(
                "Keynote audio playback update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Duplicate one slide audio control using Keynote's native placement.
    ///
    /// The audio, title/caption stand-ins, and automatic Start Audio build
    /// receive fresh object identifiers and UUIDs. The clone shares its
    /// embedded audio asset with the source, exactly as with Keynote's
    /// Duplicate command.
    pub fn duplicate_slide_audio(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
    ) -> Result<KeynoteSlideAudioInfo> {
        let source = require_audio(self, slide_index, source_drawable_object_id)?;
        let media = self.duplicate_slide_media(
            slide_index,
            source_drawable_object_id,
            KeynoteSlideMovieKind::Audio,
        )?;
        let created = require_audio(self, slide_index, media.drawable_object_id)?;
        let media_position = media.geometry.position.ok_or_else(|| {
            Error::InvalidFormat("Keynote audio clone has no position".to_owned())
        })?;
        if created.audio_data_identifier != source.audio_data_identifier
            || created.position != media_position
            || created.duration != source.duration
        {
            return Err(Error::InvalidFormat(
                "Keynote audio duplication produced an inconsistent graph".to_owned(),
            ));
        }
        Ok(created)
    }

    /// Replace the bytes referenced by one slide-owned audio clip.
    ///
    /// Audio controls duplicated with [`Self::duplicate_slide_audio`] share
    /// their embedded asset, matching Keynote's native Duplicate behavior.
    pub fn replace_slide_audio_data(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &[u8],
    ) -> Result<Vec<u8>> {
        let source = require_audio(self, slide_index, drawable_object_id)?;
        self.replace_media(source.audio_data_identifier, replacement)
    }

    /// Remove an audio clip, its automatic build, private graph, and unshared asset.
    pub fn remove_slide_audio(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideAudio> {
        let audio = require_audio(self, slide_index, drawable_object_id)?;
        let removed = self.remove_slide_media(
            slide_index,
            drawable_object_id,
            KeynoteSlideMovieKind::Audio,
        )?;
        if removed.movie.movie_data_identifier != Some(audio.audio_data_identifier) {
            return Err(Error::InvalidFormat(
                "Keynote audio deletion removed a mismatched media graph".to_owned(),
            ));
        }
        Ok(RemovedKeynoteSlideAudio {
            audio,
            removed_data_identifiers: removed.removed_data_identifiers,
        })
    }
}

fn require_audio(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<KeynoteSlideAudioInfo> {
    let graph = editor.slide_movie_graph(slide_index, drawable_object_id)?;
    if graph.info.kind != KeynoteSlideMovieKind::Audio {
        return Err(Error::ParseError(format!(
            "Keynote media {drawable_object_id} is {:?}, not slide audio",
            graph.info.kind
        )));
    }
    audio_info(editor, slide_index, drawable_object_id)
}

fn audio_info(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<KeynoteSlideAudioInfo> {
    let graph = ObjectGraph::read(editor.package())?;
    let audio: tsd::MovieArchive = graph.decode_type(
        drawable_object_id,
        AUDIO_ARCHIVE_MESSAGE_TYPE,
        "TSD.MovieArchive",
    )?;
    if audio.audio_only != Some(true) {
        return Err(Error::ParseError(format!(
            "Keynote media {drawable_object_id} is not an audio-only archive"
        )));
    }
    let audio_data_identifier = audio
        .movie_data
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote audio {drawable_object_id} has no data reference"
            ))
        })?
        .identifier;
    let position = geometry_from_drawable(&audio.super_)?
        .position
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote audio {drawable_object_id} has no position"
            ))
        })?;
    let playback = media_playback_settings(&audio).map_err(|error| {
        Error::InvalidFormat(format!(
            "Keynote audio {drawable_object_id} has invalid playback settings: {error}"
        ))
    })?;
    Ok(KeynoteSlideAudioInfo {
        slide_index,
        drawable_object_id,
        audio_data_identifier,
        position,
        properties: drawable_properties(&audio.super_),
        playback,
        duration: playback.duration(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::{MediaLoopMode, MediaVolume};

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 960.0, y: 540.0 };

    fn properties(description: &str) -> DrawableProperties {
        DrawableProperties {
            hyperlink_url: Some("https://example.test/keynote-audio".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some(description.to_owned()),
        }
    }

    #[test]
    fn scratch_presentation_supports_slide_audio_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch audio")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        let options = KeynoteSlideAudioOptions::new(POSITION, Duration::from_millis(1_375));

        let created = editor
            .add_slide_audio(0, "audio.aiff", AUDIO, options)
            .unwrap();
        assert!(editor.slide_movies(0).unwrap().is_empty());
        assert_eq!(
            editor.slide_audio(0).unwrap(),
            std::slice::from_ref(&created)
        );
        assert_eq!(created.position, POSITION);
        assert_eq!(
            editor.extract_media(created.audio_data_identifier).unwrap(),
            AUDIO
        );
        let builds = editor.slide_builds(0).unwrap();
        assert_eq!(builds.len(), 1);
        assert_eq!(builds[0].drawable_object_id, created.drawable_object_id);
        assert_eq!(builds[0].settings, KeynoteBuildSettings::audio_start());
        assert_eq!(builds[0].chunks.len(), 1);

        let roundtripped = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            roundtripped.slide_audio(0).unwrap(),
            std::slice::from_ref(&created)
        );

        let changed_playback = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::Repeat),
            volume: Some(MediaVolume::new(0.75).unwrap()),
            ..created.playback
        };
        editor
            .set_slide_audio_playback_settings(0, created.drawable_object_id, changed_playback)
            .unwrap();
        assert_eq!(
            editor
                .slide_audio_playback_settings(0, created.drawable_object_id)
                .unwrap(),
            changed_playback
        );
        editor
            .set_slide_audio_playback_settings(0, created.drawable_object_id, created.playback)
            .unwrap();

        let changed_properties = properties("Accessible Keynote audio");
        editor
            .set_slide_audio_properties(0, created.drawable_object_id, changed_properties.clone())
            .unwrap();
        assert_eq!(
            editor
                .slide_audio_properties(0, created.drawable_object_id)
                .unwrap(),
            changed_properties
        );
        editor
            .set_slide_audio_properties(
                0,
                created.drawable_object_id,
                DrawableProperties::default(),
            )
            .unwrap();
        assert_eq!(
            editor
                .slide_audio_properties(0, created.drawable_object_id)
                .unwrap(),
            DrawableProperties::default()
        );

        let moved = DrawablePoint { x: 320.0, y: 240.0 };
        editor
            .set_slide_audio_position(0, created.drawable_object_id, moved)
            .unwrap();
        assert_eq!(
            editor
                .slide_audio_position(0, created.drawable_object_id)
                .unwrap(),
            moved
        );
        assert_eq!(
            editor
                .replace_slide_audio_data(0, created.drawable_object_id, REPLACEMENT_AUDIO)
                .unwrap(),
            AUDIO
        );
        assert_eq!(
            editor.extract_media(created.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );

        let removed = editor
            .remove_slide_audio(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.audio.drawable_object_id, created.drawable_object_id);
        assert_eq!(
            removed.removed_data_identifiers,
            [created.audio_data_identifier]
        );
        assert!(editor.slide_audio(0).unwrap().is_empty());
        assert!(editor.slide_builds(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn scratch_presentation_supports_native_audio_duplication() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch audio")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        let source = editor
            .add_slide_audio(
                0,
                "audio.aiff",
                AUDIO,
                KeynoteSlideAudioOptions::new(POSITION, Duration::from_millis(1_375)),
            )
            .unwrap();
        let source_properties = properties("Duplicated Keynote audio");
        editor
            .set_slide_audio_properties(0, source.drawable_object_id, source_properties.clone())
            .unwrap();

        let duplicate = editor
            .duplicate_slide_audio(0, source.drawable_object_id)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        let source_graph = editor
            .slide_movie_graph(0, source.drawable_object_id)
            .unwrap();
        let duplicate_graph = editor
            .slide_movie_graph(0, duplicate.drawable_object_id)
            .unwrap();
        assert!(
            source_graph
                .object_ids
                .iter()
                .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
        );
        assert!(
            source_graph
                .uuid_object_ids
                .iter()
                .all(|identifier| !duplicate_graph.uuid_object_ids.contains(identifier))
        );
        assert_eq!(
            duplicate.audio_data_identifier,
            source.audio_data_identifier
        );
        assert_eq!(
            duplicate.position,
            DrawablePoint {
                x: source.position.x + DRAWABLE_DUPLICATE_OFFSET,
                y: source.position.y + DRAWABLE_DUPLICATE_OFFSET,
            }
        );
        assert_eq!(duplicate.duration, source.duration);
        assert_eq!(duplicate.properties, source_properties);
        let duplicate_builds = editor
            .slide_builds(0)
            .unwrap()
            .into_iter()
            .filter(|build| build.drawable_object_id == duplicate.drawable_object_id)
            .collect::<Vec<_>>();
        assert_eq!(duplicate_builds.len(), 1);
        assert_eq!(
            duplicate_builds[0].settings,
            KeynoteBuildSettings::audio_start()
        );
        assert_eq!(duplicate_builds[0].chunks.len(), 1);

        let moved_duplicate = DrawablePoint { x: 320.0, y: 240.0 };
        editor
            .set_slide_audio_position(0, duplicate.drawable_object_id, moved_duplicate)
            .unwrap();
        assert_eq!(
            editor
                .slide_audio_position(0, source.drawable_object_id)
                .unwrap(),
            source.position
        );
        assert_eq!(
            editor
                .slide_audio_position(0, duplicate.drawable_object_id)
                .unwrap(),
            moved_duplicate
        );
        assert_eq!(
            editor
                .replace_slide_audio_data(0, duplicate.drawable_object_id, REPLACEMENT_AUDIO)
                .unwrap(),
            AUDIO
        );
        assert_eq!(
            editor.extract_media(source.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.slide_audio(0).unwrap().len(), 2);
        assert_eq!(
            reopened
                .slide_audio(0)
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .position,
            moved_duplicate
        );
        assert_eq!(
            reopened
                .slide_builds(0)
                .unwrap()
                .into_iter()
                .filter(|build| build.drawable_object_id == duplicate.drawable_object_id)
                .count(),
            1
        );

        let removed_source = editor
            .remove_slide_audio(0, source.drawable_object_id)
            .unwrap();
        assert!(removed_source.removed_data_identifiers.is_empty());
        assert_eq!(editor.slide_audio(0).unwrap().len(), 1);
        let removed_duplicate = editor
            .remove_slide_audio(0, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(
            removed_duplicate.removed_data_identifiers,
            [source.audio_data_identifier]
        );
        assert!(editor.slide_audio(0).unwrap().is_empty());
        assert!(editor.slide_builds(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_slide_audio_creation_and_cross_type_edits_are_transactional() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(editor.duplicate_slide_audio(0, 999).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        for result in [
            editor.add_slide_audio(
                0,
                "payload.bin",
                b"not audio",
                KeynoteSlideAudioOptions::new(POSITION, Duration::from_secs(1)),
            ),
            editor.add_slide_audio(
                1,
                "audio.aiff",
                AUDIO,
                KeynoteSlideAudioOptions::new(POSITION, Duration::from_secs(1)),
            ),
            editor.add_slide_audio(
                0,
                "audio.aiff",
                AUDIO,
                KeynoteSlideAudioOptions::new(POSITION, Duration::ZERO),
            ),
            editor.add_slide_audio(
                0,
                "audio.aiff",
                AUDIO,
                KeynoteSlideAudioOptions::new(
                    DrawablePoint {
                        x: f32::NAN,
                        y: 10.0,
                    },
                    Duration::from_secs(1),
                ),
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }

        let audio = editor
            .add_slide_audio(
                0,
                "audio.aiff",
                AUDIO,
                KeynoteSlideAudioOptions::new(POSITION, Duration::from_secs(1)),
            )
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_movie_geometry(0, audio.drawable_object_id, DrawableGeometry::default(),)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_slide_audio_position(
                    0,
                    audio.drawable_object_id,
                    DrawablePoint {
                        x: f32::INFINITY,
                        y: 10.0,
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .replace_slide_audio_data(
                    0,
                    audio.drawable_object_id,
                    b"\x89PNG\r\n\x1a\nnot audio",
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
