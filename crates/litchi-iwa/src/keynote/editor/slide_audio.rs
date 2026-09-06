//! Independently positioned audio-object creation and editing for Keynote slides.

use std::time::Duration;

use litchi_iwa_common::media::Type as MediaType;
use litchi_keynote::slide::audio::Options as SlideAudioOptions;
use litchi_keynote::slide::media::MovieKind;

use super::slide_movies::geometry::{set_movie_geometry, set_movie_properties};
use super::slide_movies::graph::{
    MovieObjectIds, audio_creation_values, audio_objects, movie_creation_context,
};
use super::slide_movies::movie_playback_wire_limits;
use super::*;
use crate::data_reference_registry::add_component_data_reference;
use crate::media::MediaAssetId;
use crate::shapes::{
    DrawablePoint, DrawableProperties, drawable_properties, geometry_from_drawable,
};
use litchi_iwa_common::media::playback::MediaPlaybackSettings;

const AUDIO_ARCHIVE_MESSAGE_TYPE: u32 = 3_007;

/// One independently positioned audio clip owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideAudioInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub audio_data_identifier: MediaAssetId,
    pub position: DrawablePoint,
    /// Shared drawable metadata, including accessibility description and lock state.
    pub properties: DrawableProperties,
    /// Trim, poster, repeat, and volume settings.
    pub playback: MediaPlaybackSettings,
    pub duration: Duration,
}

impl KeynoteEditor {
    /// List independently positioned audio clips owned by one slide.
    pub fn slide_audio(&self, slide_index: usize) -> Result<Vec<KeynoteSlideAudioInfo>> {
        self.slide_media_infos(slide_index)?
            .into_iter()
            .filter(|media| media.kind == MovieKind::Audio)
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
        options: SlideAudioOptions,
    ) -> Result<KeynoteSlideAudioInfo> {
        let (geometry, duration_seconds) = audio_creation_values(options)?;
        let context = movie_creation_context(self, slide_index)?;
        let ids = MovieObjectIds::allocate(next_object_identifier(self.package())?)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let asset = media.insert_unreferenced(preferred_filename, data)?;
        if asset.media_type != MediaType::Audio {
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
            asset.data_identifier.get(),
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
            asset.data_identifier.get(),
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
            || created.position != options.position()
            || created.duration != expected_duration
            || created_graph.info.kind != MovieKind::Audio
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
        if source.info.kind != MovieKind::Audio {
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
}

fn require_audio(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<KeynoteSlideAudioInfo> {
    let graph = editor.slide_movie_graph(slide_index, drawable_object_id)?;
    if graph.info.kind != MovieKind::Audio {
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
    let audio_data_identifier = MediaAssetId::try_from(
        audio
            .movie_data
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote audio {drawable_object_id} has no data reference"
                ))
            })?
            .identifier,
    )?;
    let position = geometry_from_drawable(&audio.super_)?
        .position
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote audio {drawable_object_id} has no position"
            ))
        })?;
    let raw = graph.message_data_type(
        drawable_object_id,
        AUDIO_ARCHIVE_MESSAGE_TYPE,
        "TSD.MovieArchive",
    )?;
    let playback_limits = movie_playback_wire_limits(editor.package())?;
    let playback =
        litchi_keynote::__decode_movie_playback_payload(raw, playback_limits).map_err(|error| {
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
    use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};
    use litchi_keynote::slide::audio::Options as SlideAudioOptions;
    use litchi_keynote::slide::media::{
        MediaLoopMode as KeynoteMediaLoopMode,
        MediaPlaybackSettings as KeynoteMediaPlaybackSettings, MediaVolume as KeynoteMediaVolume,
    };
    use litchi_keynote::{MovieSelector, Package as KeynotePackage, SlideSelector};

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 960.0, y: 540.0 };
    const DUPLICATE_OFFSET: f32 = 10.0;

    fn properties(description: &str) -> DrawableProperties {
        DrawableProperties {
            hyperlink_url: Some("https://example.test/keynote-audio".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some(description.to_owned()),
        }
    }

    fn keynote_playback(value: MediaPlaybackSettings) -> KeynoteMediaPlaybackSettings {
        KeynoteMediaPlaybackSettings {
            start_time: value.start_time,
            end_time: value.end_time,
            poster_time: value.poster_time,
            loop_mode: value
                .loop_mode
                .map(|mode| KeynoteMediaLoopMode::from_raw(mode.as_raw())),
            volume: value.volume.map(|volume| {
                KeynoteMediaVolume::new(volume.as_f32())
                    .expect("common playback volume is already validated")
            }),
        }
        .canonicalize()
        .expect("common playback values are valid Keynote settings")
    }

    fn focused_audio_package(editor: &KeynoteEditor) -> KeynotePackage {
        KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap()
    }

    fn replace_with_focused_audio_package(editor: &mut KeynoteEditor, package: &KeynotePackage) {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).unwrap();
        *editor = KeynoteEditor::from_bytes(&bytes).unwrap();
    }

    fn duplicate_audio(editor: &mut KeynoteEditor, movie: MovieSelector) -> KeynoteSlideAudioInfo {
        let package = focused_audio_package(editor);
        let commit = package
            .duplicate_slide_audio(SlideSelector::index(0), movie)
            .unwrap();
        replace_with_focused_audio_package(editor, commit.package());
        editor.slide_audio(0).unwrap().into_iter().last().unwrap()
    }

    fn remove_audio(editor: &mut KeynoteEditor, movie: MovieSelector) -> Vec<MediaAssetId> {
        let before = editor
            .media_assets()
            .unwrap()
            .into_iter()
            .map(|asset| asset.data_identifier)
            .collect::<Vec<_>>();
        let package = focused_audio_package(editor);
        let commit = package
            .remove_slide_audio(SlideSelector::index(0), movie)
            .unwrap();
        replace_with_focused_audio_package(editor, commit.package());
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

    #[test]
    fn scratch_presentation_supports_slide_audio_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch audio")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        let options = SlideAudioOptions::new(POSITION, Duration::from_millis(1_375)).unwrap();

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
        let package = KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = package
            .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap()
            .set(keynote_playback(changed_playback))
            .unwrap()
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        editor = KeynoteEditor::from_bytes(&bytes).unwrap();
        assert_eq!(
            editor.slide_audio(0).unwrap().first().unwrap().playback,
            changed_playback,
        );
        let package = KeynotePackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let commit = package
            .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap()
            .set(keynote_playback(created.playback))
            .unwrap()
            .commit()
            .unwrap();
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        editor = KeynoteEditor::from_bytes(&bytes).unwrap();

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
                .replace_media(created.audio_data_identifier, REPLACEMENT_AUDIO)
                .unwrap(),
            AUDIO
        );
        assert_eq!(
            editor.extract_media(created.audio_data_identifier).unwrap(),
            REPLACEMENT_AUDIO
        );

        let removed = remove_audio(&mut editor, MovieSelector::index(0));
        assert_eq!(removed, [created.audio_data_identifier]);
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
                SlideAudioOptions::new(POSITION, Duration::from_millis(1_375)).unwrap(),
            )
            .unwrap();
        let source_properties = properties("Duplicated Keynote audio");
        editor
            .set_slide_audio_properties(0, source.drawable_object_id, source_properties.clone())
            .unwrap();

        let duplicate = duplicate_audio(&mut editor, MovieSelector::index(0));
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
        assert_eq!(
            source_graph.object_ids.len(),
            duplicate_graph.object_ids.len()
        );
        assert_eq!(
            duplicate.audio_data_identifier,
            source.audio_data_identifier
        );
        assert_eq!(
            duplicate.position,
            DrawablePoint {
                x: source.position.x + DUPLICATE_OFFSET,
                y: source.position.y + DUPLICATE_OFFSET,
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
                .replace_media(source.audio_data_identifier, REPLACEMENT_AUDIO)
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

        let removed_source = remove_audio(&mut editor, MovieSelector::index(0));
        assert!(removed_source.is_empty());
        assert_eq!(editor.slide_audio(0).unwrap().len(), 1);
        let removed_duplicate = remove_audio(&mut editor, MovieSelector::index(0));
        assert_eq!(removed_duplicate, [source.audio_data_identifier]);
        assert!(editor.slide_audio(0).unwrap().is_empty());
        assert!(editor.slide_builds(0).unwrap().is_empty());
        assert!(editor.media_assets().unwrap().is_empty());
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    }

    #[test]
    fn invalid_slide_audio_creation_and_cross_type_edits_are_transactional() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            focused_audio_package(&editor)
                .duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        for result in [
            editor.add_slide_audio(
                0,
                "payload.bin",
                b"not audio",
                SlideAudioOptions::new(POSITION, Duration::from_secs(1)).unwrap(),
            ),
            editor.add_slide_audio(
                1,
                "audio.aiff",
                AUDIO,
                SlideAudioOptions::new(POSITION, Duration::from_secs(1)).unwrap(),
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes().unwrap(), baseline);
        }
        assert!(SlideAudioOptions::new(POSITION, Duration::ZERO).is_err());
        assert!(
            SlideAudioOptions::new(
                DrawablePoint {
                    x: f32::NAN,
                    y: 10.0,
                },
                Duration::from_secs(1),
            )
            .is_err()
        );

        let audio = editor
            .add_slide_audio(
                0,
                "audio.aiff",
                AUDIO,
                SlideAudioOptions::new(POSITION, Duration::from_secs(1)).unwrap(),
            )
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            KeynotePackage::from_bytes(&before)
                .unwrap()
                .edit_slide_movie_geometry(SlideSelector::index(0), MovieSelector::index(0))
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
                .replace_media(audio.audio_data_identifier, b"\x89PNG\r\n\x1a\nnot audio",)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
