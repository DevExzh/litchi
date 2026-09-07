//! Read-only projections of independently positioned Keynote slide audio.

use std::time::Duration;

use litchi_keynote::slide::media::MovieKind;

use super::slide_movies::movie_playback_wire_limits;
use super::*;
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
        MediaPlaybackSettings as KeynoteMediaPlaybackSettings,
        MediaProperties as KeynoteMediaProperties, MediaVolume as KeynoteMediaVolume,
        Point as KeynotePoint,
    };
    use litchi_keynote::{MovieSelector, Package as KeynotePackage, SlideSelector};

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 960.0, y: 540.0 };
    const DUPLICATE_OFFSET: f32 = 10.0;

    fn properties(description: &str) -> KeynoteMediaProperties {
        KeynoteMediaProperties::new()
            .with_hyperlink_url(Some("https://example.test/keynote-audio".to_owned()))
            .with_locked(Some(true))
            .with_aspect_ratio_locked(Some(true))
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

    fn set_audio_properties(editor: &mut KeynoteEditor, properties: KeynoteMediaProperties) {
        let package = focused_audio_package(editor);
        let commit = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap()
            .set(properties)
            .unwrap()
            .commit()
            .unwrap();
        replace_with_focused_audio_package(editor, commit.package());
    }

    #[test]
    fn source_built_audio_properties_project_through_focused_updates() {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Audio properties projection")
            .build()
            .unwrap();
        let created = add_audio(
            &mut seed,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375)).unwrap(),
        );
        let baseline = seed.slide_audio(0).unwrap().remove(0);
        let mut editor = KeynoteEditor::from_bytes(&seed.to_bytes().unwrap()).unwrap();
        let focused_before = focused_audio_package(&editor)
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap();
        assert_eq!(baseline.properties, raw_properties(&focused_before));

        for expected in [
            KeynoteMediaProperties::new()
                .with_hyperlink_url(Some(String::new()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("音声 🎵".to_owned())),
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
            set_audio_properties(&mut editor, expected.clone());
            let actual = editor.slide_audio(0).unwrap().remove(0);
            assert_eq!(actual.properties, raw_properties(&expected));
            assert_eq!(actual.audio_data_identifier, baseline.audio_data_identifier);
            assert_eq!(actual.position, baseline.position);
            assert_eq!(actual.playback, baseline.playback);
            assert_eq!(actual.duration, baseline.duration);
        }

        assert_eq!(
            created.audio_data_identifier,
            baseline.audio_data_identifier
        );
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

    fn add_audio(
        editor: &mut KeynoteEditor,
        preferred_filename: &str,
        data: &[u8],
        options: SlideAudioOptions,
    ) -> KeynoteSlideAudioInfo {
        let package = focused_audio_package(editor);
        let commit = package
            .add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)
            .unwrap();
        replace_with_focused_audio_package(editor, commit.package());
        editor.slide_audio(0).unwrap().into_iter().last().unwrap()
    }

    fn audio_position(editor: &KeynoteEditor, movie: MovieSelector) -> KeynotePoint {
        focused_audio_package(editor)
            .slide_audio_position(SlideSelector::index(0), movie)
            .unwrap()
    }

    fn set_audio_position(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
        position: DrawablePoint,
    ) {
        let package = focused_audio_package(editor);
        let point = KeynotePoint {
            x: position.x,
            y: position.y,
        };
        let commit = package
            .edit_slide_audio_position(SlideSelector::index(0), movie)
            .unwrap()
            .set(point)
            .unwrap()
            .commit()
            .unwrap();
        replace_with_focused_audio_package(editor, commit.package());
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

        let created = add_audio(&mut editor, "audio.aiff", AUDIO, options);
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
        set_audio_properties(&mut editor, changed_properties.clone());
        assert_eq!(
            editor.slide_audio(0).unwrap().first().unwrap().properties,
            raw_properties(&changed_properties)
        );
        let cleared_properties = KeynoteMediaProperties::default();
        set_audio_properties(&mut editor, cleared_properties.clone());
        assert_eq!(
            editor.slide_audio(0).unwrap().first().unwrap().properties,
            raw_properties(&cleared_properties)
        );

        let moved = DrawablePoint { x: 320.0, y: 240.0 };
        set_audio_position(&mut editor, MovieSelector::index(0), moved);
        assert_eq!(
            audio_position(&editor, MovieSelector::index(0)),
            KeynotePoint {
                x: moved.x,
                y: moved.y,
            }
        );
        assert_eq!(
            editor
                .slide_audio(0)
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == created.drawable_object_id)
                .unwrap()
                .position,
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
        let source = add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375)).unwrap(),
        );
        let source_properties = properties("Duplicated Keynote audio");
        set_audio_properties(&mut editor, source_properties.clone());

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
        assert_eq!(duplicate.properties, raw_properties(&source_properties));
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
        set_audio_position(&mut editor, MovieSelector::index(1), moved_duplicate);
        assert_eq!(
            audio_position(&editor, MovieSelector::index(0)),
            KeynotePoint {
                x: source.position.x,
                y: source.position.y,
            }
        );
        assert_eq!(
            audio_position(&editor, MovieSelector::index(1)),
            KeynotePoint {
                x: moved_duplicate.x,
                y: moved_duplicate.y,
            }
        );
        assert_eq!(
            editor
                .slide_audio(0)
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == source.drawable_object_id)
                .unwrap()
                .position,
            source.position
        );
        assert_eq!(
            editor
                .slide_audio(0)
                .unwrap()
                .into_iter()
                .find(|audio| audio.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .position,
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
            focused_audio_package(&editor).add_slide_audio(
                SlideSelector::index(0),
                "payload.bin",
                b"not audio",
                SlideAudioOptions::new(POSITION, Duration::from_secs(1)).unwrap(),
            ),
            focused_audio_package(&editor).add_slide_audio(
                SlideSelector::index(1),
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

        let audio = add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_secs(1)).unwrap(),
        );
        let before = editor.to_bytes().unwrap();
        assert!(
            KeynotePackage::from_bytes(&before)
                .unwrap()
                .edit_slide_movie_geometry(SlideSelector::index(0), MovieSelector::index(0))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        let focused_source = KeynotePackage::from_bytes(&before).unwrap();
        let invalid_position = focused_source
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap()
            .set(KeynotePoint {
                x: f32::INFINITY,
                y: 10.0,
            });
        assert!(invalid_position.is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            KeynotePackage::from_bytes(&before)
                .unwrap()
                .slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                .is_ok()
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
