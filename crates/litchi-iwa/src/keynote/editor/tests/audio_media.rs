//! Focused package coverage for slide-owned Keynote audio.
//!
//! The editor is retained here only as the host-side builder and graph/build
//! oracle.  Audio selection, properties, playback, position, media data, and
//! lifecycle assertions go through [`litchi_keynote::Package`].  The small
//! generated-protobuf projection is private to these tests and exists only to
//! correlate a focused media entry with the host build and graph observations.

#[cfg(test)]
mod cases {
    use std::collections::{HashSet, VecDeque};
    use std::error::Error;
    use std::io;
    use std::time::Duration;

    use litchi_iwa_archive::{iwa::Archive, package::Catalog};
    use litchi_iwa_protos::{kn, tsd, tsp};
    use litchi_keynote::slide::audio::Options as SlideAudioOptions;
    use litchi_keynote::slide::media::{
        MediaLoopMode, MediaPlaybackSettings, MediaProperties, MediaVolume, MovieKind,
        Point as KeynotePoint,
    };
    use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector};
    use prost::Message as _;

    use super::super::*;
    use crate::keynote::{KeynoteBuildSettings, KeynoteDocumentBuilder, KeynoteEditor};
    use crate::shapes::DrawablePoint;

    const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-audio";
    const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFFreplacement-audio";
    const POSITION: DrawablePoint = DrawablePoint { x: 960.0, y: 540.0 };
    const DUPLICATE_OFFSET: f32 = 10.0;

    type TestResult<T = ()> = std::result::Result<T, Box<dyn Error>>;

    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    struct NativeAudio {
        drawable_object_id: u64,
        data_identifier: u64,
    }

    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    struct NativeMedia {
        drawable_object_id: u64,
        data_identifier: Option<u64>,
        kind: MovieKind,
    }

    #[derive(Debug, Clone, PartialEq)]
    struct AudioSnapshot {
        native: NativeAudio,
        content: Vec<u8>,
        position: KeynotePoint,
        properties: MediaProperties,
        playback: MediaPlaybackSettings,
        duration: Duration,
    }

    fn properties(description: &str) -> MediaProperties {
        MediaProperties::new()
            .with_hyperlink_url(Some("https://example.test/keynote-audio".to_owned()))
            .with_locked(Some(true))
            .with_aspect_ratio_locked(Some(true))
            .with_accessibility_description(Some(description.to_owned()))
    }

    fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }

    /// Read native identities only for assertions that still observe the host's
    /// build and graph APIs.  This is deliberately private test code: focused
    /// callers continue to select media by source-order semantic selectors.
    fn native_media_ids(source: &[u8]) -> TestResult<Vec<NativeMedia>> {
        let catalog = Catalog::from_bytes(source)?;
        let mut movies = Vec::new();
        let mut slides = Vec::new();
        for entry in catalog
            .iter()
            .filter(|entry| entry.name().ends_with(".iwa"))
        {
            let stream = match litchi_iwa_archive::iwa::SnappyStream::decompress(entry.data()) {
                Ok(stream) => stream.into_bytes(),
                Err(_) => continue,
            };
            let archive = match Archive::parse(&stream) {
                Ok(archive) => archive,
                Err(_) => continue,
            };
            for object in &archive.objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                for message in &object.messages {
                    match message.type_ {
                        5 => {
                            if let Ok(slide) = kn::SlideArchive::decode(message.data.as_slice()) {
                                slides.push((identifier, slide));
                            }
                        },
                        3_007 => {
                            if let Ok(movie) = tsd::MovieArchive::decode(message.data.as_slice()) {
                                let kind = if movie.is_live_video == Some(true) {
                                    MovieKind::LiveVideo
                                } else if movie.audio_only == Some(true) {
                                    MovieKind::Audio
                                } else if movie.flags.is_some_and(|flags| flags & 1 != 0) {
                                    MovieKind::Placeholder
                                } else {
                                    MovieKind::File
                                };
                                movies.push((identifier, movie, kind));
                            }
                        },
                        _ => {},
                    }
                }
            }
        }

        let mut fallback = None;
        for (slide_identifier, slide) in slides {
            let candidate = slide
                .owned_drawables
                .into_iter()
                .filter_map(|reference| {
                    movies
                        .iter()
                        .find(|(identifier, movie, _)| {
                            *identifier == reference.identifier
                                && movie
                                    .super_
                                    .parent
                                    .as_ref()
                                    .is_some_and(|parent| parent.identifier == slide_identifier)
                        })
                        .map(|(identifier, movie, kind)| NativeMedia {
                            drawable_object_id: *identifier,
                            data_identifier: movie
                                .movie_data
                                .as_ref()
                                .map(|reference| reference.identifier),
                            kind: *kind,
                        })
                })
                .collect::<Vec<_>>();
            if candidate.is_empty() {
                continue;
            }
            let has_audio = candidate.iter().any(|media| media.kind == MovieKind::Audio);
            let has_file = candidate.iter().any(|media| media.kind == MovieKind::File);
            if has_audio && has_file {
                return Ok(candidate);
            }
            fallback.get_or_insert(candidate);
        }
        fallback.ok_or_else(|| io::Error::other("native package has no slide-owned media").into())
    }

    /// Count metadata-backed media entries for the one assertion that verifies
    /// lifecycle cleanup of the final shared audio asset.  This remains a private
    /// generated-fixture oracle because the focused package intentionally exposes
    /// media through slide selectors and materialized data, rather than metadata
    /// identifiers.
    fn native_asset_count(source: &[u8]) -> TestResult<usize> {
        let catalog = Catalog::from_bytes(source)?;
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == "Index/Metadata.iwa")
            .ok_or_else(|| io::Error::other("missing metadata component"))?;
        let stream = litchi_iwa_archive::iwa::SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&stream)?;
        let message = archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .find(|message| message.type_ == 11_006)
            .ok_or_else(|| io::Error::other("missing package metadata"))?;
        Ok(tsp::PackageMetadata::decode(message.data.as_slice())?
            .datas
            .len())
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct NativeAudioGraph {
        object_ids: Vec<u64>,
        data_identifier: u64,
    }

    /// Return the private object closure rooted at one native media drawable.
    /// This mirrors the graph traversal used by the removed host reader, but is
    /// intentionally a test-only oracle and never becomes part of the editor API.
    fn private_graph_ids(archive: &Archive, root: u64) -> TestResult<Vec<u64>> {
        let internal = archive
            .objects
            .iter()
            .map(|object| {
                object
                    .archive_info
                    .identifier
                    .ok_or_else(|| io::Error::other("native graph object has no identifier"))
            })
            .collect::<std::result::Result<HashSet<_>, _>>()?;
        let mut selected = HashSet::new();
        let mut pending = VecDeque::from([root]);
        while let Some(identifier) = pending.pop_front() {
            if !internal.contains(&identifier) {
                return Err(io::Error::other("native graph reference leaves its component").into());
            }
            if !selected.insert(identifier) {
                continue;
            }
            let object = archive
                .object(identifier)
                .ok_or_else(|| io::Error::other("native graph object is missing"))?;
            for reference in object.archive_info.message_infos.iter().flat_map(|info| {
                info.object_references.iter().chain(
                    info.field_infos
                        .iter()
                        .flat_map(|field| &field.object_references),
                )
            }) {
                if internal.contains(reference) && !selected.contains(reference) {
                    pending.push_back(*reference);
                }
            }
        }
        Ok(archive
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .filter(|identifier| selected.contains(identifier))
            .collect())
    }

    fn native_audio_graph(
        editor: &KeynoteEditor,
        drawable_object_id: u64,
    ) -> TestResult<NativeAudioGraph> {
        let slides = editor.slides()?;
        let slide = slides
            .first()
            .ok_or_else(|| io::Error::other("audio fixture has no slide"))?;
        let slide_id = slide.native_ids()?.slide.get();
        let graph = ObjectGraph::read(editor.package())?;
        let native: kn::SlideArchive = graph.decode_type(slide_id, 5, "KN.SlideArchive")?;
        if !native
            .owned_drawables
            .iter()
            .any(|reference| reference.identifier == drawable_object_id)
        {
            return Err(io::Error::other("native audio is not owned by its slide").into());
        }
        let archive_name = graph.archive_name(slide_id)?.to_owned();
        if graph.archive_name(drawable_object_id)? != archive_name {
            return Err(io::Error::other("native audio is outside its slide component").into());
        }
        let archive = editor.package().archive(&archive_name)?;
        let object_ids = private_graph_ids(&archive, drawable_object_id)?;
        if object_ids.contains(&slide_id) {
            return Err(io::Error::other("native audio graph reaches its owning slide").into());
        }
        let raw = graph.message_data_type(drawable_object_id, 3_007, "TSD.MovieArchive")?;
        let movie = tsd::MovieArchive::decode(raw)?;
        if movie.audio_only != Some(true) {
            return Err(io::Error::other("native graph root is not audio").into());
        }
        let data_identifier = movie
            .movie_data
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| io::Error::other("native audio has no data identifier"))?;
        Ok(NativeAudioGraph {
            object_ids,
            data_identifier,
        })
    }

    fn focused_audio_package(editor: &KeynoteEditor) -> TestResult<Package> {
        Ok(Package::from_bytes(&editor.to_bytes()?)?)
    }

    fn replace_with_focused_package(editor: &mut KeynoteEditor, package: &Package) -> TestResult {
        let bytes = package_bytes(package)?;
        *editor = KeynoteEditor::from_bytes(&bytes)?;
        Ok(())
    }

    fn audio_snapshot(editor: &KeynoteEditor, movie_position: usize) -> TestResult<AudioSnapshot> {
        let package = focused_audio_package(editor)?;
        let slide = package
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("audio fixture has no slide"))?;
        let movie = *slide
            .movies()
            .get(movie_position)
            .ok_or_else(|| io::Error::other("audio fixture has no selected media"))?;
        if !movie.is_audio() {
            return Err(io::Error::other("selected media is not audio").into());
        }
        let selector = MovieSelector::index(movie_position);
        let content = package
            .slide_media_data(SlideSelector::index(0), selector, MediaPart::Content)?
            .to_vec();
        let position = package.slide_audio_position(SlideSelector::index(0), selector)?;
        let properties = package.slide_media_properties(SlideSelector::index(0), selector)?;
        let playback = package
            .slide_movie_playback_settings(SlideSelector::index(0), selector)?
            .ok_or_else(|| io::Error::other("audio fixture has no playback settings"))?;
        let native = native_media_ids(&editor.to_bytes()?)?
            .get(movie_position)
            .copied()
            .ok_or_else(|| io::Error::other("audio fixture has no native audio identity"))?;
        if native.kind != MovieKind::Audio {
            return Err(io::Error::other("selected native media is not audio").into());
        }
        let data_identifier = native
            .data_identifier
            .ok_or_else(|| io::Error::other("native audio has no data identifier"))?;
        Ok(AudioSnapshot {
            native: NativeAudio {
                drawable_object_id: native.drawable_object_id,
                data_identifier,
            },
            content,
            position,
            properties,
            playback,
            duration: movie
                .duration()
                .ok_or_else(|| io::Error::other("audio fixture has no duration"))?,
        })
    }

    fn audio_positions(editor: &KeynoteEditor) -> TestResult<Vec<usize>> {
        let package = focused_audio_package(editor)?;
        Ok(package
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("audio fixture has no slide"))?
            .movies()
            .iter()
            .enumerate()
            .filter_map(|(position, movie)| movie.is_audio().then_some(position))
            .collect())
    }

    fn audio_snapshots(editor: &KeynoteEditor) -> TestResult<Vec<AudioSnapshot>> {
        audio_positions(editor)?
            .into_iter()
            .map(|position| audio_snapshot(editor, position))
            .collect()
    }

    fn audio_data_ids(editor: &KeynoteEditor) -> TestResult<Vec<u64>> {
        if audio_positions(editor)?.is_empty() {
            return Ok(Vec::new());
        }
        let mut identifiers = Vec::new();
        for media in native_media_ids(&editor.to_bytes()?)? {
            if media.kind == MovieKind::Audio {
                identifiers.push(
                    media
                        .data_identifier
                        .ok_or_else(|| io::Error::other("native audio has no data identifier"))?,
                );
            }
        }
        identifiers.sort_unstable();
        identifiers.dedup();
        Ok(identifiers)
    }

    fn add_audio(
        editor: &mut KeynoteEditor,
        preferred_filename: &str,
        data: &[u8],
        options: SlideAudioOptions,
    ) -> TestResult<AudioSnapshot> {
        let package = focused_audio_package(editor)?;
        let commit =
            package.add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)?;
        replace_with_focused_package(editor, commit.package())?;
        let position = audio_positions(editor)?
            .last()
            .copied()
            .ok_or_else(|| io::Error::other("focused audio creation produced no audio"))?;
        audio_snapshot(editor, position)
    }

    fn set_audio_properties(
        editor: &mut KeynoteEditor,
        movie_position: usize,
        properties: MediaProperties,
    ) -> TestResult {
        let package = focused_audio_package(editor)?;
        let commit = package
            .edit_slide_media_properties(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
            )?
            .set(properties)?
            .commit()?;
        replace_with_focused_package(editor, commit.package())
    }

    fn set_audio_playback(
        editor: &mut KeynoteEditor,
        movie_position: usize,
        playback: MediaPlaybackSettings,
    ) -> TestResult {
        let package = focused_audio_package(editor)?;
        let commit = package
            .edit_slide_movie_playback_settings(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
            )?
            .set(playback)?
            .commit()?;
        replace_with_focused_package(editor, commit.package())
    }

    fn set_audio_position(
        editor: &mut KeynoteEditor,
        movie_position: usize,
        position: DrawablePoint,
    ) -> TestResult {
        let package = focused_audio_package(editor)?;
        let commit = package
            .edit_slide_audio_position(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
            )?
            .set(KeynotePoint {
                x: position.x,
                y: position.y,
            })?
            .commit()?;
        replace_with_focused_package(editor, commit.package())
    }

    fn replace_audio_data(
        editor: &mut KeynoteEditor,
        movie_position: usize,
        data: &[u8],
    ) -> TestResult {
        let package = focused_audio_package(editor)?;
        let commit = package
            .edit_slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
                MediaPart::Content,
            )?
            .set(data)?
            .commit()?;
        replace_with_focused_package(editor, commit.package())
    }

    fn duplicate_audio(
        editor: &mut KeynoteEditor,
        movie_position: usize,
    ) -> TestResult<AudioSnapshot> {
        let package = focused_audio_package(editor)?;
        let commit = package.duplicate_slide_audio(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
        )?;
        replace_with_focused_package(editor, commit.package())?;
        let position = audio_positions(editor)?
            .last()
            .copied()
            .ok_or_else(|| io::Error::other("focused audio duplication produced no audio"))?;
        audio_snapshot(editor, position)
    }

    fn remove_audio(
        editor: &mut KeynoteEditor,
        movie_position: usize,
    ) -> TestResult<(Vec<u64>, Vec<Vec<u8>>)> {
        let before_data_ids = audio_data_ids(editor)?;
        let before = audio_snapshots(editor)?
            .into_iter()
            .map(|audio| audio.content)
            .collect::<Vec<_>>();
        let package = focused_audio_package(editor)?;
        let commit = package.remove_slide_audio(
            SlideSelector::index(0),
            MovieSelector::index(movie_position),
        )?;
        replace_with_focused_package(editor, commit.package())?;
        let after = audio_snapshots(editor)?
            .into_iter()
            .map(|audio| audio.content)
            .collect::<Vec<_>>();
        let after_data_ids = audio_data_ids(editor)?;
        let released_data_ids = before_data_ids
            .into_iter()
            .filter(|identifier| !after_data_ids.contains(identifier))
            .collect();
        Ok((
            released_data_ids,
            before
                .into_iter()
                .filter(|content| !after.contains(content))
                .collect(),
        ))
    }

    fn audio_count(editor: &KeynoteEditor) -> TestResult<usize> {
        Ok(focused_audio_package(editor)?
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("audio fixture has no slide"))?
            .audio()
            .count())
    }

    #[test]
    fn source_built_audio_properties_project_through_focused_updates() -> TestResult {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Audio properties projection")
            .build()?;
        let created = add_audio(
            &mut seed,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375))?,
        )?;
        let baseline = audio_snapshot(&seed, 0)?;
        assert_eq!(created, baseline);
        let mut editor = KeynoteEditor::from_bytes(&seed.to_bytes()?)?;
        let focused_before = focused_audio_package(&editor)?
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))?;
        assert_eq!(baseline.properties, focused_before);

        for expected in [
            MediaProperties::new()
                .with_hyperlink_url(Some(String::new()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("音声 🎵".to_owned())),
            MediaProperties::new()
                .with_hyperlink_url(Some("opaque target 日本語".to_owned()))
                .with_locked(Some(true))
                .with_aspect_ratio_locked(Some(true))
                .with_accessibility_description(Some("locked 🔒".to_owned())),
            MediaProperties::new()
                .with_hyperlink_url(Some("second target".to_owned()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("unlocked again".to_owned())),
            MediaProperties::default(),
        ] {
            set_audio_properties(&mut editor, 0, expected.clone())?;
            let actual = audio_snapshot(&editor, 0)?;
            assert_eq!(actual.properties, expected);
            assert_eq!(actual.native, baseline.native);
            assert_eq!(actual.content, baseline.content);
            assert_eq!(actual.position, baseline.position);
            assert_eq!(actual.playback, baseline.playback);
            assert_eq!(actual.duration, baseline.duration);
        }

        assert_eq!(created.content, baseline.content);
        Ok(())
    }

    #[test]
    fn scratch_presentation_supports_slide_audio_crud() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch audio")
            .subtitle("No embedded package")
            .build()?;
        let options = SlideAudioOptions::new(POSITION, Duration::from_millis(1_375))?;

        let created = add_audio(&mut editor, "audio.aiff", AUDIO, options)?;
        let package = focused_audio_package(&editor)?;
        let slide = package.show()?.slides().first().unwrap();
        assert!(slide.video_movies().next().is_none());
        assert_eq!(slide.audio().count(), 1);
        assert_eq!(audio_snapshot(&editor, 0)?, created);
        assert_eq!(
            created.position,
            KeynotePoint {
                x: POSITION.x,
                y: POSITION.y
            }
        );
        assert_eq!(created.content, AUDIO);
        let builds = editor.slide_builds(0)?;
        assert_eq!(builds.len(), 1);
        assert_eq!(
            builds[0].drawable_object_id,
            created.native.drawable_object_id
        );
        assert_eq!(builds[0].settings, KeynoteBuildSettings::audio_start());
        assert_eq!(builds[0].chunks.len(), 1);

        let roundtripped = KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        assert_eq!(audio_snapshot(&roundtripped, 0)?, created);

        let changed_playback = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::Repeat),
            volume: Some(MediaVolume::new(0.75)?),
            ..created.playback
        };
        set_audio_playback(&mut editor, 0, changed_playback)?;
        assert_eq!(audio_snapshot(&editor, 0)?.playback, changed_playback);
        set_audio_playback(&mut editor, 0, created.playback)?;
        assert_eq!(audio_snapshot(&editor, 0)?.playback, created.playback);

        let changed_properties = properties("Accessible Keynote audio");
        set_audio_properties(&mut editor, 0, changed_properties.clone())?;
        assert_eq!(audio_snapshot(&editor, 0)?.properties, changed_properties);
        let cleared_properties = MediaProperties::default();
        set_audio_properties(&mut editor, 0, cleared_properties.clone())?;
        assert_eq!(audio_snapshot(&editor, 0)?.properties, cleared_properties);

        let moved = DrawablePoint { x: 320.0, y: 240.0 };
        set_audio_position(&mut editor, 0, moved)?;
        assert_eq!(
            focused_audio_package(&editor)?
                .slide_audio_position(SlideSelector::index(0), MovieSelector::index(0),)?,
            KeynotePoint {
                x: moved.x,
                y: moved.y,
            }
        );
        assert_eq!(
            audio_snapshot(&editor, 0)?.position,
            KeynotePoint {
                x: moved.x,
                y: moved.y
            }
        );

        replace_audio_data(&mut editor, 0, REPLACEMENT_AUDIO)?;
        let replaced = audio_snapshot(&editor, 0)?;
        assert_eq!(replaced.content, REPLACEMENT_AUDIO);
        assert_eq!(
            replaced.native.data_identifier,
            created.native.data_identifier
        );

        let (removed_data_ids, removed_content) = remove_audio(&mut editor, 0)?;
        assert_eq!(removed_data_ids, [created.native.data_identifier]);
        assert_eq!(removed_content, [REPLACEMENT_AUDIO.to_vec()]);
        assert_eq!(audio_count(&editor)?, 0);
        assert!(editor.slide_builds(0)?.is_empty());
        assert!(
            focused_audio_package(&editor)?
                .show()?
                .slides()
                .first()
                .unwrap()
                .movies()
                .is_empty()
        );
        assert!(audio_data_ids(&editor)?.is_empty());
        assert_eq!(native_asset_count(&editor.to_bytes()?)?, 0);
        KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        Ok(())
    }

    #[test]
    fn scratch_presentation_supports_native_audio_duplication() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch audio")
            .subtitle("No embedded package")
            .build()?;
        let source = add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375))?,
        )?;
        let source_properties = properties("Duplicated Keynote audio");
        set_audio_properties(&mut editor, 0, source_properties.clone())?;

        let duplicate = duplicate_audio(&mut editor, 0)?;
        assert_ne!(duplicate.native, source.native);
        assert_eq!(
            duplicate.native.data_identifier,
            source.native.data_identifier
        );
        let source_graph = native_audio_graph(&editor, source.native.drawable_object_id)?;
        let duplicate_graph = native_audio_graph(&editor, duplicate.native.drawable_object_id)?;
        assert_eq!(
            source_graph.data_identifier,
            duplicate_graph.data_identifier
        );
        assert_eq!(source_graph.data_identifier, source.native.data_identifier);
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
        assert_eq!(duplicate.content, source.content);
        assert_eq!(
            duplicate.position,
            KeynotePoint {
                x: source.position.x + DUPLICATE_OFFSET,
                y: source.position.y + DUPLICATE_OFFSET,
            }
        );
        assert_eq!(duplicate.duration, source.duration);
        assert_eq!(duplicate.properties, source_properties);
        let duplicate_builds = editor
            .slide_builds(0)?
            .into_iter()
            .filter(|build| build.drawable_object_id == duplicate.native.drawable_object_id)
            .collect::<Vec<_>>();
        assert_eq!(duplicate_builds.len(), 1);
        assert_eq!(
            duplicate_builds[0].settings,
            KeynoteBuildSettings::audio_start()
        );
        assert_eq!(duplicate_builds[0].chunks.len(), 1);

        let moved_duplicate = DrawablePoint { x: 320.0, y: 240.0 };
        set_audio_position(&mut editor, 1, moved_duplicate)?;
        assert_eq!(audio_snapshot(&editor, 0)?.position, source.position);
        assert_eq!(
            audio_snapshot(&editor, 1)?.position,
            KeynotePoint {
                x: moved_duplicate.x,
                y: moved_duplicate.y,
            }
        );
        replace_audio_data(&mut editor, 0, REPLACEMENT_AUDIO)?;
        let duplicate_after_replacement = audio_snapshot(&editor, 1)?;
        assert_eq!(duplicate_after_replacement.content, REPLACEMENT_AUDIO);
        assert_eq!(
            duplicate_after_replacement.native.data_identifier,
            source.native.data_identifier
        );

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        assert_eq!(audio_count(&reopened)?, 2);
        assert_eq!(
            audio_snapshot(&reopened, 0)?.native.data_identifier,
            source.native.data_identifier
        );
        assert_eq!(
            audio_snapshot(&reopened, 1)?.position,
            audio_snapshot(&editor, 1)?.position
        );
        assert_eq!(
            audio_snapshot(&reopened, 1)?.native.data_identifier,
            source.native.data_identifier
        );
        assert_eq!(
            reopened
                .slide_builds(0)?
                .into_iter()
                .filter(|build| build.drawable_object_id == duplicate.native.drawable_object_id)
                .count(),
            1
        );

        let (removed_source_data_ids, removed_source_content) = remove_audio(&mut editor, 0)?;
        assert!(removed_source_data_ids.is_empty());
        assert!(removed_source_content.is_empty());
        assert_eq!(audio_count(&editor)?, 1);
        let (removed_duplicate_data_ids, removed_duplicate_content) = remove_audio(&mut editor, 0)?;
        assert_eq!(removed_duplicate_data_ids, [source.native.data_identifier]);
        assert_eq!(removed_duplicate_content, [REPLACEMENT_AUDIO.to_vec()]);
        assert_eq!(audio_count(&editor)?, 0);
        assert!(editor.slide_builds(0)?.is_empty());
        assert!(
            focused_audio_package(&editor)?
                .show()?
                .slides()
                .first()
                .unwrap()
                .movies()
                .is_empty()
        );
        assert!(audio_data_ids(&editor)?.is_empty());
        assert_eq!(native_asset_count(&editor.to_bytes()?)?, 0);
        KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        Ok(())
    }

    #[test]
    fn invalid_slide_audio_creation_and_cross_type_edits_are_transactional() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new().build()?;
        let baseline = editor.to_bytes()?;
        assert!(
            focused_audio_package(&editor)?
                .duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
                .is_err()
        );
        assert_eq!(editor.to_bytes()?, baseline);
        for result in [
            focused_audio_package(&editor)?.add_slide_audio(
                SlideSelector::index(0),
                "payload.bin",
                b"not audio",
                SlideAudioOptions::new(POSITION, Duration::from_secs(1))?,
            ),
            focused_audio_package(&editor)?.add_slide_audio(
                SlideSelector::index(1),
                "audio.aiff",
                AUDIO,
                SlideAudioOptions::new(POSITION, Duration::from_secs(1))?,
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes()?, baseline);
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

        let _audio = add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_secs(1))?,
        )?;
        let before = editor.to_bytes()?;
        assert!(
            Package::from_bytes(&before)?
                .edit_slide_movie_geometry(SlideSelector::index(0), MovieSelector::index(0))
                .is_err()
        );
        assert_eq!(editor.to_bytes()?, before);
        let focused_source = Package::from_bytes(&before)?;
        let invalid_position = focused_source
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
            .set(KeynotePoint {
                x: f32::INFINITY,
                y: 10.0,
            });
        assert!(invalid_position.is_err());
        assert_eq!(editor.to_bytes()?, before);
        assert!(
            Package::from_bytes(&before)?
                .slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                .is_ok()
        );
        assert_eq!(editor.to_bytes()?, before);
        assert!(
            Package::from_bytes(&before)?
                .edit_slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(0),
                    MediaPart::Content,
                )?
                .set(b"\x89PNG\r\n\x1a\nnot audio")
                .is_err()
        );
        assert_eq!(editor.to_bytes()?, before);
        Ok(())
    }
}
