//! Keynote movie compatibility tests through the focused semantic package.
//!
//! The migration host intentionally keeps only a small native graph oracle in
//! this module.  Movie values, properties, playback, geometry, and media
//! bytes are asserted through `litchi_keynote::Package`; native identifiers
//! are used only when checking graph locality or constructing sparse physical
//! fixtures.

#[cfg(test)]
mod cases {
    use super::super::*;

    use std::time::Duration;

    use crate::archive::RawMessage;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::wire::remove_repeated_length_delimited_field_where;
    use litchi_keynote::slide::audio::Options as SlideAudioOptions;
    use litchi_keynote::slide::media::{
        MediaLoopMode, MediaPlaybackSettings, MediaProperties, MediaVolume, MovieInfo, MovieKind,
        Point as MediaPoint, Size as MediaSize,
    };
    use litchi_keynote::slide::movie::Options as SlideMovieOptions;
    use litchi_keynote::{
        MediaPart, MovieSelector, Package as FocusedKeynotePackage, SlideSelector,
    };
    use prost::Message as _;

    const MOVIE_MESSAGE_TYPE: u32 = 3_007;
    const SLIDE_MESSAGE_TYPE: u32 = 5;
    const MOVIE_DATA_FIELD: u32 = 14;
    const POSTER_IMAGE_DATA_FIELD: u32 = 15;

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
    const NATIVE_MEDIA_BASELINE: &[u8] = include_bytes!(
        "../../../../../../test-data/iwork/keynote/media-comments-baseline-native.key"
    );

    type TestResult<T = ()> = std::result::Result<T, Box<dyn std::error::Error>>;

    #[derive(Debug, Clone, PartialEq)]
    struct MovieSnapshot {
        info: MovieInfo,
        properties: MediaProperties,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct MovieGraphEntry {
        object_id: u64,
        component: String,
        parent_id: Option<u64>,
        movie_data_id: Option<u64>,
        poster_data_id: Option<u64>,
        movie_data_field_count: usize,
        poster_data_field_count: usize,
    }

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct MovieGraphOracle {
        slide_id: u64,
        slide_component: String,
        movies: Vec<MovieGraphEntry>,
    }

    fn options() -> SlideMovieOptions {
        SlideMovieOptions::new(POSITION, DISPLAY_SIZE, Duration::from_millis(1_250))
            .unwrap()
            .with_natural_size(NATURAL_SIZE)
            .unwrap()
    }

    fn focused_package(editor: &KeynoteEditor) -> TestResult<FocusedKeynotePackage> {
        Ok(FocusedKeynotePackage::from_bytes(&editor.to_bytes()?)?)
    }

    fn focused_bytes(package: &FocusedKeynotePackage) -> TestResult<Vec<u8>> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }

    fn reopen_focused_package(
        editor: &mut KeynoteEditor,
        package: &FocusedKeynotePackage,
    ) -> TestResult<()> {
        *editor = KeynoteEditor::from_bytes(&focused_bytes(package)?)?;
        Ok(())
    }

    fn movie_snapshot(package: &FocusedKeynotePackage, movie: usize) -> TestResult<MovieSnapshot> {
        let info = package
            .slides()?
            .first()
            .and_then(|slide| slide.movies().get(movie))
            .copied()
            .ok_or_else(|| format!("focused movie {movie} is missing"))?;
        let properties = package
            .slide_media_properties(SlideSelector::index(0), MovieSelector::index(movie))
            .map_err(|error| format!("properties snapshot for media {movie}: {error}"))?;
        Ok(MovieSnapshot { info, properties })
    }

    fn movie_graph(editor: &KeynoteEditor, slide_index: usize) -> TestResult<MovieGraphOracle> {
        let slides = editor.slides()?;
        let slide = slides
            .get(slide_index)
            .ok_or_else(|| format!("slide {slide_index} is missing"))?;
        let slide_id = slide.native_ids()?.slide.get();
        let graph = ObjectGraph::read(editor.package())?;
        let slide_archive: kn::SlideArchive =
            graph.decode_type(slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let slide_component = graph.archive_name(slide_id)?.to_owned();
        let mut movies = Vec::new();
        for reference in slide_archive.owned_drawables {
            let object_id = reference.identifier;
            let Some(messages) = graph.objects.get(&object_id) else {
                continue;
            };
            if !messages
                .iter()
                .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            {
                continue;
            }
            let component = graph.archive_name(object_id)?.to_owned();
            let raw = graph.message_data_type(object_id, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
            let movie = tsd::MovieArchive::decode(raw)?;
            movies.push(MovieGraphEntry {
                object_id,
                component,
                parent_id: movie.super_.parent.map(|reference| reference.identifier),
                movie_data_id: movie
                    .movie_data
                    .as_ref()
                    .map(|reference| reference.identifier),
                poster_data_id: movie
                    .poster_image_data
                    .as_ref()
                    .map(|reference| reference.identifier),
                movie_data_field_count: crate::wire::repeated_length_delimited_payloads(
                    raw,
                    MOVIE_DATA_FIELD,
                )?
                .len(),
                poster_data_field_count: crate::wire::repeated_length_delimited_payloads(
                    raw,
                    POSTER_IMAGE_DATA_FIELD,
                )?
                .len(),
            });
        }
        Ok(MovieGraphOracle {
            slide_id,
            slide_component,
            movies,
        })
    }

    fn assert_movie_graph_is_local(graph: &MovieGraphOracle) {
        assert!(!graph.movies.is_empty());
        for movie in &graph.movies {
            assert_eq!(movie.component, graph.slide_component);
            assert_eq!(movie.parent_id, Some(graph.slide_id));
            assert_eq!(
                movie.movie_data_field_count,
                usize::from(movie.movie_data_id.is_some())
            );
            assert_eq!(
                movie.poster_data_field_count,
                usize::from(movie.poster_data_id.is_some())
            );
        }
    }

    fn native_movie_properties(
        editor: &KeynoteEditor,
        object_id: u64,
    ) -> TestResult<MediaProperties> {
        let graph = ObjectGraph::read(editor.package())?;
        let raw = graph.message_data_type(object_id, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
        let movie = tsd::MovieArchive::decode(raw)?;
        Ok(MediaProperties::from_parts(
            movie.super_.hyperlink_url,
            movie.super_.locked,
            movie.super_.aspect_ratio_locked,
            movie.super_.accessibility_description,
        ))
    }

    fn add_audio(
        editor: &mut KeynoteEditor,
        preferred_filename: &str,
        data: &[u8],
        options: SlideAudioOptions,
    ) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit =
            package.add_slide_audio(SlideSelector::index(0), preferred_filename, data, options)?;
        reopen_focused_package(editor, commit.package())
    }

    fn add_movie(
        editor: &mut KeynoteEditor,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: SlideMovieOptions,
    ) -> TestResult<MovieSnapshot> {
        add_movie_at(
            editor,
            0,
            preferred_movie_filename,
            movie_data,
            preferred_poster_filename,
            poster_data,
            options,
        )
    }

    fn add_movie_at(
        editor: &mut KeynoteEditor,
        slide_index: usize,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: SlideMovieOptions,
    ) -> TestResult<MovieSnapshot> {
        let package = focused_package(editor)?;
        let commit = package
            .add_slide_movie(
                SlideSelector::index(slide_index),
                preferred_movie_filename,
                movie_data,
                preferred_poster_filename,
                poster_data,
                options,
            )
            .map_err(|error| format!("focused movie creation failed: {error}"))?;
        reopen_focused_package(editor, commit.package())?;
        let package = focused_package(editor)?;
        Ok(movie_snapshot(&package, 0)?)
    }

    fn remove_movie_poster(
        editor: &mut KeynoteEditor,
        drawable_object_id: u64,
        poster_data_identifier: u64,
    ) -> TestResult<()> {
        let archive_name = ObjectGraph::read(editor.package())?
            .archive_name(drawable_object_id)?
            .to_owned();
        let mut package = editor.package().clone();
        package.update_archive(&archive_name, |archive| {
            let object = archive
                .object_mut(drawable_object_id)
                .ok_or_else(|| Error::InvalidFormat("movie object is missing".to_owned()))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                .ok_or_else(|| Error::InvalidFormat("movie message is missing".to_owned()))?;
            let data = remove_repeated_length_delimited_field_where(
                &object.messages[message_index].data,
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
                .retain(|identifier| *identifier != poster_data_identifier);
            Ok(())
        })?;
        *editor = KeynoteEditor::from_bytes(&package.to_bytes()?)?;
        Ok(())
    }

    fn remove_movie_data(
        editor: &mut KeynoteEditor,
        drawable_object_id: u64,
        movie_data_identifier: u64,
    ) -> TestResult<()> {
        let archive_name = ObjectGraph::read(editor.package())?
            .archive_name(drawable_object_id)?
            .to_owned();
        let mut package = editor.package().clone();
        package.update_archive(&archive_name, |archive| {
            let object = archive
                .object_mut(drawable_object_id)
                .ok_or_else(|| Error::InvalidFormat("movie object is missing".to_owned()))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                .ok_or_else(|| Error::InvalidFormat("movie message is missing".to_owned()))?;
            let data = remove_repeated_length_delimited_field_where(
                &object.messages[message_index].data,
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
                .retain(|identifier| *identifier != movie_data_identifier);
            Ok(())
        })?;
        *editor = KeynoteEditor::from_bytes(&package.to_bytes()?)?;
        Ok(())
    }

    fn duplicate_movie(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
    ) -> TestResult<MovieSnapshot> {
        let package = focused_package(editor)?;
        let commit = package.duplicate_slide_movie(SlideSelector::index(0), movie)?;
        reopen_focused_package(editor, commit.package())?;
        let package = focused_package(editor)?;
        let index = package
            .slides()?
            .first()
            .map_or(0, |slide| slide.movies().len() - 1);
        Ok(movie_snapshot(&package, index)?)
    }

    fn remove_movie(editor: &mut KeynoteEditor, movie: MovieSelector) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit = package.remove_slide_movie(SlideSelector::index(0), movie)?;
        reopen_focused_package(editor, commit.package())
    }

    fn movie_title(editor: &KeynoteEditor, movie: MovieSelector) -> TestResult<Option<String>> {
        Ok(focused_package(editor)?.slide_movie_title(SlideSelector::index(0), movie)?)
    }

    fn set_movie_title(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
        title: &str,
    ) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit = package
            .edit_slide_movie_title(SlideSelector::index(0), movie)?
            .set(title)?
            .commit()?;
        reopen_focused_package(editor, commit.package())
    }

    fn remove_movie_title(editor: &mut KeynoteEditor, movie: MovieSelector) -> TestResult<bool> {
        let package = focused_package(editor)?;
        let edit = package.edit_slide_movie_title(SlideSelector::index(0), movie)?;
        let had_title = edit.before().is_some();
        let commit = edit.clear()?.commit()?;
        if !commit.patch().is_noop() {
            reopen_focused_package(editor, commit.package())?;
        }
        Ok(had_title)
    }

    fn movie_caption(editor: &KeynoteEditor, movie: MovieSelector) -> TestResult<Option<String>> {
        Ok(focused_package(editor)?.slide_movie_caption(SlideSelector::index(0), movie)?)
    }

    fn set_movie_caption(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
        caption: &str,
    ) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit = package
            .edit_slide_movie_caption(SlideSelector::index(0), movie)?
            .set(caption)?
            .commit()?;
        reopen_focused_package(editor, commit.package())
    }

    fn remove_movie_caption(editor: &mut KeynoteEditor, movie: MovieSelector) -> TestResult<bool> {
        let package = focused_package(editor)?;
        let edit = package.edit_slide_movie_caption(SlideSelector::index(0), movie)?;
        let had_caption = edit.before().is_some();
        let commit = edit.clear()?.commit()?;
        if !commit.patch().is_noop() {
            reopen_focused_package(editor, commit.package())?;
        }
        Ok(had_caption)
    }

    fn properties(description: &str) -> MediaProperties {
        MediaProperties::new()
            .with_hyperlink_url(Some("https://example.test/keynote-movie".to_owned()))
            .with_locked(Some(true))
            .with_aspect_ratio_locked(Some(false))
            .with_accessibility_description(Some(description.to_owned()))
    }

    fn set_movie_properties(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
        properties: MediaProperties,
    ) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit = package
            .edit_slide_media_properties(SlideSelector::index(0), movie)?
            .set(properties)?
            .commit()?;
        reopen_focused_package(editor, commit.package())
    }

    fn replace_movie_media(
        editor: &mut KeynoteEditor,
        movie: MovieSelector,
        part: MediaPart,
        replacement: &[u8],
    ) -> TestResult<()> {
        let package = focused_package(editor)?;
        let commit = package
            .edit_slide_media_data(SlideSelector::index(0), movie, part)?
            .set(replacement)?
            .commit()?;
        reopen_focused_package(editor, commit.package())
    }

    fn media_bytes(
        package: &FocusedKeynotePackage,
        movie: MovieSelector,
        part: MediaPart,
    ) -> TestResult<Vec<u8>> {
        Ok(package
            .slide_media_data(SlideSelector::index(0), movie, part)?
            .to_vec())
    }

    #[test]
    fn source_built_movie_properties_project_through_focused_updates() -> TestResult {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties projection")
            .build()?;
        add_movie(
            &mut seed,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;
        let mut editor = KeynoteEditor::from_bytes(&seed.to_bytes()?)?;
        let package = focused_package(&editor)?;
        let baseline = movie_snapshot(&package, 0)?;
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Content)?,
            MOVIE
        );
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Poster)?,
            POSTER
        );

        for expected in [
            MediaProperties::new()
                .with_hyperlink_url(Some(String::new()))
                .with_locked(Some(false))
                .with_aspect_ratio_locked(Some(false))
                .with_accessibility_description(Some("媒体 🎬".to_owned())),
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
            set_movie_properties(&mut editor, MovieSelector::index(0), expected.clone())?;
            let package = focused_package(&editor)?;
            let actual = movie_snapshot(&package, 0)?;
            assert_eq!(actual.properties, expected);
            assert_eq!(actual.info, baseline.info);
            assert_eq!(
                media_bytes(&package, MovieSelector::index(0), MediaPart::Content)?,
                MOVIE
            );
            assert_eq!(
                media_bytes(&package, MovieSelector::index(0), MediaPart::Poster)?,
                POSTER
            );
        }
        Ok(())
    }

    #[test]
    fn native_file_movie_properties_match_focused_source_projection() -> TestResult {
        let editor = KeynoteEditor::from_bytes(NATIVE_MEDIA_BASELINE)?;
        let package = focused_package(&editor)?;
        let graph = movie_graph(&editor, 0)?;
        assert_movie_graph_is_local(&graph);
        let movies = package
            .slides()?
            .first()
            .ok_or("focused slide is missing")?
            .movies();
        assert_eq!(movies.len(), graph.movies.len());
        let mut file_count = 0;
        for (movie_position, movie) in movies.iter().enumerate() {
            let properties = package.slide_media_properties(
                SlideSelector::index(0),
                MovieSelector::index(movie_position),
            )?;
            if movie.kind() != MovieKind::File {
                continue;
            }
            let entry = &graph.movies[movie_position];
            assert_eq!(entry.movie_data_field_count, 1);
            assert_eq!(entry.poster_data_field_count, 1);
            assert_eq!(entry.parent_id, Some(graph.slide_id));
            assert_eq!(entry.component, graph.slide_component);
            assert_eq!(
                properties,
                native_movie_properties(&editor, entry.object_id)?,
                "focused properties must match the independent native MovieArchive projection"
            );
            file_count += 1;
        }
        assert!(file_count > 0);
        Ok(())
    }

    #[test]
    fn source_built_movie_properties_support_an_absent_poster() -> TestResult {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties without poster")
            .build()?;
        add_movie(
            &mut seed,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;
        let seed_editor = KeynoteEditor::from_bytes(&seed.to_bytes()?)?;
        let graph = movie_graph(&seed_editor, 0)?;
        let movie = graph.movies.first().ok_or("movie graph is empty")?;
        let mut editor = seed_editor;
        remove_movie_poster(
            &mut editor,
            movie.object_id,
            movie.poster_data_id.ok_or("poster data is missing")?,
        )?;

        let package = focused_package(&editor)?;
        let baseline = movie_snapshot(&package, 0)?;
        assert_eq!(baseline.info.kind(), MovieKind::File);
        assert!(
            package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(0),
                    MediaPart::Poster
                )
                .is_err()
        );
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Content)?,
            MOVIE
        );

        let expected = MediaProperties::new()
            .with_hyperlink_url(Some("opaque target 日本語".to_owned()))
            .with_locked(Some(false))
            .with_aspect_ratio_locked(Some(true))
            .with_accessibility_description(Some("No poster 🎬".to_owned()));
        set_movie_properties(&mut editor, MovieSelector::index(0), expected.clone())?;

        let actual = movie_snapshot(&focused_package(&editor)?, 0)?;
        assert_eq!(actual.properties, expected);
        assert_eq!(actual.info, baseline.info);
        Ok(())
    }

    #[test]
    fn source_built_movie_properties_support_missing_content_data() -> TestResult {
        let mut seed = KeynoteDocumentBuilder::new()
            .title("Movie properties without content")
            .build()?;
        add_movie(
            &mut seed,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;
        let seed_editor = KeynoteEditor::from_bytes(&seed.to_bytes()?)?;
        let graph = movie_graph(&seed_editor, 0)?;
        let movie = graph.movies.first().ok_or("movie graph is empty")?;
        let mut editor = seed_editor;
        remove_movie_data(
            &mut editor,
            movie.object_id,
            movie.movie_data_id.ok_or("movie data is missing")?,
        )?;

        let package = focused_package(&editor)?;
        let snapshot = movie_snapshot(&package, 0)?;
        assert_eq!(snapshot.info.kind(), MovieKind::File);
        assert!(
            package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(0),
                    MediaPart::Content
                )
                .is_err()
        );
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Poster)?,
            POSTER
        );
        assert_eq!(
            snapshot.properties,
            MediaProperties::from_parts(None, Some(false), Some(true), None)
        );

        let before = focused_bytes(&package)?;
        let rejected = package
            .edit_slide_media_properties(SlideSelector::index(0), MovieSelector::index(0))
            .and_then(|edit| {
                edit.set(
                    MediaProperties::new()
                        .with_accessibility_description(Some("missing content".to_owned())),
                )
            })
            .and_then(|edit| edit.commit());
        assert!(rejected.is_err());
        assert_eq!(focused_bytes(&package)?, before);
        Ok(())
    }

    #[test]
    fn scratch_presentation_supports_movie_crud_without_a_source_drawable() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch movie")
            .subtitle("No embedded package")
            .build()?;
        assert!(movie_graph(&editor, 0)?.movies.is_empty());

        let created = add_movie(
            &mut editor,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;
        assert_eq!(created.info.kind(), MovieKind::File);
        assert_eq!(
            created.info.original_size(),
            Some(MediaSize {
                width: 1_280.0,
                height: 720.0
            })
        );
        assert_eq!(
            created.info.natural_size(),
            Some(MediaSize {
                width: 1_280.0,
                height: 720.0
            })
        );
        assert_eq!(
            created.info.position(),
            Some(MediaPoint { x: 100.0, y: 120.0 })
        );
        assert_eq!(
            created.info.size(),
            Some(MediaSize {
                width: 640.0,
                height: 360.0
            })
        );
        let graph = movie_graph(&editor, 0)?;
        assert_movie_graph_is_local(&graph);
        let original_object_id = graph.movies[0].object_id;
        let package = focused_package(&editor)?;
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Content)?,
            MOVIE
        );
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Poster)?,
            POSTER
        );
        let builds = editor.slide_builds(0)?;
        assert_eq!(builds.len(), 1);
        assert_eq!(builds[0].drawable_object_id, original_object_id);
        assert_eq!(builds[0].settings, KeynoteBuildSettings::movie_start());
        assert_eq!(builds[0].chunks.len(), 1);

        let roundtripped = KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        assert_eq!(
            movie_snapshot(&focused_package(&roundtripped)?, 0)?,
            created
        );

        let package = focused_package(&editor)?;
        let initial_playback = package
            .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
            .ok_or("playback is missing")?;
        let changed_playback = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::BackAndForth),
            volume: Some(MediaVolume::new(0.75)?),
            ..initial_playback
        };
        let commit = package
            .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
            .set(changed_playback)?
            .commit()?;
        reopen_focused_package(&mut editor, commit.package())?;
        let package = focused_package(&editor)?;
        assert_eq!(
            package
                .slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?,
            Some(changed_playback)
        );
        let commit = package
            .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
            .set(initial_playback)?
            .commit()?;
        reopen_focused_package(&mut editor, commit.package())?;

        let changed_properties = properties("Accessible Keynote movie");
        set_movie_properties(
            &mut editor,
            MovieSelector::index(0),
            changed_properties.clone(),
        )?;
        assert_eq!(
            movie_snapshot(&focused_package(&editor)?, 0)?.properties,
            changed_properties
        );
        let cleared_properties = MediaProperties::default();
        set_movie_properties(
            &mut editor,
            MovieSelector::index(0),
            cleared_properties.clone(),
        )?;
        assert_eq!(
            movie_snapshot(&focused_package(&editor)?, 0)?.properties,
            cleared_properties
        );

        let source_info = movie_snapshot(&focused_package(&editor)?, 0)?.info;
        replace_movie_media(
            &mut editor,
            MovieSelector::index(0),
            MediaPart::Content,
            REPLACEMENT_MOVIE,
        )?;
        replace_movie_media(
            &mut editor,
            MovieSelector::index(0),
            MediaPart::Poster,
            REPLACEMENT_POSTER,
        )?;
        let package = focused_package(&editor)?;
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Content)?,
            REPLACEMENT_MOVIE
        );
        assert_eq!(
            media_bytes(&package, MovieSelector::index(0), MediaPart::Poster)?,
            REPLACEMENT_POSTER
        );

        let duplicate_properties = properties("Duplicated Keynote movie");
        set_movie_properties(
            &mut editor,
            MovieSelector::index(0),
            duplicate_properties.clone(),
        )?;
        let duplicate = duplicate_movie(&mut editor, MovieSelector::index(0))?;
        let expected_duplicate = MovieInfo::from_parts(
            source_info.kind(),
            source_info.position().map(|point| MediaPoint {
                x: point.x + DUPLICATE_OFFSET,
                y: point.y + DUPLICATE_OFFSET,
            }),
            source_info.size(),
            source_info.natural_size(),
            source_info.playback(),
        )
        .with_original_size(source_info.original_size())
        .with_transform(source_info.transform());
        assert_eq!(duplicate.info, expected_duplicate);
        assert_eq!(duplicate.properties, duplicate_properties);
        assert_eq!(
            duplicate.info.position(),
            Some(MediaPoint { x: 110.0, y: 130.0 })
        );
        let graph = movie_graph(&editor, 0)?;
        assert_eq!(graph.movies.len(), 2);
        assert_ne!(graph.movies[0].object_id, graph.movies[1].object_id);
        assert_eq!(graph.movies[0].movie_data_id, graph.movies[1].movie_data_id);
        assert_eq!(
            graph.movies[0].poster_data_id,
            graph.movies[1].poster_data_id
        );
        assert_movie_graph_is_local(&graph);

        remove_movie(&mut editor, MovieSelector::index(0))?;
        assert_eq!(movie_graph(&editor, 0)?.movies.len(), 1);
        assert_eq!(editor.media_assets()?.len(), 2);
        remove_movie(&mut editor, MovieSelector::index(0))?;
        assert!(movie_graph(&editor, 0)?.movies.is_empty());
        assert!(editor.media_assets()?.is_empty());
        assert!(editor.slide_builds(0)?.is_empty());
        let package = focused_package(&editor)?;
        assert!(
            package
                .slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(0),
                    MediaPart::Content
                )
                .is_err()
        );
        KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        Ok(())
    }

    #[test]
    fn scratch_presentation_supports_native_movie_title_caption_crud() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Movie labels")
            .build()?;
        add_movie(
            &mut editor,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;

        assert_eq!(movie_title(&editor, MovieSelector::index(0))?, None);
        assert_eq!(movie_caption(&editor, MovieSelector::index(0))?, None);
        set_movie_title(&mut editor, MovieSelector::index(0), "Quarterly highlight")?;
        set_movie_caption(&mut editor, MovieSelector::index(0), "Revenue overview")?;
        assert_eq!(
            movie_title(&editor, MovieSelector::index(0))?,
            Some("Quarterly highlight".to_owned())
        );
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(0))?,
            Some("Revenue overview".to_owned())
        );

        set_movie_caption(
            &mut editor,
            MovieSelector::index(0),
            "Updated revenue overview",
        )?;
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(0))?,
            Some("Updated revenue overview".to_owned())
        );

        let duplicate = duplicate_movie(&mut editor, MovieSelector::index(0))?;
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(1))?,
            Some("Updated revenue overview".to_owned())
        );
        assert_eq!(duplicate.info.kind(), MovieKind::File);

        set_movie_title(&mut editor, MovieSelector::index(0), "Updated highlight")?;
        assert!(remove_movie_caption(&mut editor, MovieSelector::index(0))?);
        assert!(!remove_movie_caption(&mut editor, MovieSelector::index(0))?);
        assert!(remove_movie_title(&mut editor, MovieSelector::index(0))?);
        assert_eq!(movie_title(&editor, MovieSelector::index(0))?, None);
        assert_eq!(movie_caption(&editor, MovieSelector::index(0))?, None);

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes()?)?;
        assert_eq!(
            movie_caption(&reopened, MovieSelector::index(1))?,
            Some("Updated revenue overview".to_owned())
        );
        editor = reopened;
        remove_movie(&mut editor, MovieSelector::index(1))?;
        assert_eq!(movie_graph(&editor, 0)?.movies.len(), 1);
        Ok(())
    }

    #[test]
    fn movie_caption_package_preserves_movie_selector_order_after_audio() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Movie caption selector order")
            .build()?;
        add_audio(
            &mut editor,
            "audio.aiff",
            AUDIO,
            SlideAudioOptions::new(POSITION, Duration::from_millis(1_375))?,
        )?;
        add_movie(
            &mut editor,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;

        let package = focused_package(&editor)?;
        let movies = package
            .slides()?
            .first()
            .ok_or("focused slide is missing")?
            .movies();
        assert_eq!(movies.len(), 2);
        assert_eq!(movies[0].kind(), MovieKind::Audio);
        assert_eq!(movies[1].kind(), MovieKind::File);
        assert_eq!(movie_graph(&editor, 0)?.movies.len(), 2);

        set_movie_caption(&mut editor, MovieSelector::index(1), "Caption after audio")?;
        set_movie_title(&mut editor, MovieSelector::index(1), "Title after audio")?;
        assert_eq!(
            movie_title(&editor, MovieSelector::index(1))?,
            Some("Title after audio".to_owned())
        );
        assert_eq!(
            movie_caption(&editor, MovieSelector::index(1))?,
            Some("Caption after audio".to_owned())
        );
        assert!(remove_movie_caption(&mut editor, MovieSelector::index(1))?);
        assert_eq!(movie_caption(&editor, MovieSelector::index(1))?, None);
        Ok(())
    }

    #[test]
    fn invalid_movie_creation_is_transactional() -> TestResult {
        let mut editor = KeynoteDocumentBuilder::new().build()?;
        let baseline = editor.to_bytes()?;

        for result in [
            add_movie(
                &mut editor,
                "payload.bin",
                b"not video",
                "poster.png",
                POSTER,
                options(),
            ),
            add_movie(
                &mut editor,
                "movie.mov",
                MOVIE,
                "payload.bin",
                b"not image",
                options(),
            ),
            add_movie_at(
                &mut editor,
                1,
                "movie.mov",
                MOVIE,
                "poster.png",
                POSTER,
                options(),
            ),
        ] {
            assert!(result.is_err());
            assert_eq!(editor.to_bytes()?, baseline);
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
        assert_eq!(editor.to_bytes()?, baseline);

        add_movie(
            &mut editor,
            "movie.mov",
            MOVIE,
            "poster.png",
            POSTER,
            options(),
        )?;
        Ok(())
    }
}
