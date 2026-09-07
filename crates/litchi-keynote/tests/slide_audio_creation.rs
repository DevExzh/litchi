//! Fresh audio creation through semantic selectors and exact-source patches.

use std::{collections::BTreeMap, io, time::Duration};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::shape::geometry::Point;
use litchi_keynote::{
    Limits, MediaPart, MovieKind, MovieSelector, Package, ReadOptions, SemanticLimits,
    SlideAudioCreationError, SlideAudioCreationLimitKind, SlideSelector, slide::audio::Options,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");

fn audio() -> Vec<u8> {
    // A complete 100 ms, mono, signed 16-bit PCM WAV. This payload is valid
    // independently of the signature-only media admission check.
    let mut data = Vec::new();
    data.extend_from_slice(b"RIFF");
    data.extend_from_slice(&1_636u32.to_le_bytes());
    data.extend_from_slice(b"WAVEfmt ");
    data.extend_from_slice(&16u32.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&8_000u32.to_le_bytes());
    data.extend_from_slice(&16_000u32.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    data.extend_from_slice(&16u16.to_le_bytes());
    data.extend_from_slice(b"data");
    data.extend_from_slice(&1_600u32.to_le_bytes());
    for sample in 0..800i16 {
        data.extend_from_slice(&((sample % 50 - 25) * 400).to_le_bytes());
    }
    data
}

fn options() -> TestResult<Options> {
    Ok(Options::new(
        Point {
            x: 120.5,
            y: 240.25,
        },
        Duration::from_millis(100),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn members(source: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

#[test]
fn creates_audio_on_native_source_and_restores_exact_source() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let before = source.slides()?[0].movies().to_vec();
    let requested = options()?;
    let data = audio();
    let commit =
        source.add_slide_audio(SlideSelector::index(0), "fresh-pcm.wav", &data, requested)?;
    let patch = commit.patch();
    assert!(commit.diagnostics().changed());
    assert_eq!(patch.source_media_count(), before.len());
    assert_eq!(patch.target_media_count(), before.len() + 1);
    assert_eq!(patch.movie_position().get(), before.len());
    assert_eq!(patch.created_objects(), 5);
    assert_eq!(patch.created_data(), 1);
    assert_eq!(patch.options(), requested);
    let after = commit.package().slides()?[0].movies();
    assert_eq!(&after[..before.len()], before);
    let added = after
        .last()
        .ok_or_else(|| io::Error::other("missing created audio"))?;
    assert_eq!(added.kind(), MovieKind::Audio);
    assert_eq!(
        added.position().map(|point| (point.x, point.y)),
        Some((120.5, 240.25))
    );
    assert_eq!(added.duration(), Some(requested.duration()));
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::position(patch.movie_position()),
            MediaPart::Content,
        )?,
        data,
    );
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    let inverse = patch.inverse();
    assert_eq!(inverse.removed_objects(), 5);
    assert_eq!(inverse.removed_data(), 1);
    let restored = commit.package().apply_slide_audio_creation(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE_SOURCE);
    assert_eq!(
        restored.diagnostics().restored_previews(),
        patch.deleted_previews()
    );
    let replay = source.apply_slide_audio_creation(patch)?;
    assert_eq!(
        exact_bytes(replay.package())?,
        exact_bytes(commit.package())?
    );
    let twice = source.apply_slide_audio_creation(&inverse.inverse())?;
    assert_eq!(
        exact_bytes(twice.package())?,
        exact_bytes(commit.package())?
    );
    assert_eq!(twice.patch().deleted_previews(), patch.deleted_previews());
    Ok(())
}

#[test]
fn repeated_audio_content_reuses_one_data_record() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let data = audio();
    let first =
        source.add_slide_audio(SlideSelector::index(0), "fresh-pcm.wav", &data, options()?)?;
    let second = first.package().add_slide_audio(
        SlideSelector::index(0),
        "same-content.wav",
        &data,
        Options::new(Point { x: 321.0, y: 42.0 }, Duration::from_millis(100))?,
    )?;
    assert_eq!(first.patch().created_data(), 1);
    assert_eq!(second.patch().created_data(), 0);
    assert_eq!(second.patch().created_objects(), 5);
    let first_data = members(&exact_bytes(first.package())?)?
        .into_iter()
        .filter(|(name, _)| name.starts_with("Data/"))
        .collect::<BTreeMap<_, _>>();
    let second_data = members(&exact_bytes(second.package())?)?
        .into_iter()
        .filter(|(name, _)| name.starts_with("Data/"))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(first_data, second_data);
    let restored = second
        .package()
        .apply_slide_audio_creation(&second.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(first.package())?
    );
    Ok(())
}

#[test]
fn rejects_unsafe_names_or_non_audio_without_mutating_source() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let data = audio();
    for name in [
        "",
        "../audio.wav",
        "folder/audio.wav",
        "audio\\track.wav",
        "audio.png",
        "audio\0.wav",
    ] {
        assert!(matches!(
            source.add_slide_audio(SlideSelector::index(0), name, &data, options()?),
            Err(SlideAudioCreationError::InvalidFilename),
        ));
    }
    for invalid in [b"".as_slice(), b"not an audio payload".as_slice()] {
        assert!(matches!(
            source.add_slide_audio(SlideSelector::index(0), "audio.wav", invalid, options()?),
            Err(SlideAudioCreationError::UnsupportedAudio),
        ));
    }
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn patch_requires_the_exact_source_snapshot() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let first =
        source.add_slide_audio(SlideSelector::index(0), "fresh.wav", &audio(), options()?)?;
    assert!(matches!(
        first.package().apply_slide_audio_creation(first.patch()),
        Err(SlideAudioCreationError::PatchConflict),
    ));
    assert!(matches!(
        source.apply_slide_audio_creation(&first.patch().inverse()),
        Err(SlideAudioCreationError::PatchConflict),
    ));
    Ok(())
}

#[test]
fn source_selector_is_semantic_and_out_of_range_is_atomic() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    assert!(matches!(
        source.add_slide_audio(
            SlideSelector::index(usize::MAX),
            "audio.wav",
            &audio(),
            options()?
        ),
        Err(SlideAudioCreationError::SlidePositionNotFound { .. }),
    ));
    assert!(matches!(
        source.add_slide_audio(SlideSelector::name(""), "audio.wav", &audio(), options()?),
        Err(SlideAudioCreationError::EmptySlideName),
    ));
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn oversized_audio_reports_entry_limit_without_changing_source() -> TestResult {
    let largest = Catalog::from_bytes(NATIVE_SOURCE)?
        .iter()
        .map(|entry| entry.data().len())
        .max()
        .ok_or_else(|| io::Error::other("empty native fixture"))?;
    let defaults = Limits::default();
    let limits = Limits::new(
        defaults.max_input_bytes(),
        defaults.max_entries(),
        largest as u64,
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let source = Package::from_bytes_with_options(
        NATIVE_SOURCE,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let mut data = audio();
    data.resize(largest + 1, 0);
    match source.add_slide_audio(SlideSelector::index(0), "oversized.wav", &data, options()?) {
        Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::EntryBytes,
            observed,
            maximum,
        }) => {
            assert_eq!(observed, data.len() as u64);
            assert_eq!(maximum, largest as u64);
        },
        result => panic!("expected entry byte limit, got {result:?}"),
    }
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn created_audio_composes_with_existing_semantic_editors() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let creation = source.add_slide_audio(
        SlideSelector::index(0),
        "editable.wav",
        &audio(),
        options()?,
    )?;
    let selector = MovieSelector::position(creation.patch().movie_position());
    let moved = creation
        .package()
        .edit_slide_audio_position(SlideSelector::index(0), selector)?
        .set(litchi_keynote::slide::media::Point { x: 250.0, y: 125.0 })?
        .commit()?;
    let properties = moved
        .package()
        .slide_media_properties(SlideSelector::index(0), selector)?
        .with_accessibility_description(Some("Created narration".to_owned()));
    let described = moved
        .package()
        .edit_slide_media_properties(SlideSelector::index(0), selector)?
        .set(properties.clone())?
        .commit()?;
    assert_eq!(
        described
            .package()
            .slide_media_properties(SlideSelector::index(0), selector)?,
        properties,
    );
    let removed = described
        .package()
        .remove_slide_media(SlideSelector::index(0), selector)?;
    assert_eq!(
        removed.package().slides()?[0].movies(),
        source.slides()?[0].movies(),
    );
    assert_eq!(exact_bytes(&source)?, NATIVE_SOURCE);
    Ok(())
}
