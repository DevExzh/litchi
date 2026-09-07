#![no_main]

//! Bounded selector-first fuzzing for fresh Keynote slide audio.
//!
//! The target exercises [`Package::add_slide_audio`] against arbitrary
//! bounded package bytes and a mandatory checked-in native package seed. A
//! successful transaction must preserve the source, reopen and validate its
//! candidate, replay exactly, restore through its inverse, and produce the
//! same candidate through the double inverse. Failed admission must leave the
//! exact source untouched.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock, time::Duration};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_common::shape::geometry::Point;
use litchi_keynote::{
    Limits, MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::audio::Options,
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const MAX_PACKAGE_BYTES: u64 = 8 * 1024 * 1024;
const MAX_ENTRIES: usize = 256;
const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SLIDES: usize = 512;
const MAX_REFERENCES: usize = 32 * 1024;
const MAX_TEXT_STORAGES: usize = 8 * 1024;
const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;

const TARGET_INPUT: &[u8] = b"target-fresh-audio";
const NATIVE_KEYNOTE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/media-comments-baseline-native.key");

fuzz_target!(|data: &[u8]| {
    exercise_arbitrary_input(data);
    exercise_seed(data);
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_PACKAGE_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid audio-creation archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid audio-creation semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| {
                panic!("native Keynote audio-creation seed must open: {error}")
            });
        verify_seed_success_contract(&package);
        package
    })
}

fn exercise_arbitrary_input(data: &[u8]) {
    if data.len() > usize::try_from(MAX_INPUT_BYTES).unwrap_or(usize::MAX) {
        return;
    }

    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }
}

fn exercise_seed(data: &[u8]) {
    exercise_package(native_package(), data);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let audio = requested_audio(data);
    let filename = requested_filename(data);
    let options = requested_options(data);
    let slide = requested_slide(data);
    let commit = match package.add_slide_audio(slide, &filename, &audio, options) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };

    let patch = commit.patch();
    assert!(
        !patch.is_noop(),
        "fresh audio creation must change the source"
    );
    assert_eq!(patch.options(), options);
    assert_eq!(patch.created_objects(), 5);
    assert!(patch.created_data() <= 1);
    let candidate_bytes = package_bytes(commit.package());
    assert_ne!(candidate_bytes, source_bytes);
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("created audio package must validate: {error}"));
    assert_eq!(
        commit
            .package()
            .slide_media_data(
                SlideSelector::position(patch.slide_position()),
                MovieSelector::position(patch.movie_position()),
                MediaPart::Content,
            )
            .unwrap_or_else(|error| panic!("created audio data must read back: {error}")),
        audio.as_slice(),
    );
    assert_eq!(package_bytes(package), source_bytes);

    let replay = package
        .apply_slide_audio_creation(patch)
        .unwrap_or_else(|error| panic!("audio-creation forward replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    replay
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("replayed audio package must validate: {error}"));

    let inverse = patch.inverse();
    let restored = commit
        .package()
        .apply_slide_audio_creation(&inverse)
        .unwrap_or_else(|error| panic!("audio-creation inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    restored
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("inverse-restored audio package must validate: {error}"));

    let double_inverse = inverse.inverse();
    let twice = package
        .apply_slide_audio_creation(&double_inverse)
        .unwrap_or_else(|error| panic!("audio-creation double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    assert_eq!(package_bytes(package), source_bytes);

    assert!(
        commit.package().apply_slide_audio_creation(patch).is_err(),
        "a creation patch must reject its post-state as a source"
    );
    assert!(
        package.apply_slide_audio_creation(&inverse).is_err(),
        "an inverse creation patch must reject its original source"
    );
    assert_eq!(package_bytes(package), source_bytes);
}

fn verify_seed_success_contract(package: &Package) {
    let source_bytes = package_bytes(package);
    let data = seed_audio();
    let options = seed_options();
    let commit = package
        .add_slide_audio(SlideSelector::index(0), "fuzz-seed.wav", &data, options)
        .unwrap_or_else(|error| panic!("native audio-creation seed must commit: {error}"));
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().options(), options);
    assert_eq!(commit.patch().created_objects(), 5);
    assert_ne!(package_bytes(commit.package()), source_bytes);
    commit
        .package()
        .validate()
        .unwrap_or_else(|error| panic!("native audio-creation candidate must validate: {error}"));
    assert_eq!(
        commit
            .package()
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::position(commit.patch().movie_position()),
                MediaPart::Content,
            )
            .unwrap_or_else(|error| panic!("native created audio data must read back: {error}")),
        data.as_slice(),
    );

    let candidate_bytes = package_bytes(commit.package());
    let replay = package
        .apply_slide_audio_creation(commit.patch())
        .unwrap_or_else(|error| panic!("native audio-creation replay must apply: {error}"));
    assert_eq!(package_bytes(replay.package()), candidate_bytes);
    let inverse = commit.patch().inverse();
    let restored = commit
        .package()
        .apply_slide_audio_creation(&inverse)
        .unwrap_or_else(|error| panic!("native audio-creation inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    let twice = package
        .apply_slide_audio_creation(&inverse.inverse())
        .unwrap_or_else(|error| panic!("native audio-creation double inverse must apply: {error}"));
    assert_eq!(package_bytes(twice.package()), candidate_bytes);
    assert_eq!(package_bytes(package), source_bytes);

    exercise_limit_budget();
}

fn exercise_limit_budget() {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| panic!("finite audio-creation semantic limit must build: {error}"));
    let package = Package::from_bytes_with_options(
        NATIVE_KEYNOTE,
        ReadOptions::new(Limits::default(), semantic),
    )
    .unwrap_or_else(|error| panic!("native audio-creation limit seed must open: {error}"));
    let source_bytes = package_bytes(&package);
    assert!(
        package
            .add_slide_audio(
                SlideSelector::index(0),
                "limit.wav",
                &seed_audio(),
                seed_options(),
            )
            .is_err(),
        "the tight reference budget must reject creation"
    );
    assert_eq!(package_bytes(&package), source_bytes);
}

fn requested_slide(data: &[u8]) -> SlideSelector<'static> {
    if data.starts_with(TARGET_INPUT) {
        return SlideSelector::index(0);
    }
    match control(data, 0) % 6 {
        0..=3 => SlideSelector::index(0),
        4 => SlideSelector::index(usize::MAX),
        _ => SlideSelector::name(""),
    }
}

fn requested_filename(data: &[u8]) -> String {
    if data.starts_with(TARGET_INPUT) {
        return "fuzz-seed.wav".to_owned();
    }
    match control(data, 1) % 8 {
        0 => String::new(),
        1 => "../audio.wav".to_owned(),
        2 => "audio.png".to_owned(),
        3 => "folder/audio.wav".to_owned(),
        4 => "audio\0.wav".to_owned(),
        _ => format!("fuzz-{:02x}.wav", control(data, 2)),
    }
}

fn requested_options(data: &[u8]) -> Options {
    if data.starts_with(TARGET_INPUT) {
        return seed_options();
    }
    Options::new(
        Point {
            x: finite_axis(data, 3, 120.5),
            y: finite_axis(data, 7, 240.25),
        },
        Duration::from_millis(1 + (u64::from(control(data, 11)) % 4_000)),
    )
    .unwrap_or_else(|error| panic!("bounded audio-creation options must be valid: {error}"))
}

fn seed_options() -> Options {
    Options::new(
        Point {
            x: 120.5,
            y: 240.25,
        },
        Duration::from_millis(500),
    )
    .unwrap_or_else(|error| panic!("fixed audio-creation options must be valid: {error}"))
}

fn requested_audio(data: &[u8]) -> Vec<u8> {
    if data.starts_with(TARGET_INPUT) {
        return seed_audio();
    }
    match control(data, 12) % 8 {
        0 => return Vec::new(),
        1 => return b"not an audio payload".to_vec(),
        2 => return b"RIFF".to_vec(),
        _ => {},
    }
    let sample_count = 64 + usize::from(control(data, 12) % 128);
    let samples = (0..sample_count)
        .map(|index| {
            let offset = 13 + index.saturating_mul(2);
            i16::from_le_bytes([control(data, offset), control(data, offset + 1)])
        })
        .collect::<Vec<_>>();
    encode_wav(&samples)
}

fn seed_audio() -> Vec<u8> {
    let samples = (0..800)
        .map(|sample| (sample as i16 % 50 - 25) * 400)
        .collect::<Vec<_>>();
    encode_wav(&samples)
}

fn encode_wav(samples: &[i16]) -> Vec<u8> {
    let data_bytes = samples
        .len()
        .checked_mul(2)
        .and_then(|length| u32::try_from(length).ok())
        .unwrap_or_else(|| unreachable!("bounded fuzz WAV must fit u32"));
    let riff_size = 36 + data_bytes;
    let mut output = Vec::with_capacity(44 + data_bytes as usize);
    output.extend_from_slice(b"RIFF");
    output.extend_from_slice(&riff_size.to_le_bytes());
    output.extend_from_slice(b"WAVEfmt ");
    output.extend_from_slice(&16_u32.to_le_bytes());
    output.extend_from_slice(&1_u16.to_le_bytes());
    output.extend_from_slice(&1_u16.to_le_bytes());
    output.extend_from_slice(&8_000_u32.to_le_bytes());
    output.extend_from_slice(&16_000_u32.to_le_bytes());
    output.extend_from_slice(&2_u16.to_le_bytes());
    output.extend_from_slice(&16_u16.to_le_bytes());
    output.extend_from_slice(b"data");
    output.extend_from_slice(&data_bytes.to_le_bytes());
    for sample in samples {
        output.extend_from_slice(&sample.to_le_bytes());
    }
    output
}

fn finite_axis(data: &[u8], offset: usize, fallback: f32) -> f32 {
    let bits = u32::from(control(data, offset))
        | (u32::from(control(data, offset.saturating_add(1))) << 8)
        | (u32::from(control(data, offset.saturating_add(2))) << 16)
        | (u32::from(control(data, offset.saturating_add(3))) << 24);
    let value = f32::from_bits(bits);
    if value.is_finite() {
        let value = value.clamp(-1_000_000.0, 1_000_000.0);
        if value == 0.0 { 0.0 } else { value }
    } else {
        fallback
    }
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote fuzz package must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
