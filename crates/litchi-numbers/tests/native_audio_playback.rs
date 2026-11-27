//! Native Numbers integration coverage for file-backed audio playback.
//!
//! The fixtures were authored, saved, closed, and reopened by Numbers 14.4.
//! The test edits playback through the drawable-properties family, then locates
//! the audio MovieArchive by its typed `audio_only` marker to check the bytes,
//! without depending on a native drawable identifier or a particular component
//! name.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    WireLimits,
    media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume},
    wire::WireView,
};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{movie_playback_codec as codec, tsd};
use litchi_numbers::Package;
use litchi_numbers::cell::Value;
use litchi_numbers::sheet::drawable::DrawableSelector;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const MARKER_POSITION: &str = "B2";
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const MEDIA_MEMBER: &str = "Data/ringin-16.wav";
const PLAYBACK_FIELDS: [u32; 6] = [3, 4, 5, 6, 7, 24];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/audio-playback-native.numbers")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/audio-playback-native-resaved.numbers")
}

#[derive(Debug, Clone)]
struct AudioPayload {
    member: String,
    bytes: Vec<u8>,
}

fn audio_payload(source: &[u8]) -> TestResult<AudioPayload> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let mut found = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data(), Default::default()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes(), Default::default()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in archive.objects {
            for message in object.messages {
                if message.type_ != MOVIE_MESSAGE_TYPE {
                    continue;
                }
                let movie = match tsd::MovieArchive::decode(message.data.as_slice()) {
                    Ok(movie) => movie,
                    Err(_) => continue,
                };
                if movie.audio_only != Some(true) {
                    continue;
                }
                if found.is_some() {
                    return Err(io::Error::other(
                        "native fixture contains multiple ordinary audio MovieArchives",
                    )
                    .into());
                }
                found = Some(AudioPayload {
                    member: entry.name().to_owned(),
                    bytes: message.data,
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native ordinary audio MovieArchive is missing").into())
}

fn playback(package: &Package) -> TestResult<MediaPlaybackSettings> {
    package
        .sheet_drawable_properties(0, DrawableSelector::audio(0))?
        .playback()
        .ok_or_else(|| io::Error::other("native audio playback is missing").into())
}

/// Set playback through the public family and return the committed package.
fn set_playback(package: &Package, settings: MediaPlaybackSettings) -> TestResult<Package> {
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    edit.set_playback(settings)?;
    Ok(edit.commit()?.package().clone())
}

fn marker(package: &Package) -> TestResult<String> {
    let table = package
        .table(SHEET_NAME, TABLE_NAME)?
        .ok_or_else(|| io::Error::other("native marker table is missing"))?;
    match table.get_a1(MARKER_POSITION)? {
        Some(Value::Text(value)) => Ok(value.clone()),
        value => {
            Err(io::Error::other(format!("native audio marker is not text: {value:?}")).into())
        },
    }
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_member_locality(source: &[u8], target: &[u8], changed_member: &str) -> TestResult {
    let before = Catalog::from_bytes(source, Default::default())?;
    let after = Catalog::from_bytes(target, Default::default())?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native audio edit removed a package member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                entry.name()
            );
        }
    }
    changed.sort_unstable();
    assert_eq!(changed, [changed_member.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_media_asset_untouched(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source, Default::default())?;
    let after = Catalog::from_bytes(target, Default::default())?;
    let source_asset = before
        .iter()
        .find(|entry| entry.name() == MEDIA_MEMBER)
        .ok_or_else(|| io::Error::other("native audio asset is missing"))?;
    let target_asset = after
        .iter()
        .find(|entry| entry.name() == MEDIA_MEMBER)
        .ok_or_else(|| io::Error::other("candidate audio asset is missing"))?;
    assert_eq!(source_asset.data(), target_asset.data());
    assert_eq!(
        source_asset.raw_record().local_record(),
        target_asset.raw_record().local_record()
    );
    Ok(())
}

fn unknown_playback_wire(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let root = WireView::parse(source, WireLimits::DEFAULT)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("native audio playback envelope is missing"))?;
    let mut records = root
        .fields()
        .filter(|field| field.number() != 1 && !PLAYBACK_FIELDS.contains(&field.number()))
        .map(|field| field.raw().to_vec())
        .collect::<Vec<_>>();
    records.extend(
        WireView::parse(super_field.payload(), WireLimits::DEFAULT)?
            .fields()
            .map(|field| field.raw().to_vec()),
    );
    Ok(records)
}

fn replacement(settings: MediaPlaybackSettings) -> TestResult<MediaPlaybackSettings> {
    Ok(settings
        .with_loop_mode(Some(MediaLoopMode::Repeat))
        .with_volume(Some(MediaVolume::new(0.5)?)))
}

fn assert_native_source_settings(settings: MediaPlaybackSettings) {
    assert_eq!(settings.loop_mode, Some(MediaLoopMode::None));
    assert_eq!(settings.volume, Some(MediaVolume::FULL));
}

fn assert_native_resaved_settings(settings: MediaPlaybackSettings) {
    assert_eq!(settings.loop_mode, Some(MediaLoopMode::Repeat));
    assert_eq!(settings.volume, Some(MediaVolume::new(0.5).unwrap()));
}

#[test]
fn native_audio_source_reads_typed_settings_and_has_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let baseline = playback(&package)?;
    assert_native_source_settings(baseline);
    let unchanged = set_playback(&package, baseline)?;
    assert_eq!(exact_bytes(&unchanged)?, source);
    Ok(())
}

#[test]
fn native_audio_source_rewrite_preserves_wire_graph_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let source_marker = marker(&package)?;
    let location = audio_payload(&source)?;
    let baseline = playback(&package)?;
    let expected = replacement(baseline)?;

    let changed = set_playback(&package, expected)?;
    let changed_package = exact_bytes(&changed)?;
    let changed_location = audio_payload(&changed_package)?;
    assert_ne!(changed_location.bytes, location.bytes);
    assert_eq!(playback(&changed)?, expected.canonicalize()?);
    assert_eq!(
        unknown_playback_wire(&changed_location.bytes)?,
        unknown_playback_wire(&location.bytes)?
    );
    assert_eq!(
        marker(&Package::from_bytes(&changed_package)?)?,
        source_marker
    );
    assert_media_asset_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored = exact_bytes(&set_playback(&changed, baseline)?)?;
    assert_eq!(audio_payload(&restored)?.bytes, location.bytes);
    assert_eq!(marker(&Package::from_bytes(&restored)?)?, source_marker);
    assert_media_asset_untouched(&source, &restored)?;
    Ok(())
}

#[test]
fn native_audio_resaved_fixture_reads_changed_settings_and_roundtrips() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package)?, source);
    let source_marker = marker(&package)?;

    let location = audio_payload(&source)?;
    let baseline = playback(&package)?;
    assert_native_resaved_settings(baseline);
    assert_eq!(exact_bytes(&set_playback(&package, baseline)?)?, source);

    let changed_settings = baseline
        .with_loop_mode(Some(MediaLoopMode::BackAndForth))
        .with_volume(Some(MediaVolume::new(0.25)?));
    let changed = set_playback(&package, changed_settings)?;
    assert_eq!(playback(&changed)?, changed_settings.canonicalize()?);
    let changed_package = exact_bytes(&changed)?;
    let reopened = Package::from_bytes(&changed_package)?;
    assert_eq!(marker(&reopened)?, source_marker);
    assert_media_asset_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored = exact_bytes(&set_playback(&changed, baseline)?)?;
    assert_eq!(audio_payload(&restored)?.bytes, location.bytes);
    assert_media_asset_untouched(&source, &restored)?;
    Ok(())
}

#[test]
fn native_audio_playback_codec_rejects_truncated_payload_without_rewrite() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let location = audio_payload(&source)?;
    let truncated = &location.bytes[..location.bytes.len().saturating_sub(1)];
    let options = codec::DecodeOptions::for_source(truncated);
    assert!(codec::decode_movie_playback(truncated, options).is_err());
    Ok(())
}
