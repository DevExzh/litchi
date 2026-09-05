#![cfg(feature = "internal-iwork-source")]

//! Native Pages integration coverage for file-backed audio playback.
//!
//! The fixtures were authored, saved, closed, and reopened by Pages 14.4.
//! The test locates the audio MovieArchive by its typed `audio_only` marker,
//! then exercises the focused Pages payload seam without depending on a native
//! drawable identifier or a particular component name.

use std::{io, path::PathBuf};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{
    WireLimits,
    media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume},
    wire::WireView,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::tsd;
use litchi_pages::{
    __decode_movie_playback_payload, __rewrite_movie_playback_payload, MoviePlaybackError, Package,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const MEDIA_MEMBER: &str = "Data/ringin-24.wav";
const PLAYBACK_FIELDS: [u32; 6] = [3, 4, 5, 6, 7, 24];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/audio-playback-native.pages")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/audio-playback-native-resaved.pages")
}

#[derive(Debug, Clone)]
struct AudioPayload {
    member: String,
    object_id: u64,
    message_index: usize,
    bytes: Vec<u8>,
}

fn audio_payload(source: &[u8]) -> TestResult<AudioPayload> {
    let catalog = Catalog::from_bytes(source)?;
    let mut found = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in archive.objects {
            let object_id = object.archive_info.identifier;
            for (message_index, message) in object.messages.into_iter().enumerate() {
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
                    object_id: object_id
                        .ok_or_else(|| io::Error::other("native audio object has no identifier"))?,
                    message_index,
                    bytes: message.data,
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native ordinary audio MovieArchive is missing").into())
}

fn rewrite_audio_payload(
    source: &[u8],
    location: &AudioPayload,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == location.member)
        .ok_or_else(|| io::Error::other("native audio component is missing"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .object_mut(location.object_id)
        .ok_or_else(|| io::Error::other("native audio MovieArchive object is missing"))?;
    let message_type = object
        .messages
        .get(location.message_index)
        .map(|message| message.type_)
        .ok_or_else(|| io::Error::other("native audio MovieArchive message is missing"))?;
    if message_type != MOVIE_MESSAGE_TYPE {
        return Err(io::Error::other("native audio message type changed").into());
    }
    object.replace_message_preserving_header(
        location.message_index,
        RawMessage {
            type_: MOVIE_MESSAGE_TYPE,
            data: replacement.to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&location.member, &compressed)],
        Limits::default(),
    )?)
}

fn marker(package: &Package) -> TestResult<String> {
    let text = package.text()?;
    if !text.contains("Native Pages playback marker") {
        return Err(io::Error::other("native Pages playback marker is missing").into());
    }
    Ok(text)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_member_locality(source: &[u8], target: &[u8], changed_member: &str) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
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
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
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
    let root = WireView::parse(source)?;
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
        WireView::parse(super_field.payload())?
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

    let location = audio_payload(&source)?;
    let baseline = __decode_movie_playback_payload(&location.bytes, WireLimits::default())?;
    assert_native_source_settings(baseline);
    assert_eq!(
        __rewrite_movie_playback_payload(&location.bytes, baseline, WireLimits::default())?,
        location.bytes
    );
    Ok(())
}

#[test]
fn native_audio_source_rewrite_preserves_wire_graph_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let source_marker = marker(&Package::from_bytes(&source)?)?;
    let location = audio_payload(&source)?;
    let baseline = __decode_movie_playback_payload(&location.bytes, WireLimits::default())?;
    let expected = replacement(baseline)?;
    let changed_payload =
        __rewrite_movie_playback_payload(&location.bytes, expected, WireLimits::default())?;

    assert_ne!(changed_payload, location.bytes);
    assert_eq!(
        __decode_movie_playback_payload(&changed_payload, WireLimits::default())?,
        expected.canonicalize()?
    );
    assert_eq!(
        unknown_playback_wire(&changed_payload)?,
        unknown_playback_wire(&location.bytes)?
    );

    let changed_package = rewrite_audio_payload(&source, &location, &changed_payload)?;
    let changed_location = audio_payload(&changed_package)?;
    assert_eq!(changed_location.bytes, changed_payload);
    assert_eq!(
        marker(&Package::from_bytes(&changed_package)?)?,
        source_marker
    );
    assert_media_asset_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored_payload =
        __rewrite_movie_playback_payload(&changed_payload, baseline, WireLimits::default())?;
    assert_eq!(restored_payload, location.bytes);
    let restored_package =
        rewrite_audio_payload(&changed_package, &changed_location, &restored_payload)?;
    assert_eq!(audio_payload(&restored_package)?.bytes, location.bytes);
    assert_eq!(
        marker(&Package::from_bytes(&restored_package)?)?,
        source_marker
    );
    assert_media_asset_untouched(&source, &restored_package)?;
    Ok(())
}

#[test]
fn native_audio_resaved_fixture_reads_changed_settings_and_roundtrips() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package)?, source);
    let source_marker = marker(&package)?;

    let location = audio_payload(&source)?;
    let baseline = __decode_movie_playback_payload(&location.bytes, WireLimits::default())?;
    assert_native_resaved_settings(baseline);
    assert_eq!(
        __rewrite_movie_playback_payload(&location.bytes, baseline, WireLimits::default())?,
        location.bytes
    );

    let changed = baseline
        .with_loop_mode(Some(MediaLoopMode::BackAndForth))
        .with_volume(Some(MediaVolume::new(0.25)?));
    let changed_payload =
        __rewrite_movie_playback_payload(&location.bytes, changed, WireLimits::default())?;
    assert_eq!(
        __decode_movie_playback_payload(&changed_payload, WireLimits::default())?,
        changed.canonicalize()?
    );
    let changed_package = rewrite_audio_payload(&source, &location, &changed_payload)?;
    let reopened = Package::from_bytes(&changed_package)?;
    assert_eq!(marker(&reopened)?, source_marker);
    assert_media_asset_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored_payload =
        __rewrite_movie_playback_payload(&changed_payload, baseline, WireLimits::default())?;
    assert_eq!(restored_payload, location.bytes);
    let restored = rewrite_audio_payload(
        &changed_package,
        &audio_payload(&changed_package)?,
        &restored_payload,
    )?;
    assert_eq!(audio_payload(&restored)?.bytes, location.bytes);
    assert_media_asset_untouched(&source, &restored)?;
    Ok(())
}

#[test]
fn native_audio_playback_rejects_truncated_payload_without_rewrite() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let location = audio_payload(&source)?;
    let truncated = &location.bytes[..location.bytes.len().saturating_sub(1)];
    assert!(matches!(
        __decode_movie_playback_payload(truncated, WireLimits::default()),
        Err(MoviePlaybackError::Codec(_))
    ));
    assert!(matches!(
        __rewrite_movie_playback_payload(
            truncated,
            MediaPlaybackSettings::new(std::time::Duration::from_secs(1)),
            WireLimits::default()
        ),
        Err(MoviePlaybackError::Codec(_))
    ));
    Ok(())
}
