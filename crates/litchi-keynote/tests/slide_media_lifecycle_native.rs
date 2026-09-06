//! Admission tests for native Keynote slide-media lifecycle oracles.
//!
//! The three packages in this test were authored and reopened by Keynote after
//! duplicating a movie and removing one or both movies.  The focused media
//! transaction tests use synthetic packages because they need deterministic
//! mutation and inverse assertions.  These tests keep the native packages as
//! independent read-only evidence: the public API must expose the same
//! source-order media values, while the fixture-only checks pin the native
//! ownership, build, UUID, and ZIP data graph that those values came from.

use std::{collections::BTreeMap, io, time::Duration};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{kn, package_metadata_media_codec, tsa, tsd, tsp, tswp};
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-native.key");
const DUPLICATE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-duplicate-native.key");
const REMOVE_FINAL: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-remove-native.key");

const SLIDE_COMPONENT: u64 = 2_652_150;
const MOVIE_A: u64 = 2_653_286;
const MOVIE_B: u64 = 2_653_610;
const DUPLICATE_MOVIE: u64 = 2_653_696;
const DUPLICATE_BUILD: u64 = 2_653_697;
const BASELINE_DATA_METADATA_MAP: u64 = 2_653_651;
const AUDIO_A: u64 = 2_652_595;
const AUDIO_B: u64 = 2_652_622;
const AUDIO_DATA: u64 = 9_075;
const MOVIE_DATA: u64 = 9_085;
const POSTER_DATA: u64 = 9_086;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const BUILD_MESSAGE_TYPE: u32 = 8;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const AUDIO_MEMBER: &str = "Data/keynote-coral-9075.wav";
const MOVIE_MEMBER: &str = "Data/keynote-selfauthored-coral-mjpeg-9085.mov";
const POSTER_MEMBER: &str = "Data/posterImage-9086.png";
const CAPTION: &str = "Shared native movie caption";

#[derive(Debug, Clone, PartialEq)]
struct MovieGraph {
    movie_data: Option<u64>,
    poster_data: Option<u64>,
    title: Option<u64>,
    caption: Option<u64>,
    parent: Option<u64>,
    position: Option<(f32, f32)>,
    size: Option<(f32, f32)>,
    original_size: Option<(f32, f32)>,
    natural_size: Option<(f32, f32)>,
    start_time: Option<f32>,
    end_time: Option<f32>,
    poster_time: Option<f32>,
    loop_option: Option<i32>,
    volume: Option<f32>,
    audio_only: Option<bool>,
    plays_across_slides: Option<bool>,
}

fn member_bytes(source: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Catalog::from_bytes(source)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| io::Error::other(format!("missing native package member {name}")))
        .map_err(Into::into)
}

fn has_member(source: &[u8], name: &str) -> TestResult<bool> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .any(|entry| entry.name() == name))
}

fn component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        archives.push((entry.name().to_owned(), archive));
    }
    Ok(archives)
}

fn component_containing_object(source: &[u8], identifier: u64) -> TestResult<(String, Archive)> {
    component_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")))
        .map_err(Into::into)
}

fn object_message<'a>(archive: &'a Archive, identifier: u64, type_: u32) -> TestResult<&'a [u8]> {
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .map(|message| message.data.as_slice())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier} type {type_}")))
        .map_err(Into::into)
}

fn slide_archive(source: &[u8]) -> TestResult<kn::SlideArchive> {
    let (_, archive) = component_containing_object(source, SLIDE_COMPONENT)?;
    Ok(kn::SlideArchive::decode(object_message(
        &archive,
        SLIDE_COMPONENT,
        SLIDE_MESSAGE_TYPE,
    )?)?)
}

fn movie_archive(source: &[u8], identifier: u64) -> TestResult<tsd::MovieArchive> {
    let (_, archive) = component_containing_object(source, identifier)?;
    Ok(tsd::MovieArchive::decode(object_message(
        &archive,
        identifier,
        MOVIE_MESSAGE_TYPE,
    )?)?)
}

#[allow(
    deprecated,
    reason = "The native fixture oracle inspects both storage encodings retained by Keynote."
)]
fn caption_text(source: &[u8], movie_identifier: u64) -> TestResult<Option<String>> {
    let movie = movie_archive(source, movie_identifier)?;
    let Some(caption_identifier) = movie
        .super_
        .caption
        .as_ref()
        .map(|reference| reference.identifier)
    else {
        return Ok(None);
    };
    let (_, archive) = component_containing_object(source, caption_identifier)?;
    let caption = tsa::CaptionInfoArchive::decode(object_message(
        &archive,
        caption_identifier,
        CAPTION_INFO_MESSAGE_TYPE,
    )?)?;
    let storage_identifier = caption
        .super_
        .owned_storage
        .as_ref()
        .or(caption.super_.deprecated_storage.as_ref())
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("native caption has no text storage"))?;
    let storage = tsa_storage_text(source, storage_identifier)?;
    Ok(storage)
}

fn tsa_storage_text(source: &[u8], storage_identifier: u64) -> TestResult<Option<String>> {
    let (_, archive) = component_containing_object(source, storage_identifier)?;
    Ok(tswp::StorageArchive::decode(object_message(
        &archive,
        storage_identifier,
        STORAGE_MESSAGE_TYPE,
    )?)?
    .text
    .into_iter()
    .next())
}

fn movie_graph(source: &[u8], identifier: u64) -> TestResult<MovieGraph> {
    let movie = movie_archive(source, identifier)?;
    let geometry = movie.super_.geometry.as_ref();
    Ok(MovieGraph {
        movie_data: movie.movie_data.map(|reference| reference.identifier),
        poster_data: movie
            .poster_image_data
            .map(|reference| reference.identifier),
        title: movie.super_.title.map(|reference| reference.identifier),
        caption: movie.super_.caption.map(|reference| reference.identifier),
        parent: movie.super_.parent.map(|reference| reference.identifier),
        position: geometry
            .and_then(|geometry| geometry.position.as_ref())
            .map(|point| (point.x, point.y)),
        size: geometry
            .and_then(|geometry| geometry.size.as_ref())
            .map(|size| (size.width, size.height)),
        original_size: movie
            .original_size
            .as_ref()
            .map(|size| (size.width, size.height)),
        natural_size: movie
            .natural_size
            .as_ref()
            .map(|size| (size.width, size.height)),
        start_time: movie.start_time,
        end_time: movie.end_time,
        poster_time: movie.poster_time,
        loop_option: movie.loop_option,
        volume: movie.volume,
        audio_only: movie.audio_only,
        plays_across_slides: movie.plays_across_slides,
    })
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut matches = component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| {
            archive.objects.into_iter().flat_map(|object| {
                object.messages.into_iter().filter_map(|message| {
                    (message.type_ == METADATA_MESSAGE_TYPE).then_some(message.data)
                })
            })
        });
    let payload = matches
        .next()
        .ok_or_else(|| io::Error::other("missing native PackageMetadata"))?;
    if matches.next().is_some() {
        return Err(io::Error::other("native package has multiple PackageMetadata roots").into());
    }
    Ok(payload)
}

fn metadata(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let payload = metadata_payload(source)?;
    let metadata = tsp::PackageMetadata::decode(payload.as_slice())?;
    Ok(metadata)
}

fn metadata_root_field(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn data_metadata_map_payload(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let (_, archive) = component_containing_object(source, identifier)?;
    archive
        .object(identifier)
        .and_then(|object| object.messages.first())
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing native DataMetadataMap payload"))
        .map_err(Into::into)
}

fn metadata_from_payload(payload: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let metadata = tsp::PackageMetadata::decode(payload)?;
    Ok(metadata)
}

fn document_component<'a>(
    metadata: &'a tsp::PackageMetadata,
) -> TestResult<&'a tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| component.identifier == SLIDE_COMPONENT)
        .ok_or_else(|| io::Error::other("missing native current slide component"))
        .map_err(Into::into)
}

fn data_record(metadata: &tsp::PackageMetadata, identifier: u64) -> TestResult<&tsp::DataInfo> {
    metadata
        .datas
        .iter()
        .find(|data| data.identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing native DataInfo {identifier}")))
        .map_err(Into::into)
}

fn owner_list(metadata: &tsp::PackageMetadata, identifier: u64) -> TestResult<Vec<(u64, u32)>> {
    let component = document_component(metadata)?;
    let record = component
        .data_references
        .iter()
        .find(|record| record.data_identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing native DataInfo owners {identifier}")))?;
    let mut owners = record
        .object_reference_list
        .iter()
        .map(|owner| (owner.object_identifier, owner.count))
        .collect::<Vec<_>>();
    owners.sort_unstable();
    Ok(owners)
}

fn uuid_map(metadata: &tsp::PackageMetadata) -> BTreeMap<u64, (u64, u64)> {
    document_component(metadata)
        .map(|component| {
            component
                .object_uuid_map_entries
                .iter()
                .map(|entry| (entry.identifier, (entry.uuid.lower, entry.uuid.upper)))
                .collect()
        })
        .unwrap_or_default()
}

fn refs(references: &[tsp::Reference]) -> Vec<u64> {
    references
        .iter()
        .map(|reference| reference.identifier)
        .collect()
}

fn build_drawable(source: &[u8], identifier: u64) -> TestResult<Option<u64>> {
    let (_, archive) = component_containing_object(source, identifier)?;
    Ok(
        kn::BuildArchive::decode(object_message(&archive, identifier, BUILD_MESSAGE_TYPE)?)?
            .drawable
            .map(|reference| reference.identifier),
    )
}

fn assert_semantic_inventory(
    source: &[u8],
    expected_count: usize,
    expected_video_positions: &[(f32, f32)],
) -> TestResult {
    let package = Package::from_bytes(source)?;
    let slide = package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("native lifecycle fixture has no first slide"))?;
    assert_eq!(slide.movies().len(), expected_count);
    assert_eq!(
        slide.audio().count(),
        expected_count - expected_video_positions.len()
    );
    assert_eq!(slide.video_movies().count(), expected_video_positions.len());

    let videos = slide.video_movies().collect::<Vec<_>>();
    for (movie, &(x, y)) in videos.iter().zip(expected_video_positions) {
        assert_eq!(
            movie.position(),
            Some(litchi_keynote::slide::media::Point { x, y })
        );
        assert_eq!(
            movie.size(),
            Some(litchi_keynote::slide::media::Size {
                width: 320.0,
                height: 180.0,
            })
        );
        assert_eq!(movie.duration(), Some(Duration::from_secs(2)));
    }
    Ok(())
}

fn assert_media_data(
    package: &Package,
    movie: usize,
    part: MediaPart,
    expected: &[u8],
) -> TestResult {
    assert_eq!(
        package.slide_media_data(SlideSelector::index(0), MovieSelector::index(movie), part,)?,
        expected,
        "native lifecycle movie {movie} {part:?} bytes",
    );
    Ok(())
}

#[test]
fn native_lifecycle_oracles_expose_strict_source_order_and_shared_bytes() -> TestResult {
    let audio = member_bytes(BASELINE, AUDIO_MEMBER)?;
    let movie = member_bytes(BASELINE, MOVIE_MEMBER)?;
    let poster = member_bytes(BASELINE, POSTER_MEMBER)?;

    assert_semantic_inventory(BASELINE, 4, &[(200.0, 700.0), (700.0, 700.0)])?;
    assert_semantic_inventory(
        DUPLICATE,
        5,
        &[(200.0, 700.0), (700.0, 700.0), (210.0, 710.0)],
    )?;
    assert_semantic_inventory(REMOVE_FINAL, 2, &[])?;

    for (source, has_duplicate) in [(BASELINE, false), (DUPLICATE, true)] {
        let package = Package::from_bytes(source)?;
        assert_media_data(&package, 0, MediaPart::Content, &audio)?;
        assert_media_data(&package, 1, MediaPart::Content, &audio)?;
        assert_media_data(&package, 2, MediaPart::Content, &movie)?;
        assert_media_data(&package, 3, MediaPart::Content, &movie)?;
        assert_media_data(&package, 2, MediaPart::Poster, &poster)?;
        assert_media_data(&package, 3, MediaPart::Poster, &poster)?;
        if has_duplicate {
            assert_media_data(&package, 4, MediaPart::Content, &movie)?;
            assert_media_data(&package, 4, MediaPart::Poster, &poster)?;
        }
    }

    let removed = Package::from_bytes(REMOVE_FINAL)?;
    assert_media_data(&removed, 0, MediaPart::Content, &audio)?;
    assert_media_data(&removed, 1, MediaPart::Content, &audio)?;
    assert!(matches!(
        removed.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Poster,
        ),
        Err(litchi_keynote::SlideMediaDataError::AudioPoster)
    ));
    assert!(matches!(
        removed.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Content,
        ),
        Err(litchi_keynote::SlideMediaDataError::MoviePositionNotFound { .. })
    ));
    Ok(())
}

#[test]
fn native_duplicate_preserves_movie_graph_playback_titles_captions_and_metadata() -> TestResult {
    let baseline_metadata = metadata(BASELINE)?;
    let duplicate_metadata = metadata(DUPLICATE)?;
    assert_eq!(
        data_record(&baseline_metadata, AUDIO_DATA)?,
        data_record(&duplicate_metadata, AUDIO_DATA)?
    );
    assert_eq!(
        data_record(&baseline_metadata, MOVIE_DATA)?,
        data_record(&duplicate_metadata, MOVIE_DATA)?
    );
    assert_eq!(
        data_record(&baseline_metadata, POSTER_DATA)?,
        data_record(&duplicate_metadata, POSTER_DATA)?
    );
    assert_eq!(
        owner_list(&baseline_metadata, AUDIO_DATA)?,
        vec![(AUDIO_A, 1), (AUDIO_B, 1)]
    );
    assert_eq!(
        owner_list(&duplicate_metadata, AUDIO_DATA)?,
        vec![(AUDIO_A, 1), (AUDIO_B, 1)]
    );
    assert_eq!(
        owner_list(&baseline_metadata, MOVIE_DATA)?,
        vec![(MOVIE_A, 1), (MOVIE_B, 1)]
    );
    assert_eq!(
        owner_list(&duplicate_metadata, MOVIE_DATA)?,
        vec![(MOVIE_A, 1), (MOVIE_B, 1), (DUPLICATE_MOVIE, 1)]
    );
    assert_eq!(
        owner_list(&baseline_metadata, POSTER_DATA)?,
        vec![(MOVIE_A, 1), (MOVIE_B, 1)]
    );
    assert_eq!(
        owner_list(&duplicate_metadata, POSTER_DATA)?,
        vec![(MOVIE_A, 1), (MOVIE_B, 1), (DUPLICATE_MOVIE, 1)]
    );

    let baseline_slide = slide_archive(BASELINE)?;
    let duplicate_slide = slide_archive(DUPLICATE)?;
    assert_eq!(
        refs(&baseline_slide.owned_drawables),
        vec![
            2_652_163, 2_652_173, 2_652_180, AUDIO_A, AUDIO_B, MOVIE_A, MOVIE_B,
        ]
    );
    assert_eq!(
        refs(&duplicate_slide.owned_drawables),
        vec![
            2_652_163,
            2_652_173,
            2_652_180,
            AUDIO_A,
            AUDIO_B,
            MOVIE_A,
            MOVIE_B,
            DUPLICATE_MOVIE,
        ]
    );
    assert_eq!(
        refs(&duplicate_slide.drawables_z_order),
        refs(&duplicate_slide.owned_drawables)
    );
    assert_eq!(
        duplicate_slide.builds.len(),
        baseline_slide.builds.len() + 1
    );
    assert_eq!(
        build_drawable(DUPLICATE, DUPLICATE_BUILD)?,
        Some(DUPLICATE_MOVIE)
    );

    assert_eq!(
        movie_graph(BASELINE, MOVIE_A)?,
        movie_graph(DUPLICATE, MOVIE_A)?
    );
    assert_eq!(
        movie_graph(BASELINE, MOVIE_B)?,
        movie_graph(DUPLICATE, MOVIE_B)?
    );
    assert_eq!(caption_text(BASELINE, MOVIE_A)?.as_deref(), Some(CAPTION));
    assert_eq!(caption_text(BASELINE, MOVIE_B)?.as_deref(), Some(CAPTION));
    assert_eq!(caption_text(DUPLICATE, MOVIE_A)?.as_deref(), Some(CAPTION));
    assert_eq!(caption_text(DUPLICATE, MOVIE_B)?.as_deref(), Some(CAPTION));
    assert_eq!(
        caption_text(DUPLICATE, DUPLICATE_MOVIE)?.as_deref(),
        Some(CAPTION)
    );
    let duplicate_graph = movie_graph(DUPLICATE, DUPLICATE_MOVIE)?;
    assert_eq!(duplicate_graph.position, Some((210.0, 710.0)));
    assert_eq!(duplicate_graph.size, Some((320.0, 180.0)));
    assert_eq!(duplicate_graph.movie_data, Some(MOVIE_DATA));
    assert_eq!(duplicate_graph.poster_data, Some(POSTER_DATA));

    let baseline_uuids = uuid_map(&baseline_metadata);
    let duplicate_uuids = uuid_map(&duplicate_metadata);
    assert_eq!(baseline_uuids.get(&MOVIE_A), duplicate_uuids.get(&MOVIE_A));
    assert_eq!(baseline_uuids.get(&MOVIE_B), duplicate_uuids.get(&MOVIE_B));
    let duplicate_uuid = duplicate_uuids
        .get(&DUPLICATE_MOVIE)
        .copied()
        .ok_or_else(|| io::Error::other("native duplicate movie UUID is missing"))?;
    assert_ne!(duplicate_uuid, (0, 0));
    assert!(!baseline_uuids.values().any(|uuid| *uuid == duplicate_uuid));
    let duplicate_build_uuid = duplicate_uuids
        .get(&DUPLICATE_BUILD)
        .copied()
        .ok_or_else(|| io::Error::other("native duplicate build UUID is missing"))?;
    assert_ne!(duplicate_build_uuid, (0, 0));
    assert_ne!(duplicate_build_uuid, duplicate_uuid);

    assert_eq!(
        member_bytes(BASELINE, AUDIO_MEMBER)?,
        member_bytes(DUPLICATE, AUDIO_MEMBER)?
    );
    assert_eq!(
        member_bytes(BASELINE, MOVIE_MEMBER)?,
        member_bytes(DUPLICATE, MOVIE_MEMBER)?
    );
    assert_eq!(
        member_bytes(BASELINE, POSTER_MEMBER)?,
        member_bytes(DUPLICATE, POSTER_MEMBER)?
    );
    Ok(())
}

#[test]
fn native_remove_final_culls_movie_assets_but_retains_audio_graph_and_zip_locality() -> TestResult {
    let metadata = metadata(REMOVE_FINAL)?;
    assert_eq!(data_record(&metadata, AUDIO_DATA)?.identifier, AUDIO_DATA);
    assert!(data_record(&metadata, MOVIE_DATA).is_err());
    assert!(data_record(&metadata, POSTER_DATA).is_err());
    assert_eq!(
        owner_list(&metadata, AUDIO_DATA)?,
        vec![(AUDIO_A, 1), (AUDIO_B, 1)]
    );
    let component = document_component(&metadata)?;
    assert_eq!(
        component
            .data_references
            .iter()
            .filter(|record| matches!(
                record.data_identifier,
                AUDIO_DATA | MOVIE_DATA | POSTER_DATA
            ))
            .map(|record| record.data_identifier)
            .collect::<Vec<_>>(),
        vec![AUDIO_DATA]
    );
    assert!(has_member(REMOVE_FINAL, AUDIO_MEMBER)?);
    assert!(!has_member(REMOVE_FINAL, MOVIE_MEMBER)?);
    assert!(!has_member(REMOVE_FINAL, POSTER_MEMBER)?);

    let slide = slide_archive(REMOVE_FINAL)?;
    assert_eq!(
        refs(&slide.owned_drawables),
        vec![2_652_163, 2_652_173, 2_652_180, AUDIO_A, AUDIO_B]
    );
    assert_eq!(refs(&slide.drawables_z_order), refs(&slide.owned_drawables));
    assert_eq!(slide.builds.len(), 2);
    assert!(component_containing_object(REMOVE_FINAL, MOVIE_A).is_err());
    assert!(component_containing_object(REMOVE_FINAL, MOVIE_B).is_err());
    assert!(component_containing_object(REMOVE_FINAL, DUPLICATE_BUILD).is_err());

    let uuids = uuid_map(&metadata);
    assert!(!uuids.contains_key(&MOVIE_A));
    assert!(!uuids.contains_key(&MOVIE_B));
    assert!(uuids.contains_key(&AUDIO_A));
    assert!(uuids.contains_key(&AUDIO_B));
    Ok(())
}

#[test]
fn native_baseline_metadata_removal_uses_map_witness_and_matches_final_selected_state() -> TestResult
{
    // The metadata codec is exercised against the native baseline payload
    // directly.  This remains a codec admission test: no invalid document
    // graph is published, and the native remove-final package is the semantic
    // oracle for only the selected records.
    let source = metadata_payload(BASELINE)?;
    let source_metadata = metadata_from_payload(&source)?;
    let map_payload = data_metadata_map_payload(BASELINE, BASELINE_DATA_METADATA_MAP)?;
    let map_source = package_metadata_media_codec::DataMetadataMapSource::from_source(
        BASELINE_DATA_METADATA_MAP,
        &map_payload,
        package_metadata_media_codec::DecodeOptions::for_source(&map_payload),
    )?;
    assert!(!map_source.has_unknown_fields());
    assert!(map_source.entries() > 0);

    let component =
        package_metadata_media_codec::ComponentSelector::new(SLIDE_COMPONENT, "Slide-2652150");
    let data_removals = [
        package_metadata_media_codec::DataInfoRemoval::new(MOVIE_DATA),
        package_metadata_media_codec::DataInfoRemoval::new(POSTER_DATA),
    ];
    let owner_removals = [
        package_metadata_media_codec::DataReferenceOwnerRemoval::new(
            component, MOVIE_DATA, MOVIE_A, 1,
        ),
        package_metadata_media_codec::DataReferenceOwnerRemoval::new(
            component, MOVIE_DATA, MOVIE_B, 1,
        ),
        package_metadata_media_codec::DataReferenceOwnerRemoval::new(
            component,
            POSTER_DATA,
            MOVIE_A,
            1,
        ),
        package_metadata_media_codec::DataReferenceOwnerRemoval::new(
            component,
            POSTER_DATA,
            MOVIE_B,
            1,
        ),
    ];
    let batch = package_metadata_media_codec::MediaRewriteBatch::new(
        &[],
        &data_removals,
        &[],
        &owner_removals,
    )
    .with_data_metadata_map_source(map_source);
    let output = package_metadata_media_codec::rewrite_package_metadata_media(
        &source,
        batch,
        package_metadata_media_codec::DecodeOptions::for_source(&source),
    )?;
    assert_eq!(output.report().data_removals(), 2);
    assert_eq!(output.report().owner_removals(), 4);

    let candidate = metadata_from_payload(output.bytes())?;
    let native_final = metadata(REMOVE_FINAL)?;
    assert_eq!(
        data_record(&candidate, AUDIO_DATA)?,
        data_record(&native_final, AUDIO_DATA)?
    );
    assert!(data_record(&candidate, MOVIE_DATA).is_err());
    assert!(data_record(&candidate, POSTER_DATA).is_err());
    assert_eq!(
        owner_list(&candidate, AUDIO_DATA)?,
        owner_list(&native_final, AUDIO_DATA)?
    );
    assert!(owner_list(&candidate, MOVIE_DATA).is_err());
    assert!(owner_list(&candidate, POSTER_DATA).is_err());

    // Every untouched DataInfo remains source-authoritative.  This guards
    // against a codec that accidentally rebuilds the full repeated registry.
    for record in &source_metadata.datas {
        if matches!(record.identifier, MOVIE_DATA | POSTER_DATA) {
            continue;
        }
        assert_eq!(data_record(&candidate, record.identifier)?, record);
    }
    assert_eq!(candidate.datas.len(), source_metadata.datas.len() - 2);

    // DataMetadataMap is an external object.  Its root edge and exact map
    // payload are witnesses only and must stay untouched by DataInfo removal.
    assert_eq!(
        metadata_root_field(&source, 10)?,
        metadata_root_field(output.bytes(), 10)?
    );
    assert_eq!(
        map_payload,
        data_metadata_map_payload(BASELINE, BASELINE_DATA_METADATA_MAP)?
    );
    assert_eq!(
        source_metadata.data_metadata_map,
        candidate.data_metadata_map
    );
    Ok(())
}
