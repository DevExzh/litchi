//! Native graph oracle for fresh, slide-owned Keynote audio.
//!
//! The public transaction is selector-first and archive-free.  This test keeps
//! the native inspection on the fixture side of the integration boundary and
//! uses the generated Prost types only as a differential oracle for the exact
//! graph Keynote expects.

use std::{
    collections::{BTreeMap, BTreeSet},
    io,
    time::Duration,
};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{kn, tsd, tsp, tss};
use litchi_keynote::{Package, SlideSelector, slide::audio::Options};
use prost::Message as _;
use sha1::{Digest as _, Sha1};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");

const SLIDE_COMPONENT: u64 = 2_652_150;
const SLIDE_NODE: u64 = 2_652_149;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct ObjectKey {
    component: String,
    identifier: u64,
    message_type: u32,
    message_index: usize,
}

fn valid_wav() -> Vec<u8> {
    // A complete, mono, signed 16-bit PCM WAV.  Keeping a real RIFF payload
    // here exercises the same media admission path as a caller's file bytes.
    let sample_count = 800u32;
    let data_bytes = sample_count * 2;
    let mut data = Vec::with_capacity(44 + data_bytes as usize);
    data.extend_from_slice(b"RIFF");
    data.extend_from_slice(&(36 + data_bytes).to_le_bytes());
    data.extend_from_slice(b"WAVEfmt ");
    data.extend_from_slice(&16u32.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&8_000u32.to_le_bytes());
    data.extend_from_slice(&16_000u32.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    data.extend_from_slice(&16u16.to_le_bytes());
    data.extend_from_slice(b"data");
    data.extend_from_slice(&data_bytes.to_le_bytes());
    for sample in 0..sample_count {
        let sample = (sample as i16 % 50 - 25) * 400;
        data.extend_from_slice(&sample.to_le_bytes());
    }
    data
}

fn options() -> TestResult<Options> {
    Ok(Options::new(
        litchi_iwa_common::shape::geometry::Point {
            x: 320.5,
            y: 180.25,
        },
        Duration::from_millis(500),
    )?)
}

fn native_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
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

fn object_message(source: &[u8], identifier: u64, message_type: u32) -> TestResult<Vec<u8>> {
    native_archives(source)?
        .into_iter()
        .find_map(|(_, archive)| {
            archive.object(identifier).and_then(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == message_type)
                    .map(|message| message.data.clone())
            })
        })
        .ok_or_else(|| io::Error::other(format!("missing object {identifier} type {message_type}")))
        .map_err(Into::into)
}

fn component_containing_object(source: &[u8], identifier: u64) -> TestResult<String> {
    native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .map(|(component, _)| component)
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")))
        .map_err(Into::into)
}

fn object_payloads(source: &[u8]) -> TestResult<BTreeMap<ObjectKey, Vec<u8>>> {
    let mut payloads = BTreeMap::new();
    for (component, archive) in native_archives(source)? {
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            for (message_index, message) in object.messages.into_iter().enumerate() {
                let key = ObjectKey {
                    component: component.clone(),
                    identifier,
                    message_type: message.type_,
                    message_index,
                };
                assert!(
                    payloads.insert(key, message.data).is_none(),
                    "native fixture has duplicate object/message key"
                );
            }
        }
    }
    Ok(payloads)
}

fn object_ids(source: &[u8], component: &str) -> TestResult<BTreeSet<u64>> {
    Ok(native_archives(source)?
        .into_iter()
        .find(|(name, _)| name == component)
        .map(|(_, archive)| {
            archive
                .objects
                .into_iter()
                .filter_map(|object| object.archive_info.identifier)
                .collect()
        })
        .ok_or_else(|| io::Error::other(format!("missing component {component}")))?)
}

fn all_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for (_, archive) in native_archives(source)? {
        identifiers.extend(
            archive
                .objects
                .into_iter()
                .filter_map(|object| object.archive_info.identifier),
        );
    }
    Ok(identifiers)
}

fn native_metadata(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let payload = native_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| archive.objects)
        .flat_map(|object| object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data)
        .ok_or_else(|| io::Error::other("missing native PackageMetadata"))?;
    Ok(tsp::PackageMetadata::decode(payload.as_slice())?)
}

fn metadata_component<'a>(
    metadata: &'a tsp::PackageMetadata,
    identifier: u64,
) -> TestResult<&'a tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| component.identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing metadata component {identifier}")))
        .map_err(Into::into)
}

fn data_record<'a>(
    metadata: &'a tsp::PackageMetadata,
    identifier: u64,
) -> TestResult<&'a tsp::DataInfo> {
    metadata
        .datas
        .iter()
        .find(|data| data.identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing DataInfo {identifier}")))
        .map_err(Into::into)
}

fn data_owners(
    metadata: &tsp::PackageMetadata,
    component_identifier: u64,
    data_identifier: u64,
) -> TestResult<Vec<(u64, u32)>> {
    let component = metadata_component(metadata, component_identifier)?;
    let reference = component
        .data_references
        .iter()
        .find(|reference| reference.data_identifier == data_identifier)
        .ok_or_else(|| io::Error::other(format!("missing data owners {data_identifier}")))?;
    let mut owners = reference
        .object_reference_list
        .iter()
        .map(|owner| (owner.object_identifier, owner.count))
        .collect::<Vec<_>>();
    owners.sort_unstable();
    Ok(owners)
}

fn uuid_map(
    metadata: &tsp::PackageMetadata,
    component_identifier: u64,
) -> BTreeMap<u64, (u64, u64)> {
    metadata_component(metadata, component_identifier)
        .map(|component| {
            component
                .object_uuid_map_entries
                .iter()
                .map(|entry| (entry.identifier, (entry.uuid.lower, entry.uuid.upper)))
                .collect()
        })
        .unwrap_or_default()
}

fn reference_ids(references: &[tsp::Reference]) -> Vec<u64> {
    references
        .iter()
        .map(|reference| reference.identifier)
        .collect()
}

fn strip_fields(payload: &[u8], excluded: &[u32]) -> TestResult<Vec<u8>> {
    let mut output = Vec::with_capacity(payload.len());
    for field in WireView::parse(payload)?.fields() {
        if !excluded.contains(&field.number()) {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

fn catalog_member(source: &[u8], name: &str) -> TestResult<Option<Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec()))
}

fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// Replace one decoded IWA member while retaining the source ZIP's physical
/// ordering, names, extras, comments, and untouched member records.
fn replace_component(source: &[u8], name: &str, archive: &Archive) -> TestResult<Vec<u8>> {
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entry_name = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.name().to_owned())
        .ok_or_else(|| io::Error::other(format!("missing native package member {name}")))?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&entry_name, &replacement)],
        Limits::default(),
    )?)
}

fn mutate_package_metadata(
    source: &[u8],
    edit: impl FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(name, archive)| {
            name == METADATA_COMPONENT
                && archive.objects.iter().any(|object| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
                })
        })
        .ok_or_else(|| io::Error::other("native package metadata component is missing"))?;
    let mut selected = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected.replace((object_index, message_index)).is_some() {
                return Err(
                    io::Error::other("native package has multiple PackageMetadata roots").into(),
                );
            }
        }
    }
    let (object_index, message_index) =
        selected.ok_or_else(|| io::Error::other("native PackageMetadata payload is missing"))?;
    let message = archive
        .objects
        .get_mut(object_index)
        .and_then(|object| object.messages.get_mut(message_index))
        .ok_or_else(|| io::Error::other("native PackageMetadata payload is missing"))?;
    let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
    edit(&mut metadata)?;
    message.data = metadata.encode_to_vec();
    replace_component(source, &component_name, &archive)
}

fn fresh_ids(source: &[u8], candidate: &[u8], component: &str) -> TestResult<Vec<u64>> {
    let source_ids = object_ids(source, component)?;
    let candidate_ids = object_ids(candidate, component)?;
    Ok(candidate_ids.difference(&source_ids).copied().collect())
}

fn fresh_ids_with_message_type(
    source: &[u8],
    candidate: &[u8],
    component: &str,
    message_type: u32,
) -> TestResult<Vec<u64>> {
    let ids = fresh_ids(source, candidate, component)?;
    let payloads = object_payloads(candidate)?;
    Ok(ids
        .into_iter()
        .filter(|identifier| {
            payloads.keys().any(|key| {
                key.component == component
                    && key.identifier == *identifier
                    && key.message_type == message_type
            })
        })
        .collect())
}

fn effective_locator(component: &tsp::ComponentInfo) -> &str {
    component
        .locator
        .as_deref()
        .unwrap_or(&component.preferred_locator)
}

fn metadata_component_for_locator<'a>(
    metadata: &'a tsp::PackageMetadata,
    locator: &str,
) -> TestResult<&'a tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| effective_locator(component) == locator)
        .ok_or_else(|| io::Error::other(format!("missing metadata component {locator}")))
        .map_err(Into::into)
}

#[allow(deprecated)]
#[test]
fn fresh_audio_matches_native_graph_and_preserves_unrelated_source_payloads() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let data = valid_wav();
    let requested = options()?;
    let commit = source_package.add_slide_audio(
        SlideSelector::index(0),
        "fresh-pcm.wav",
        &data,
        requested,
    )?;
    let candidate_package = commit.package();
    let candidate = {
        let mut bytes = Vec::new();
        candidate_package.write_to(&mut bytes)?;
        bytes
    };

    let slide_component = component_containing_object(NATIVE_SOURCE, SLIDE_COMPONENT)?;
    let node_component = component_containing_object(NATIVE_SOURCE, SLIDE_NODE)?;
    let source_ids = object_ids(NATIVE_SOURCE, &slide_component)?;
    let candidate_ids = object_ids(&candidate, &slide_component)?;
    let source_payloads = object_payloads(NATIVE_SOURCE)?;
    let new_ids = candidate_ids
        .difference(&source_ids)
        .copied()
        .collect::<Vec<_>>();
    let candidate_payloads = object_payloads(&candidate)?;
    let source_object_keys = source_payloads
        .keys()
        .map(|key| (key.component.clone(), key.identifier))
        .collect::<BTreeSet<_>>();
    let candidate_object_keys = candidate_payloads
        .keys()
        .map(|key| (key.component.clone(), key.identifier))
        .collect::<BTreeSet<_>>();
    let all_new_object_keys = candidate_object_keys
        .difference(&source_object_keys)
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(
        all_new_object_keys.len(),
        5,
        "the candidate must not create objects outside the selected component"
    );
    assert!(
        all_new_object_keys
            .iter()
            .all(|(component, _)| component == &slide_component)
    );
    let source_max_object_identifier = *all_object_ids(NATIVE_SOURCE)?
        .iter()
        .max()
        .ok_or_else(|| io::Error::other("native source has no objects"))?;
    let expected_new_ids =
        (source_max_object_identifier + 1..=source_max_object_identifier + 5).collect::<Vec<_>>();
    assert_eq!(new_ids, expected_new_ids);
    assert_eq!(
        new_ids.len(),
        5,
        "fresh audio must add exactly five native objects"
    );
    assert_eq!(
        new_ids
            .iter()
            .filter_map(|identifier| {
                candidate_payloads.keys().find_map(|key| {
                    (key.identifier == *identifier && key.component == slide_component)
                        .then_some(key.message_type)
                })
            })
            .collect::<BTreeSet<_>>(),
        BTreeSet::from([
            MOVIE_MESSAGE_TYPE,
            STANDIN_MESSAGE_TYPE,
            BUILD_MESSAGE_TYPE,
            BUILD_CHUNK_MESSAGE_TYPE,
        ]),
        "new objects must be one movie, two stand-ins, one build, and one chunk"
    );

    let created_movies = new_ids
        .iter()
        .copied()
        .filter(|identifier| {
            candidate_payloads.keys().any(|key| {
                key.component == slide_component
                    && key.identifier == *identifier
                    && key.message_type == MOVIE_MESSAGE_TYPE
            })
        })
        .collect::<Vec<_>>();
    let created_standins = new_ids
        .iter()
        .copied()
        .filter(|identifier| {
            candidate_payloads.keys().any(|key| {
                key.component == slide_component
                    && key.identifier == *identifier
                    && key.message_type == STANDIN_MESSAGE_TYPE
            })
        })
        .collect::<Vec<_>>();
    let created_builds = new_ids
        .iter()
        .copied()
        .filter(|identifier| {
            candidate_payloads.keys().any(|key| {
                key.component == slide_component
                    && key.identifier == *identifier
                    && key.message_type == BUILD_MESSAGE_TYPE
            })
        })
        .collect::<Vec<_>>();
    let created_chunks = new_ids
        .iter()
        .copied()
        .filter(|identifier| {
            candidate_payloads.keys().any(|key| {
                key.component == slide_component
                    && key.identifier == *identifier
                    && key.message_type == BUILD_CHUNK_MESSAGE_TYPE
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(created_movies.len(), 1);
    assert_eq!(created_standins.len(), 2);
    assert_eq!(created_builds.len(), 1);
    assert_eq!(created_chunks.len(), 1);
    let movie_id = created_movies[0];
    let build_id = created_builds[0];
    let chunk_id = created_chunks[0];

    let source_slide = kn::SlideArchive::decode(
        object_message(NATIVE_SOURCE, SLIDE_COMPONENT, SLIDE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let candidate_slide = kn::SlideArchive::decode(
        object_message(&candidate, SLIDE_COMPONENT, SLIDE_MESSAGE_TYPE)?.as_slice(),
    )?;
    assert_eq!(
        reference_ids(&candidate_slide.owned_drawables),
        source_slide
            .owned_drawables
            .iter()
            .map(|reference| reference.identifier)
            .chain(std::iter::once(movie_id))
            .collect::<Vec<_>>()
    );
    assert_eq!(
        reference_ids(&candidate_slide.drawables_z_order),
        source_slide
            .drawables_z_order
            .iter()
            .map(|reference| reference.identifier)
            .chain(std::iter::once(movie_id))
            .collect::<Vec<_>>()
    );
    assert_eq!(
        reference_ids(&candidate_slide.builds),
        source_slide
            .builds
            .iter()
            .map(|reference| reference.identifier)
            .chain(std::iter::once(build_id))
            .collect::<Vec<_>>()
    );
    assert_eq!(
        reference_ids(&candidate_slide.build_chunks),
        source_slide
            .build_chunks
            .iter()
            .map(|reference| reference.identifier)
            .chain(std::iter::once(chunk_id))
            .collect::<Vec<_>>()
    );

    let movie = tsd::MovieArchive::decode(
        object_message(&candidate, movie_id, MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let movie_data_id = movie
        .movie_data
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("created movie has no content data"))?;
    assert_eq!(
        movie
            .super_
            .parent
            .as_ref()
            .map(|reference| reference.identifier),
        Some(SLIDE_COMPONENT)
    );
    assert_eq!(
        movie
            .super_
            .title
            .as_ref()
            .map(|reference| reference.identifier),
        Some(created_standins[0])
    );
    assert_eq!(
        movie
            .super_
            .caption
            .as_ref()
            .map(|reference| reference.identifier),
        Some(created_standins[1])
    );
    assert_eq!(movie.super_.locked, Some(false));
    assert_eq!(movie.super_.aspect_ratio_locked, Some(true));
    let geometry = movie
        .super_
        .geometry
        .as_ref()
        .ok_or_else(|| io::Error::other("created movie has no geometry"))?;
    let position = geometry
        .position
        .as_ref()
        .ok_or_else(|| io::Error::other("created movie has no position"))?;
    assert_eq!(
        (position.x, position.y),
        (requested.position().x, requested.position().y)
    );
    let size = geometry
        .size
        .as_ref()
        .ok_or_else(|| io::Error::other("created movie has no size"))?;
    assert_eq!((size.width, size.height), (0.0, 0.0));
    assert_eq!(geometry.flags, Some(3));
    assert_eq!(geometry.angle, Some(0.0));
    let document = kn::DocumentArchive::decode(object_message(NATIVE_SOURCE, 1, 1)?.as_slice())?;
    let show = kn::ShowArchive::decode(
        object_message(NATIVE_SOURCE, document.show.identifier, 2)?.as_slice(),
    )?;
    let stylesheet = tss::StylesheetArchive::decode(
        object_message(NATIVE_SOURCE, show.stylesheet.identifier, 401)?.as_slice(),
    )?;
    let source_style = stylesheet
        .styles
        .iter()
        .map(|reference| reference.identifier)
        .find(|identifier| object_message(NATIVE_SOURCE, *identifier, 3_016).is_ok())
        .ok_or_else(|| io::Error::other("native stylesheet has no media style"))?;
    assert_eq!(
        movie.style.as_ref().map(|reference| reference.identifier),
        Some(source_style)
    );
    assert_eq!(movie.start_time, Some(0.0));
    assert_eq!(movie.end_time, Some(0.5));
    assert_eq!(movie.poster_time, Some(0.0));
    assert_eq!(movie.loop_option, Some(0));
    assert_eq!(movie.volume, Some(1.0));
    assert_eq!(movie.audio_only, Some(true));
    assert_eq!(movie.streaming, Some(false));
    assert_eq!(movie.plays_across_slides, Some(true));
    assert_eq!(movie.poster_image_data, None);
    assert_eq!(movie.poster_image_generated_with_alpha_support, Some(false));
    assert_eq!(movie.flags, Some(0));
    assert_eq!(
        movie.original_size.map(|size| (size.width, size.height)),
        Some((0.0, 0.0))
    );
    assert_eq!(
        movie.natural_size.map(|size| (size.width, size.height)),
        Some((0.0, 0.0))
    );

    for &identifier in &created_standins {
        assert!(object_message(&candidate, identifier, STANDIN_MESSAGE_TYPE)?.is_empty());
    }

    let build = kn::BuildArchive::decode(
        object_message(&candidate, build_id, BUILD_MESSAGE_TYPE)?.as_slice(),
    )?;
    assert_eq!(
        build
            .drawable
            .as_ref()
            .map(|reference| reference.identifier),
        Some(movie_id)
    );
    assert_eq!(build.delivery, "All at Once");
    assert_eq!(build.duration, Some(0.0));
    assert_eq!(build.chunk_id_seed, Some(1));
    assert_eq!(build.attributes.event_trigger, Some(1));
    assert_eq!(build.attributes.chart_rotation3_d, Some(60.0));
    let animation = build
        .attributes
        .animation_attributes
        .as_ref()
        .ok_or_else(|| io::Error::other("created build has no animation attributes"))?;
    assert_eq!(animation.animation_type.as_deref(), Some("In"));
    assert_eq!(animation.effect.as_deref(), Some("apple:audio-start"));
    assert_eq!(animation.duration, Some(0.5));
    assert_eq!(animation.delay, Some(0.0));
    assert_eq!(animation.writing_direction_is_rtl, Some(false));

    let metadata_source = native_metadata(NATIVE_SOURCE)?;
    let metadata_candidate = native_metadata(&candidate)?;
    let source_uuid_map = uuid_map(&metadata_source, SLIDE_COMPONENT);
    let candidate_uuid_map = uuid_map(&metadata_candidate, SLIDE_COMPONENT);
    let uuid_additions = candidate_uuid_map
        .keys()
        .filter(|identifier| !source_uuid_map.contains_key(identifier))
        .copied()
        .collect::<BTreeSet<_>>();
    assert_eq!(
        uuid_additions,
        BTreeSet::from([movie_id, created_standins[0], created_standins[1], build_id])
    );
    assert!(!candidate_uuid_map.contains_key(&chunk_id));
    let build_uuid = candidate_uuid_map
        .get(&build_id)
        .copied()
        .ok_or_else(|| io::Error::other("created build has no UUID map entry"))?;
    assert_ne!(
        build_uuid,
        (0, 0),
        "created build must have a nonzero metadata UUID"
    );
    let chunk = kn::BuildChunkArchive::decode(
        object_message(&candidate, chunk_id, BUILD_CHUNK_MESSAGE_TYPE)?.as_slice(),
    )?;
    assert_eq!(
        chunk.build.as_ref().map(|reference| reference.identifier),
        Some(build_id)
    );
    assert_eq!(chunk.delay, Some(0.0));
    assert_eq!(chunk.duration, Some(0.5));
    assert_eq!(chunk.automatic, Some(false));
    assert_eq!(chunk.referent, Some(true));
    let chunk_build_uuid = chunk
        .build_id
        .ok_or_else(|| io::Error::other("created chunk has no playback build UUID"))?;
    let chunk_identifier_uuid = chunk
        .build_chunk_identifier
        .as_ref()
        .and_then(|identifier| identifier.build_id)
        .ok_or_else(|| io::Error::other("created chunk identifier has no playback build UUID"))?;
    assert_ne!(
        (chunk_build_uuid.lower, chunk_build_uuid.upper),
        (0, 0),
        "created chunk must have a nonzero playback build UUID"
    );
    assert_eq!(
        chunk_build_uuid, chunk_identifier_uuid,
        "chunk and chunk-identifier playback UUIDs must agree"
    );
    assert_eq!(
        chunk
            .build_chunk_identifier
            .as_ref()
            .and_then(|identifier| identifier.build_chunk_id),
        Some(1)
    );

    let source_node = kn::SlideNodeArchive::decode(
        object_message(NATIVE_SOURCE, SLIDE_NODE, SLIDE_NODE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let candidate_node = kn::SlideNodeArchive::decode(
        object_message(&candidate, SLIDE_NODE, SLIDE_NODE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let expected_event_count = u32::try_from(candidate_slide.builds.len())?;
    assert_eq!(candidate_node.build_event_count, Some(expected_event_count));
    assert_eq!(candidate_node.build_event_count_cache_version, Some(2));
    assert_eq!(candidate_node.has_explicit_builds, Some(true));
    assert_eq!(candidate_node.has_explicit_builds_cache_version, Some(2));
    assert_eq!(candidate_node.children, source_node.children);
    assert_eq!(candidate_node.slide, source_node.slide);
    assert_eq!(candidate_node.thumbnails, source_node.thumbnails);
    assert_eq!(candidate_node.thumbnail_sizes, source_node.thumbnail_sizes);
    assert_eq!(
        strip_fields(
            &object_message(NATIVE_SOURCE, SLIDE_NODE, SLIDE_NODE_MESSAGE_TYPE)?,
            &[15, 26, 20, 27],
        )?,
        strip_fields(
            &object_message(&candidate, SLIDE_NODE, SLIDE_NODE_MESSAGE_TYPE)?,
            &[15, 26, 20, 27],
        )?,
        "node-cache rewrite must preserve unknown and unrelated wire fields"
    );

    let source_data_ids = metadata_source
        .datas
        .iter()
        .map(|data| data.identifier)
        .collect::<BTreeSet<_>>();
    let candidate_data_ids = metadata_candidate
        .datas
        .iter()
        .map(|data| data.identifier)
        .collect::<BTreeSet<_>>();
    let new_data_ids = candidate_data_ids
        .difference(&source_data_ids)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(
        new_data_ids.len(),
        1,
        "fresh WAV must add one DataInfo record"
    );
    assert_eq!(movie_data_id, new_data_ids[0]);
    let data_info = data_record(&metadata_candidate, movie_data_id)?;
    assert_eq!(data_info.digest, Sha1::digest(&data).to_vec());
    assert_eq!(data_info.materialized_length, Some(data.len() as u64));
    assert_eq!(
        data_owners(&metadata_candidate, SLIDE_COMPONENT, movie_data_id)?,
        vec![(movie_id, 1)]
    );
    let data_name = data_info
        .file_name
        .as_deref()
        .filter(|name| !name.is_empty())
        .unwrap_or(&data_info.preferred_file_name);
    assert_eq!(
        catalog_member(&candidate, &format!("Data/{data_name}"))?,
        Some(data)
    );

    let source_metadata_component = metadata_component(&metadata_source, SLIDE_COMPONENT)?;
    let candidate_metadata_component = metadata_component(&metadata_candidate, SLIDE_COMPONENT)?;
    assert!(
        candidate_metadata_component
            .data_references
            .iter()
            .any(|reference| reference.data_identifier == movie_data_id)
    );
    for source_reference in &source_metadata_component.data_references {
        if source_reference.data_identifier == movie_data_id {
            continue;
        }
        let candidate_reference = candidate_metadata_component
            .data_references
            .iter()
            .find(|reference| reference.data_identifier == source_reference.data_identifier)
            .ok_or_else(|| io::Error::other("unrelated DataInfo owner was dropped"))?;
        assert_eq!(candidate_reference, source_reference);
    }
    for source_data in &metadata_source.datas {
        let candidate_data = data_record(&metadata_candidate, source_data.identifier)?;
        assert_eq!(candidate_data, source_data, "unrelated DataInfo changed");
    }
    assert_eq!(
        metadata_candidate.last_object_identifier,
        source_max_object_identifier + 5,
        "metadata watermark must cover exactly the five fresh object IDs"
    );
    assert!(
        new_ids
            .iter()
            .all(|identifier| *identifier <= metadata_candidate.last_object_identifier)
    );

    for (key, source_payload) in source_payloads {
        if key.component == slide_component && key.identifier == SLIDE_COMPONENT
            || key.component == node_component && key.identifier == SLIDE_NODE
            || key.message_type == PACKAGE_METADATA_MESSAGE_TYPE
        {
            continue;
        }
        assert_eq!(
            candidate_payloads.get(&key),
            Some(&source_payload),
            "unrelated native payload changed: {key:?}"
        );
    }

    let source_catalog = Catalog::from_bytes(NATIVE_SOURCE)?;
    let deleted_previews = source_catalog
        .iter()
        .filter(|entry| entry.name().starts_with("preview"))
        .map(|entry| entry.name().to_owned())
        .collect::<BTreeSet<_>>();
    for entry in source_catalog.iter() {
        if entry.name() == slide_component
            || entry.name() == "Index/Metadata.iwa"
            || entry.name() == node_component
            || deleted_previews.contains(entry.name())
            || entry.name().starts_with("Data/fresh-pcm-")
        {
            continue;
        }
        assert_eq!(
            catalog_member(&candidate, entry.name())?,
            Some(entry.data().to_vec()),
            "unrelated package member changed: {}",
            entry.name()
        );
    }
    assert!(commit.diagnostics().changed());
    Ok(())
}

#[test]
fn fresh_audio_allocates_above_metadata_watermark() -> TestResult {
    let physical_max = *all_object_ids(NATIVE_SOURCE)?
        .iter()
        .max()
        .ok_or_else(|| io::Error::other("native source has no objects"))?;
    let watermark = physical_max
        .checked_add(10_000)
        .ok_or_else(|| io::Error::other("metadata watermark overflowed"))?;
    let source = mutate_package_metadata(NATIVE_SOURCE, |metadata| {
        metadata.last_object_identifier = watermark;
        Ok(())
    })?;
    let source_metadata = native_metadata(&source)?;
    assert_eq!(source_metadata.last_object_identifier, watermark);

    let package = Package::from_bytes(&source)?;
    let commit = package.add_slide_audio(
        SlideSelector::index(0),
        "watermark-pcm.wav",
        &valid_wav(),
        options()?,
    )?;
    let candidate = package_bytes(commit.package())?;
    let slide_component = component_containing_object(&source, SLIDE_COMPONENT)?;
    let new_ids = fresh_ids(&source, &candidate, &slide_component)?;
    let expected_ids = (1..=5).map(|offset| watermark + offset).collect::<Vec<_>>();
    assert_eq!(new_ids, expected_ids);
    assert!(new_ids.iter().all(|identifier| *identifier > physical_max));
    let candidate_metadata = native_metadata(&candidate)?;
    assert_eq!(
        candidate_metadata.last_object_identifier,
        watermark + 5,
        "the metadata watermark must cover IDs allocated above the physical maximum"
    );
    assert!(commit.diagnostics().changed());
    Ok(())
}

#[test]
fn fresh_audio_uses_metadata_component_identity_and_rejects_duplicate_locator_atomically()
-> TestResult {
    let slide_component = component_containing_object(NATIVE_SOURCE, SLIDE_COMPONENT)?;
    let slide_component_locator = slide_component
        .strip_prefix("Index/")
        .unwrap_or(slide_component.as_str());
    let slide_locator = slide_component_locator
        .strip_suffix(".iwa")
        .unwrap_or(slide_component_locator)
        .to_owned();
    let source_metadata = native_metadata(NATIVE_SOURCE)?;
    let metadata_identifier_max = source_metadata
        .components
        .iter()
        .chain(&source_metadata.versioned_components)
        .map(|component| component.identifier)
        .max()
        .ok_or_else(|| io::Error::other("native metadata has no components"))?;
    let metadata_component_identifier = metadata_identifier_max
        .checked_add(10_000)
        .ok_or_else(|| io::Error::other("metadata component identifier overflowed"))?;
    let mismatched_source = mutate_package_metadata(NATIVE_SOURCE, |metadata| {
        let component = metadata
            .components
            .iter_mut()
            .find(|component| effective_locator(component) == slide_locator.as_str())
            .ok_or_else(|| io::Error::other("native slide metadata component is missing"))?;
        component.identifier = metadata_component_identifier;
        Ok(())
    })?;
    let mismatched_metadata = native_metadata(&mismatched_source)?;
    let source_component = metadata_component_for_locator(&mismatched_metadata, &slide_locator)?;
    assert_ne!(source_component.identifier, SLIDE_COMPONENT);
    assert_eq!(source_component.identifier, metadata_component_identifier);

    let package = Package::from_bytes(&mismatched_source)?;
    let commit = package.add_slide_audio(
        SlideSelector::index(0),
        "component-id-pcm.wav",
        &valid_wav(),
        options()?,
    )?;
    let candidate = package_bytes(commit.package())?;
    let candidate_metadata = native_metadata(&candidate)?;
    let candidate_component = metadata_component_for_locator(&candidate_metadata, &slide_locator)?;
    assert_eq!(
        candidate_component.identifier,
        metadata_component_identifier
    );
    let source_component_token = mismatched_metadata.save_token.unwrap_or(0);
    assert_eq!(
        candidate_component.save_token,
        Some(
            source_component_token
                .checked_add(1)
                .ok_or_else(|| io::Error::other("component save token overflowed"))?,
        ),
        "the save-token selector must target the metadata component identity"
    );
    let source_root_token = mismatched_metadata.save_token.unwrap_or(0);
    assert_eq!(
        candidate_metadata.save_token,
        Some(
            source_root_token
                .checked_add(1)
                .ok_or_else(|| io::Error::other("root save token overflowed"))?,
        )
    );

    let created_movie = fresh_ids_with_message_type(
        &mismatched_source,
        &candidate,
        &slide_component,
        MOVIE_MESSAGE_TYPE,
    )?;
    let created_standins = fresh_ids_with_message_type(
        &mismatched_source,
        &candidate,
        &slide_component,
        STANDIN_MESSAGE_TYPE,
    )?;
    let created_build = fresh_ids_with_message_type(
        &mismatched_source,
        &candidate,
        &slide_component,
        BUILD_MESSAGE_TYPE,
    )?;
    assert_eq!(created_movie.len(), 1);
    assert_eq!(created_standins.len(), 2);
    assert_eq!(created_build.len(), 1);
    let movie_id = created_movie[0];
    let build_id = created_build[0];

    let source_uuid_map = uuid_map(&mismatched_metadata, metadata_component_identifier);
    let candidate_uuid_map = uuid_map(&candidate_metadata, metadata_component_identifier);
    let uuid_additions = candidate_uuid_map
        .keys()
        .filter(|identifier| !source_uuid_map.contains_key(identifier))
        .copied()
        .collect::<BTreeSet<_>>();
    assert_eq!(
        uuid_additions,
        BTreeSet::from([movie_id, created_standins[0], created_standins[1], build_id,])
    );

    let movie = tsd::MovieArchive::decode(
        object_message(&candidate, movie_id, MOVIE_MESSAGE_TYPE)?.as_slice(),
    )?;
    let movie_data_id = movie
        .movie_data
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("created movie has no content data"))?;
    assert_eq!(
        data_owners(
            &candidate_metadata,
            metadata_component_identifier,
            movie_data_id
        )?,
        vec![(movie_id, 1)],
        "the media owner selector must target the metadata component identity"
    );

    // A second current component with the selected effective locator makes
    // the metadata selector deliberately ambiguous. The package must reject
    // before publishing any candidate bytes.
    let invalid_source = mutate_package_metadata(&mismatched_source, |metadata| {
        let mut duplicate = metadata_component_for_locator(metadata, &slide_locator)?.clone();
        let duplicate_identifier = metadata
            .components
            .iter()
            .chain(&metadata.versioned_components)
            .map(|component| component.identifier)
            .max()
            .ok_or_else(|| io::Error::other("native metadata has no components"))?
            .checked_add(1)
            .ok_or_else(|| io::Error::other("duplicate component identifier overflowed"))?;
        duplicate.identifier = duplicate_identifier;
        metadata.components.push(duplicate);
        Ok(())
    })?;
    let invalid_package = Package::from_bytes(&invalid_source)?;
    assert_eq!(package_bytes(&invalid_package)?, invalid_source);
    let rejected = invalid_package.add_slide_audio(
        SlideSelector::index(0),
        "duplicate-locator.wav",
        &valid_wav(),
        options()?,
    );
    assert!(rejected.is_err());
    assert_eq!(
        package_bytes(&invalid_package)?,
        invalid_source,
        "invalid metadata rejection must preserve the exact source ZIP"
    );
    assert!(commit.diagnostics().changed());
    Ok(())
}

#[test]
fn audio_position_validates_build_payload_before_admitting_its_header_edge() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let creation = source.add_slide_audio(
        SlideSelector::index(0),
        "build-edge.wav",
        &valid_wav(),
        options()?,
    )?;
    let candidate = package_bytes(creation.package())?;
    let original_ids = all_object_ids(NATIVE_SOURCE)?;
    let (component, archive) = native_archives(&candidate)?
        .into_iter()
        .find(|(_, archive)| {
            archive.objects.iter().any(|object| {
                object
                    .archive_info
                    .identifier
                    .is_some_and(|id| !original_ids.contains(&id))
                    && object
                        .messages
                        .iter()
                        .any(|message| message.type_ == BUILD_MESSAGE_TYPE)
            })
        })
        .ok_or_else(|| io::Error::other("created audio build is missing"))?;
    let build_id = archive
        .objects
        .iter()
        .find(|object| {
            object
                .archive_info
                .identifier
                .is_some_and(|id| !original_ids.contains(&id))
                && object
                    .messages
                    .iter()
                    .any(|message| message.type_ == BUILD_MESSAGE_TYPE)
        })
        .and_then(|object| object.archive_info.identifier)
        .ok_or_else(|| io::Error::other("created audio build has no identifier"))?;
    let selector = litchi_keynote::MovieSelector::position(creation.patch().movie_position());

    // Sparse producer headers were admitted before fresh creation existed.
    let mut sparse_archive = archive.clone();
    sparse_archive
        .object_mut(build_id)
        .ok_or_else(|| io::Error::other("missing build"))?
        .archive_info
        .message_infos[0]
        .object_references
        .clear();
    let sparse = replace_component(&candidate, &component, &sparse_archive)?;
    let sparse_package = Package::from_bytes(&sparse)?;
    let _moved = sparse_package
        .edit_slide_audio_position(SlideSelector::index(0), selector)?
        .set(litchi_keynote::slide::media::Point { x: 27.0, y: 49.0 })?
        .commit()?;

    // A header pointing to the selected audio cannot authorize a build whose
    // actual payload points elsewhere, even when it belongs to this slide.
    let mut forged_archive = archive;
    let object = forged_archive
        .object_mut(build_id)
        .ok_or_else(|| io::Error::other("missing build"))?;
    let mut build = kn::BuildArchive::decode(object.messages[0].data.as_slice())?;
    build.drawable = Some(tsp::Reference {
        identifier: SLIDE_NODE,
        ..Default::default()
    });
    object.replace_message_preserving_header(
        0,
        litchi_iwa_core::RawMessage {
            type_: BUILD_MESSAGE_TYPE,
            data: build.encode_to_vec(),
        },
    )?;
    let forged = replace_component(&candidate, &component, &forged_archive)?;
    let forged_package = Package::from_bytes(&forged)?;
    assert!(matches!(
        forged_package.edit_slide_audio_position(SlideSelector::index(0), selector),
        Err(litchi_keynote::SlideAudioPositionError::UnsupportedDependency),
    ));
    assert_eq!(package_bytes(&forged_package)?, forged);
    Ok(())
}
