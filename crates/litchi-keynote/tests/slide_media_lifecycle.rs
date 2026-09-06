//! Focused selector-first Keynote media lifecycle coverage.
//!
//! The lifecycle owner is exercised through semantic slide and media
//! selectors.  These tests intentionally keep native identifiers, object
//! graphs, and metadata records behind fixture-only helpers; the public
//! transaction only exposes a reopened package and an exact-source patch.

use std::{
    collections::{BTreeMap, BTreeSet},
    env, fs, io,
    path::PathBuf,
};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_media_lifecycle_codec as lifecycle_codec, kn, tsa, tsd, tsp, tswp,
};
use litchi_keynote::{
    MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideMediaLifecycleError,
    SlideMediaLifecycleLimitKind, SlideSelector,
};
use prost::Message as _;

#[path = "support/slide_media_fixture.rs"]
#[allow(
    dead_code,
    reason = "Shared fixture also supports media replacement tests."
)]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-native.key");
const NATIVE_DUPLICATE_MOVIE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-duplicate-native.key");
const NATIVE_REMOVE_MOVIE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-remove-native.key");
const NATIVE_DUPLICATE_AUDIO: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-audio-duplicate-native.key");
const NATIVE_REMOVE_AUDIO: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-lifecycle-audio-remove-native.key");

const NATIVE_SLIDE_COMPONENT: u64 = 2_652_150;
const NATIVE_SLIDE_NODE: u64 = 2_652_149;
const NATIVE_MOVIE_A: u64 = 2_653_286;
const NATIVE_MOVIE_B: u64 = 2_653_610;
const NATIVE_AUDIO_A: u64 = 2_652_595;
const NATIVE_AUDIO_B: u64 = 2_652_622;
const NATIVE_AUDIO_DATA: u64 = 9_075;
const NATIVE_MOVIE_DATA: u64 = 9_085;
const NATIVE_POSTER_DATA: u64 = 9_086;
const NATIVE_SLIDE_MESSAGE_TYPE: u32 = 5;
const NATIVE_BUILD_MESSAGE_TYPE: u32 = 8;
const NATIVE_BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const NATIVE_MOVIE_MESSAGE_TYPE: u32 = 3_007;
const NATIVE_CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const NATIVE_STORAGE_MESSAGE_TYPE: u32 = 2_001;
const NATIVE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const NATIVE_AUDIO_MEMBER: &str = "Data/keynote-coral-9075.wav";
const NATIVE_MOVIE_MEMBER: &str = "Data/keynote-selfauthored-coral-mjpeg-9085.mov";
const NATIVE_POSTER_MEMBER: &str = "Data/posterImage-9086.png";
const NATIVE_CAPTION: &str = "Shared native movie caption";

#[derive(Debug, Clone, PartialEq)]
struct NativeMovieGraph {
    movie_data: Option<u64>,
    poster_data: Option<u64>,
    title: Option<u64>,
    caption: Option<u64>,
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

fn package() -> TestResult<Package> {
    Ok(Package::from_bytes(&lifecycle_source()?)?)
}

/// Add the type-11015 DataMetadataMap witness required when a lifecycle edit
/// removes the final physical owner of a DataInfo record.  The existing media
/// replacement fixture intentionally omits that optional graph because it
/// never removes records; lifecycle tests keep the two admission profiles
/// separate instead of weakening the shared fixture.
fn with_data_metadata_map(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut metadata = tsp::PackageMetadata::decode(fixture::metadata_stream(source)?.as_slice())?;
    metadata.data_metadata_map = Some(tsp::Reference {
        identifier: 500,
        ..Default::default()
    });
    let source = fixture::replace_metadata_payload(source, metadata.encode_to_vec())?;
    let document_stream =
        SnappyStream::decompress(&fixture::member_bytes(&source, fixture::DOCUMENT_MEMBER)?)?;
    let mut archive = Archive::parse(document_stream.as_bytes())?;

    // Selected media are absent from the map, as in the native source.
    // A mapped DataInfo is deliberately outside final-GC admission.
    let map = Vec::new();
    archive.insert_object(ArchiveObject::new(
        500,
        vec![RawMessage {
            type_: 11_015,
            data: map,
        }],
    )?)?;
    fixture::replace_document_archive(&source, archive)
}

fn with_repeated_content_owner(source: &[u8], repeats: usize) -> TestResult<Vec<u8>> {
    let repeats = u32::try_from(repeats)?;
    let mut archive = Archive::parse(fixture::document_stream(source)?.as_slice())?;
    let movie = archive
        .object_mut(fixture::MOVIES[0])
        .ok_or_else(|| io::Error::other("missing repeated-owner movie"))?;
    let info = movie
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing repeated-owner message metadata"))?;
    let mut data_references = Vec::with_capacity(repeats as usize + 1);
    data_references.extend(std::iter::repeat(fixture::CONTENT_DATA).take(repeats as usize));
    data_references.push(fixture::POSTER_DATA);
    info.data_references = data_references;
    let source = fixture::replace_document_archive(source, archive)?;

    let mut metadata = tsp::PackageMetadata::decode(fixture::metadata_stream(&source)?.as_slice())?;
    let component = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == fixture::DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing repeated-owner component"))?;
    let content = component
        .data_references
        .iter_mut()
        .find(|reference| reference.data_identifier == fixture::CONTENT_DATA)
        .ok_or_else(|| io::Error::other("missing repeated-owner content record"))?;
    let owner = content
        .object_reference_list
        .iter_mut()
        .find(|owner| owner.object_identifier == fixture::MOVIES[0])
        .ok_or_else(|| io::Error::other("missing repeated-owner movie record"))?;
    owner.count = repeats;
    fixture::replace_metadata_payload(&source, metadata.encode_to_vec())
}

fn with_surviving_private_child_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(fixture::document_stream(source)?.as_slice())?;
    let movie = archive
        .object_mut(fixture::MOVIES[1])
        .ok_or_else(|| io::Error::other("missing surviving movie"))?;
    let info = movie
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing surviving movie metadata"))?;
    info.object_references.push(fixture::CAPTIONS[0]);
    fixture::replace_document_archive(source, archive)
}

fn with_slide_header_object_references(
    source: &[u8],
    update: impl FnOnce(&mut Vec<u64>),
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(fixture::document_stream(source)?.as_slice())?;
    let slide = archive
        .object_mut(fixture::SLIDE)
        .ok_or_else(|| io::Error::other("missing slide object"))?;
    let info = slide
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing slide message metadata"))?;
    update(&mut info.object_references);
    fixture::replace_document_archive(source, archive)
}

fn with_slide_header_field_info(
    source: &[u8],
    references: Vec<u64>,
    field_type: FieldType,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(fixture::document_stream(source)?.as_slice())?;
    let slide = archive
        .object_mut(fixture::SLIDE)
        .ok_or_else(|| io::Error::other("missing slide object"))?;
    let info = slide
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing slide message metadata"))?;
    info.field_infos.push(FieldInfo {
        path: vec![7].into(),
        r#type: Some(field_type),
        object_references: references,
        ..FieldInfo::default()
    });
    fixture::replace_document_archive(source, archive)
}

fn replace_native_component_object_payload(
    source: &[u8],
    identifier: u64,
    type_: u32,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) = native_component_containing_object(source, identifier)?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| {
            io::Error::other(format!("missing native object {identifier} type {type_}"))
        })?;
    message.data = payload;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
                (entry.name(), replacement.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn native_selected_chunk_for_movie(source: &[u8], movie: u64) -> TestResult<(u64, Vec<u8>)> {
    let builds = native_build_edges(source)?
        .into_iter()
        .filter(|(_, drawable)| *drawable == movie)
        .map(|(identifier, _)| identifier)
        .collect::<BTreeSet<_>>();
    let (chunk, _, chunk_uuid, build_uuid) = native_chunk_edges(source)?
        .into_iter()
        .find(|(_, build, _, _)| builds.contains(build))
        .ok_or_else(|| io::Error::other("missing selected native build chunk"))?;
    assert!(chunk_uuid.is_some(), "selected native chunk lacks its UUID");
    assert!(
        build_uuid.is_some(),
        "selected native chunk lacks its build UUID"
    );
    let (_, archive) = native_component_containing_object(source, chunk)?;
    Ok((
        chunk,
        native_object_message(&archive, chunk, NATIVE_BUILD_CHUNK_MESSAGE_TYPE)?.to_vec(),
    ))
}

fn physical_uncompressed_total(source: &[u8]) -> TestResult<u64> {
    Catalog::from_bytes(source)?
        .iter()
        .try_fold(0_u64, |total, entry| {
            let size = if entry.name().ends_with(".iwa") {
                u64::try_from(SnappyStream::decompress(entry.data())?.as_bytes().len())?
                    .max(entry.metadata().uncompressed_size())
            } else {
                entry.metadata().uncompressed_size()
            };
            total
                .checked_add(size)
                .ok_or_else(|| io::Error::other("physical total overflowed").into())
        })
}

fn lifecycle_source() -> TestResult<Vec<u8>> {
    with_data_metadata_map(&fixture::synthetic_package()?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn movie_count(package: &Package) -> TestResult<usize> {
    Ok(package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("fixture has no slide"))?
        .movies()
        .len())
}

fn data_member(source: &[u8], name: &str) -> TestResult<Option<Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec()))
}

fn export_if_requested(package: &Package, name: &str) -> TestResult<()> {
    let Some(directory) = env::var_os("LITCHI_KEYNOTE_MEDIA_LIFECYCLE_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    fs::create_dir_all(&directory)?;
    let path = directory.join(name);
    fs::write(&path, exact_bytes(package)?)?;
    eprintln!("exported Keynote lifecycle candidate to {}", path.display());
    Ok(())
}

fn assert_member_unchanged(before: &[u8], after: &[u8], name: &str) -> TestResult<()> {
    assert_eq!(
        data_member(before, name)?,
        data_member(after, name)?,
        "unrelated or shared member {name} changed",
    );
    Ok(())
}

fn assert_no_member(source: &[u8], name: &str) -> TestResult<()> {
    assert!(
        data_member(source, name)?.is_none(),
        "member {name} should have been reclaimed",
    );
    Ok(())
}

fn native_member_bytes(source: &[u8], name: &str) -> TestResult<Vec<u8>> {
    data_member(source, name)?
        .ok_or_else(|| io::Error::other(format!("missing native package member {name}")).into())
}

fn native_component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
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

fn native_component_containing_object(
    source: &[u8],
    identifier: u64,
) -> TestResult<(String, Archive)> {
    native_component_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .ok_or_else(|| io::Error::other(format!("missing native object {identifier}")))
        .map_err(Into::into)
}

fn native_object_message<'a>(
    archive: &'a Archive,
    identifier: u64,
    type_: u32,
) -> TestResult<&'a [u8]> {
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

#[derive(Debug, Clone, PartialEq, Eq)]
struct NativeTextObjectSemantics {
    /// Caption/title references are private graph edges.  Canonicalizing
    /// their identifiers lets the assertion compare the decoded shape and
    /// scalar fields while still rejecting an accidental field loss.
    caption_payload: Vec<u8>,
    /// Text storage has its own private references (style and rich-text
    /// attribute tables).  Those references are canonicalized for the same
    /// reason, while text and every scalar field remain in the encoded view.
    storage_payload: Option<Vec<u8>>,
    text: Option<Vec<String>>,
}

fn canonical_native_reference(reference: &mut Option<tsp::Reference>) {
    if let Some(reference) = reference {
        reference.identifier = 0;
    }
}

fn canonical_native_drawable(drawable: &mut tsd::DrawableArchive) {
    canonical_native_reference(&mut drawable.parent);
    canonical_native_reference(&mut drawable.comment);
    for reference in &mut drawable.pencil_annotations {
        reference.identifier = 0;
    }
    canonical_native_reference(&mut drawable.title);
    canonical_native_reference(&mut drawable.caption);
}

fn canonical_native_shape(shape: &mut tsd::ShapeArchive) {
    canonical_native_drawable(&mut shape.super_);
    canonical_native_reference(&mut shape.style);
}

#[allow(
    deprecated,
    reason = "The native fixture intentionally exercises both storage edges."
)]
fn canonical_native_shape_info(shape: &mut tswp::ShapeInfoArchive) {
    canonical_native_shape(&mut shape.super_);
    canonical_native_reference(&mut shape.deprecated_storage);
    canonical_native_reference(&mut shape.text_flow);
    canonical_native_reference(&mut shape.owned_storage);
}

#[allow(
    deprecated,
    reason = "The native fixture intentionally exercises both storage edges."
)]
fn canonical_native_caption(caption: &mut tsa::CaptionInfoArchive) {
    canonical_native_shape_info(&mut caption.super_);
    canonical_native_reference(&mut caption.placement);
}

fn canonical_native_object_attributes(table: &mut tswp::ObjectAttributeTable) {
    for entry in &mut table.entries {
        canonical_native_reference(&mut entry.object);
    }
}

fn canonical_native_overlapping_attributes(table: &mut tswp::OverlappingFieldAttributeTable) {
    for entry in &mut table.entries {
        entry.field.identifier = 0;
    }
}

fn canonical_native_storage(storage: &mut tswp::StorageArchive) {
    canonical_native_reference(&mut storage.style_sheet);
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_attachment,
        &mut storage.table_smartfield,
        &mut storage.table_layout_style,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_section,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_highlight,
        &mut storage.table_tatechuyoko,
        &mut storage.table_drop_cap_style,
    ] {
        if let Some(table) = table.as_mut() {
            canonical_native_object_attributes(table);
        }
    }
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ] {
        if let Some(table) = table.as_mut() {
            canonical_native_overlapping_attributes(table);
        }
    }
}

#[allow(
    deprecated,
    reason = "The native fixture intentionally exercises both storage edges."
)]
fn native_text_object_semantics(
    source: &[u8],
    identifier: u64,
) -> TestResult<NativeTextObjectSemantics> {
    let (_, archive) = native_component_containing_object(source, identifier)?;
    let payload = native_object_message(&archive, identifier, NATIVE_CAPTION_INFO_MESSAGE_TYPE)?;
    let mut caption = tsa::CaptionInfoArchive::decode(payload)?;
    let storage_identifier = caption
        .super_
        .owned_storage
        .as_ref()
        .or(caption.super_.deprecated_storage.as_ref())
        .map(|reference| reference.identifier);
    let (storage_payload, text) = if let Some(storage_identifier) = storage_identifier {
        let (_, archive) = native_component_containing_object(source, storage_identifier)?;
        let payload =
            native_object_message(&archive, storage_identifier, NATIVE_STORAGE_MESSAGE_TYPE)?;
        let mut storage = tswp::StorageArchive::decode(payload)?;
        let text = storage.text.clone();
        canonical_native_storage(&mut storage);
        (Some(storage.encode_to_vec()), Some(text))
    } else {
        (None, None)
    };
    canonical_native_caption(&mut caption);
    Ok(NativeTextObjectSemantics {
        caption_payload: caption.encode_to_vec(),
        storage_payload,
        text,
    })
}

fn native_slide_archive(source: &[u8]) -> TestResult<kn::SlideArchive> {
    let (_, archive) = native_component_containing_object(source, NATIVE_SLIDE_COMPONENT)?;
    Ok(kn::SlideArchive::decode(native_object_message(
        &archive,
        NATIVE_SLIDE_COMPONENT,
        NATIVE_SLIDE_MESSAGE_TYPE,
    )?)?)
}

fn native_movie_archive(source: &[u8], identifier: u64) -> TestResult<tsd::MovieArchive> {
    let (_, archive) = native_component_containing_object(source, identifier)?;
    Ok(tsd::MovieArchive::decode(native_object_message(
        &archive,
        identifier,
        NATIVE_MOVIE_MESSAGE_TYPE,
    )?)?)
}

fn native_movie_graph(source: &[u8], identifier: u64) -> TestResult<NativeMovieGraph> {
    let movie = native_movie_archive(source, identifier)?;
    let geometry = movie.super_.geometry.as_ref();
    Ok(NativeMovieGraph {
        movie_data: movie.movie_data.map(|reference| reference.identifier),
        poster_data: movie
            .poster_image_data
            .map(|reference| reference.identifier),
        title: movie.super_.title.map(|reference| reference.identifier),
        caption: movie.super_.caption.map(|reference| reference.identifier),
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

#[allow(
    deprecated,
    reason = "The native fixture has both current and deprecated caption storage edges."
)]
fn native_caption_text(source: &[u8], movie_identifier: u64) -> TestResult<Option<String>> {
    let movie = native_movie_archive(source, movie_identifier)?;
    let Some(caption_identifier) = movie
        .super_
        .caption
        .as_ref()
        .map(|reference| reference.identifier)
    else {
        return Ok(None);
    };
    let (_, archive) = native_component_containing_object(source, caption_identifier)?;
    if let Some(object) = archive.object(caption_identifier) {
        if object.messages.len() == 1
            && object.messages[0].type_ == 3097
            && object.messages[0].data.is_empty()
        {
            return Ok(None);
        }
    }
    let caption = tsa::CaptionInfoArchive::decode(native_object_message(
        &archive,
        caption_identifier,
        NATIVE_CAPTION_INFO_MESSAGE_TYPE,
    )?)?;
    let storage_identifier = caption
        .super_
        .owned_storage
        .as_ref()
        .or(caption.super_.deprecated_storage.as_ref())
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("native caption has no text storage"))?;
    let (_, archive) = native_component_containing_object(source, storage_identifier)?;
    Ok(tswp::StorageArchive::decode(native_object_message(
        &archive,
        storage_identifier,
        NATIVE_STORAGE_MESSAGE_TYPE,
    )?)?
    .text
    .into_iter()
    .next())
}

fn native_metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut matches = native_component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| {
            archive.objects.into_iter().flat_map(|object| {
                object.messages.into_iter().filter_map(|message| {
                    (message.type_ == NATIVE_METADATA_MESSAGE_TYPE).then_some(message.data)
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

fn native_metadata(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        native_metadata_payload(source)?.as_slice(),
    )?)
}

fn native_document_component<'a>(
    metadata: &'a tsp::PackageMetadata,
) -> TestResult<&'a tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| component.identifier == NATIVE_SLIDE_COMPONENT)
        .ok_or_else(|| io::Error::other("missing native slide component"))
        .map_err(Into::into)
}

fn native_data_record(
    metadata: &tsp::PackageMetadata,
    identifier: u64,
) -> TestResult<&tsp::DataInfo> {
    metadata
        .datas
        .iter()
        .find(|data| data.identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing native DataInfo {identifier}")))
        .map_err(Into::into)
}

fn native_owner_list(
    metadata: &tsp::PackageMetadata,
    identifier: u64,
) -> TestResult<Vec<(u64, u32)>> {
    let component = native_document_component(metadata)?;
    let record = component
        .data_references
        .iter()
        .find(|record| record.data_identifier == identifier)
        .ok_or_else(|| io::Error::other(format!("missing native data owners {identifier}")))?;
    let mut owners = record
        .object_reference_list
        .iter()
        .map(|owner| (owner.object_identifier, owner.count))
        .collect::<Vec<_>>();
    owners.sort_unstable();
    Ok(owners)
}

fn native_uuid_map(metadata: &tsp::PackageMetadata) -> BTreeMap<u64, (u64, u64)> {
    native_document_component(metadata)
        .map(|component| {
            component
                .object_uuid_map_entries
                .iter()
                .map(|entry| (entry.identifier, (entry.uuid.lower, entry.uuid.upper)))
                .collect()
        })
        .unwrap_or_default()
}

fn native_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let slide = native_slide_archive(source)?;
    Ok(slide
        .owned_drawables
        .into_iter()
        .filter_map(|reference| {
            let identifier = reference.identifier;
            native_movie_archive(source, identifier)
                .ok()
                .map(|_| identifier)
        })
        .collect())
}

fn native_object_graph_ids(source: &[u8], root: u64) -> TestResult<BTreeSet<u64>> {
    let (owned_drawables, z_order, builds, chunks) = native_slide_lifecycle_ids(source)?;
    let mut shared = BTreeSet::from([NATIVE_SLIDE_COMPONENT]);
    shared.extend(owned_drawables);
    shared.extend(z_order);
    shared.extend(builds);
    shared.extend(chunks);
    let mut edges = BTreeMap::<u64, BTreeSet<u64>>::new();
    for (_, archive) in native_component_archives(source)? {
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            let references = edges.entry(identifier).or_default();
            for message_info in &object.archive_info.message_infos {
                references.extend(message_info.object_references.iter().copied());
                for field_info in &message_info.field_infos {
                    references.extend(field_info.object_references.iter().copied());
                }
            }
        }
    }

    let mut output = BTreeSet::from([root]);
    let mut pending = vec![root];
    while let Some(identifier) = pending.pop() {
        for referenced in edges.get(&identifier).into_iter().flatten().copied() {
            if shared.contains(&referenced) {
                continue;
            }
            if output.insert(referenced) {
                pending.push(referenced);
            }
        }
    }
    Ok(output)
}

fn native_slide_lifecycle_ids(
    source: &[u8],
) -> TestResult<(Vec<u64>, Vec<u64>, Vec<u64>, Vec<u64>)> {
    let (_, archive) = native_component_containing_object(source, NATIVE_SLIDE_COMPONENT)?;
    let payload =
        native_object_message(&archive, NATIVE_SLIDE_COMPONENT, NATIVE_SLIDE_MESSAGE_TYPE)?;
    let slide = lifecycle_codec::decode_slide_lifecycle(
        payload,
        lifecycle_codec::DecodeOptions::for_source(payload),
    )?;
    Ok((
        slide
            .owned_drawables()
            .map(|reference| reference.identifier())
            .collect(),
        slide
            .drawables_z_order()
            .map(|reference| reference.identifier())
            .collect(),
        slide
            .builds()
            .map(|reference| reference.identifier())
            .collect(),
        slide
            .build_chunks()
            .map(|reference| reference.identifier())
            .collect(),
    ))
}

fn native_build_edges(source: &[u8]) -> TestResult<Vec<(u64, u64)>> {
    let (_, _, build_ids, _) = native_slide_lifecycle_ids(source)?;
    let mut edges = Vec::new();
    for identifier in build_ids {
        let (_, archive) = native_component_containing_object(source, identifier)?;
        let payload = native_object_message(&archive, identifier, NATIVE_BUILD_MESSAGE_TYPE)?;
        let build = lifecycle_codec::decode_build(
            payload,
            lifecycle_codec::DecodeOptions::for_source(payload),
        )?;
        if let Some(drawable) = build.drawable() {
            edges.push((identifier, drawable.identifier()));
        }
    }
    Ok(edges)
}

type NativeChunkEdge = (u64, u64, Option<(u64, u64)>, Option<(u64, u64)>);

fn native_chunk_edges(source: &[u8]) -> TestResult<Vec<NativeChunkEdge>> {
    let (_, _, _, chunk_ids) = native_slide_lifecycle_ids(source)?;
    let mut edges = Vec::new();
    for identifier in chunk_ids {
        let (_, archive) = native_component_containing_object(source, identifier)?;
        let payload = native_object_message(&archive, identifier, NATIVE_BUILD_CHUNK_MESSAGE_TYPE)?;
        let chunk = lifecycle_codec::decode_build_chunk(
            payload,
            lifecycle_codec::DecodeOptions::for_source(payload),
        )?;
        let uuid = |snapshot: Option<lifecycle_codec::UuidSnapshot<'_>>| {
            snapshot.map(|snapshot| {
                let uuid = snapshot.uuid();
                (uuid.lower(), uuid.upper())
            })
        };
        edges.push((
            identifier,
            chunk.build().identifier(),
            uuid(chunk.chunk_identifier()),
            uuid(chunk.build_id()),
        ));
    }
    Ok(edges)
}

fn native_candidate_media_id(source: &[u8], known: &[u64]) -> TestResult<u64> {
    let known = known.iter().copied().collect::<BTreeSet<_>>();
    let candidates = native_media_ids(source)?
        .into_iter()
        .filter(|identifier| !known.contains(identifier))
        .collect::<Vec<_>>();
    assert_eq!(
        candidates.len(),
        1,
        "expected exactly one newly cloned media object"
    );
    Ok(candidates[0])
}

fn assert_native_text_clone(
    source: &[u8],
    candidate: &[u8],
    old: u64,
    new: u64,
) -> TestResult<bool> {
    let (_, source_archive) = native_component_containing_object(source, old)?;
    let (_, clone_archive) = native_component_containing_object(candidate, new)?;
    let source_object = source_archive
        .object(old)
        .ok_or_else(|| io::Error::other("missing source text object"))?;
    let clone_object = clone_archive
        .object(new)
        .ok_or_else(|| io::Error::other("missing cloned text object"))?;
    if source_object.messages.len() == 1 && source_object.messages[0].type_ == 3097 {
        assert!(source_object.messages[0].data.is_empty());
        assert_eq!(
            clone_object.messages, source_object.messages,
            "empty stand-in must be preserved"
        );
        return Ok(false);
    }
    assert_ne!(
        native_object_message(&source_archive, old, NATIVE_CAPTION_INFO_MESSAGE_TYPE)?,
        native_object_message(&clone_archive, new, NATIVE_CAPTION_INFO_MESSAGE_TYPE)?,
        "caption storage reference must be remapped"
    );
    assert_eq!(
        native_text_object_semantics(source, old)?,
        native_text_object_semantics(candidate, new)?,
        "caption text/style/scalars must be preserved"
    );
    Ok(true)
}

fn assert_native_movie_clone_graph(
    source: &[u8],
    candidate: &[u8],
    source_movie: u64,
    cloned_movie: u64,
    cloned_data: &[u64],
) -> TestResult<(Vec<u64>, Vec<u64>)> {
    let before = native_movie_graph(source, source_movie)?;
    let after = native_movie_graph(candidate, cloned_movie)?;
    assert_eq!(after.movie_data, before.movie_data);
    assert_eq!(after.poster_data, before.poster_data);
    assert_eq!(after.size, before.size);
    assert_eq!(after.original_size, before.original_size);
    assert_eq!(after.natural_size, before.natural_size);
    assert_eq!(after.start_time, before.start_time);
    assert_eq!(after.end_time, before.end_time);
    assert_eq!(after.poster_time, before.poster_time);
    assert_eq!(after.loop_option, before.loop_option);
    assert_eq!(after.volume, before.volume);
    assert_eq!(after.audio_only, before.audio_only);
    assert_eq!(after.plays_across_slides, before.plays_across_slides);
    assert_ne!(
        after.position, before.position,
        "a duplicate must receive a fresh position"
    );

    if let (Some(source_title), Some(clone_title)) = (before.title, after.title) {
        assert_ne!(
            source_title, clone_title,
            "title objects must be privately cloned"
        );
        assert_native_text_clone(source, candidate, source_title, clone_title)?;
    }
    if let (Some(source_caption), Some(clone_caption)) = (before.caption, after.caption) {
        assert_ne!(
            source_caption, clone_caption,
            "caption must be privately cloned"
        );
        if assert_native_text_clone(source, candidate, source_caption, clone_caption)? {
            assert_eq!(
                native_caption_text(source, source_movie)?,
                native_caption_text(candidate, cloned_movie)?
            );
            assert_eq!(
                native_caption_text(candidate, cloned_movie)?.as_deref(),
                Some(NATIVE_CAPTION)
            );
        }
    }

    let source_graph = native_object_graph_ids(source, source_movie)?;
    let clone_graph = native_object_graph_ids(candidate, cloned_movie)?;
    assert!(source_graph.contains(&source_movie));
    assert!(clone_graph.contains(&cloned_movie));
    assert!(!source_graph.contains(&cloned_movie));
    assert!(!clone_graph.contains(&source_movie));
    assert_eq!(
        source_graph.len(),
        clone_graph.len(),
        "private clone graph changed shape"
    );

    let source_builds = native_build_edges(source)?
        .into_iter()
        .filter(|(_, drawable)| *drawable == source_movie)
        .map(|(identifier, _)| identifier)
        .collect::<Vec<_>>();
    let clone_builds = native_build_edges(candidate)?
        .into_iter()
        .filter(|(_, drawable)| *drawable == cloned_movie)
        .map(|(identifier, _)| identifier)
        .collect::<Vec<_>>();
    assert_eq!(clone_builds.len(), source_builds.len());
    assert!(
        !source_builds.is_empty(),
        "native media fixture must exercise builds"
    );
    assert!(
        source_builds
            .iter()
            .all(|identifier| !clone_builds.contains(identifier))
    );

    let source_chunks = native_chunk_edges(source)?
        .into_iter()
        .filter(|(_, build, _, _)| source_builds.contains(build))
        .collect::<Vec<_>>();
    let clone_chunks = native_chunk_edges(candidate)?
        .into_iter()
        .filter(|(_, build, _, _)| clone_builds.contains(build))
        .collect::<Vec<_>>();
    assert_eq!(clone_chunks.len(), source_chunks.len());
    assert!(
        !source_chunks.is_empty(),
        "native media fixture must exercise build chunks"
    );
    assert!(source_chunks.iter().all(|(identifier, _, _, _)| {
        !clone_chunks
            .iter()
            .any(|(clone, _, _, _)| clone == identifier)
    }));

    let source_metadata = native_metadata(source)?;
    let candidate_metadata = native_metadata(candidate)?;
    let source_uuids = native_uuid_map(&source_metadata);
    let candidate_uuids = native_uuid_map(&candidate_metadata);

    // Keynote's ObjectUuidMap owns object/build UUIDs, but build chunks are
    // identified through the UUID pair copied into each chunk.  The native
    // package has both chunk fields pointing at the UUID of their referenced
    // build, so validate that relationship directly instead of requiring a
    // (non-native) chunk registry entry.
    for (_, build, chunk_uuid, build_uuid) in &source_chunks {
        let expected = source_uuids.get(build).copied().ok_or_else(|| {
            io::Error::other(format!("native source build UUID is missing {build}"))
        })?;
        assert_eq!(*chunk_uuid, Some(expected));
        assert_eq!(*build_uuid, Some(expected));
    }
    for (chunk, build, chunk_uuid, build_uuid) in &clone_chunks {
        let expected = candidate_uuids.get(build).copied().ok_or_else(|| {
            io::Error::other(format!("native cloned build UUID is missing {build}"))
        })?;
        assert_eq!(*chunk_uuid, Some(expected));
        assert_eq!(*build_uuid, Some(expected));
        assert!(
            !candidate_uuids.contains_key(chunk),
            "native build chunks must not be ObjectUuidMap entries"
        );
    }
    for (identifier, uuid) in &source_uuids {
        assert_eq!(
            candidate_uuids.get(identifier),
            Some(uuid),
            "old UUID changed for {identifier}",
        );
    }
    let source_uuid_values = source_uuids.values().copied().collect::<BTreeSet<_>>();
    for identifier in clone_builds.iter().copied().chain([cloned_movie]) {
        let uuid = candidate_uuids.get(&identifier).copied().ok_or_else(|| {
            io::Error::other(format!("missing clone UUID registry entry {identifier}"))
        })?;
        assert_ne!(uuid, (0, 0));
        assert!(!source_uuid_values.contains(&uuid));
    }

    for data_identifier in cloned_data {
        let mut expected = native_owner_list(&source_metadata, *data_identifier)?;
        expected.push((cloned_movie, 1));
        expected.sort_unstable();
        assert_eq!(
            native_owner_list(&candidate_metadata, *data_identifier)?,
            expected,
        );
        assert_eq!(
            native_data_record(&candidate_metadata, *data_identifier)?,
            native_data_record(&source_metadata, *data_identifier)?,
        );
    }
    Ok((
        clone_builds,
        clone_chunks
            .into_iter()
            .map(|(identifier, _, _, _)| identifier)
            .collect(),
    ))
}

fn assert_native_audio_clone_ownership(
    source: &[u8],
    candidate: &[u8],
    cloned_audio: u64,
) -> TestResult<()> {
    let source_metadata = native_metadata(source)?;
    let candidate_metadata = native_metadata(candidate)?;
    let mut expected = native_owner_list(&source_metadata, NATIVE_AUDIO_DATA)?;
    expected.push((cloned_audio, 1));
    expected.sort_unstable();
    assert_eq!(
        native_owner_list(&candidate_metadata, NATIVE_AUDIO_DATA)?,
        expected
    );
    assert_eq!(
        native_data_record(&candidate_metadata, NATIVE_AUDIO_DATA)?,
        native_data_record(&source_metadata, NATIVE_AUDIO_DATA)?,
    );
    assert_eq!(
        native_owner_list(&candidate_metadata, NATIVE_MOVIE_DATA)?,
        native_owner_list(&source_metadata, NATIVE_MOVIE_DATA)?,
    );
    assert_eq!(
        native_owner_list(&candidate_metadata, NATIVE_POSTER_DATA)?,
        native_owner_list(&source_metadata, NATIVE_POSTER_DATA)?,
    );
    Ok(())
}

fn assert_native_media_zip_unchanged(source: &[u8], candidate: &[u8]) -> TestResult<()> {
    for name in [
        NATIVE_AUDIO_MEMBER,
        NATIVE_MOVIE_MEMBER,
        NATIVE_POSTER_MEMBER,
    ] {
        assert_eq!(
            native_member_bytes(source, name)?,
            native_member_bytes(candidate, name)?
        );
    }
    Ok(())
}

fn assert_native_final_data_and_zip(
    source: &[u8],
    candidate: &[u8],
    oracle: &[u8],
    removed_data: &[u64],
    removed_members: &[&str],
    retained_members: &[&str],
) -> TestResult {
    let candidate_metadata = native_metadata(candidate)?;
    let oracle_metadata = native_metadata(oracle)?;
    // Native saves may refresh unrelated thumbnail DataInfo records. Compare
    // the selected materialized media and their ownership by stable identity.
    for identifier in [NATIVE_AUDIO_DATA, NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA] {
        let actual = candidate_metadata
            .datas
            .iter()
            .find(|data| data.identifier == identifier);
        let expected = oracle_metadata
            .datas
            .iter()
            .find(|data| data.identifier == identifier);
        assert_eq!(actual, expected, "media DataInfo {identifier}");
        if actual.is_some() {
            assert_eq!(
                native_owner_list(&candidate_metadata, identifier)?,
                native_owner_list(&oracle_metadata, identifier)?
            );
        }
    }
    for identifier in removed_data {
        assert!(native_data_record(&candidate_metadata, *identifier).is_err());
    }
    for name in removed_members {
        assert_no_member(candidate, name)?;
        assert_no_member(oracle, name)?;
    }
    for name in retained_members {
        assert_eq!(
            native_member_bytes(candidate, name)?,
            native_member_bytes(oracle, name)?
        );
        assert_eq!(
            native_member_bytes(candidate, name)?,
            native_member_bytes(source, name)?
        );
    }
    Ok(())
}

fn assert_native_removed_graph(
    source: &[u8],
    candidate: &[u8],
    removed_media: &[u64],
) -> TestResult {
    let removed_media = removed_media.iter().copied().collect::<BTreeSet<_>>();
    let mut removed_private = BTreeSet::new();
    for identifier in &removed_media {
        removed_private.extend(native_object_graph_ids(source, *identifier)?);
    }
    let source_builds = native_build_edges(source)?
        .into_iter()
        .filter(|(_, drawable)| removed_media.contains(drawable))
        .map(|(identifier, _)| identifier)
        .collect::<BTreeSet<_>>();
    let source_chunks = native_chunk_edges(source)?
        .into_iter()
        .filter(|(_, build, _, _)| source_builds.contains(build))
        .map(|(identifier, _, _, _)| identifier)
        .collect::<BTreeSet<_>>();
    assert!(
        native_build_edges(candidate)?
            .iter()
            .all(|(identifier, drawable)| {
                !removed_media.contains(drawable) && !source_builds.contains(identifier)
            })
    );
    assert!(
        native_chunk_edges(candidate)?
            .iter()
            .all(|(identifier, build, _, _)| {
                !source_builds.contains(build) && !source_chunks.contains(identifier)
            })
    );

    let source_metadata = native_metadata(source)?;
    let candidate_metadata = native_metadata(candidate)?;
    let source_uuids = native_uuid_map(&source_metadata);
    let candidate_uuids = native_uuid_map(&candidate_metadata);
    for (identifier, uuid) in source_uuids {
        if removed_private.contains(&identifier)
            || source_builds.contains(&identifier)
            || source_chunks.contains(&identifier)
        {
            assert!(!candidate_uuids.contains_key(&identifier));
        } else {
            assert_eq!(candidate_uuids.get(&identifier), Some(&uuid));
        }
    }
    Ok(())
}

fn assert_native_inventory(
    package: &Package,
    total: usize,
    audio: usize,
    video: usize,
) -> TestResult {
    let slide = package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("native lifecycle package has no first slide"))?;
    assert_eq!(slide.movies().len(), total);
    assert_eq!(slide.audio().count(), audio);
    assert_eq!(slide.video_movies().count(), video);
    Ok(())
}

#[test]
fn native_lifecycle_oracles_match_the_public_media_projection() -> TestResult {
    assert_eq!(movie_count(&Package::from_bytes(NATIVE_SOURCE)?)?, 4);
    assert_eq!(
        movie_count(&Package::from_bytes(NATIVE_DUPLICATE_MOVIE)?)?,
        5
    );
    assert_eq!(movie_count(&Package::from_bytes(NATIVE_REMOVE_MOVIE)?)?, 2);
    assert_eq!(
        movie_count(&Package::from_bytes(NATIVE_DUPLICATE_AUDIO)?)?,
        5
    );
    assert_eq!(movie_count(&Package::from_bytes(NATIVE_REMOVE_AUDIO)?)?, 2);

    assert_eq!(
        native_media_ids(NATIVE_DUPLICATE_MOVIE)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_MOVIE_A,
            NATIVE_MOVIE_B,
            2_653_696,
        ]
    );
    assert_eq!(
        native_build_edges(NATIVE_DUPLICATE_MOVIE)?
            .into_iter()
            .find(|(_, drawable)| *drawable == 2_653_696),
        Some((2_653_697, 2_653_696))
    );
    assert_eq!(
        native_media_ids(NATIVE_DUPLICATE_AUDIO)?,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_MOVIE_A,
            NATIVE_MOVIE_B,
            2_653_700,
        ]
    );
    assert_eq!(
        native_build_edges(NATIVE_DUPLICATE_AUDIO)?
            .into_iter()
            .find(|(_, drawable)| *drawable == 2_653_700),
        Some((2_653_701, 2_653_700))
    );
    assert!(
        native_chunk_edges(NATIVE_DUPLICATE_AUDIO)?
            .iter()
            .any(|(identifier, build, _, _)| *identifier == 2_653_703 && *build == 2_653_701)
    );

    let duplicate_movie = Package::from_bytes(NATIVE_DUPLICATE_MOVIE)?;
    assert_member_unchanged(
        NATIVE_SOURCE,
        &exact_bytes(&duplicate_movie)?,
        "Data/keynote-selfauthored-coral-mjpeg-9085.mov",
    )?;
    assert_member_unchanged(
        NATIVE_SOURCE,
        &exact_bytes(&duplicate_movie)?,
        "Data/posterImage-9086.png",
    )?;

    let remove_movie = Package::from_bytes(NATIVE_REMOVE_MOVIE)?;
    assert_no_member(
        &exact_bytes(&remove_movie)?,
        "Data/keynote-selfauthored-coral-mjpeg-9085.mov",
    )?;
    assert_no_member(&exact_bytes(&remove_movie)?, "Data/posterImage-9086.png")?;

    let duplicate_audio = Package::from_bytes(NATIVE_DUPLICATE_AUDIO)?;
    assert_member_unchanged(
        NATIVE_SOURCE,
        &exact_bytes(&duplicate_audio)?,
        "Data/keynote-coral-9075.wav",
    )?;
    let remove_audio = Package::from_bytes(NATIVE_REMOVE_AUDIO)?;
    assert_no_member(&exact_bytes(&remove_audio)?, "Data/keynote-coral-9075.wav")?;
    Ok(())
}

#[test]
fn duplicate_file_movie_is_source_bound_and_shares_materialized_assets() -> TestResult {
    let source = lifecycle_source()?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert_eq!(movie_count(&package)?, 3);

    let commit = package.duplicate_slide_media(
        SlideSelector::name("Media replacement"),
        MovieSelector::index(0),
    )?;
    let candidate = commit.package();
    assert_eq!(movie_count(candidate)?, 4);
    assert_eq!(
        exact_bytes(&package)?,
        before,
        "source snapshot was mutated"
    );
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/movie.mov")?;
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/poster.png")?;
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/audio.m4a")?;
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/sentinel.bin")?;
    assert_eq!(
        fixture::movie_payload_from_package(&source, fixture::MOVIES[0])?,
        fixture::movie_payload_from_package(&exact_bytes(candidate)?, fixture::MOVIES[0])?,
        "source movie payload and opaque extensions changed",
    );
    export_if_requested(candidate, "focused-duplicate-movie.key")?;

    let restored = candidate.apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn duplicate_audio_is_source_bound_and_keeps_audio_bytes_shared() -> TestResult {
    let source = lifecycle_source()?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;

    let commit = package.duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(2))?;
    let candidate = commit.package();
    assert_eq!(movie_count(candidate)?, 4);
    assert_eq!(
        exact_bytes(&package)?,
        before,
        "source snapshot was mutated"
    );
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/audio.m4a")?;
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/movie.mov")?;
    assert_member_unchanged(&source, &exact_bytes(candidate)?, "Data/poster.png")?;
    export_if_requested(candidate, "focused-duplicate-audio.key")?;

    let restored = candidate.apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn removing_shared_movie_reclaims_data_only_at_the_final_owner() -> TestResult {
    let source = lifecycle_source()?;
    let package = Package::from_bytes(&source)?;

    let first = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let after_first = exact_bytes(first.package())?;
    assert_eq!(movie_count(first.package())?, 2);
    assert_member_unchanged(&source, &after_first, "Data/movie.mov")?;
    assert_member_unchanged(&source, &after_first, "Data/poster.png")?;
    assert_member_unchanged(&source, &after_first, "Data/audio.m4a")?;
    assert_eq!(
        fixture::movie_payload_from_package(&source, fixture::MOVIES[1])?,
        fixture::movie_payload_from_package(&after_first, fixture::MOVIES[1])?,
        "surviving movie payload and opaque extensions changed",
    );
    export_if_requested(first.package(), "focused-remove-movie-shared.key")?;

    let second = first
        .package()
        .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let after_second = exact_bytes(second.package())?;
    assert_eq!(movie_count(second.package())?, 1);
    assert_no_member(&after_second, "Data/movie.mov")?;
    assert_no_member(&after_second, "Data/poster.png")?;
    assert_member_unchanged(&source, &after_second, "Data/audio.m4a")?;
    assert_member_unchanged(&source, &after_second, "Data/sentinel.bin")?;
    export_if_requested(second.package(), "focused-remove-movie-final.key")?;

    let restored_first = second
        .package()
        .apply_slide_media_lifecycle(&second.patch().inverse())?;
    let restored_source = restored_first
        .package()
        .apply_slide_media_lifecycle(&first.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn removing_shared_audio_reclaims_audio_data_at_the_final_owner() -> TestResult {
    let source = lifecycle_source()?;
    let package = Package::from_bytes(&source)?;

    let first = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(2))?;
    let after_first = exact_bytes(first.package())?;
    assert_eq!(movie_count(first.package())?, 2);
    assert_no_member(&after_first, "Data/audio.m4a")?;
    assert_member_unchanged(&source, &after_first, "Data/movie.mov")?;
    assert_member_unchanged(&source, &after_first, "Data/poster.png")?;

    let restored = first
        .package()
        .apply_slide_media_lifecycle(&first.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn lifecycle_patch_rejects_a_foreign_exact_source_atomically() -> TestResult {
    let source = lifecycle_source()?;
    let package = Package::from_bytes(&source)?;
    let commit = package.duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;

    let foreign_source =
        with_data_metadata_map(&fixture::synthetic_package_with_shared_image_poster_owner()?)?;
    let foreign = Package::from_bytes(&foreign_source)?;
    let foreign_before = exact_bytes(&foreign)?;
    assert!(foreign.apply_slide_media_lifecycle(commit.patch()).is_err());
    assert_eq!(exact_bytes(&foreign)?, foreign_before);
    Ok(())
}

#[test]
fn lifecycle_selectors_fail_atomically_for_missing_slide_or_media() -> TestResult {
    let package = package()?;
    let before = exact_bytes(&package)?;

    assert!(
        package
            .duplicate_slide_media(SlideSelector::index(9), MovieSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(9))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn duplicate_movie_reference_is_refused_without_mutating_source() -> TestResult {
    let source = lifecycle_source()?;
    let payload = fixture::movie_payload_from_package(&source, fixture::MOVIES[0])?;
    let mut malformed = payload.clone();
    append_length_delimited_field(
        &mut malformed,
        14,
        &tsp::DataReference {
            identifier: fixture::CONTENT_DATA,
        }
        .encode_to_vec(),
    )?;
    let hostile = fixture::with_movie_payload(&source, fixture::MOVIES[0], malformed)?;
    let package = Package::from_bytes(&hostile)?;
    assert!(
        package
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, hostile);
    Ok(())
}

#[test]
fn remove_reclaims_repeated_owner_counts_without_premature_shared_data_gc() -> TestResult {
    let source = with_repeated_content_owner(&lifecycle_source()?, 2)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let commit = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let candidate = exact_bytes(commit.package())?;

    assert_eq!(movie_count(commit.package())?, 2);
    assert_member_unchanged(&source, &candidate, "Data/movie.mov")?;
    assert_member_unchanged(&source, &candidate, "Data/poster.png")?;
    assert_member_unchanged(&source, &candidate, "Data/audio.m4a")?;

    let metadata = tsp::PackageMetadata::decode(fixture::metadata_stream(&candidate)?.as_slice())?;
    let component = metadata
        .components
        .iter()
        .find(|component| component.identifier == fixture::DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing repeated-owner candidate component"))?;
    let content = component
        .data_references
        .iter()
        .find(|reference| reference.data_identifier == fixture::CONTENT_DATA)
        .ok_or_else(|| io::Error::other("shared content DataInfo was reclaimed early"))?;
    assert_eq!(
        content.object_reference_list,
        vec![tsp::component_data_reference::ObjectReference {
            object_identifier: fixture::MOVIES[1],
            count: 1,
        }]
    );
    assert!(
        metadata
            .datas
            .iter()
            .any(|data| data.identifier == fixture::CONTENT_DATA)
    );

    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn removal_refuses_surviving_reference_to_selected_private_child_atomically() -> TestResult {
    let source = with_surviving_private_child_reference(&lifecycle_source()?)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn removal_refuses_slide_header_missing_selected_aggregate_atomically() -> TestResult {
    let source = with_slide_header_object_references(&lifecycle_source()?, |references| {
        let position = references
            .iter()
            .position(|identifier| *identifier == fixture::MOVIES[0])
            .expect("fixture movie must be a slide aggregate reference");
        references.remove(position);
    })?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn removal_refuses_duplicate_selected_slide_header_aggregate_atomically() -> TestResult {
    let source = with_slide_header_object_references(&lifecycle_source()?, |references| {
        references.push(fixture::MOVIES[0]);
    })?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn removal_refuses_partial_or_wrong_kind_slide_field_info_atomically() -> TestResult {
    let source = with_slide_header_field_info(
        &lifecycle_source()?,
        vec![fixture::MOVIES[0]],
        FieldType::Value,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn final_shared_poster_owner_prevents_poster_gc_without_blocking_movie_remove() -> TestResult {
    let source =
        with_data_metadata_map(&fixture::synthetic_package_with_shared_image_poster_owner()?)?;
    let package = Package::from_bytes(&source)?;
    let first = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let before_final = exact_bytes(first.package())?;
    assert_member_unchanged(&source, &before_final, "Data/poster.png")?;

    let final_movie = first
        .package()
        .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let after_final = exact_bytes(final_movie.package())?;
    assert_no_member(&after_final, "Data/movie.mov")?;
    assert_member_unchanged(&source, &after_final, "Data/poster.png")?;
    assert_member_unchanged(&source, &after_final, "Data/audio.m4a")?;
    assert_ne!(before_final, after_final);
    Ok(())
}

#[test]
fn native_focused_movie_duplicate_clones_graph_and_preserves_source() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let source_media = native_media_ids(&source)?;
    assert_eq!(
        source_media,
        vec![
            NATIVE_AUDIO_A,
            NATIVE_AUDIO_B,
            NATIVE_MOVIE_A,
            NATIVE_MOVIE_B
        ]
    );

    let commit =
        source_package.duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(2))?;
    let candidate = exact_bytes(commit.package())?;
    assert_native_inventory(commit.package(), 5, 2, 3)?;
    assert_native_media_zip_unchanged(&source, &candidate)?;
    assert_eq!(
        exact_bytes(&source_package)?,
        source,
        "source snapshot was mutated"
    );
    let cloned_movie = native_candidate_media_id(&candidate, &source_media)?;
    assert_native_movie_clone_graph(
        &source,
        &candidate,
        NATIVE_MOVIE_A,
        cloned_movie,
        &[NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA],
    )?;
    let candidate_package = Package::from_bytes(&candidate)?;
    for index in [2, 3, 4] {
        assert_eq!(
            candidate_package.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Content,
            )?,
            native_member_bytes(&source, NATIVE_MOVIE_MEMBER)?,
        );
        assert_eq!(
            candidate_package.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Poster,
            )?,
            native_member_bytes(&source, NATIVE_POSTER_MEMBER)?,
        );
    }
    export_if_requested(commit.package(), "focused-native-duplicate-movie.key")?;

    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn native_focused_movie_removal_reclaims_only_final_owner() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let first =
        source_package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(2))?;
    let first_bytes = exact_bytes(first.package())?;
    assert_native_inventory(first.package(), 3, 2, 1)?;
    assert_eq!(
        native_media_ids(&first_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_MOVIE_MEMBER)?,
        native_member_bytes(&source, NATIVE_MOVIE_MEMBER)?
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_POSTER_MEMBER)?,
        native_member_bytes(&source, NATIVE_POSTER_MEMBER)?
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_AUDIO_MEMBER)?,
        native_member_bytes(&source, NATIVE_AUDIO_MEMBER)?
    );
    assert_eq!(
        native_movie_graph(&first_bytes, NATIVE_MOVIE_B)?,
        native_movie_graph(&source, NATIVE_MOVIE_B)?
    );
    assert_eq!(
        native_caption_text(&first_bytes, NATIVE_MOVIE_B)?,
        native_caption_text(&source, NATIVE_MOVIE_B)?
    );
    let first_metadata = native_metadata(&first_bytes)?;
    assert_eq!(
        native_owner_list(&first_metadata, NATIVE_MOVIE_DATA)?,
        vec![(NATIVE_MOVIE_B, 1)]
    );
    assert_eq!(
        native_owner_list(&first_metadata, NATIVE_POSTER_DATA)?,
        vec![(NATIVE_MOVIE_B, 1)]
    );
    assert!(native_data_record(&first_metadata, NATIVE_MOVIE_DATA).is_ok());
    assert!(native_data_record(&first_metadata, NATIVE_POSTER_DATA).is_ok());
    export_if_requested(first.package(), "focused-native-remove-movie-shared.key")?;

    let second = first
        .package()
        .remove_slide_media(SlideSelector::index(0), MovieSelector::index(2))?;
    let final_bytes = exact_bytes(second.package())?;
    assert_native_inventory(second.package(), 2, 2, 0)?;
    assert_eq!(
        native_media_ids(&final_bytes)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B]
    );
    assert_native_final_data_and_zip(
        &source,
        &final_bytes,
        NATIVE_REMOVE_MOVIE,
        &[NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA],
        &[NATIVE_MOVIE_MEMBER, NATIVE_POSTER_MEMBER],
        &[NATIVE_AUDIO_MEMBER],
    )?;
    assert_native_removed_graph(&source, &final_bytes, &[NATIVE_MOVIE_A, NATIVE_MOVIE_B])?;
    let final_metadata = native_metadata(&final_bytes)?;
    let source_metadata = native_metadata(&source)?;
    let retained = source_metadata
        .datas
        .into_iter()
        .filter(|data| ![NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA].contains(&data.identifier))
        .collect::<Vec<_>>();
    assert_eq!(
        final_metadata.datas, retained,
        "unrelated source DataInfo records must be preserved exactly"
    );
    for identifier in [NATIVE_AUDIO_A, NATIVE_AUDIO_B] {
        assert_eq!(
            native_uuid_map(&final_metadata).get(&identifier),
            native_uuid_map(&native_metadata(&source)?).get(&identifier),
        );
    }
    export_if_requested(second.package(), "focused-native-remove-movie-final.key")?;

    let restored_first = second
        .package()
        .apply_slide_media_lifecycle(&second.patch().inverse())?;
    let restored_source = restored_first
        .package()
        .apply_slide_media_lifecycle(&first.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn native_focused_audio_duplicate_clones_private_graph_and_shared_data() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let source_media = native_media_ids(&source)?;
    let commit =
        source_package.duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let candidate = exact_bytes(commit.package())?;
    assert_native_inventory(commit.package(), 5, 3, 2)?;
    assert_native_media_zip_unchanged(&source, &candidate)?;
    assert_eq!(
        exact_bytes(&source_package)?,
        source,
        "source snapshot was mutated"
    );
    let cloned_audio = native_candidate_media_id(&candidate, &source_media)?;
    assert_native_movie_clone_graph(&source, &candidate, NATIVE_AUDIO_A, cloned_audio, &[])?;
    assert_native_audio_clone_ownership(&source, &candidate, cloned_audio)?;
    let candidate_package = Package::from_bytes(&candidate)?;
    for index in [0, 1, 4] {
        assert_eq!(
            candidate_package.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Content,
            )?,
            native_member_bytes(&source, NATIVE_AUDIO_MEMBER)?,
        );
    }
    for index in [2, 3] {
        assert_eq!(
            candidate_package.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                MediaPart::Content,
            )?,
            native_member_bytes(&source, NATIVE_MOVIE_MEMBER)?,
        );
    }
    assert_eq!(
        native_caption_text(&source, NATIVE_AUDIO_A)?,
        native_caption_text(&candidate, cloned_audio)?
    );
    export_if_requested(commit.package(), "focused-native-duplicate-audio.key")?;

    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn native_focused_audio_removal_reclaims_audio_at_final_owner() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let first =
        source_package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let first_bytes = exact_bytes(first.package())?;
    assert_native_inventory(first.package(), 3, 1, 2)?;
    assert_eq!(
        native_media_ids(&first_bytes)?,
        vec![NATIVE_AUDIO_B, NATIVE_MOVIE_A, NATIVE_MOVIE_B]
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_AUDIO_MEMBER)?,
        native_member_bytes(&source, NATIVE_AUDIO_MEMBER)?
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_MOVIE_MEMBER)?,
        native_member_bytes(&source, NATIVE_MOVIE_MEMBER)?
    );
    assert_eq!(
        native_member_bytes(&first_bytes, NATIVE_POSTER_MEMBER)?,
        native_member_bytes(&source, NATIVE_POSTER_MEMBER)?
    );
    let first_metadata = native_metadata(&first_bytes)?;
    assert_eq!(
        native_owner_list(&first_metadata, NATIVE_AUDIO_DATA)?,
        vec![(NATIVE_AUDIO_B, 1)]
    );
    assert!(native_data_record(&first_metadata, NATIVE_AUDIO_DATA).is_ok());
    export_if_requested(first.package(), "focused-native-remove-audio-shared.key")?;

    let second = first
        .package()
        .remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let final_bytes = exact_bytes(second.package())?;
    assert_native_inventory(second.package(), 2, 0, 2)?;
    assert_eq!(
        native_media_ids(&final_bytes)?,
        vec![NATIVE_MOVIE_A, NATIVE_MOVIE_B]
    );
    assert_native_final_data_and_zip(
        &source,
        &final_bytes,
        NATIVE_REMOVE_AUDIO,
        &[NATIVE_AUDIO_DATA],
        &[NATIVE_AUDIO_MEMBER],
        &[NATIVE_MOVIE_MEMBER, NATIVE_POSTER_MEMBER],
    )?;
    assert_native_removed_graph(&source, &final_bytes, &[NATIVE_AUDIO_A, NATIVE_AUDIO_B])?;
    let final_metadata = native_metadata(&final_bytes)?;
    let oracle_metadata = native_metadata(NATIVE_REMOVE_AUDIO)?;
    assert_eq!(final_metadata.datas, oracle_metadata.datas);
    export_if_requested(second.package(), "focused-native-remove-audio-final.key")?;

    let restored_first = second
        .package()
        .apply_slide_media_lifecycle(&second.patch().inverse())?;
    let restored_source = restored_first
        .package()
        .apply_slide_media_lifecycle(&first.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn duplicate_build_chunk_uuid_fields_refuse_lifecycle_atomically() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let (_, payload) = native_selected_chunk_for_movie(&source, NATIVE_MOVIE_A)?;
    let mut malformed = payload;
    let contradictory_uuid = tsp::Uuid {
        lower: 0xfeed_face_dead_beef,
        upper: 0x0123_4567_89ab_cdef,
    }
    .encode_to_vec();
    append_length_delimited_field(&mut malformed, 8, &contradictory_uuid)?;
    let hostile = replace_native_component_object_payload(
        &source,
        native_selected_chunk_for_movie(&source, NATIVE_MOVIE_A)?.0,
        NATIVE_BUILD_CHUNK_MESSAGE_TYPE,
        malformed,
    )?;
    let package = Package::from_bytes(&hostile)?;
    let before = exact_bytes(&package)?;

    assert!(
        package
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(2))
            .is_err()
    );
    assert!(
        package
            .remove_slide_media(SlideSelector::index(0), MovieSelector::index(2))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn explicit_tiny_media_byte_ceiling_rejects_lifecycle_reads_atomically() -> TestResult {
    let original = package()?;
    let large_audio = vec![17; 128 * 1024];
    let replacement = original
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Content,
        )?
        .set(&large_audio)?
        .commit()?;
    let source = exact_bytes(replacement.package())?;
    let total = physical_uncompressed_total(&source)?;
    let defaults = Limits::default();
    let tight = Limits::new(
        defaults.max_input_bytes(),
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        total,
        defaults.max_iwa_stream_bytes(),
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(tight, SemanticLimits::default()),
    )?;
    let before = exact_bytes(&package)?;
    let error = package
        .remove_slide_media(SlideSelector::index(0), MovieSelector::index(2))
        .err()
        .ok_or_else(|| io::Error::other("tiny media-byte ceiling must reject lifecycle reads"))?;
    match error {
        SlideMediaLifecycleError::LimitExceeded {
            kind: SlideMediaLifecycleLimitKind::MediaBytes,
            observed,
            maximum,
        } => assert!(observed > maximum),
        other => panic!("expected a media-byte limit, got {other:?}"),
    }
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_saved_candidates_are_read_back_without_rewriting_them() -> TestResult {
    let Some(directory) = env::var_os("LITCHI_KEYNOTE_MEDIA_LIFECYCLE_NATIVE_SAVED_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    let read_candidate = |name: &str| -> TestResult<Vec<u8>> {
        let bytes = fs::read(directory.join(name))?;
        Package::from_bytes(&bytes)?;
        Ok(bytes)
    };

    let source_package = Package::from_bytes(NATIVE_SOURCE)?;
    let source = exact_bytes(&source_package)?;
    let source_media = native_media_ids(&source)?;

    let duplicate_movie = read_candidate("focused-native-duplicate-movie.key")?;
    let duplicate_movie_package = Package::from_bytes(&duplicate_movie)?;
    assert_native_inventory(&duplicate_movie_package, 5, 2, 3)?;
    assert_native_media_zip_unchanged(&source, &duplicate_movie)?;
    let duplicate_movie_id = native_candidate_media_id(&duplicate_movie, &source_media)?;
    assert_native_movie_clone_graph(
        &source,
        &duplicate_movie,
        NATIVE_MOVIE_A,
        duplicate_movie_id,
        &[NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA],
    )?;

    let duplicate_audio = read_candidate("focused-native-duplicate-audio.key")?;
    let duplicate_audio_package = Package::from_bytes(&duplicate_audio)?;
    assert_native_inventory(&duplicate_audio_package, 5, 3, 2)?;
    assert_native_media_zip_unchanged(&source, &duplicate_audio)?;
    let duplicate_audio_id = native_candidate_media_id(&duplicate_audio, &source_media)?;
    assert_native_movie_clone_graph(
        &source,
        &duplicate_audio,
        NATIVE_AUDIO_A,
        duplicate_audio_id,
        &[],
    )?;
    assert_native_audio_clone_ownership(&source, &duplicate_audio, duplicate_audio_id)?;

    let shared_movie = read_candidate("focused-native-remove-movie-shared.key")?;
    let shared_movie_package = Package::from_bytes(&shared_movie)?;
    assert_native_inventory(&shared_movie_package, 3, 2, 1)?;
    assert_eq!(
        native_media_ids(&shared_movie)?,
        vec![NATIVE_AUDIO_A, NATIVE_AUDIO_B, NATIVE_MOVIE_B]
    );
    let shared_movie_metadata = native_metadata(&shared_movie)?;
    assert_eq!(
        native_owner_list(&shared_movie_metadata, NATIVE_MOVIE_DATA)?,
        vec![(NATIVE_MOVIE_B, 1)]
    );
    assert_eq!(
        native_owner_list(&shared_movie_metadata, NATIVE_POSTER_DATA)?,
        vec![(NATIVE_MOVIE_B, 1)]
    );
    assert_eq!(
        native_caption_text(&shared_movie, NATIVE_MOVIE_B)?,
        native_caption_text(&source, NATIVE_MOVIE_B)?
    );

    let final_movie = read_candidate("focused-native-remove-movie-final.key")?;
    let final_movie_package = Package::from_bytes(&final_movie)?;
    assert_native_inventory(&final_movie_package, 2, 2, 0)?;
    assert_native_final_data_and_zip(
        &source,
        &final_movie,
        NATIVE_REMOVE_MOVIE,
        &[NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA],
        &[NATIVE_MOVIE_MEMBER, NATIVE_POSTER_MEMBER],
        &[NATIVE_AUDIO_MEMBER],
    )?;
    assert_native_removed_graph(&source, &final_movie, &[NATIVE_MOVIE_A, NATIVE_MOVIE_B])?;

    let shared_audio = read_candidate("focused-native-remove-audio-shared.key")?;
    let shared_audio_package = Package::from_bytes(&shared_audio)?;
    assert_native_inventory(&shared_audio_package, 3, 1, 2)?;
    assert_eq!(
        native_media_ids(&shared_audio)?,
        vec![NATIVE_AUDIO_B, NATIVE_MOVIE_A, NATIVE_MOVIE_B]
    );
    let shared_audio_metadata = native_metadata(&shared_audio)?;
    assert_eq!(
        native_owner_list(&shared_audio_metadata, NATIVE_AUDIO_DATA)?,
        vec![(NATIVE_AUDIO_B, 1)]
    );

    let final_audio = read_candidate("focused-native-remove-audio-final.key")?;
    let final_audio_package = Package::from_bytes(&final_audio)?;
    assert_native_inventory(&final_audio_package, 2, 0, 2)?;
    assert_native_final_data_and_zip(
        &source,
        &final_audio,
        NATIVE_REMOVE_AUDIO,
        &[NATIVE_AUDIO_DATA],
        &[NATIVE_AUDIO_MEMBER],
        &[NATIVE_MOVIE_MEMBER, NATIVE_POSTER_MEMBER],
    )?;
    assert_native_removed_graph(&source, &final_audio, &[NATIVE_AUDIO_A, NATIVE_AUDIO_B])?;
    Ok(())
}

#[test]
fn lifecycle_preserves_selector_order_with_unsupported_media_siblings() -> TestResult {
    for live_video in [false, true] {
        let original = lifecycle_source()?;
        let mut movie = tsd::MovieArchive::decode(
            fixture::movie_payload_from_package(&original, fixture::MOVIES[0])?.as_slice(),
        )?;
        movie.is_live_video = live_video.then_some(true);
        movie.flags = (!live_video).then_some(1);
        let opaque_sibling = movie.encode_to_vec();
        let source =
            fixture::with_movie_payload(&original, fixture::MOVIES[0], opaque_sibling.clone())?;
        let package = Package::from_bytes(&source)?;
        let duplicate =
            package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(1))?;
        assert_eq!(duplicate.diagnostics().source_media_count(), 3);
        assert_eq!(duplicate.diagnostics().target_media_count(), 4);
        assert_eq!(movie_count(duplicate.package())?, 4);
        assert_eq!(
            fixture::movie_payload_from_package(
                &exact_bytes(duplicate.package())?,
                fixture::MOVIES[0]
            )?,
            opaque_sibling,
        );
        let restored = duplicate
            .package()
            .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source);
        let removed =
            package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(1))?;
        assert_eq!(removed.diagnostics().source_media_count(), 3);
        assert_eq!(removed.diagnostics().target_media_count(), 2);
        assert_eq!(movie_count(removed.package())?, 2);
        assert_eq!(
            fixture::movie_payload_from_package(
                &exact_bytes(removed.package())?,
                fixture::MOVIES[0]
            )?,
            opaque_sibling,
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn typed_lifecycle_kind_mismatch_precedes_metadata_rewrite() -> TestResult {
    let source = lifecycle_source()?;
    let mut metadata = tsp::PackageMetadata::decode(fixture::metadata_stream(&source)?.as_slice())?;
    metadata.last_object_identifier = 0;
    let source = fixture::replace_metadata_payload(&source, metadata.encode_to_vec())?;
    let package = Package::from_bytes(&source)?;
    for error in [
        package
            .duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))
            .unwrap_err(),
        package
            .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))
            .unwrap_err(),
        package
            .duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap_err(),
        package
            .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(0))
            .unwrap_err(),
    ] {
        assert!(
            matches!(error, SlideMediaLifecycleError::KindMismatch { expected, actual } if expected != actual)
        );
    }
    assert!(
        package
            .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(2))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
#[allow(deprecated)]
fn lifecycle_invalidates_materialized_node_cache_and_restores_exact_source() -> TestResult {
    let (_, archive) = native_component_containing_object(NATIVE_SOURCE, NATIVE_SLIDE_NODE)?;
    let mut node =
        kn::SlideNodeArchive::decode(native_object_message(&archive, NATIVE_SLIDE_NODE, 4)?)?;
    node.build_event_count = Some(99);
    node.build_event_count_cache_version = Some(2);
    node.build_event_count_is_up_to_date = Some(true);
    node.has_explicit_builds = Some(true);
    node.has_explicit_builds_cache_version = Some(2);
    node.has_explicit_builds_is_up_to_date = Some(true);
    let mut payload = node.encode_to_vec();
    append_length_delimited_field(&mut payload, 199, b"opaque node witness")?;
    let source =
        replace_native_component_object_payload(NATIVE_SOURCE, NATIVE_SLIDE_NODE, 4, payload)?;
    let package = Package::from_bytes(&source)?;
    for duplicate in [true, false] {
        let commit = if duplicate {
            package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?
        } else {
            package.remove_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?
        };
        let candidate = exact_bytes(commit.package())?;
        let (_, archive) = native_component_containing_object(&candidate, NATIVE_SLIDE_NODE)?;
        let payload = native_object_message(&archive, NATIVE_SLIDE_NODE, 4)?;
        let cache = kn::SlideNodeArchive::decode(payload)?;
        assert_eq!(cache.build_event_count, None);
        assert_eq!(cache.has_explicit_builds, None);
        assert_eq!(cache.build_event_count_cache_version, Some(u32::MAX));
        assert_eq!(cache.has_explicit_builds_cache_version, Some(u32::MAX));
        assert_ne!(cache.build_event_count_is_up_to_date, Some(true));
        assert_ne!(cache.has_explicit_builds_is_up_to_date, Some(true));
        assert!(
            payload
                .windows(b"opaque node witness".len())
                .any(|span| span == b"opaque node witness")
        );
        let mut expected = node.clone();
        expected.build_event_count = cache.build_event_count;
        expected.has_explicit_builds = cache.has_explicit_builds;
        expected.build_event_count_cache_version = cache.build_event_count_cache_version;
        expected.has_explicit_builds_cache_version = cache.has_explicit_builds_cache_version;
        expected.build_event_count_is_up_to_date = cache.build_event_count_is_up_to_date;
        expected.has_explicit_builds_is_up_to_date = cache.has_explicit_builds_is_up_to_date;
        assert_eq!(cache, expected);
        let restored = commit
            .package()
            .apply_slide_media_lifecycle(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source);
        assert_eq!(exact_bytes(&package)?, source);
        if duplicate {
            export_if_requested(commit.package(), "focused-materialized-cache-duplicate.key")?;
        }
    }
    Ok(())
}

#[test]
fn lifecycle_keeps_already_invalidated_native_node_component_exact() -> TestResult {
    let (node_member, _) = native_component_containing_object(NATIVE_SOURCE, NATIVE_SLIDE_NODE)?;
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let commit = package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    assert_member_unchanged(NATIVE_SOURCE, &exact_bytes(commit.package())?, &node_member)?;
    Ok(())
}

#[test]
fn native_saved_materialized_cache_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_CACHE_NATIVE_SAVED_PATH") else {
        return Ok(());
    };
    let saved = fs::read(path)?;
    let package = Package::from_bytes(&saved)?;
    package.validate()?;
    assert_native_inventory(&package, 5, 2, 3)?;
    assert_native_media_zip_unchanged(NATIVE_SOURCE, &saved)?;
    let clone = native_candidate_media_id(&saved, &native_media_ids(NATIVE_SOURCE)?)?;
    assert_native_movie_clone_graph(
        NATIVE_SOURCE,
        &saved,
        NATIVE_MOVIE_A,
        clone,
        &[NATIVE_MOVIE_DATA, NATIVE_POSTER_DATA],
    )?;
    let (_, archive) = native_component_containing_object(&saved, NATIVE_SLIDE_NODE)?;
    let node =
        kn::SlideNodeArchive::decode(native_object_message(&archive, NATIVE_SLIDE_NODE, 4)?)?;
    assert_eq!(node.build_event_count, None);
    // Keynote recomputes the explicit-builds cache while saving this changed
    // candidate. The focused transaction's invalidated state is checked
    // separately before native save.
    assert_eq!(node.has_explicit_builds, Some(true));
    assert_eq!(node.build_event_count_cache_version, Some(u32::MAX));
    assert_eq!(node.has_explicit_builds_cache_version, Some(2));
    assert_eq!(exact_bytes(&package)?, saved);
    Ok(())
}

#[test]
fn duplicate_refuses_unproven_transitive_movie_header_reference() -> TestResult {
    let source = lifecycle_source()?;
    let mut archive = Archive::parse(fixture::document_stream(&source)?.as_slice())?;
    archive
        .object_mut(fixture::MOVIES[0])
        .ok_or_else(|| io::Error::other("missing selected movie"))?
        .archive_info
        .message_infos[0]
        .object_references
        .push(fixture::CAPTIONS[1]);
    let source = fixture::replace_document_archive(&source, archive)?;
    let package = Package::from_bytes(&source)?;
    assert!(
        package
            .duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}
