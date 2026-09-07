//! Fresh movie creation reserves identifiers mentioned only by DataMetadataMap.
//!
//! Unmapped keys occupy the data identifier namespace even though the media
//! closure rejects them. Creation must not repair such a dangling association
//! accidentally by assigning its key to a fresh asset. Valid maps retain their
//! exact payload and root association through creation and inverse restoration.

use std::{io, time::Duration};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::tsp;
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector, slide::movie::Options};
use prost::Message as _;
use sha1::{Digest as _, Sha1};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const SOURCE_MOVIE_POSITION: usize = 2;
const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
const DATA_METADATA_MESSAGE_TYPE: u32 = 11_014;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn source_assets(source: &Package) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let movie = source
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Content,
        )?
        .to_vec();
    let poster = source
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Poster,
        )?
        .to_vec();
    if movie.is_empty() || poster.is_empty() {
        return Err(io::Error::other("native source movie assets are empty").into());
    }
    Ok((movie, poster))
}

fn options() -> TestResult<Options> {
    Ok(Options::new(
        Point { x: 321.0, y: 42.0 },
        Size {
            width: 640.0,
            height: 360.0,
        },
        Duration::from_millis(1_250),
    )?
    .with_natural_size(Size {
        width: 320.0,
        height: 180.0,
    })?)
}

fn fresh_movie(source: &[u8]) -> Vec<u8> {
    let mut movie = source.to_vec();
    movie.extend_from_slice(&16u32.to_be_bytes());
    movie.extend_from_slice(b"free");
    movie.extend_from_slice(b"litchi!!");
    movie
}

fn png_chunk(kind: &[u8; 4], payload: &[u8]) -> Vec<u8> {
    let length = u32::try_from(payload.len()).expect("test PNG chunk fits u32");
    let mut crc_input = Vec::with_capacity(kind.len() + payload.len());
    crc_input.extend_from_slice(kind);
    crc_input.extend_from_slice(payload);
    let mut crc = 0xffff_ffffu32;
    for byte in crc_input {
        crc ^= u32::from(byte);
        for _ in 0..8 {
            let mask = 0u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    let mut output = Vec::with_capacity(12 + payload.len());
    output.extend_from_slice(&length.to_be_bytes());
    output.extend_from_slice(kind);
    output.extend_from_slice(payload);
    output.extend_from_slice(&(!crc).to_be_bytes());
    output
}

fn fresh_poster(source: &[u8]) -> TestResult<Vec<u8>> {
    let type_offset = source
        .windows(4)
        .position(|window| window == b"IEND")
        .ok_or_else(|| io::Error::other("native source poster has no IEND chunk"))?;
    let chunk_start = type_offset
        .checked_sub(4)
        .ok_or_else(|| io::Error::other("native source poster has a truncated IEND chunk"))?;
    let mut poster = Vec::with_capacity(source.len() + 32);
    poster.extend_from_slice(&source[..chunk_start]);
    poster.extend_from_slice(&png_chunk(b"tEXt", b"Comment\0litchi fresh poster"));
    poster.extend_from_slice(&source[chunk_start..]);
    Ok(poster)
}

fn metadata_from_source(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing PackageMetadata member"))?;
    let decoded = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&decoded)?;
    let mut payloads = archive.objects.iter().flat_map(|object| {
        object.messages.iter().filter_map(|message| {
            (message.type_ == METADATA_MESSAGE_TYPE).then_some(message.data.as_slice())
        })
    });
    let payload = payloads
        .next()
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload"))?;
    if payloads.next().is_some() {
        return Err(io::Error::other("multiple PackageMetadata payloads").into());
    }
    Ok(tsp::PackageMetadata::decode(payload)?)
}

fn data_identifier(package: &Package, bytes: &[u8]) -> TestResult<u64> {
    let digest = Sha1::digest(bytes).to_vec();
    metadata_from_source(&exact_bytes(package)?)?
        .datas
        .into_iter()
        .find(|record| {
            record.digest == digest && record.materialized_length == Some(bytes.len() as u64)
        })
        .map(|record| record.identifier)
        .ok_or_else(|| io::Error::other("fresh movie DataInfo is missing").into())
}

fn member_bytes(source: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Catalog::from_bytes(source)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| io::Error::other(format!("missing package member {name}")))
        .map_err(Into::into)
}

fn rewrite_member(source: &[u8], name: &str, replacement: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(name, replacement)], Limits::default())?)
}

fn rewrite_archive_member(
    source: &[u8],
    name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let decoded = SnappyStream::decompress(&member_bytes(source, name)?)?.into_bytes();
    let mut archive = Archive::parse(&decoded)?;
    mutate(&mut archive)?;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    rewrite_member(source, name, &replacement)
}

fn fresh_object_identifier(source: &[u8]) -> TestResult<u64> {
    let mut maximum = 0u64;
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let decoded = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decoded)?;
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("native object has no identifier"))?;
            maximum = maximum.max(identifier);
        }
    }
    maximum
        .checked_add(1)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| io::Error::other("native object identifier space is exhausted").into())
}

fn fresh_object_identifier_avoiding(source: &[u8], avoid: &[u64]) -> TestResult<u64> {
    let mut identifier = fresh_object_identifier(source)?;
    while avoid.contains(&identifier) {
        identifier = identifier
            .checked_add(1)
            .filter(|candidate| *candidate != 0)
            .ok_or_else(|| io::Error::other("native object identifier space is exhausted"))?;
    }
    Ok(identifier)
}

fn next_identifier_avoiding(identifier: u64, avoid: &[u64]) -> TestResult<u64> {
    let mut candidate = identifier;
    while avoid.contains(&candidate) {
        candidate = candidate
            .checked_add(1)
            .filter(|value| *value != 0)
            .ok_or_else(|| io::Error::other("native object identifier space is exhausted"))?;
    }
    Ok(candidate)
}

fn varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    varint(output, u64::from(number) << 3);
    varint(output, value);
}

fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    varint(output, (u64::from(number) << 3) | 2);
    varint(
        output,
        u64::try_from(payload.len()).expect("test payload fits in a varint"),
    );
    output.extend_from_slice(payload);
}

fn map_payload(data_identifiers: &[u64], metadata_identifier: u64) -> Vec<u8> {
    let mut output = Vec::new();
    for &data_identifier in data_identifiers {
        let mut entry = Vec::new();
        varint_field(&mut entry, 1, data_identifier);
        let mut reference = Vec::new();
        varint_field(&mut reference, 1, metadata_identifier);
        bytes_field(&mut entry, 2, &reference);
        // Keep an opaque extension in the source map. The creation rewrite
        // must leave this archive object untouched while reserving its key.
        varint_field(&mut entry, 99, 1);
        bytes_field(&mut output, 1, &entry);
    }
    output
}

fn map_payload_from_source(source: &[u8], map_identifier: u64) -> TestResult<Vec<u8>> {
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let decoded = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decoded)?;
        for object in archive.objects {
            if object.archive_info.identifier != Some(map_identifier) {
                continue;
            }
            let mut payload = None;
            for message in object.messages {
                if message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE {
                    if payload.replace(message.data).is_some() {
                        return Err(io::Error::other("duplicate DataMetadataMap payload").into());
                    }
                }
            }
            return payload
                .ok_or_else(|| io::Error::other("missing DataMetadataMap payload"))
                .map_err(Into::into);
        }
    }
    Err(io::Error::other("missing DataMetadataMap object").into())
}

fn with_opaque_map(
    source: &[u8],
    map_identifier: u64,
    metadata_identifier: u64,
    payload: &[u8],
) -> TestResult<Vec<u8>> {
    let source = rewrite_archive_member(source, DOCUMENT_MEMBER, |archive| {
        archive.insert_object(ArchiveObject::new(
            map_identifier,
            vec![RawMessage {
                type_: DATA_METADATA_MAP_MESSAGE_TYPE,
                data: payload.to_vec(),
            }],
        )?)?;
        archive.insert_object(ArchiveObject::new(
            metadata_identifier,
            vec![RawMessage {
                type_: DATA_METADATA_MESSAGE_TYPE,
                data: Vec::new(),
            }],
        )?)?;
        Ok(())
    })?;
    let mut metadata = metadata_from_source(&source)?;
    metadata.data_metadata_map = Some(tsp::Reference {
        identifier: map_identifier,
        ..tsp::Reference::default()
    });
    let metadata_payload = metadata.encode_to_vec();
    rewrite_archive_member(&source, METADATA_MEMBER, |archive| {
        let mut location = None;
        for (object_index, object) in archive.objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != METADATA_MESSAGE_TYPE {
                    continue;
                }
                if location.replace((object_index, message_index)).is_some() {
                    return Err(io::Error::other("duplicate PackageMetadata payload").into());
                }
            }
        }
        let (object_index, message_index) =
            location.ok_or_else(|| io::Error::other("missing PackageMetadata payload"))?;
        archive.objects[object_index].messages[message_index].data = metadata_payload;
        Ok(())
    })
}

#[test]
fn unmapped_opaque_map_keys_reject_creation_atomically() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (source_movie, source_poster) = source_assets(&source)?;
    let movie = fresh_movie(&source_movie);
    let poster = fresh_poster(&source_poster)?;
    // Probe the identifiers the allocator would choose without a map. The
    // negative fixture puts exactly those keys in a valid map, but leaves the
    // corresponding DataInfo records absent. With map reservation, creation
    // chooses later IDs and the closure validator rejects the dangling keys;
    // without reservation it would capture both keys and incorrectly pass.
    let probe = source.add_slide_movie(
        SlideSelector::index(0),
        "map-probe.mov",
        &movie,
        "map-probe.png",
        &poster,
        options()?,
    )?;
    assert_eq!(probe.patch().created_data(), 2);
    let probe_content_identifier = data_identifier(probe.package(), &movie)?;
    let probe_poster_identifier = data_identifier(probe.package(), &poster)?;
    let source_bytes = exact_bytes(&source)?;
    let map_identifier = fresh_object_identifier_avoiding(
        &source_bytes,
        &[probe_content_identifier, probe_poster_identifier],
    )?;
    let metadata_identifier = next_identifier_avoiding(
        map_identifier
            .checked_add(1)
            .filter(|identifier| *identifier != 0)
            .ok_or_else(|| io::Error::other("metadata object identifier space is exhausted"))?,
        &[probe_content_identifier, probe_poster_identifier],
    )?;
    let opaque_map_payload = map_payload(
        &[probe_content_identifier, probe_poster_identifier],
        metadata_identifier,
    );
    let mapped_source = with_opaque_map(
        &source_bytes,
        map_identifier,
        metadata_identifier,
        &opaque_map_payload,
    )?;
    let mapped_metadata = metadata_from_source(&mapped_source)?;
    assert_eq!(
        mapped_metadata
            .data_metadata_map
            .as_ref()
            .map(|reference| reference.identifier),
        Some(map_identifier)
    );
    assert_eq!(
        map_payload_from_source(&mapped_source, map_identifier)?,
        opaque_map_payload
    );
    let mapped = Package::from_bytes(&mapped_source)?;
    let result = mapped.add_slide_movie(
        SlideSelector::index(0),
        "map-rejected.mov",
        &movie,
        "map-rejected.png",
        &poster,
        options()?,
    );
    assert!(
        matches!(
            result,
            Err(litchi_keynote::SlideMovieCreationError::InvalidSource
                | litchi_keynote::SlideMovieCreationError::Verification)
        ),
        "dangling map keys must reject creation without a resource-limit failure"
    );
    assert_eq!(exact_bytes(&mapped)?, mapped_source);
    assert_eq!(
        map_payload_from_source(&mapped_source, map_identifier)?,
        opaque_map_payload
    );
    Ok(())
}

#[test]
fn mapped_existing_assets_allow_fresh_movie_creation_and_inverse() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (source_movie, source_poster) = source_assets(&source)?;
    let existing_movie_identifier = data_identifier(&source, &source_movie)?;
    let existing_poster_identifier = data_identifier(&source, &source_poster)?;
    assert_ne!(existing_movie_identifier, existing_poster_identifier);
    let source_bytes = exact_bytes(&source)?;
    let map_identifier = fresh_object_identifier_avoiding(
        &source_bytes,
        &[existing_movie_identifier, existing_poster_identifier],
    )?;
    let metadata_identifier = next_identifier_avoiding(
        map_identifier
            .checked_add(1)
            .filter(|identifier| *identifier != 0)
            .ok_or_else(|| io::Error::other("metadata object identifier space is exhausted"))?,
        &[existing_movie_identifier, existing_poster_identifier],
    )?;
    let opaque_map_payload = map_payload(
        &[existing_movie_identifier, existing_poster_identifier],
        metadata_identifier,
    );
    let mapped_source = with_opaque_map(
        &source_bytes,
        map_identifier,
        metadata_identifier,
        &opaque_map_payload,
    )?;
    let mapped_metadata = metadata_from_source(&mapped_source)?;
    assert_eq!(
        mapped_metadata
            .data_metadata_map
            .as_ref()
            .map(|reference| reference.identifier),
        Some(map_identifier)
    );
    assert_eq!(
        map_payload_from_source(&mapped_source, map_identifier)?,
        opaque_map_payload
    );

    let movie = fresh_movie(&source_movie);
    let poster = fresh_poster(&source_poster)?;
    let mapped = Package::from_bytes(&mapped_source)?;
    let second = mapped.add_slide_movie(
        SlideSelector::index(0),
        "map-second.mov",
        &movie,
        "map-second.png",
        &poster,
        options()?,
    )?;
    assert_eq!(second.patch().created_data(), 2);
    let second_content_identifier = data_identifier(second.package(), &movie)?;
    let second_poster_identifier = data_identifier(second.package(), &poster)?;
    assert_ne!(second_content_identifier, existing_movie_identifier);
    assert_ne!(second_poster_identifier, existing_poster_identifier);
    assert_ne!(second_content_identifier, existing_poster_identifier);
    assert_ne!(second_poster_identifier, existing_movie_identifier);

    let second_source = exact_bytes(second.package())?;
    assert_eq!(
        map_payload_from_source(&second_source, map_identifier)?,
        opaque_map_payload
    );
    let second_metadata = metadata_from_source(&second_source)?;
    assert_eq!(
        second_metadata
            .data_metadata_map
            .as_ref()
            .map(|reference| reference.identifier),
        Some(map_identifier)
    );

    let restored = second
        .package()
        .apply_slide_movie_creation(&second.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, mapped_source);
    Ok(())
}
