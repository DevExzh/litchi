//! Bounded generated-name collision coverage for fresh slide movies.
//!
//! The preferred filename is valid at the package boundary but too long for
//! a generated `stem-id.ext` member.  When the first short fallback is
//! occupied by an unrelated physical member, creation must keep the bounded
//! fallback stem and choose its next suffix.

use std::{io, time::Duration};

use litchi_iwa_archive::package::EntryInsertion;
use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tsp;
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector, slide::movie::Options};
use prost::Message as _;
use sha1::{Digest as _, Sha1};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const SOURCE_MOVIE_POSITION: usize = 2;
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_MESSAGE_TYPE: u32 = 11_006;

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

fn metadata(package: &Package) -> TestResult<tsp::PackageMetadata> {
    let bytes = exact_bytes(package)?;
    let catalog = Catalog::from_bytes(&bytes)?;
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
    metadata(package)?
        .datas
        .into_iter()
        .find(|record| {
            record.digest == digest && record.materialized_length == Some(bytes.len() as u64)
        })
        .map(|record| record.identifier)
        .ok_or_else(|| io::Error::other("fresh movie DataInfo is missing").into())
}

fn append_orphan(source: &Package, name: &str, bytes: &[u8]) -> TestResult<Vec<u8>> {
    let source_bytes = exact_bytes(source)?;
    let catalog = Catalog::from_bytes(&source_bytes)?;
    Ok(catalog.reassemble_with_insertions_to_bytes(
        &[EntryInsertion::new(name, bytes)],
        Limits::default(),
    )?)
}

fn member_bytes(package: &Package, name: &str) -> TestResult<Vec<u8>> {
    let bytes = exact_bytes(package)?;
    Catalog::from_bytes(&bytes)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| io::Error::other(format!("missing package member {name}")))
        .map_err(Into::into)
}

#[test]
fn oversized_preferred_movie_uses_short_collision_suffix_and_inverse() -> TestResult {
    let source = Package::from_bytes(NATIVE_SOURCE)?;
    let (source_movie, source_poster) = source_assets(&source)?;
    let movie = fresh_movie(&source_movie);
    let poster = fresh_poster(&source_poster)?;
    let first = source.add_slide_movie(
        SlideSelector::index(0),
        "first-fresh.mov",
        &movie,
        "first-fresh.png",
        &poster,
        options()?,
    )?;
    assert_eq!(first.patch().created_data(), 2);
    let content_identifier = data_identifier(first.package(), &movie)?;
    let collision_name = format!("Data/litchi-{content_identifier}.mov");
    let source_with_orphan_bytes = append_orphan(&source, &collision_name, &movie)?;
    let source_with_orphan = Package::from_bytes(&source_with_orphan_bytes)?;

    let long_filename = format!("{}.mov", "m".repeat(4_092));
    assert_eq!(long_filename.len(), 4_096);
    let second = source_with_orphan.add_slide_movie(
        SlideSelector::index(0),
        &long_filename,
        &movie,
        "suffix-poster.png",
        &poster,
        options()?,
    )?;
    assert_eq!(second.patch().created_data(), 2);

    let short_suffix_name = format!("Data/litchi-{content_identifier}-1.mov");
    assert_eq!(member_bytes(second.package(), &short_suffix_name)?, movie);
    assert_eq!(member_bytes(second.package(), &collision_name)?, movie);

    let restored = second
        .package()
        .apply_slide_movie_creation(&second.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source_with_orphan_bytes);
    Ok(())
}
