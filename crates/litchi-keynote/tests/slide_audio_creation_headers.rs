//! Fresh slide-audio creation preserves unknown and non-canonical slide headers.

use std::{io, time::Duration};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_keynote::{Package, SlideSelector, slide::audio::Options};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const SLIDE_IDENTIFIER: u64 = 2_652_150;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const UNKNOWN_HEADER_FIELD: u32 = 97;
const UNKNOWN_HEADER_MARKER: &[u8] = b"fresh-audio-slide-header-unknown";

fn audio() -> Vec<u8> {
    let sample_count = 80u32;
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
        data.extend_from_slice(&((sample as i16 % 20 - 10) * 400).to_le_bytes());
    }
    data
}

fn options() -> TestResult<Options> {
    Ok(Options::new(
        litchi_iwa_common::shape::geometry::Point {
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

fn push_varint_width(mut value: u64, width: usize, output: &mut Vec<u8>) {
    assert!((1..=10).contains(&width));
    for index in 0..width {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if index + 1 != width {
            byte |= 0x80;
        }
        output.push(byte);
    }
    assert_eq!(value, 0, "requested varint width is too narrow");
}

fn slide_header(source: &[u8], component_name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or_else(|| io::Error::other("native slide component is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object(SLIDE_IDENTIFIER)
        .ok_or_else(|| io::Error::other("native slide object is missing"))?;
    let object_start = usize::try_from(object.header_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(
        stream
            .as_bytes()
            .get(object_start..)
            .ok_or_else(|| io::Error::other("slide object offset is invalid"))?,
    )?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("slide header start overflows"))?;
    let header_end = header_start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("slide header end overflows"))?;
    if Some(header_end) != usize::try_from(object.data_offset).ok() {
        return Err(io::Error::other("slide header framing is inconsistent").into());
    }
    Ok(stream.as_bytes()[header_start..header_end].to_vec())
}

fn mutate_slide_header(source: &[u8]) -> TestResult<(Vec<u8>, String, Vec<u8>, Vec<u8>)> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected_component = None;
    let mut selected_data = None;
    let mut unknown_field = Vec::new();

    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        let Some(object) = archive.object(SLIDE_IDENTIFIER) else {
            continue;
        };
        if !object
            .messages
            .iter()
            .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        {
            continue;
        }

        let object_start = usize::try_from(object.header_offset)?;
        let (header_length, prefix_length) = decode_varint_from_bytes(
            stream
                .as_bytes()
                .get(object_start..)
                .ok_or_else(|| io::Error::other("slide object offset is invalid"))?,
        )?;
        let header_start = object_start
            .checked_add(prefix_length)
            .ok_or_else(|| io::Error::other("slide header start overflows"))?;
        let header_end = header_start
            .checked_add(usize::try_from(header_length)?)
            .ok_or_else(|| io::Error::other("slide header end overflows"))?;
        if Some(header_end) != usize::try_from(object.data_offset).ok() {
            return Err(io::Error::other("slide header framing is inconsistent").into());
        }
        let header = &stream.as_bytes()[header_start..header_end];
        let mut rewritten_header = Vec::with_capacity(header.len() + 64);
        let mut rewrote_identifier = false;
        for field in WireView::parse(header)?.fields() {
            if field.number() == 1 && field.wire_type() == 0 && !rewrote_identifier {
                let (identifier, consumed) = decode_varint_from_bytes(field.payload())?;
                if consumed != field.payload().len() {
                    return Err(
                        io::Error::other("slide identifier field has trailing bytes").into(),
                    );
                }
                let mut encoded_identifier = Vec::with_capacity(7);
                push_varint_width(
                    (u64::from(field.number()) << 3) | u64::from(field.wire_type()),
                    2,
                    &mut encoded_identifier,
                );
                push_varint_width(identifier, 5, &mut encoded_identifier);
                rewritten_header.extend_from_slice(&encoded_identifier);
                rewrote_identifier = true;
            } else {
                rewritten_header.extend_from_slice(field.raw());
            }
        }
        if !rewrote_identifier {
            return Err(io::Error::other("slide identifier field is missing").into());
        }

        unknown_field.reserve(16 + UNKNOWN_HEADER_MARKER.len());
        push_varint_width(
            (u64::from(UNKNOWN_HEADER_FIELD) << 3) | 2,
            3,
            &mut unknown_field,
        );
        push_varint_width(
            u64::try_from(UNKNOWN_HEADER_MARKER.len())?,
            2,
            &mut unknown_field,
        );
        unknown_field.extend_from_slice(UNKNOWN_HEADER_MARKER);
        rewritten_header.extend_from_slice(&unknown_field);

        let mut rewritten_stream = Vec::with_capacity(
            stream
                .as_bytes()
                .len()
                .saturating_add(rewritten_header.len())
                .saturating_sub(header.len()),
        );
        rewritten_stream.extend_from_slice(&stream.as_bytes()[..object_start]);
        encode_varint_into(
            &mut rewritten_stream,
            u64::try_from(rewritten_header.len())?,
        );
        rewritten_stream.extend_from_slice(&rewritten_header);
        rewritten_stream.extend_from_slice(&stream.as_bytes()[header_end..]);
        Archive::parse(&rewritten_stream)?;

        selected_component = Some(entry.name().to_owned());
        selected_data = Some(SnappyStream::compress(&rewritten_stream)?);
        break;
    }

    let component_name =
        selected_component.ok_or_else(|| io::Error::other("native slide component is missing"))?;
    let compressed = selected_data.ok_or_else(|| io::Error::other("missing rewritten slide"))?;
    let edit = EntryEdit::new(&component_name, &compressed);
    let rewritten_package = catalog.reassemble_to_bytes(&[edit], Limits::default())?;
    let header = slide_header(&rewritten_package, &component_name)?;
    if !WireView::parse(&header)?.fields().any(|field| {
        field.number() == UNKNOWN_HEADER_FIELD && field.raw() == unknown_field.as_slice()
    }) {
        return Err(io::Error::other("rewritten slide header lost its unknown field").into());
    }
    let identifier_field = WireView::parse(&header)?
        .fields()
        .find(|field| field.number() == 1 && field.wire_type() == 0)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("rewritten slide identifier field is missing"))?;
    if identifier_field.len() <= 1 {
        return Err(io::Error::other("rewritten slide identifier was canonicalized").into());
    }
    Ok((
        rewritten_package,
        component_name,
        unknown_field,
        identifier_field,
    ))
}

#[test]
fn fresh_audio_preserves_unknown_and_noncanonical_slide_header_bytes() -> TestResult {
    let (source_bytes, component_name, unknown_field, noncanonical_identifier) =
        mutate_slide_header(NATIVE_SOURCE)?;
    let source = Package::from_bytes(&source_bytes)?;
    assert_eq!(exact_bytes(&source)?, source_bytes);

    let commit = source.add_slide_audio(
        SlideSelector::index(0),
        "fresh-pcm.wav",
        &audio(),
        options()?,
    )?;
    let candidate_bytes = exact_bytes(commit.package())?;
    let candidate_header = slide_header(&candidate_bytes, &component_name)?;
    let candidate_view = WireView::parse(&candidate_header)?;
    assert!(candidate_view.fields().any(|field| {
        field.number() == UNKNOWN_HEADER_FIELD && field.raw() == unknown_field.as_slice()
    }));
    assert!(candidate_view.fields().any(|field| {
        field.number() == 1 && field.wire_type() == 0 && field.raw() == noncanonical_identifier
    }));

    let restored = commit
        .package()
        .apply_slide_audio_creation(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source_bytes);
    Ok(())
}
