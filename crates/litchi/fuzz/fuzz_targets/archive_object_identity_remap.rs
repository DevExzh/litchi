#![no_main]

//! Bounded fuzzing for the neutral `ArchiveObject` identity/remap primitive.
//!
//! The target deliberately works below any iWork format owner.  It builds one
//! small object with aggregate and nested object references, independent data
//! references, and two payloads.  Successful cases verify remapped identities,
//! payload replacement, source immutability, deterministic output, and
//! untouched data references.  Rejected maps, replacements, and finite
//! profiles are observed without publishing a partial object.

use std::{fmt::Debug, fmt::Display, hint::black_box};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, Limits, RawMessage};

const MAX_PAYLOAD_BYTES: usize = 256;
const SOURCE_IDENTIFIER: u64 = 17;
const FIRST_OBJECT_REFERENCE: u64 = 101;
const SECOND_OBJECT_REFERENCE: u64 = 102;
const THIRD_OBJECT_REFERENCE: u64 = 103;
const FIRST_DATA_REFERENCE: u64 = 201;
const SECOND_DATA_REFERENCE: u64 = 202;
const THIRD_DATA_REFERENCE: u64 = 203;

fuzz_target!(|data: &[u8]| {
    let source = source_object(data);
    exercise_arbitrary_rejections(&source, data);
    exercise_limits(&source, data);

    if control(data, 0) & 1 == 0 {
        exercise_successful_changed_clone(&source, data);
        exercise_successful_identity_clone(&source);
    }
});

fn source_object(data: &[u8]) -> ArchiveObject {
    let mut object = ArchiveObject::new(
        SOURCE_IDENTIFIER,
        vec![
            RawMessage {
                type_: 7,
                data: payload(data, 0, b'A'),
            },
            RawMessage {
                type_: 8,
                data: payload(data, 1, b'B'),
            },
        ],
    )
    .unwrap_or_else(|error| panic!("bounded fuzz object must build: {error}"));

    let first = object
        .archive_info
        .message_infos
        .get_mut(0)
        .expect("the synthetic object has a first message");
    first.object_references = vec![
        FIRST_OBJECT_REFERENCE,
        SECOND_OBJECT_REFERENCE,
        FIRST_OBJECT_REFERENCE,
    ];
    first.data_references = vec![FIRST_DATA_REFERENCE, SECOND_DATA_REFERENCE];
    let mut nested = FieldInfo::new(FieldPath::new(vec![1, 2, 3]));
    nested.object_references = vec![SECOND_OBJECT_REFERENCE, THIRD_OBJECT_REFERENCE];
    nested.data_references = vec![THIRD_DATA_REFERENCE, FIRST_DATA_REFERENCE];
    first.field_infos.push(nested);

    let second = object
        .archive_info
        .message_infos
        .get_mut(1)
        .expect("the synthetic object has a second message");
    second.object_references = vec![THIRD_OBJECT_REFERENCE, FIRST_OBJECT_REFERENCE];
    second.data_references = vec![SECOND_DATA_REFERENCE];
    let encoded = Archive {
        objects: vec![object],
    }
    .to_bytes()
    .unwrap_or_else(|error| panic!("bounded fuzz archive must encode: {error}"));
    let encoded = append_unknown_header(&encoded);
    Archive::parse(&encoded)
        .unwrap_or_else(|error| panic!("bounded fuzz archive must reopen: {error}"))
        .objects
        .into_iter()
        .next()
        .unwrap_or_else(|| panic!("bounded fuzz archive must retain its object"))
}

fn exercise_successful_changed_clone(source: &ArchiveObject, data: &[u8]) {
    let before = source.clone();
    let replacement = replacement_messages(source, data);
    let remap = [
        (FIRST_OBJECT_REFERENCE, 1_001),
        (SECOND_OBJECT_REFERENCE, 1_002),
        (THIRD_OBJECT_REFERENCE, 1_003),
    ];
    let new_identifier = 9_000_u64.saturating_add(u64::from(control(data, 2)));

    let cloned = match source.clone_with_identity_remap(new_identifier, &remap, &replacement) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            assert_eq!(source, &before);
            return;
        },
    };
    assert_clone_invariants(source, &cloned, new_identifier, &replacement, &remap);
    assert_eq!(source, &before);

    let repeat = source
        .clone_with_identity_remap(new_identifier, &remap, &replacement)
        .unwrap_or_else(|error| panic!("deterministic identity remap must repeat: {error}"));
    assert!(cloned.same_content_ignoring_offsets(&repeat));
    assert_eq!(source, &before);
}

fn exercise_successful_identity_clone(source: &ArchiveObject) {
    let before = source.clone();
    let cloned = match source.clone_with_identity_remap(
        SOURCE_IDENTIFIER.saturating_add(1_000),
        &[],
        &source.messages,
    ) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            assert_eq!(source, &before);
            return;
        },
    };
    assert_eq!(
        cloned.archive_info.identifier,
        Some(SOURCE_IDENTIFIER + 1_000)
    );
    assert_eq!(cloned.messages, source.messages);
    assert_eq!(
        cloned.archive_info.message_infos,
        source.archive_info.message_infos
    );
    assert_eq!(source, &before);
}

fn exercise_arbitrary_rejections(source: &ArchiveObject, data: &[u8]) {
    let before = source.clone();
    let valid_replacement = replacement_messages(source, data);

    let malformed_replacement = if control(data, 1) & 1 == 0 {
        valid_replacement[..1].to_vec()
    } else {
        vec![RawMessage {
            type_: source.messages[0].type_,
            data: vec![0; MAX_PAYLOAD_BYTES + 1],
        }]
    };
    let result = source.clone_with_identity_remap(
        8_000,
        &[(FIRST_OBJECT_REFERENCE, 1_001)],
        &malformed_replacement,
    );
    if let Err(error) = result {
        observe_error(error);
    }
    assert_eq!(source, &before);

    let duplicate_map = [
        (FIRST_OBJECT_REFERENCE, 1_001),
        (FIRST_OBJECT_REFERENCE, 1_002),
    ];
    if let Err(error) = source.clone_with_identity_remap(8_001, &duplicate_map, &valid_replacement)
    {
        observe_error(error);
    }
    assert_eq!(source, &before);

    if let Err(error) = source.clone_with_identity_remap(0, &[], &source.messages) {
        observe_error(error);
    }
    assert_eq!(source, &before);
}

fn exercise_limits(source: &ArchiveObject, data: &[u8]) {
    let replacement = replacement_messages(source, data);
    let remap = [
        (FIRST_OBJECT_REFERENCE, 1_001),
        (SECOND_OBJECT_REFERENCE, 1_002),
        (THIRD_OBJECT_REFERENCE, 1_003),
    ];
    let before = source.clone();

    let metadata_limit = Limits::default()
        .with_metadata_items(1)
        .unwrap_or_else(|error| panic!("metadata limit must construct: {error}"));
    if let Err(error) =
        source.clone_with_identity_remap_with_limits(9_100, &remap, &replacement, metadata_limit)
    {
        observe_error(error);
    }
    assert_eq!(source, &before);

    let object_limit = Limits::default()
        .with_object_bytes(1)
        .unwrap_or_else(|error| panic!("object limit must construct: {error}"));
    if let Err(error) =
        source.clone_with_identity_remap_with_limits(9_101, &remap, &replacement, object_limit)
    {
        observe_error(error);
    }
    assert_eq!(source, &before);
}

fn assert_clone_invariants(
    source: &ArchiveObject,
    cloned: &ArchiveObject,
    new_identifier: u64,
    replacement: &[RawMessage],
    remap: &[(u64, u64)],
) {
    assert_eq!(cloned.archive_info.identifier, Some(new_identifier));
    assert_eq!(cloned.messages, replacement);
    assert_eq!(cloned.messages.len(), source.messages.len());
    assert_eq!(
        cloned
            .archive_info
            .message_infos
            .iter()
            .map(|info| info.type_)
            .collect::<Vec<_>>(),
        source
            .archive_info
            .message_infos
            .iter()
            .map(|info| info.type_)
            .collect::<Vec<_>>()
    );
    for ((index, before), after) in source
        .archive_info
        .message_infos
        .iter()
        .enumerate()
        .zip(&cloned.archive_info.message_infos)
    {
        assert_eq!(
            after.length,
            u32::try_from(replacement[index].data.len()).unwrap_or(u32::MAX)
        );
        assert_eq!(after.data_references, before.data_references);
        assert_eq!(after.field_infos.len(), before.field_infos.len());
        assert_eq!(
            after.object_references,
            remap_references(&before.object_references, remap)
        );
        for (before_field, after_field) in before.field_infos.iter().zip(&after.field_infos) {
            assert_eq!(after_field.path, before_field.path);
            assert_eq!(after_field.data_references, before_field.data_references);
            assert_eq!(
                after_field.object_references,
                remap_references(&before_field.object_references, remap)
            );
        }
    }
    cloned
        .validate()
        .unwrap_or_else(|error| panic!("remapped clone must validate: {error}"));
}

fn remap_references(references: &[u64], remap: &[(u64, u64)]) -> Vec<u64> {
    references
        .iter()
        .copied()
        .map(|identifier| {
            remap
                .iter()
                .find_map(|(before, after)| (*before == identifier).then_some(*after))
                .unwrap_or(identifier)
        })
        .collect()
}

fn replacement_messages(source: &ArchiveObject, data: &[u8]) -> Vec<RawMessage> {
    source
        .messages
        .iter()
        .enumerate()
        .map(|(index, message)| RawMessage {
            type_: message.type_,
            data: replacement_payload(data, index, message.type_),
        })
        .collect()
}

fn replacement_payload(data: &[u8], index: usize, message_type: u32) -> Vec<u8> {
    let mut output = vec![b'R', b'0' + u8::try_from(index).unwrap_or_default()];
    output.extend_from_slice(&message_type.to_le_bytes());
    let start = index.min(data.len());
    output.extend(
        data.get(start..)
            .unwrap_or_default()
            .iter()
            .take(MAX_PAYLOAD_BYTES.saturating_sub(output.len()))
            .map(|byte| byte.rotate_left(1)),
    );
    output
}

fn payload(data: &[u8], offset: usize, marker: u8) -> Vec<u8> {
    let mut output = vec![marker];
    output.extend(
        data.get(offset..)
            .unwrap_or_default()
            .iter()
            .take(MAX_PAYLOAD_BYTES.saturating_sub(1).min(32))
            .copied(),
    );
    if output.len() == 1 {
        output.push(marker ^ 0xff);
    }
    output
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn append_unknown_header(source: &[u8]) -> Vec<u8> {
    let (header_length, prefix_length) = read_varint(source)
        .unwrap_or_else(|| panic!("bounded fuzz archive must have a header length"));
    let header_start = prefix_length;
    let header_end = header_start
        .checked_add(header_length)
        .unwrap_or_else(|| panic!("bounded fuzz archive header length must fit"));
    assert!(header_end <= source.len());

    let mut header = source[header_start..header_end].to_vec();
    // Field 1000, wire type 0, value 1.  The core parser retains this unknown
    // ArchiveInfo field as raw provenance while its typed projection ignores
    // it, which is the important path for this focused clone/remap primitive.
    header.extend_from_slice(&[0xb8, 0x3e, 0x01]);
    let encoded_length = encode_varint(header.len());
    let mut output =
        Vec::with_capacity(encoded_length.len() + header.len() + source.len() - header_end);
    output.extend_from_slice(&encoded_length);
    output.extend_from_slice(&header);
    output.extend_from_slice(&source[header_end..]);
    output
}

fn read_varint(source: &[u8]) -> Option<(usize, usize)> {
    let mut value = 0usize;
    let mut shift = 0usize;
    for (index, byte) in source.iter().copied().enumerate().take(10) {
        let part = usize::from(byte & 0x7f).checked_shl(u32::try_from(shift).ok()?)?;
        value = value.checked_add(part)?;
        if byte & 0x80 == 0 {
            return Some((value, index + 1));
        }
        shift = shift.checked_add(7)?;
    }
    None
}

fn encode_varint(mut value: usize) -> Vec<u8> {
    let mut output = Vec::with_capacity(10);
    loop {
        let mut byte = u8::try_from(value & 0x7f).unwrap_or_default();
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return output;
        }
    }
}
