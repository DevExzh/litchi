use super::{Kind, Link, Times};
use crate::property_set::Guid;
use std::sync::Arc;

fn append_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn append_u64(output: &mut Vec<u8>, value: u64) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn linked_wire() -> Vec<u8> {
    let mut output = Vec::new();
    append_u32(&mut output, Link::VERSION);
    append_u32(&mut output, 0x1001);
    append_u32(&mut output, 7);
    append_u32(&mut output, 0);
    append_u32(&mut output, 0);

    let relative = [0x10; 16];
    append_u32(&mut output, (relative.len() + 4) as u32);
    output.extend_from_slice(&relative);

    let absolute = [0x20; 16];
    append_u32(&mut output, (absolute.len() + 4) as u32);
    output.extend_from_slice(&absolute);

    append_u32(&mut output, u32::MAX);
    output.extend_from_slice(&[0x30; 16]);
    append_u32(&mut output, 4);
    output.extend_from_slice(&[0, 0, 0, 0]);
    append_u32(&mut output, 0xDEAD_BEEF);
    append_u64(&mut output, 11);
    append_u64(&mut output, 22);
    append_u64(&mut output, 33);
    output.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
    output
}

#[test]
fn embedded_header_round_trips_without_inventing_optional_fields() {
    let bytes = [
        0x01, 0x00, 0x00, 0x02, // version
        0, 0, 0, 0, // embedded
        0, 0, 0, 0, // update option
        0, 0, 0, 0, // reserved
    ];
    let link = Link::parse(&bytes).expect("embedded OLE link should parse");
    assert_eq!(link.kind(), Kind::Embedded);
    assert!(link.kind().is_embedded());
    assert!(!link.cache_hint());
    assert!(link.absolute_source().is_none());
    assert_eq!(link.bytes(), bytes);
    assert_eq!(link.to_bytes(), bytes);
}

#[test]
fn linked_metadata_is_typed_and_unknown_tail_is_retained() {
    let bytes = linked_wire();
    let shared: Arc<[u8]> = Arc::from(bytes.clone());
    let link = Link::parse_shared(Arc::clone(&shared)).expect("linked OLE stream should parse");
    assert!(Arc::ptr_eq(&shared, &link.bytes_shared()));
    assert_eq!(link.kind(), Kind::Linked);
    assert_eq!(link.flags(), 0x1001);
    assert!(link.cache_hint());
    assert_eq!(link.link_update_option(), 7);
    assert_eq!(
        link.relative_source().unwrap().class_id(),
        Guid::from_bytes([0x10; 16])
    );
    assert_eq!(link.source().unwrap().data(), &[]);
    assert_eq!(
        link.absolute_source().unwrap().class_id(),
        Guid::from_bytes([0x20; 16])
    );
    assert_eq!(link.class_id(), Some(Guid::from_bytes([0x30; 16])));
    assert_eq!(link.reserved_display_name(), Some(&[0, 0, 0, 0][..]));
    assert_eq!(link.reserved2(), Some(0xDEAD_BEEF));
    assert_eq!(link.times(), Some(Times::new(11, 22, 33)));
    assert_eq!(link.unknown_tail(), &[0xAA, 0xBB, 0xCC]);
    assert_eq!(link.to_bytes(), bytes);
}

#[test]
fn typed_edits_change_only_owned_fields() {
    let original = linked_wire();
    let mut link = Link::parse(&original).expect("linked OLE stream should parse");
    link.set_cache_hint(false);
    link.set_link_update_option(19);
    link.set_class_id(Guid::from_bytes([0x44; 16]))
        .expect("linked class identifier should be editable");
    link.set_times(Times::new(101, 202, 303))
        .expect("linked timestamps should be editable");
    let changed = link.to_bytes();
    assert_eq!(&changed[16..64], &original[16..64]);
    assert_eq!(&changed[80..92], &original[80..92]);
    assert_eq!(&changed[116..], &original[116..]);
    let parsed = Link::parse(&changed).expect("edited link should parse");
    assert!(!parsed.cache_hint());
    assert_eq!(parsed.link_update_option(), 19);
    assert_eq!(parsed.class_id(), Some(Guid::from_bytes([0x44; 16])));
    assert_eq!(parsed.times(), Some(Times::new(101, 202, 303)));
    assert_eq!(parsed.unknown_tail(), &[0xAA, 0xBB, 0xCC]);
}

#[test]
fn malformed_known_fields_are_rejected() {
    let mut invalid_version = vec![0; 16];
    invalid_version[0] = 1;
    assert!(Link::parse(&invalid_version).is_err());

    let invalid_reserved = [0x01, 0x00, 0x00, 0x02, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0];
    assert!(Link::parse(&invalid_reserved).is_err());

    let mut missing_absolute = vec![
        0x01, 0x00, 0x00, 0x02, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    ];
    append_u32(&mut missing_absolute, 0);
    assert!(Link::parse(&missing_absolute).is_err());

    let mut invalid_indicator = linked_wire();
    let indicator = invalid_indicator
        .windows(4)
        .position(|bytes| bytes == [0xFF; 4])
        .expect("test indicator");
    invalid_indicator[indicator] = 0;
    assert!(Link::parse(&invalid_indicator).is_err());
}
