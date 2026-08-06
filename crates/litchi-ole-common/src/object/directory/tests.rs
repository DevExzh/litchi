use super::codec::parse_class_id;
use super::{EntryKind, Metadata, Sid, decode};
use crate::property_set::Guid;
use litchi_cfb::DirectoryEntry;

fn entry(entry_type: u8) -> DirectoryEntry {
    DirectoryEntry {
        sid: 7,
        name: "Payload".into(),
        entry_type,
        sid_left: 0xFFFF_FFFF,
        sid_right: 9,
        sid_child: 0xFFFF_FFFF,
        clsid: "00112233-4455-6677-8899-AABBCCDDEEFF".into(),
        start_sector: 12,
        size: 42,
        is_minifat: true,
        children: Vec::new(),
    }
}

#[test]
fn decodes_stream_identity_and_links_without_payload_allocation() {
    let mut source = entry(0x02);
    source.clsid.clear();
    let metadata = decode(&source).expect("stream metadata should decode");
    assert_eq!(metadata.sid(), Sid::new(7).unwrap());
    assert_eq!(metadata.kind(), EntryKind::Stream);
    assert_eq!(metadata.links().left(), None);
    assert_eq!(metadata.links().right(), Some(Sid::new(9).unwrap()));
    assert_eq!(metadata.links().child(), None);
    assert_eq!(metadata.start_sector(), 12);
    assert_eq!(metadata.stream_size(), 42);
    assert!(metadata.uses_mini_stream());
}

#[test]
fn decodes_storage_class_id_and_rejects_stream_only_fields() {
    let mut source = entry(0x01);
    source.start_sector = 0;
    source.size = 0;
    source.is_minifat = false;
    source.sid_child = 9;
    let metadata = decode(&source).expect("storage metadata should decode");
    assert_eq!(metadata.kind(), EntryKind::Storage);
    assert_eq!(metadata.class_id(), parse_class_id(&source.clsid).unwrap());
    assert!(decode(&entry(0x01)).is_err());
}

#[test]
fn class_id_codec_preserves_cfb_byte_order() {
    let value = Guid::from_bytes([
        0x33, 0x22, 0x11, 0x00, 0x55, 0x44, 0x77, 0x66, 0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE,
        0xFF,
    ]);
    assert_eq!(
        parse_class_id("00112233-4455-6677-8899-AABBCCDDEEFF").unwrap(),
        Some(value)
    );
    assert!(parse_class_id("bad-guid").is_err());
}

#[test]
fn invalid_relationships_are_rejected_before_publication() {
    let mut source = entry(0x02);
    source.clsid.clear();
    source.sid_child = source.sid;
    assert!(decode(&source).is_err());

    let mut source = entry(0x02);
    source.clsid.clear();
    source.sid_left = 0xFFFF_FFFB;
    assert!(decode(&source).is_err());
}

#[test]
fn metadata_is_small_and_copyable() {
    let metadata = Metadata::new(
        Sid::new(1).unwrap(),
        EntryKind::Storage,
        None,
        Default::default(),
        0,
        0,
        false,
    );
    assert_eq!(metadata, metadata);
}
