use super::codec::parse_class_id;
use super::{Catalog, EntryKind, Limits, Links, Metadata, Sid, Snapshot, decode};
use crate::property_set::Guid;
use litchi_cfb::DirectoryEntry;
use std::sync::Arc;

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

fn catalog_entries() -> Vec<DirectoryEntry> {
    vec![
        DirectoryEntry {
            sid: 0,
            name: "Root Entry".into(),
            entry_type: 0x05,
            sid_left: super::NOSTREAM,
            sid_right: super::NOSTREAM,
            sid_child: 1,
            clsid: "00112233-4455-6677-8899-AABBCCDDEEFF".into(),
            start_sector: litchi_cfb::consts::ENDOFCHAIN,
            size: 0,
            is_minifat: false,
            children: Vec::new(),
        },
        DirectoryEntry {
            sid: 1,
            name: "Storage".into(),
            entry_type: 0x01,
            sid_left: super::NOSTREAM,
            sid_right: super::NOSTREAM,
            sid_child: 2,
            clsid: String::new(),
            start_sector: 0,
            size: 0,
            is_minifat: false,
            children: Vec::new(),
        },
        DirectoryEntry {
            sid: 2,
            name: "Payload".into(),
            entry_type: 0x02,
            sid_left: super::NOSTREAM,
            sid_right: super::NOSTREAM,
            sid_child: super::NOSTREAM,
            clsid: String::new(),
            start_sector: 12,
            size: 42,
            is_minifat: true,
            children: Vec::new(),
        },
        DirectoryEntry {
            sid: 3,
            name: "Future".into(),
            entry_type: 0x7F,
            sid_left: super::NOSTREAM,
            sid_right: super::NOSTREAM,
            sid_child: super::NOSTREAM,
            clsid: "producer-defined".into(),
            start_sector: 91,
            size: 7,
            is_minifat: false,
            children: vec![DirectoryEntry {
                sid: 4,
                name: "Opaque Child".into(),
                entry_type: 0x7E,
                sid_left: super::NOSTREAM,
                sid_right: super::NOSTREAM,
                sid_child: super::NOSTREAM,
                clsid: "raw".into(),
                start_sector: 0,
                size: 0,
                is_minifat: false,
                children: Vec::new(),
            }],
        },
    ]
}

fn catalog_snapshot() -> Snapshot {
    Snapshot::from_entries(catalog_entries(), Limits::default())
        .expect("directory catalog should parse")
}

#[test]
fn catalog_preserves_unknown_raw_entries_and_exact_no_op_source() {
    let entries = catalog_entries();
    let shared: Arc<[DirectoryEntry]> = Arc::from(entries.into_boxed_slice());
    let source = Snapshot::from_entries_shared(Arc::clone(&shared), Limits::default())
        .expect("shared directory catalog should parse");
    let future = source
        .catalog()
        .get(Sid::new(3).unwrap())
        .expect("opaque entry should remain addressable");
    assert!(future.is_unknown());
    assert!(future.metadata().is_none());
    assert_eq!(future.raw().clsid, "producer-defined");
    assert_eq!(future.raw().children[0].name, "Opaque Child");

    let commit = source.edit().commit().expect("no-op should commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert!(commit.patch().change().is_none());
    assert!(Arc::ptr_eq(
        &source.raw_entries_shared(),
        &commit.snapshot().raw_entries_shared()
    ));
}

#[test]
fn metadata_and_containment_edits_are_typed_and_raw_preserving() {
    let source = catalog_snapshot();
    let payload = Sid::new(2).unwrap();
    let storage = Sid::new(1).unwrap();
    let mut transaction = source.edit();
    transaction
        .set_stream_size(payload, 84)
        .expect("stream size should be editable")
        .set_class_id(storage, Some(Guid::from_bytes([0xAA; 16])))
        .expect("storage class identifier should be editable")
        .set_links(storage, Links::new(None, None, Some(payload)))
        .expect("containment links should be editable");
    let commit = transaction.commit().expect("metadata edit should commit");
    let edited = commit.snapshot();
    assert_eq!(edited.metadata(payload).unwrap().stream_size(), 84);
    assert_eq!(
        edited.metadata(storage).unwrap().class_id(),
        Some(Guid::from_bytes([0xAA; 16]))
    );
    assert_eq!(edited.links(storage).unwrap().child(), Some(payload));
    assert_eq!(edited.raw_entries()[2].name, "Payload");
    assert_eq!(
        edited.raw_entries()[1].clsid,
        super::codec::format_class_id(Guid::from_bytes([0xAA; 16]))
    );
    assert_eq!(edited.raw_entries()[3].clsid, "producer-defined");
    assert!(commit.patch().change().is_some());
    assert_eq!(
        commit.patch().change().unwrap().before().metadata(payload),
        source.metadata(payload)
    );
}

#[test]
fn directory_patch_rejects_stale_sources_and_round_trips_inverse() {
    let source = catalog_snapshot();
    let payload = Sid::new(2).unwrap();
    let mut transaction = source.edit();
    transaction
        .set_stream_size(payload, 99)
        .expect("stream size should be editable");
    let commit = transaction.commit().expect("edit should commit");

    let mut stale_entries = catalog_entries();
    stale_entries[2].name = "Other Payload".into();
    let stale = Snapshot::from_entries(stale_entries, Limits::default())
        .expect("stale catalog should still be valid");
    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
}

#[test]
fn catalog_validation_is_bounded_and_failure_atomic() {
    let entries = catalog_entries();
    assert!(
        Snapshot::parse(
            &entries,
            Limits {
                max_entries: 2,
                ..Limits::default()
            }
        )
        .is_err()
    );

    let mut duplicate = entries.clone();
    duplicate[3].sid = duplicate[2].sid;
    assert!(Catalog::parse(&duplicate, Limits::default()).is_err());

    let mut invalid = entries.clone();
    invalid[2].sid_child = invalid[2].sid;
    assert!(Catalog::parse(&invalid, Limits::default()).is_err());

    let mut invalid_storage = entries.clone();
    invalid_storage[1].size = 1;
    assert!(Catalog::parse(&invalid_storage, Limits::default()).is_err());

    let source = catalog_snapshot();
    let before = source.raw_entries().to_vec();
    let mut transaction = source.edit();
    assert!(
        transaction
            .update_metadata(Sid::new(2).unwrap(), |metadata| {
                metadata.set_stream_size(1);
                metadata.set_kind(EntryKind::Storage);
                Ok(())
            })
            .is_err()
    );
    assert!(!transaction.is_changed());
    assert!(super::codec::raw_catalog_equal(
        transaction.catalog().raw_entries(),
        &before
    ));
}
