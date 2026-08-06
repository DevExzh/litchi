use std::sync::Arc;

use super::{Blob, Error, Kind, Limits};

fn push_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn property_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    push_u32(&mut bytes, 0, 48);
    push_u32(&mut bytes, 4, 8);
    push_u32(&mut bytes, 8, 3);
    push_u32(&mut bytes, 12, 44);
    push_u32(&mut bytes, 16, 4);
    push_u32(&mut bytes, 20, 48);
    push_u32(&mut bytes, 24, 0);
    push_u32(&mut bytes, 28, 52);
    push_u32(&mut bytes, 32, 0xA5A5_5A5A);
    push_u32(&mut bytes, 36, 0);
    push_u32(&mut bytes, 40, 54);
    bytes[44..47].copy_from_slice(&[1, 2, 3]);
    bytes[47] = 0xCC;
    bytes[48..52].copy_from_slice(&[4, 5, 6, 7]);
    bytes
}

fn word_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    bytes[0..2].copy_from_slice(&27u16.to_le_bytes());
    push_u32(&mut bytes, 2, 45);
    push_u32(&mut bytes, 6, 8);
    push_u32(&mut bytes, 10, 2);
    push_u32(&mut bytes, 14, 44);
    push_u32(&mut bytes, 18, 3);
    push_u32(&mut bytes, 22, 46);
    push_u32(&mut bytes, 26, 0);
    push_u32(&mut bytes, 30, 49);
    push_u32(&mut bytes, 34, 0);
    push_u32(&mut bytes, 38, 0);
    push_u32(&mut bytes, 42, 51);
    bytes[46..48].copy_from_slice(&[9, 8]);
    bytes[48..51].copy_from_slice(&[7, 6, 5]);
    bytes[55] = 0xEE;
    bytes
}

#[test]
fn property_blob_exposes_payloads_and_preserves_unknown_bytes() {
    let source = property_blob();
    let blob = Blob::parse_property(&source).unwrap();

    assert_eq!(blob.kind(), Kind::Property);
    assert_eq!(blob.info().signature(), [1, 2, 3]);
    assert_eq!(blob.info().certificate_store(), [4, 5, 6, 7]);
    assert_eq!(blob.info().reserved_project_name(), [0, 0]);
    assert_eq!(blob.info().reserved_timestamp_url(), [0, 0]);
    assert_eq!(blob.info().reserved_timestamp_marker(), 0xA5A5_5A5A);
    assert!(blob.info().padding().is_empty());
    assert_eq!(blob.bytes()[47], 0xCC);
    assert_eq!(blob.bytes(), source);
}

#[test]
fn word_blob_uses_its_distinct_offset_base_and_keeps_padding() {
    let source: Arc<[u8]> = word_blob().into();
    let blob = Blob::parse_shared(Arc::clone(&source), Kind::Word, Limits::default()).unwrap();

    assert!(Arc::ptr_eq(&source, &blob.bytes_shared()));
    assert_eq!(blob.info().signature(), [9, 8]);
    assert_eq!(blob.info().certificate_store(), [7, 6, 5]);
    assert_eq!(blob.info().padding(), [0xEE]);
    assert_eq!(&*blob.into_bytes(), &*source);
}

#[test]
fn malformed_boundaries_reserved_fields_and_overlaps_are_rejected() {
    let mut bad_pointer = property_blob();
    push_u32(&mut bad_pointer, 4, 12);
    assert!(matches!(
        Blob::parse_property(&bad_pointer),
        Err(Error::Invalid(_))
    ));

    let mut reserved_length = property_blob();
    push_u32(&mut reserved_length, 24, 1);
    assert!(matches!(
        Blob::parse_property(&reserved_length),
        Err(Error::Invalid(_))
    ));

    let mut overlap = property_blob();
    push_u32(&mut overlap, 20, 46);
    assert!(matches!(
        Blob::parse_property(&overlap),
        Err(Error::Invalid(_))
    ));

    let truncated = &property_blob()[..55];
    assert!(matches!(
        Blob::parse_property(truncated),
        Err(Error::Truncated(_))
    ));
}

#[test]
fn configured_payload_budgets_are_enforced_before_retention() {
    let source = property_blob();
    let limits = Limits {
        max_blob_bytes: source.len(),
        max_signature_bytes: 2,
        max_certificate_store_bytes: 4,
    };
    assert_eq!(
        Blob::parse_with(&source, Kind::Property, limits),
        Err(Error::Limit("signature"))
    );

    let limits = Limits {
        max_blob_bytes: source.len() - 1,
        ..Limits::default()
    };
    assert_eq!(
        Blob::parse_with(&source, Kind::Property, limits),
        Err(Error::Limit("blob byte"))
    );
}
