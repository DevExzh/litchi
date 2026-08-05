use std::borrow::Cow;

use super::{Compression, Kind, Storage};
use crate::records::Record;

const MAX_DECLARED_BYTES: u32 = 256 * 1_048_576;

#[test]
fn every_storage_kind_and_compression_roundtrips_exactly() {
    for kind in [Kind::OleObject, Kind::VbaProject, Kind::ActiveXControl] {
        let uncompressed = Storage::uncompressed(kind, vec![1, 2, 3]).unwrap();
        assert_eq!(uncompressed.kind(), kind);
        assert_eq!(uncompressed.compression(), Compression::Uncompressed);
        let record = uncompressed.to_record().unwrap();
        assert_eq!(Storage::parse_as(&record, kind).unwrap(), uncompressed);

        let compressed = Storage::compressed(kind, 4096, vec![0x78, 0x9c, 1, 2, 3, 4]).unwrap();
        assert_eq!(compressed.compression(), Compression::Zlib);
        assert_eq!(compressed.declared_uncompressed_len(), Some(4096));
        let record = compressed.to_record().unwrap();
        assert_eq!(record.instance, 1);
        assert_eq!(Storage::parse_as(&record, kind).unwrap(), compressed);
    }
}

#[test]
fn constructors_reject_contradictory_or_oversized_state() {
    assert!(Storage::compressed(Kind::VbaProject, MAX_DECLARED_BYTES + 1, Vec::new()).is_err());
    assert!(Storage::uncompressed(Kind::OleObject, Vec::new()).is_ok());
}

#[test]
fn storage_rejects_invalid_instance_and_truncated_compressed_header() {
    let value = Storage::uncompressed(Kind::OleObject, Vec::new()).unwrap();
    let mut bytes = value.to_record_bytes().unwrap();
    bytes[0..2].copy_from_slice(&(2u16 << 4).to_le_bytes());
    assert!(Storage::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
    bytes.truncate(8);
    bytes[0..2].copy_from_slice(&(1u16 << 4).to_le_bytes());
    assert!(Storage::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

#[test]
fn borrowed_storage_keeps_uncompressed_payload_borrowed() {
    let value =
        Storage::uncompressed(Kind::VbaProject, b"borrowed compound bytes".to_vec()).unwrap();
    let record = value.to_record_bytes().unwrap();
    let storage = super::Ref::parse_at(&record, 0, Kind::VbaProject).unwrap();

    assert!(storage.check_stored_limit(0).is_err());
    let cfb = storage
        .check_stored_limit(value.stored_payload_len())
        .unwrap()
        .decompressed_bytes(value.stored_payload_len())
        .unwrap();
    assert!(matches!(cfb, Cow::Borrowed(_)));
    assert_eq!(cfb.as_ptr(), record[8..].as_ptr());
}

#[test]
fn bounded_zlib_decompression_requires_exact_size_and_no_trailing_data() {
    use std::io::Write;

    let original = b"compound storage bytes".repeat(100);
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
    encoder.write_all(&original).unwrap();
    let compressed = encoder.finish().unwrap();
    let storage =
        Storage::compressed(Kind::VbaProject, original.len() as u32, compressed.clone()).unwrap();
    assert_eq!(
        storage.decompressed_bytes(original.len()).unwrap(),
        original
    );
    assert!(storage.decompressed_bytes(original.len() - 1).is_err());

    let wrong_size = Storage::compressed(
        Kind::VbaProject,
        original.len() as u32 + 1,
        compressed.clone(),
    )
    .unwrap();
    assert!(wrong_size.decompressed_bytes(original.len() + 1).is_err());

    let mut trailing_bytes = compressed;
    trailing_bytes.push(0);
    let trailing =
        Storage::compressed(Kind::VbaProject, original.len() as u32, trailing_bytes).unwrap();
    assert!(trailing.decompressed_bytes(original.len()).is_err());
}
