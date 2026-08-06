//! Regression tests for the DOC embedded-object model and codec.

use super::Limits;
use super::codec::{OBJECT_POOL, discover_targets, is_object_storage_name};
use super::model::Info;
use litchi_cfb::OleWriter;
use std::io::Cursor;

#[test]
fn object_pool_target_names_follow_decimal_storage_form() {
    assert!(is_object_storage_name("_0"));
    assert!(is_object_storage_name("_00042"));
    assert!(is_object_storage_name("_-1"));
    assert!(!is_object_storage_name("Object"));
    assert!(!is_object_storage_name("_"));
    assert!(!is_object_storage_name("_+1"));
    assert!(!is_object_storage_name("_42x"));
}

#[test]
fn target_discovery_keeps_exact_object_pool_storage_names() {
    let mut writer = OleWriter::new();
    writer.create_storage(&[OBJECT_POOL, "_00042"]).unwrap();
    writer.create_storage(&[OBJECT_POOL, "_-1"]).unwrap();
    writer.create_storage(&[OBJECT_POOL, "not-an-id"]).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let (targets, object_pool_exists) = discover_targets(&bytes.into_inner(), Limits::default())
        .expect("ObjectPool target discovery should succeed");
    assert!(object_pool_exists);
    assert_eq!(targets.len(), 2);
    assert!(
        targets.get("_00042").is_some_and(|target| {
            target.path() == [OBJECT_POOL.to_owned(), "_00042".to_owned()]
        })
    );
    assert!(
        targets
            .get("_-1")
            .is_some_and(|target| { target.path() == [OBJECT_POOL.to_owned(), "_-1".to_owned()] })
    );
}

#[test]
fn obj_info_reads_the_doc_opaque_stream_shape() {
    let info = Info::read(&[0x00, 0x82, 0x03, 0x00, 0x00, 0x00]).unwrap();
    assert!(info.recompose_on_resize);
    assert!(info.view_object);
    assert_eq!(info.clipboard_format, 3);
    assert!(info.persist2_present);
    assert_eq!(
        info.to_bytes().unwrap(),
        [0x00, 0x82, 0x03, 0x00, 0x00, 0x00]
    );
    assert!(Info::read(&[0x00, 0x04, 0x00, 0x00]).is_err());
}

#[test]
fn obj_info_preserves_undefined_bits_and_optional_presence() {
    let bytes = [0x2D, 0x40, 0x14, 0x00, 0xF0, 0x00];
    let info = Info::read(&bytes).unwrap();
    assert_eq!(info.reserved_persist1, 0x402D);
    assert_eq!(info.reserved_persist2, 0x00F0);
    assert!(info.persist2_present);
    assert_eq!(info.to_bytes().unwrap(), bytes);

    let without_optional = Info::read(&[0x00, 0x00, 0x03, 0x00]).unwrap();
    assert!(!without_optional.persist2_present);
    assert_eq!(
        without_optional.to_bytes().unwrap(),
        [0x00, 0x00, 0x03, 0x00]
    );

    let explicit_zero_optional = Info::read(&[0x00, 0x00, 0x03, 0x00, 0x00, 0x00]).unwrap();
    assert!(explicit_zero_optional.persist2_present);
    assert_eq!(
        explicit_zero_optional.to_bytes().unwrap(),
        [0x00, 0x00, 0x03, 0x00, 0x00, 0x00]
    );
}

#[test]
fn obj_info_rejects_invalid_required_bits_without_ole_access() {
    assert!(Info::read(&[0x00, 0x08, 0x00, 0x00]).is_err());
    assert!(Info::read(&[0x00, 0x00, 0x00, 0x00, 0x02, 0x00]).is_err());
    assert!(Info::read(&[0x00, 0x20, 0x00, 0x00]).is_err());

    let mut info = Info::read(&[0x00, 0x00, 0x00, 0x00]).unwrap();
    info.reserved_persist1 = 1 << 1;
    assert!(info.to_bytes().is_err());
    info.reserved_persist1 = 0;
    info.reserved_persist2 = 1 << 1;
    assert!(info.to_bytes().is_err());
}
