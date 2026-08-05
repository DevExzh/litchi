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
    assert!(Info::read(&[0x00, 0x04, 0x00, 0x00]).is_err());
}
