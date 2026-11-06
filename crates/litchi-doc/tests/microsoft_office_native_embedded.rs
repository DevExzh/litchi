//! Native ObjectPool storage reads remain distinct from managed-field editing.
use litchi_core::edit::RefusalKind;
use litchi_doc::{
    Package,
    embedded_object::{Error, Metadata},
};
fn source() -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word/word-visual-objects.doc"),
    )
    .unwrap()
}
fn unsupported<T>(result: Result<T, Error>) {
    assert!(matches!(result,Err(Error::Refused(r)) if r.kind()==RefusalKind::UnsupportedContent));
}
#[test]
fn native_unmanaged_object_edits_refuse_without_misreporting_corruption() {
    let p = source();
    let bytes = p.to_bytes().unwrap();
    let s = p.embedded_objects().unwrap();
    assert_eq!(s.storages().len(), 2);
    // The positive native storage inventory is independent of the managed-field index.
    assert!(s.objects().unwrap().is_empty());
    let mut edit = p.edit_embedded_objects().unwrap();
    for (id, key) in [(1853584064, "_1853584064"), (1853584157, "_1853584157")] {
        let object = s.storages().get(key).unwrap();
        let metadata = Metadata::of(object).unwrap();
        let called = std::cell::Cell::new(false);
        unsupported(edit.update_info(id, |info| {
            called.set(true);
            info.recompose_on_resize = true;
        }));
        assert!(!called.get());
        unsupported(edit.set_info(id, *metadata.obj_info().unwrap()));
        unsupported(edit.replace_storage(id, object.compound().to_vec()));
        unsupported(edit.remove(id));
        unsupported(edit.update_link(
            id,
            litchi_ole_common::object::link::Limits::default(),
            |_| panic!("unmanaged callback"),
        ));
        unsupported(edit.update_presentation(
            id,
            0,
            litchi_ole_common::ole_streams::Limits::default(),
            |_| panic!("unmanaged callback"),
        ));
        unsupported(edit.update_native(
            id,
            litchi_ole_common::ole_streams::Limits::default(),
            |_| panic!("unmanaged callback"),
        ));
    }
    unsupported(edit.reorder(&[1853584157, 1853584064]));
    let commit = edit.commit().unwrap();
    assert!(commit.package().to_bytes().unwrap() == bytes);
    assert!(p.to_bytes().unwrap() == bytes);
    assert_eq!(
        commit
            .package()
            .embedded_objects()
            .unwrap()
            .storages()
            .len(),
        2
    );
}
#[test]
fn native_storage_compound_export_retains_child_streams_and_class_identity() {
    let p = source();
    let s = p.embedded_objects().unwrap();
    assert_eq!(s.storages().len(), 2);
    for object in s.storages().iter() {
        let mut ole = litchi_cfb::OleFile::open(
            std::io::Cursor::new(object.compound()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        assert_eq!(ole.list_streams().len(), object.streams().len());
        for stream in object.streams() {
            let path = stream.path().iter().map(String::as_str).collect::<Vec<_>>();
            assert!(ole.open_stream(&path).unwrap() == stream.bytes());
        }
        let expected = if object.key() == "_1853584064" {
            "00020830-0000-0000-C000-000000000046"
        } else {
            "00020821-0000-0000-C000-000000000046"
        };
        assert_eq!(object.storage().clsid(), Some(expected));
        assert_eq!(ole.root_entry().unwrap().clsid, expected);
    }
}
