//! Native binary Word visual metadata, with payloads kept inert.
use litchi_doc::Package;
use sha2::{Digest, Sha256};
fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn hash(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

#[test]
fn native_binary_vector_group_exposes_all_three_children_without_flattening() {
    use litchi_doc::shape::{Kind, ShapeId};
    let source = native("word-vector-diagram.doc");
    let bytes = source.to_bytes().unwrap();
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(&bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let shapes = litchi_doc::shape::extract_drawing_shapes(&mut ole).unwrap();
    assert_eq!(shapes.len(), 1);
    let root = &shapes[0];
    assert_eq!(root.id(), 1026);
    assert_eq!(root.shape_type, Kind::Group);
    assert_eq!(root.native_shape_type, Some(0));
    assert_eq!(root.office_art_bytes().len(), 5636);
    assert_eq!(root.unknown_properties().len(), 12);
    let group = root.group().unwrap();
    let bounds = group.bounds().unwrap();
    assert_eq!(
        (bounds.left, bounds.top, bounds.right, bounds.bottom),
        (0, 0, 3048000, 1714500)
    );
    assert_eq!(group.children().len(), 3);
    for (shape, (id, kind, native, rect, size)) in group.children().iter().zip([
        (1027, Kind::Rectangle, 1, (0, 0, 3048000, 1714500), 1140),
        (
            1028,
            Kind::AutoShape,
            2,
            (609600, 342900, 2438400, 1371600),
            1169,
        ),
        (
            1029,
            Kind::TextBox,
            202,
            (548520, 696225, 2499875, 1143265),
            1139,
        ),
    ]) {
        assert_eq!(shape.id(), id);
        assert_eq!(shape.shape_type, kind);
        assert_eq!(shape.native_shape_type, Some(native));
        assert_eq!(shape.office_art_bytes().len(), size);
        let anchor = shape.anchor().unwrap();
        assert_eq!((anchor.left, anchor.top, anchor.right, anchor.bottom), rect);
        assert_eq!(group.find(ShapeId::from_raw(id)).unwrap().id(), id);
    }
    assert_eq!(group.children()[0].fill_color, Some((40, 121, 216)));
    assert_eq!(group.children()[1].fill_color, Some((255, 193, 42)));
    assert_eq!(
        group.children()[2].text.as_deref(),
        Some("SVG_MARKER_007\r")
    );
}

#[test]
fn native_binary_ole_inventory_keeps_two_separate_storages_and_payloads() {
    let package = native("word-visual-objects.doc");
    let snapshot = package.embedded_objects().unwrap();
    let storages = snapshot.storages();
    assert_eq!(storages.len(), 2);
    assert_eq!(
        storages.iter().map(|e| e.key()).collect::<Vec<_>>(),
        ["_1853584064", "_1853584157"]
    );
    let inventory = storages
        .iter()
        .map(litchi_doc::embedded_object::Metadata::of)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    for (entry, user_type) in inventory
        .iter()
        .zip(["Microsoft Excel Worksheet", "Microsoft Excel Chart"])
    {
        assert!(!entry.is_linked());
        assert!(!entry.is_activex());
        let comp = entry.comp_obj().unwrap();
        assert_eq!(comp.ansi_user_type(), user_type);
        assert_eq!(
            comp.ansi_clipboard(),
            &litchi_doc::embedded_object::Clipboard::None
        );
        let info = entry.obj_info().unwrap();
        assert!(info.persist2_present && info.queried_enhanced_metafile);
        assert_eq!(info.clipboard_format, 3);
        assert!(
            !info.display_as_icon
                && !info.ole1
                && !info.enhanced_metafile
                && !info.stored_as_enhanced_metafile
        );
    }
    assert!(!inventory[0].obj_info().unwrap().recompose_on_resize);
    assert!(inventory[1].obj_info().unwrap().recompose_on_resize);
    let payload = inventory[0]
        .unknown()
        .iter()
        .find(|p| p.name() == Some("Package"))
        .unwrap();
    assert_eq!(payload.bytes().len(), 18372);
    assert_eq!(
        hash(payload.bytes()),
        "c4937954a07d192200a9a4cb17d016c345be8405dff727f25e39e2d1cf65e866"
    );
    let chart = inventory[1]
        .unknown()
        .iter()
        .find(|p| p.name() == Some("Workbook"))
        .unwrap();
    assert_eq!(chart.bytes().len(), 30785);
}

#[test]
fn native_binary_body_edit_preserves_all_object_storage_metadata_and_payloads() {
    let package = native("word-visual-objects.doc");
    let before = package.to_bytes().unwrap();
    let snapshot = package.embedded_objects().unwrap();
    let metadata = snapshot
        .storages()
        .iter()
        .map(litchi_doc::embedded_object::Metadata::of)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph(litchi_core::Position::new(0), "LEGACY_VISUALS_REVIEWED")
        .unwrap();
    let commit = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("reviewed.doc");
    commit.package().save(&path).unwrap();
    let reopened = Package::open(path).unwrap();
    let after = reopened.embedded_objects().unwrap();
    assert_eq!(
        after
            .storages()
            .iter()
            .map(litchi_doc::embedded_object::Metadata::of)
            .collect::<Result<Vec<_>, _>>()
            .unwrap(),
        metadata
    );
    assert_eq!(
        reopened
            .document()
            .body_paragraph_text(litchi_core::Position::new(0))
            .unwrap(),
        Some("LEGACY_VISUALS_REVIEWED")
    );
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
    assert!(package.to_bytes().unwrap() == before);
}
