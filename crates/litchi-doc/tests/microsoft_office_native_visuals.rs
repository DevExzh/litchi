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

fn image_runs(document: &litchi_doc::Document) -> Vec<litchi_doc::Run> {
    document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|p| p.runs().unwrap())
        .filter(|r| r.image().is_some())
        .collect()
}

#[test]
fn native_picf_and_floating_store_images_resolve_to_their_own_payloads() {
    use litchi_odraw::image::{Compression, Kind};
    use std::io::Read;
    let package = native("word-visual-objects.doc");
    let document = package.document();
    let runs = image_runs(&document);
    assert_eq!(runs.len(), 4);
    for (run, (cp, shape_id, kind, size, digest)) in runs.iter().zip([
        (
            76,
            None,
            Kind::Jpeg,
            83204,
            "b6f8d168d41e0fc296eb1e0e3855d9669584b7528c2ef4966817443d9bc9b929",
        ),
        (
            97,
            Some(1026),
            Kind::Png,
            1698,
            "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8",
        ),
        (
            143,
            None,
            Kind::Emf,
            36576,
            "b9d997d9601b0fe88fdb111b55c39b9946abe54fc2d3544a5ff5729cc86f1a9f",
        ),
        (
            183,
            None,
            Kind::Png,
            33196,
            "ded298bda4611ec0d174be16f34afd51b0ecfe4f57252c8bec9fd28791f7f542",
        ),
    ]) {
        assert_eq!(run.cp_range().unwrap().start(), cp);
        let image = run.image().unwrap();
        assert_eq!(image.drawing_shape_id(), shape_id);
        let file = document.image_data(image).unwrap();
        assert_eq!(file.kind(), kind);
        assert_eq!(file.data().unwrap().len(), size);
        assert_eq!(hash(file.data().unwrap()), digest);
    }
    let floating = runs[1].image().unwrap();
    let placement = document
        .shape_positions()
        .iter()
        .find(|a| a.spa.shape_id == 1026)
        .unwrap();
    assert_eq!(placement.cp, 97);
    assert_eq!(
        (
            placement.spa.right - placement.spa.left,
            placement.spa.bottom - placement.spa.top
        ),
        (2880, 1620)
    );
    assert_eq!(placement.spa.wrap, litchi_doc::ShapeTextWrap::Square);
    assert_eq!(
        floating.pic_offset(),
        0,
        "inherited CPicLocation is only a hint"
    );
    assert!(matches!(
        floating.data(document.data_stream().unwrap(), document.word_document()),
        Err(litchi_doc::image::ImageError::DrawingStoreRequired(1026))
    ));
    let anchors = document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|p| p.runs().unwrap())
        .filter(|r| r.text().unwrap() == "\u{0008}")
        .collect::<Vec<_>>();
    assert_eq!(anchors.len(), 2);
    assert!(
        anchors[0].image().is_none(),
        "the rounded rectangle is not the inline JPEG"
    );
    assert!(anchors[1].image().is_some());
    let data = document.data_stream().unwrap();
    for (offset, length, goal, scale) in [
        (0, 83693, (10240, 7680), (422, 422)),
        (83693, 36802, (3920, 1620), (1000, 1000)),
        (120495, 33443, (9040, 5290), (960, 960)),
    ] {
        let picf = litchi_doc::image::PictureFields::try_parse(data, offset).unwrap();
        assert_eq!((picf.lcb, picf.cb_header, picf.mm), (length, 68, 100));
        assert_eq!((picf.dxa_goal, picf.dya_goal), goal);
        assert_eq!((picf.mx, picf.my), scale);
    }
    let preview = document.image_data(runs[2].image().unwrap()).unwrap();
    let blip = preview.blip().unwrap();
    let meta = blip.meta().unwrap().header();
    assert_eq!(
        (meta.size, meta.saved, meta.compression),
        (74428, 36576, Compression::Deflate)
    );
    let mut decoded = Vec::new();
    flate2::read::ZlibDecoder::new(preview.data().unwrap())
        .take(u64::from(meta.size) + 1)
        .read_to_end(&mut decoded)
        .unwrap();
    assert_eq!(decoded.len(), 74428);
    assert_eq!(
        hash(&decoded),
        "2585e48c4e95e526a3ca2cf32e14354d2bbb43b0d98bff7d4906870071a2c4a3"
    );
}

#[test]
fn native_vector_picf_placeholders_stay_distinct_from_groups_and_real_images() {
    let package = native("word-vector-diagram.doc");
    let document = package.document();
    let runs = image_runs(&document);
    assert_eq!(runs.len(), 3);
    assert_eq!(
        runs.iter()
            .map(|r| r.image().unwrap().pic_offset())
            .collect::<Vec<_>>(),
        [0, 27846, 27990]
    );
    assert!(
        runs.iter()
            .all(|r| r.image().unwrap().drawing_shape_id().is_none())
    );
    for (index, size, digest) in [
        (
            0,
            10853,
            "544cbfb9063b539645ef785be27aaba8717299e6719c48272d858b00c5060da8",
        ),
        (
            2,
            15111,
            "106ccd6deda8b393dade58844244222279a8dbf3940e7667142fa58af596b987",
        ),
    ] {
        let image = document.image_data(runs[index].image().unwrap()).unwrap();
        assert_eq!(image.kind(), litchi_odraw::image::Kind::Png);
        assert_eq!(image.data().unwrap().len(), size);
        assert_eq!(hash(image.data().unwrap()), digest);
    }
    let placeholder =
        litchi_doc::image::PictureFields::try_parse(document.data_stream().unwrap(), 27846)
            .unwrap();
    assert_eq!(
        (
            placeholder.lcb,
            placeholder.cb_header,
            placeholder.dxa_goal,
            placeholder.dya_goal
        ),
        (144, 68, 4800, 2700)
    );
    assert!(matches!(
        document.image_data(runs[1].image().unwrap()),
        Err(litchi_doc::image::ImageError::NoPicture)
    ));
    let group_anchor = document
        .paragraphs()
        .unwrap()
        .into_iter()
        .flat_map(|p| p.runs().unwrap())
        .find(|r| r.text().unwrap() == "\u{0008}")
        .unwrap();
    assert!(
        group_anchor.image().is_none(),
        "a group must not alias the first PICF"
    );
}

#[test]
fn native_picture_payloads_survive_body_edits_and_shape_anchor_relocation() {
    let signature = |package: &Package| {
        let document = package.document();
        image_runs(&document)
            .iter()
            .map(|run| {
                let image = run.image().unwrap();
                let payload = match document.image_data(image) {
                    Ok(file) => Some((file.kind(), hash(file.data().unwrap()))),
                    Err(litchi_doc::image::ImageError::NoPicture) => None,
                    Err(error) => panic!("unexpected image error: {error}"),
                };
                (image.pic_offset(), image.drawing_shape_id(), payload)
            })
            .collect::<Vec<_>>()
    };
    for name in ["word-visual-objects.doc", "word-vector-diagram.doc"] {
        let source = native(name);
        let before = source.to_bytes().unwrap();
        let expected = signature(&source);
        let mut edit = source.edit().unwrap();
        edit.replace_paragraph(litchi_core::Position::new(0), "PICTURE_REVIEWED")
            .unwrap();
        let commit = edit.commit().unwrap();
        let directory = tempfile::tempdir().unwrap();
        let destination = directory.path().join(name);
        commit.package().save(&destination).unwrap();
        let reopened = Package::open(destination).unwrap();
        assert_eq!(signature(&reopened), expected);
        assert!(
            reopened
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == before
        );
        assert!(source.to_bytes().unwrap() == before);
    }
}
