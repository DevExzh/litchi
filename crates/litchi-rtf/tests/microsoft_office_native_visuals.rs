//! Native RTF visual payloads are metadata and inert bytes, never activated.
use litchi_rtf::{ObjectKind, Package};
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
fn canonical(package: &Package) -> Package {
    let mut bytes = Vec::new();
    litchi_rtf::write::Writer::new(&mut bytes)
        .write_document(package.raw())
        .unwrap();
    Package::from_bytes(&bytes).unwrap()
}

#[test]
fn native_rtf_ole_headers_payloads_and_previews_survive_model_writing() {
    let source = native("word-visual-objects.rtf");
    let verify = |package: &Package| {
        let objects = package.raw().objects();
        assert_eq!(objects.len(), 2);
        for (object, (class, width, height, length, digest, native_length, native_hash)) in
            objects.iter().zip([
                (
                    "Excel.Sheet.12",
                    3920,
                    1620,
                    39153,
                    "368367529b1ebf212145b1f211bc7916e593f7f03f0a9e02c1ec06fb98a56151",
                    20992,
                    "b1be607ac5c402d3c6a0bf002a675d0e0eef93ab9ecf86752ac0b5903efc45dc",
                ),
                (
                    "Excel.Chart.8",
                    9041,
                    5291,
                    42730,
                    "104c33abd4f333e120bd55b551c3c97e8430c9c2a5e79f0db745648e9ffef9c8",
                    34816,
                    "827c8e3396ef7757d2785b7e35d0de13672a0000b807fdcc2ee65cb758da05de",
                ),
            ])
        {
            assert_eq!(object.kind, ObjectKind::Embedded);
            assert_eq!(object.class_name, class);
            assert_eq!((object.width, object.height), (width, height));
            assert_eq!(object.data.len(), length);
            assert_eq!(hash(&object.data), digest);
            let header = object.ole_header().unwrap();
            assert_eq!(header.ole_version, 0x501);
            assert_eq!(header.format_id, 2);
            assert_eq!(header.class_name, class.as_bytes());
            assert!(header.topic_name.is_empty() && header.item_name.is_empty());
            assert!(header.is_compound_file());
            assert_eq!(header.native_data.len(), native_length);
            assert_eq!(hash(header.native_data), native_hash);
            assert_eq!(object.result_picture_indices.len(), 2);
            assert_eq!(
                object.result_picture_wrappers,
                [
                    Some(litchi_rtf::PictureCompatibilityKind::ShapePicture),
                    Some(litchi_rtf::PictureCompatibilityKind::NonShapePicture)
                ]
            );
            assert!(
                !package.raw().pictures()[object.result_picture_indices[0]]
                    .data
                    .is_empty()
            );
        }
        assert!(objects[1].set_size);
        assert_eq!(
            (objects[1].scale_x, objects[1].scale_y),
            (Some(96), Some(96))
        );
    };
    verify(&source);
    let written = canonical(&source);
    verify(&written);
    assert_eq!(written.raw().pictures(), source.raw().pictures());
    assert_eq!(written.text(), source.text());
}

#[test]
fn native_rtf_vector_picture_payloads_remain_distinct_from_declared_types() {
    let source = native("word-vector-diagram.rtf");
    let pictures = source.raw().pictures();
    assert_eq!(pictures.len(), 4);
    for (pair, (size, digest, width, height)) in pictures.chunks_exact(2).zip([
        (
            10853,
            "544cbfb9063b539645ef785be27aaba8717299e6719c48272d858b00c5060da8",
            8855,
            4957,
        ),
        (
            15111,
            "106ccd6deda8b393dade58844244222279a8dbf3940e7667142fa58af596b987",
            15910,
            2593,
        ),
    ]) {
        assert_eq!(pair[0].image_type, litchi_rtf::ImageType::Png);
        assert_eq!(pair[1].image_type, litchi_rtf::ImageType::Wmf);
        assert_eq!(pair[0].data, pair[1].data);
        for picture in pair {
            assert_eq!(picture.data.len(), size);
            assert_eq!(hash(&picture.data), digest);
            assert_eq!((picture.width, picture.height), (Some(width), Some(height)));
            assert!(picture.data.starts_with(b"\x89PNG\r\n\x1a\n"));
        }
    }
    let structural = source
        .raw()
        .opaque_nodes()
        .iter()
        .filter(|n| matches!(n.anchor(), litchi_rtf::opaque::Anchor::Structural { .. }))
        .collect::<Vec<_>>();
    assert_eq!(structural.len(), 1);
    assert_eq!(structural[0].source(), b"\\defshp");
    assert!(matches!(
        structural[0].anchor(),
        litchi_rtf::opaque::Anchor::Structural {
            context: litchi_rtf::opaque::Context::Field,
            ..
        }
    ));
    let mut output = Vec::new();
    let error = litchi_rtf::write::Writer::new(&mut output)
        .write_document(source.raw())
        .unwrap_err();
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    assert_eq!(
        error.to_string(),
        "RTF canonical rewrite cannot preserve a structurally scoped opaque node"
    );
    assert!(output.is_empty());
}

#[test]
fn native_rtf_visual_snapshots_preserve_bytes_and_refuse_ordinary_run_edits() {
    // The ordinary-run mapper explicitly excludes picture/object/field owners;
    // it does not yet prove their positional dependencies after a run edit.
    for name in ["word-visual-objects.rtf", "word-vector-diagram.rtf"] {
        let source = native(name);
        let before = source.to_bytes().unwrap();
        let directory = tempfile::tempdir().unwrap();
        let destination = directory.path().join(name);
        source.save(&destination).unwrap();
        assert!(std::fs::read(&destination).unwrap() == before);
        let reopened = Package::open(destination).unwrap();
        assert_eq!(reopened.raw().pictures(), source.raw().pictures());
        assert_eq!(
            reopened
                .raw()
                .objects()
                .iter()
                .map(|o| (
                    &o.data,
                    &o.result_picture_indices,
                    &o.result_picture_wrappers
                ))
                .collect::<Vec<_>>(),
            source
                .raw()
                .objects()
                .iter()
                .map(|o| (
                    &o.data,
                    &o.result_picture_indices,
                    &o.result_picture_wrappers
                ))
                .collect::<Vec<_>>()
        );
        let mut edit = source.edit();
        edit.replace_body_run_texts(&[litchi_rtf::edit::RunTextReplacement::new(
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            "LEGACY_VISUALS_REVIEWED",
        )])
        .unwrap();
        let Err(litchi_rtf::edit::Error::Refused(refusal)) = edit.commit() else {
            panic!("ordinary-run source requires a typed refusal")
        };
        assert_eq!(
            refusal.kind(),
            litchi_core::edit::RefusalKind::UnsupportedContent
        );
        assert_eq!(
            refusal.reason(),
            "run-format editing refuses dependent body destinations or review content"
        );
        assert!(source.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_rtf_object_transfer_preserves_payload_and_preview_wrappers() {
    let source = native("word-visual-objects.rtf");
    let source_bytes = source.to_bytes().unwrap();
    let native_target = native("word-rtf-minimal.rtf");
    let native_before = native_target.to_bytes().unwrap();
    let Err(litchi_rtf::edit::Error::Refused(refusal)) =
        litchi_rtf::edit::TransferPlan::object(&source, 0, &native_target)
    else {
        panic!("opaque target requires a typed refusal")
    };
    assert_eq!(
        refusal.kind(),
        litchi_core::edit::RefusalKind::UnsupportedContent
    );
    assert_eq!(
        refusal.reason(),
        "ordinary-root transfer refuses unknown target destinations"
    );
    assert!(native_target.to_bytes().unwrap() == native_before);
    // The supported transfer receiver is an ordinary generated RTF root.
    // The native object source and all its payloads stay unchanged.
    for index in 0..2 {
        let target = Package::from_bytes(br"{\rtf1\ansi Target}").unwrap();
        let before = target.to_bytes().unwrap();
        let commit = litchi_rtf::edit::TransferPlan::object(&source, index, &target)
            .unwrap()
            .commit()
            .unwrap();
        let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.raw().objects().len(), 1);
        let original = &source.raw().objects()[index];
        let copied = &reopened.raw().objects()[0];
        assert_eq!(copied.data, original.data);
        assert_eq!(copied.class_name, original.class_name);
        assert_eq!(
            copied.result_picture_wrappers,
            original.result_picture_wrappers
        );
        assert_eq!(
            copied
                .result_picture_indices
                .iter()
                .map(|i| &reopened.raw().pictures()[*i])
                .collect::<Vec<_>>(),
            original
                .result_picture_indices
                .iter()
                .map(|i| &source.raw().pictures()[*i])
                .collect::<Vec<_>>()
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
        assert!(target.to_bytes().unwrap() == before);
    }
    assert!(source.to_bytes().unwrap() == source_bytes);
}
