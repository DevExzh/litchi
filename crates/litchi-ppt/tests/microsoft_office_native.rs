//! Regressions backed by native Microsoft PowerPoint specimens.

use litchi_ppt::{Package, shapes::ShapeEnum};

const VISUAL: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-visual-objects.ppt");

#[test]
fn native_powerpoint_group_keeps_both_text_shapes_and_the_connector() {
    let package = Package::from_bytes(VISUAL).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slide_count(), 5);
    let slide = presentation.slide_at(4).unwrap().unwrap();
    let shapes = slide.shapes().unwrap();
    let group = shapes
        .iter()
        .find_map(|shape| match shape {
            ShapeEnum::Group(group) => Some(group),
            _ => None,
        })
        .expect("native Group 6");
    assert_eq!(group.children().len(), 3);
    let texts = group
        .children()
        .iter()
        .map(|shape| shape.text().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(texts, ["SOURCE_002", "TARGET_002", ""]);
    assert!(matches!(group.children()[2], ShapeEnum::Line(_)));
    assert_eq!(package.to_bytes().unwrap(), VISUAL);
}

#[test]
fn native_powerpoint_decks_materialize_all_slide_shapes_without_dropping_any() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    let mut failures = Vec::new();
    for (name, count) in [
        ("powerpoint-text-table.ppt", 3),
        ("powerpoint-visual-objects.ppt", 5),
        ("powerpoint-visual-objects.pps", 5),
        ("powerpoint-visual-objects.pot", 5),
        ("powerpoint-navigation-review.ppt", 6),
        ("powerpoint-master-motion.ppt", 7),
        ("powerpoint-media-objects.ppt", 4),
        ("powerpoint-modern-media.ppt", 8),
        ("powerpoint-fonts-math-handout.ppt", 4),
        ("powerpoint-passive-macro.ppt", 1),
    ] {
        let source = std::fs::read(root.join(name)).unwrap();
        let package = Package::from_bytes(&source).unwrap();
        let presentation = package.presentation().unwrap();
        assert_eq!(presentation.slide_count(), count, "{name}");
        for slide in presentation.slides().unwrap() {
            match slide.shapes() {
                Ok(shapes) => {
                    assert!(!shapes.is_empty(), "{name}, slide {}", slide.slide_number());
                    for defect in slide.tolerance_report().unwrap().defects() {
                        assert!(
                            !matches!(
                                defect.defect,
                                litchi_ppt::Defect::Shape(litchi_ppt::ShapeDefect::UnreadableShape)
                            ),
                            "{name}, slide {}: {defect:?}",
                            slide.slide_number()
                        );
                    }
                },
                Err(error) => {
                    failures.push(format!("{name}, slide {}: {error:?}", slide.slide_number()))
                },
            }
        }
        assert_eq!(package.to_bytes().unwrap(), source);
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn native_powerpoint_media_previews_resolve_by_slot_and_retain_exact_payloads() {
    use sha2::{Digest, Sha256};
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-media-objects.ppt"
    );
    let package = Package::from_bytes(bytes).unwrap();
    let presentation = package.presentation().unwrap();
    let images = presentation.images().unwrap();
    assert_eq!(images.len(), 3);
    for (index, (image, (extension, len, sha256))) in images
        .iter()
        .zip([
            (
                "png",
                1698,
                "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8",
            ),
            (
                "png",
                4717,
                "b0151c2c2e3cf64bc37a7bb9d8b8b98d4c4fccf7b6af4c08c4f847a79f9db0da",
            ),
            (
                "emf",
                23407,
                "e084b41f70fdca72da3dbaa85d8278353a7a7fc06c2a53411eaf34026f6e4a42",
            ),
        ])
        .enumerate()
    {
        assert_eq!(image.extension(), extension);
        assert_eq!(image.index(), index);
        assert_eq!(image.data().unwrap().len(), len);
        assert_eq!(
            Sha256::digest(image.data().unwrap())
                .iter()
                .map(|byte| format!("{byte:02x}"))
                .collect::<String>(),
            sha256
        );
        assert_eq!(
            presentation
                .image_at(index as u32 + 1)
                .unwrap()
                .unwrap()
                .data()
                .unwrap(),
            image.data().unwrap()
        );
    }
    assert_eq!(package.to_bytes().unwrap(), bytes);
}

#[test]
fn native_powerpoint_placeholder_table_exposes_all_twelve_cells() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-text-table.ppt");
    let package = Package::from_bytes(bytes).unwrap();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slide_at(2).unwrap().unwrap();
    let table = slide
        .shapes()
        .unwrap()
        .iter()
        .find_map(|shape| match shape {
            ShapeEnum::Table(table) => Some(table),
            _ => None,
        })
        .expect("native content-placeholder table");
    assert_eq!((table.rows(), table.columns()), (4, 3));
    for (row, values) in [
        ["Item", "Quantity", "Total"],
        ["Apples", "3", "3.75"],
        ["Pears", "5", "12.50"],
        ["Tea", "2", "9.50"],
    ]
    .iter()
    .enumerate()
    {
        for (column, value) in values.iter().enumerate() {
            assert_eq!(table.cell(row, column), Some(*value));
        }
    }
}

#[test]
fn native_powerpoint_media_and_ole_placeholders_keep_typed_frames() {
    use litchi_ppt::shapes::PictureFrameKind;
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-media-objects.ppt"
    );
    let package = Package::from_bytes(bytes).unwrap();
    let presentation = package.presentation().unwrap();
    // The downgraded video retains a media click action but no ExObjRef;
    // this identifies its preview without asserting a playable video payload.
    for (index, kind, id) in [
        (1, PictureFrameKind::Media, None),
        (2, PictureFrameKind::Media, Some(4)),
        (3, PictureFrameKind::OleObject, Some(5)),
    ] {
        let slide = presentation.slide_at(index).unwrap().unwrap();
        let picture = slide
            .shapes()
            .unwrap()
            .iter()
            .find_map(|shape| match shape {
                ShapeEnum::Picture(picture) => Some(picture),
                _ => None,
            })
            .expect("native preview frame");
        assert_eq!(picture.frame_kind(), kind);
        assert_eq!(picture.external_object_id(), id);
        assert!(picture.blip_id().is_some());
    }
    let media = presentation.external_media().unwrap().unwrap();
    assert_eq!(media.objects.len(), 1);
    let litchi_ppt::external_media::Object::EmbeddedWav(audio) = media.objects[0] else {
        panic!("native WAV")
    };
    assert_eq!(
        (audio.media.id, audio.sound_id, audio.duration_ms),
        (4, Some(1), 2000)
    );
    assert!(audio.media.loop_playback && audio.media.rewind_after_playing);
    assert!(!audio.media.narration);
    let sounds = presentation.embedded_sounds().unwrap().unwrap();
    assert_eq!(sounds.sounds.len(), 1);
    let sound = sounds.get(1).unwrap();
    assert!(sound.data.starts_with(b"RIFF"));
    assert_eq!(sound.name, "75ACA65A.wav");
    assert_eq!(sound.data.len(), 192044);
    let objects = presentation.ole_objects().unwrap().unwrap();
    assert_eq!(objects.objects.len(), 1);
    let litchi_ppt::embedded::object::ExternalObject::Object(object) = &objects.objects[0] else {
        panic!("native Excel object")
    };
    assert_eq!(object.object.id, 5);
    assert_eq!(object.program_id.as_deref(), Some("Excel.Sheet.12"));
    assert_eq!(object.menu_name.as_deref(), Some("Worksheet"));
}

fn shape_texts(package: &Package) -> Vec<Vec<String>> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .unwrap()
                .iter()
                .map(|shape| shape.text().unwrap())
                .collect()
        })
        .collect()
}

#[test]
fn native_powerpoint_text_and_table_edits_reopen_and_invert_without_losing_other_content() {
    use litchi_ppt::package::edit::{Position, Target};
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-text-table.ppt");
    let package = Package::from_bytes(bytes).unwrap();
    let before = shape_texts(&package);
    assert_eq!(before[0][0], "Native PowerPoint feature fixture");
    assert!(before[0][1].contains("café 中文 日本語 العربية 😀"));
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(
        Target::new(Position::new(0), Position::new(0)),
        "Native PowerPoint edited title",
    )
    .unwrap();
    edit.set_table_cell_text(
        Target::new(Position::new(2), Position::new(1)),
        Position::new(1),
        Position::new(0),
        "Oranges",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let published = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&published).unwrap();
    let mut expected = before;
    expected[0][0] = "Native PowerPoint edited title".into();
    expected[2][1] = expected[2][1].replace("Apples", "Oranges");
    assert_eq!(shape_texts(&reopened), expected);
    let presentation = reopened.presentation().unwrap();
    let slide = presentation.slide_at(0).unwrap().unwrap();
    assert_eq!(
        slide.speaker_notes().unwrap().unwrap().text().unwrap(),
        "Speaker note for slide 1. Expected marker: PPT_NOTES_001."
    );
    for name in ["\u{5}SummaryInformation", "\u{5}DocumentSummaryInformation"] {
        assert_eq!(
            reopened.shared_ole_file().open_stream(&[name]).unwrap(),
            package.shared_ole_file().open_stream(&[name]).unwrap()
        );
    }
    assert_eq!(package.to_bytes().unwrap(), bytes);
    assert_eq!(
        package
            .apply(commit.patch())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        published
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

#[cfg(feature = "encryption")]
#[test]
fn native_powerpoint_cryptoapi_requires_the_password_and_preserves_encrypted_source() {
    use litchi_ppt::{EncryptionProfile, Error, OpenOptions};
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-encrypted.ppt");
    assert!(matches!(
        Package::from_bytes(bytes),
        Err(Error::PasswordRequired)
    ));
    assert!(matches!(
        Package::from_bytes_with_options(bytes, OpenOptions::new().password("wrong")),
        Err(Error::InvalidPassword)
    ));
    let package =
        Package::from_bytes_with_options(bytes, OpenOptions::new().password("OfficeFixture42!"))
            .unwrap();
    assert_eq!(
        package.encryption(),
        Some(EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
    );
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slide_count(), 1);
    assert_eq!(
        presentation.text().unwrap(),
        "PPT_ENCRYPTED_008\nFixture ID: powerpoint-encrypted-008\rNative password-to-open specimen. VALUE_008: 42"
    );
    assert_eq!(
        presentation
            .slide_at(0)
            .unwrap()
            .unwrap()
            .shapes()
            .unwrap()
            .len(),
        2
    );
    assert!(matches!(package.to_bytes(), Err(Error::EncryptionPolicy)));
    assert_eq!(package.to_encrypted_bytes().unwrap(), bytes);
}

#[cfg(not(feature = "encryption"))]
#[test]
fn native_powerpoint_encryption_without_the_feature_is_explicitly_unsupported() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-encrypted.ppt");
    assert!(matches!(
        Package::from_bytes_with_options(
            bytes,
            litchi_ppt::OpenOptions::new().password("OfficeFixture42!")
        ),
        Err(litchi_ppt::Error::UnsupportedEncryption(
            litchi_ppt::package::EncryptionKind::CryptoApi
        ))
    ));
}

#[test]
fn native_powerpoint_custom_shows_follow_current_order_after_summary_insertion() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    for (name, alpha, beta) in [
        (
            "powerpoint-navigation-review.ppt",
            vec![0, 1, 4],
            vec![5, 3, 0],
        ),
        ("powerpoint-master-motion.ppt", vec![1, 2, 5], vec![6, 4, 1]),
    ] {
        let package = Package::open(root.join(name)).unwrap();
        let presentation = package.presentation().unwrap();
        let shows = presentation.custom_shows().unwrap();
        assert_eq!(shows.len(), 2, "{name}");
        assert_eq!(shows[0].name, "SHOW_ALPHA_003");
        assert_eq!(shows[0].slide_indices, alpha);
        assert_eq!(shows[1].name, "SHOW_BETA_003");
        assert_eq!(shows[1].slide_indices, beta);

        // A new append-only publication must read only the live document,
        // and membership follows stable IDs after the first slide moves.
        let before = package.to_bytes().unwrap();
        let count = presentation.slide_count();
        let mut edit = package.edit().unwrap();
        edit.move_slide(
            litchi_ppt::package::edit::Position::new(0),
            litchi_ppt::package::edit::Position::new(count - 1),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let changed = commit
            .package()
            .presentation()
            .unwrap()
            .custom_shows()
            .unwrap();
        assert_eq!(changed.len(), 2);
        for (actual, original) in changed.iter().zip(&shows) {
            assert_eq!(actual.name, original.name);
            assert_eq!(
                actual.slide_indices,
                original
                    .slide_indices
                    .iter()
                    .map(|index| if *index == 0 { count - 1 } else { index - 1 })
                    .collect::<Vec<_>>()
            );
        }
        assert_eq!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            before
        );
    }
}

#[test]
fn native_powerpoint_image_catalogs_resolve_every_stored_preview() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    for (name, count) in [
        ("powerpoint-visual-objects.ppt", 3),
        ("powerpoint-visual-objects.pps", 3),
        ("powerpoint-visual-objects.pot", 3),
        ("powerpoint-media-objects.ppt", 3),
        ("powerpoint-modern-media.ppt", 7),
        ("powerpoint-master-motion.ppt", 10),
        ("powerpoint-fonts-math-handout.ppt", 3),
    ] {
        let package = Package::open(root.join(name)).unwrap();
        let presentation = package.presentation().unwrap();
        let images = presentation.images().unwrap();
        assert_eq!(images.len(), count, "{name}");
        assert!(
            images.iter().all(|image| !image.data().unwrap().is_empty()),
            "{name}"
        );
    }
}

#[test]
fn native_powerpoint_malformed_custom_show_name_is_an_error_not_an_empty_result() {
    use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
    use std::io::Cursor;
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-navigation-review.ppt"
    );
    let mut source =
        OleFile::open(Cursor::new(bytes.as_slice()), OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let mut data = source.open_stream(&refs).unwrap();
        if refs == ["PowerPoint Document"] {
            let needle = "SHOW_ALPHA_003"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>();
            let offsets = data
                .windows(needle.len())
                .enumerate()
                .filter_map(|(offset, value)| (value == needle).then_some(offset))
                .collect::<Vec<_>>();
            assert_eq!(offsets.len(), 1);
            data[offsets[0]..offsets[0] + 2].copy_from_slice(&0xd800u16.to_le_bytes());
        }
        writer.create_stream(&refs, &data).unwrap();
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let package = Package::from_bytes(out.get_ref()).unwrap();
    assert!(
        matches!(package.presentation().unwrap().custom_shows(),Err(litchi_ppt::Error::Corrupted(message)) if message.contains("invalid UTF-16"))
    );
}

#[cfg(feature = "vba-inspection")]
#[test]
fn native_powerpoint_passive_macro_retains_module_source_without_execution() {
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-passive-macro.ppt"
    );
    let package = Package::from_bytes(bytes).unwrap();
    let presentation = package.presentation().unwrap();
    let storage = presentation.vba_project_storage().unwrap().unwrap();
    assert!(storage.has_macros() && storage.has_persisted_storage());
    assert_eq!(
        storage.compression(),
        Some(litchi_ppt::embedded::storage::Compression::ZlibSyncFlush)
    );
    let project = presentation.vba().unwrap().unwrap();
    assert_eq!(project.name(), "VBAProject");
    assert_eq!(project.modules().len(), 1);
    let module = &project.modules()[0];
    assert_eq!(module.name(), "Module1");
    let text = module.source().text().replace("\r\n", "\n");
    assert_eq!(
        text,
        "Attribute VB_Name = \"Module1\"\nOption Explicit\n\nPublic Sub FixtureNoOp()\n    Dim fixtureMarker As String\n    fixtureMarker = \"PPT_MACRO_005\"\n    \nEnd Sub\n"
    );
    assert_eq!(package.to_bytes().unwrap(), bytes);
}
