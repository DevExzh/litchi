//! Corpus representation audit only; a picture does NOT qualify SmartArt semantics.
use litchi_core::{Position, edit::RefusalKind};
use litchi_ppt::{Package, RecordType, shapes::ShapeEnum};
use sha2::{Digest, Sha256};
#[test]
fn native_smartart_derivatives_have_a_picture_without_editable_diagram_builds() {
    for name in [
        "powerpoint-visual-objects.ppt",
        "powerpoint-visual-objects.pps",
        "powerpoint-visual-objects.pot",
    ] {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/microsoft-office/powerpoint")
                .join(name),
        )
        .unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let presentation = p.presentation().unwrap();
        let slide = presentation.slide_at(2).unwrap().unwrap();
        assert_eq!(slide.text().unwrap(), "SMARTART_PROCESS_002");
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 2);
        assert_eq!(shapes[0].id(), 4097);
        assert_eq!(shapes[1].id(), 4098);
        let ShapeEnum::Picture(picture) = &shapes[1] else {
            panic!("reviewed legacy picture representation changed")
        };
        let image = picture.extract_image(&presentation).unwrap().unwrap();
        let data = image.data().unwrap();
        assert_eq!(image.extension(), "png");
        assert_eq!(data.len(), 47633);
        assert_eq!(
            Sha256::digest(data)
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect::<String>(),
            "a018e672221c348c6b90c9ed6611f42e2a2333f5b668a3425843bdf2cf348ca7"
        );
        let tags = slide
            .record()
            .versioned_binary_tag_records(10, &litchi_ppt::SourceLimits::default())
            .unwrap();
        let expected: &[RecordType] = if name.ends_with(".ppt") {
            &[RecordType::SlideTime10Atom]
        } else {
            &[]
        };
        assert_eq!(
            tags.iter().map(|r| r.record_type).collect::<Vec<_>>(),
            expected
        );
        let edit = p.edit_diagrams(Position::new(2)).unwrap();
        assert!(
            matches!(edit.builds(),Err(litchi_ppt::diagram::Error::Refused(ref r)) if r.kind()==RefusalKind::UnsupportedTarget)
        );
        assert!(p.to_bytes().unwrap() == bytes);
    }
}

#[cfg(feature = "encryption")]
#[test]
fn native_binary_powerpoint_corpus_lacks_advisor_flags_and_smart_tag_stores() {
    use litchi_ppt::{OpenOptions, ProgBinaryTagVersion, ProgTag};
    let root =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut files = 0;
    let mut encrypted = 0;
    let mut navigation = false;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("powerpoint/")
            || !matches!(
                std::path::Path::new(name)
                    .extension()
                    .and_then(|v| v.to_str()),
                Some("ppt" | "pps" | "pot")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let p = if name == "powerpoint/powerpoint-encrypted.ppt" {
            encrypted += 1;
            Package::from_bytes_with_options(
                &bytes,
                OpenOptions::new().password("OfficeFixture42!"),
            )
            .unwrap()
        } else {
            Package::from_bytes(&bytes).unwrap()
        };
        let d = p.presentation().unwrap();
        assert!(d.slide_count() > 0);
        assert!(
            d.presentation_advisor_settings().unwrap().is_none(),
            "{name}: new advisor evidence; update the review"
        );
        assert!(
            d.smart_tags().unwrap().is_none(),
            "{name}: new smart-tag evidence; update the review"
        );
        let tags = d.programmable_tags().unwrap().unwrap();
        let mut count = 0;
        for tag in &tags.tags {
            if let ProgTag::Binary(tag) = tag {
                assert_ne!(
                    tag.version,
                    ProgBinaryTagVersion::Unknown,
                    "{name}: inspect unknown binary payload before revising this audit"
                );
                for record in tag.records().unwrap() {
                    assert!(
                        !matches!(record.record_type_raw, 0x177a | 0x36b3),
                        "{name}: new target record"
                    );
                    count += 1;
                }
            }
        }
        assert!(
            count >= 2,
            "{name}: meaningful document extension inventory required"
        );
        if name == "powerpoint/powerpoint-navigation-review.ppt" {
            let pp9 = tags.binary_tag(ProgBinaryTagVersion::PowerPoint9).unwrap();
            assert_eq!(
                pp9.records()
                    .unwrap()
                    .iter()
                    .map(|r| r.record_type)
                    .collect::<Vec<_>>(),
                [RecordType::ExternalHyperlink9; 8]
            );
            navigation = true;
        }
        files += 1;
    }
    assert_eq!((files, encrypted, navigation), (11, 1, true));
}
