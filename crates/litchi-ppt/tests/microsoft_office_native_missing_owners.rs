//! Evidence-only audits: missing native owners do not qualify feature support.
#![cfg(feature = "encryption")]
use litchi_core::edit::RefusalKind;
use litchi_ppt::{OpenOptions, Package, Record, RecordType};
use std::{io::Cursor, path::Path};

fn corpus(mut check: impl FnMut(&str, &[u8], &Package)) {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut files = 0;
    let mut encrypted = 0;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("powerpoint/")
            || !matches!(
                Path::new(name).extension().and_then(|e| e.to_str()),
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
        assert!(p.presentation().unwrap().slide_count() > 0, "{name}");
        check(name, &bytes, &p);
        if p.presentation().unwrap().is_encrypted() {
            assert!(matches!(
                p.to_bytes(),
                Err(litchi_ppt::package::Error::EncryptionPolicy)
            ));
        } else {
            assert!(p.to_bytes().unwrap() == bytes, "{name}: source changed");
        }
        assert!(std::fs::read(root.join(name)).unwrap() == bytes);
        files += 1;
    }
    assert_eq!((files, encrypted), (11, 1));
}

#[test]
fn native_binary_powerpoint_corpus_has_no_custom_xml_data_store() {
    use litchi_ole_common::custom_xml::{self, Limits};
    corpus(|name, bytes, _| {
        let mut ole =
            litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
                .unwrap();
        assert!(ole.open_stream(&["PowerPoint Document"]).unwrap().len() > 100);
        for storage in [
            custom_xml::STORE_STORAGE,
            custom_xml::REDUNDANT_PROMOTION_STORAGE,
            custom_xml::MODIFIED_PROMOTION_STORAGE,
        ] {
            assert!(
                !ole.directory_exists(&[storage]),
                "{name}: new {storage} evidence"
            );
            assert!(!ole.exists(&[storage]), "{name}: new {storage} entry");
        }
        assert!(
            custom_xml::inspect(&mut ole, Limits::default())
                .unwrap()
                .is_none(),
            "{name}"
        );
    });
}

#[test]
fn native_binary_powerpoint_corpus_has_no_reviewer_trees_or_modify_password_atoms() {
    let mut pp10_records = 0;
    let mut toolbars = 0;
    corpus(|name, _, p| {
        let d = p.presentation().unwrap();
        assert!(
            d.modify_password().unwrap().is_none(),
            "{name}: modify password is distinct from open encryption"
        );
        let snapshot = d.document_comparison().unwrap();
        assert_eq!(snapshot.review().unwrap().diff_tree_count(), 0, "{name}");
        assert!(snapshot.diff_tree(0).unwrap().is_none());
        toolbars += usize::from(snapshot.toolbar().unwrap().is_some());
        let (document, consumed) =
            Record::parse(snapshot.bytes(), 0, &litchi_ppt::SourceLimits::default()).unwrap();
        assert_eq!(consumed, snapshot.bytes().len());
        assert_eq!(document.record_type, RecordType::Document);
        // These owners are direct members of the live PP10 document extension,
        // whose BinaryTagData atom is opaque to an ordinary children walk.
        let records = document
            .versioned_binary_tag_records(10, &litchi_ppt::SourceLimits::default())
            .unwrap();
        for r in &records {
            assert_ne!(r.record_type, RecordType::DiffTree10, "{name}");
            assert!(
                !(r.record_type == RecordType::CString && r.instance == 3),
                "{name}: ModifyPasswordAtom"
            );
        }
        pp10_records += records.len();
        if !d.is_encrypted() {
            let mut edit = p.edit_document_comparison().unwrap();
            assert!(
                matches!(edit.set_reviewer_name(0, "Absent reviewer"), Err(litchi_ppt::document_comparison::Error::Refused(ref r)) if r.kind() == RefusalKind::NotFound)
            );
            assert!(edit.changes().unwrap().is_empty());
            assert!(edit.commit().unwrap().package().to_bytes().unwrap() == p.to_bytes().unwrap());
        }
    });
    assert_eq!((pp10_records, toolbars), (12, 0));
}

// PPT and OfficeArt use the same eight-byte record framing. Walk exact spans,
// including OfficeArtClientData (whose payload is PPT records), and explicitly
// enter BinaryTagData. Do not interpret arbitrary atom/image payloads as records.
fn wire_records(bytes: &[u8], depth: usize, visit: &mut impl FnMut(u16)) {
    assert!(depth < 128);
    let mut offset = 0;
    while offset < bytes.len() {
        let (record, consumed) =
            litchi_odraw::Record::parse(bytes, offset, litchi_odraw::Limits::default()).unwrap();
        assert!(consumed >= 8);
        visit(record.raw_kind());
        if record.is_container() || record.raw_kind() == RecordType::BinaryTagData.as_u16() {
            wire_records(record.data(), depth + 1, visit);
        }
        offset += consumed;
    }
    assert_eq!(offset, bytes.len());
}

#[test]
fn native_binary_powerpoint_images_and_slide_records_do_not_evidence_wmf_recoloring() {
    let mut formats = std::collections::BTreeMap::new();
    let mut slide_drawings = 0;
    let mut plain_streams = 0;
    corpus(|name, bytes, p| {
        let d = p.presentation().unwrap();
        for image in d.images().unwrap() {
            let extension = image.extension();
            assert_ne!(extension, "wmf", "{name}: new native WMF evidence");
            assert!(!image.data().unwrap().is_empty());
            *formats.entry(extension.to_string()).or_insert(0) += 1;
        }
        for slide in d.slides().unwrap() {
            wire_records(&slide.record().data, 0, &mut |kind| {
                assert_ne!(kind, RecordType::RecolorInfoAtom.as_u16(), "{name}");
                slide_drawings += usize::from(kind == RecordType::PPDrawing.as_u16());
            });
        }
        if !d.is_encrypted() {
            let mut ole =
                litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
                    .unwrap();
            let stream = ole.open_stream(&["PowerPoint Document"]).unwrap();
            // Also include unreferenced, historical and master/notes owners.
            wire_records(&stream, 0, &mut |kind| {
                assert_ne!(kind, RecordType::RecolorInfoAtom.as_u16(), "{name}")
            });
            plain_streams += 1;
        }
    });
    assert_eq!(
        formats,
        [
            ("emf".to_string(), 2),
            ("jpg".to_string(), 4),
            ("png".to_string(), 26)
        ]
        .into()
    );
    assert_eq!((slide_drawings, plain_streams), (49, 10));
}

#[test]
fn native_binary_powerpoint_slides_have_no_diagram_build_records() {
    use litchi_core::Position;
    let mut slides = 0;
    let mut paragraph_builds = 0;
    let mut empty_build_lists = 0;
    let mut missing_build_lists = 0;
    corpus(|name, _, p| {
        let d = p.presentation().unwrap();
        for (index, slide) in d.slides().unwrap().iter().enumerate() {
            let records = slide
                .record()
                .versioned_binary_tag_records(10, &litchi_ppt::SourceLimits::default())
                .unwrap();
            fn inspect(record: &Record, name: &str, paragraphs: &mut usize) {
                assert!(
                    !matches!(
                        record.record_type,
                        RecordType::DiagramBuild | RecordType::DiagramBuildAtom
                    ),
                    "{name}: new diagram build evidence"
                );
                *paragraphs += usize::from(record.record_type == RecordType::ParaBuild);
                for child in &record.children {
                    inspect(child, name, paragraphs);
                }
            }
            for r in &records {
                inspect(r, name, &mut paragraph_builds);
            }
            // Independently scan exact envelopes so a tolerant child projection
            // cannot hide a diagram record from this absence assertion.
            wire_records(&slide.record().data, 0, &mut |kind| {
                assert!(
                    ![
                        RecordType::DiagramBuild.as_u16(),
                        RecordType::DiagramBuildAtom.as_u16()
                    ]
                    .contains(&kind),
                    "{name}"
                );
            });
            if !d.is_encrypted() {
                let edit = p.edit_diagrams(Position::new(index)).unwrap();
                match edit.builds() {
                    Ok(builds) => {
                        assert_eq!(builds.len(), 0, "{name}");
                        empty_build_lists += 1;
                    },
                    Err(litchi_ppt::diagram::Error::Refused(r))
                        if r.kind() == RefusalKind::UnsupportedTarget =>
                    {
                        missing_build_lists += 1;
                    },
                    Err(e) => panic!("{name} slide {index}: unexpected {e}"),
                }
            }
            slides += 1;
        }
    });
    assert_eq!(
        (
            slides,
            paragraph_builds,
            empty_build_lists,
            missing_build_lists
        ),
        (49, 1, 6, 42)
    );
}
