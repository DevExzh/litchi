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
