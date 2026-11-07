//! Native Word package dialects and the distinction between passwords and IRM.
#![cfg(feature = "encryption")]
use litchi_docx::{Package, package::StoryDialect};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::{
    collections::BTreeMap,
    path::{Path, PathBuf},
};

fn root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office")
}
fn fixtures() -> Vec<String> {
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root().join("manifest.json")).unwrap()).unwrap();
    let names = manifest["files"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|f| {
            let name = f["path"].as_str().unwrap();
            matches!(
                Path::new(name).extension().and_then(|s| s.to_str()),
                Some("docx" | "docm" | "dotx" | "dotm")
            )
            .then(|| name.to_owned())
        })
        .collect::<Vec<_>>();
    assert_eq!(names.len(), 24);
    names
}
fn open(name: &str, bytes: &[u8]) -> Package {
    if name == "word/word-encrypted.docx" {
        let encrypted = Package::from_bytes_with_password(bytes, "Office42!").unwrap();
        // The public story inventory requires an explicit plaintext boundary.
        // Keep that clear package in memory; the native ciphertext is unchanged.
        Package::from_bytes(&encrypted.to_plain_bytes().unwrap()).unwrap()
    } else {
        Package::from_bytes(bytes).unwrap()
    }
}
fn namespace(xml: &[u8]) -> Vec<u8> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (ns, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return match ns {
                    ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                    _ => panic!("Word story has no resolved root namespace"),
                };
            },
            Event::Eof => panic!("empty Word story"),
            _ => {},
        }
    }
}

#[test]
fn native_word_package_variants_have_transitional_main_and_story_dialects() {
    let mut counts = BTreeMap::new();
    let mut secondary = 0;
    let mut encrypted = 0;
    for name in fixtures() {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        let p = open(&name, &bytes);
        let extension = Path::new(&name).extension().unwrap().to_str().unwrap();
        *counts.entry(extension.to_owned()).or_insert(0) += 1;
        let expected = match extension {
            "docx" => ct::WML_DOCUMENT_MAIN,
            "docm" => ct::WML_DOCUMENT_MACRO_MAIN,
            "dotx" => ct::WML_TEMPLATE_MAIN,
            "dotm" => ct::WML_TEMPLATE_MACRO_MAIN,
            _ => unreachable!(),
        };
        let opc = p.opc_package();
        let main = opc.main_document_part().unwrap();
        assert_eq!(main.content_type(), expected);
        let owners = opc
            .rels()
            .iter()
            .filter(|r| r.reltype() == rt::OFFICE_DOCUMENT)
            .collect::<Vec<_>>();
        assert_eq!(owners.len(), 1);
        assert_eq!(owners[0].target_partname().unwrap(), *main.partname());
        assert!(!owners[0].is_external());
        let stories = p.story_inventory().unwrap();
        assert_eq!(stories.dialect(), StoryDialect::Transitional);
        assert_eq!(stories.main(), main.partname());
        for story in stories.stories() {
            assert_eq!(
                namespace(story.source()),
                b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            );
        }
        secondary += stories.stories().len() - 1;
        if name == "word/word-encrypted.docx" {
            assert_eq!(stories.stories().len(), 1);
            encrypted += 1;
        } else {
            assert_eq!(p.to_bytes().unwrap(), bytes);
        }
        assert_eq!(std::fs::read(root().join(name)).unwrap(), bytes);
    }
    assert_eq!(
        counts,
        BTreeMap::from([
            ("docx".into(), 19),
            ("docm".into(), 1),
            ("dotx".into(), 3),
            ("dotm".into(), 1)
        ])
    );
    assert_eq!(secondary, 15);
    assert_eq!(encrypted, 1);
}

#[test]
fn native_word_password_and_edit_protection_do_not_supply_irm_fixture_evidence() {
    use litchi_crypto::spaces;
    let mut encrypted = 0;
    let mut protected = 0;
    for name in fixtures() {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        if name == "word/word-encrypted.docx" {
            encrypted += 1;
            use litchi_crypto::ooxml::{self, Kind, Mode};
            assert_eq!(
                ooxml::inspect(&bytes, &ooxml::Limits::default()).unwrap(),
                Kind::Encrypted(Mode::AgileAes256Sha512)
            );
            let mut ole = litchi_cfb::OleFile::open(
                std::io::Cursor::new(&bytes),
                litchi_cfb::OleFileLimits::default(),
            )
            .unwrap();
            let mut paths = ole.list_streams();
            paths.sort();
            let mut expected = vec![
                vec![spaces::STORAGE, "Version"],
                vec![spaces::STORAGE, "DataSpaceMap"],
                vec![
                    spaces::STORAGE,
                    "DataSpaceInfo",
                    "StrongEncryptionDataSpace",
                ],
                vec![
                    spaces::STORAGE,
                    "TransformInfo",
                    "StrongEncryptionTransform",
                    spaces::PRIMARY,
                ],
                vec!["EncryptionInfo"],
                vec!["EncryptedPackage"],
            ]
            .into_iter()
            .map(|p| p.into_iter().map(str::to_owned).collect::<Vec<_>>())
            .collect::<Vec<_>>();
            expected.sort();
            assert_eq!(paths, expected);
            let limits = spaces::Limits::default();
            let map = spaces::parse_map(
                &ole.open_stream(&[spaces::STORAGE, "DataSpaceMap"]).unwrap(),
                &limits,
            )
            .unwrap();
            assert_eq!(map.entries.len(), 1);
            let entry = &map.entries[0];
            assert_eq!(entry.data_space_name, "StrongEncryptionDataSpace");
            assert_eq!(entry.references.len(), 1);
            assert_eq!(entry.references[0].component, "EncryptedPackage");
            assert_eq!(entry.references[0].kind, spaces::ReferenceKind::Stream);
            let definition = spaces::parse_definition(
                &ole.open_stream(&[
                    spaces::STORAGE,
                    "DataSpaceInfo",
                    "StrongEncryptionDataSpace",
                ])
                .unwrap(),
                &limits,
            )
            .unwrap();
            assert_eq!(definition.transforms, ["StrongEncryptionTransform"]);
            let primary = ole
                .open_stream(&[
                    spaces::STORAGE,
                    "TransformInfo",
                    "StrongEncryptionTransform",
                    spaces::PRIMARY,
                ])
                .unwrap();
            let (header, consumed) = spaces::parse_transform_header(&primary, &limits).unwrap();
            assert_eq!(header.transform_id, spaces::ENCRYPTION_ID);
            assert_eq!(header.transform_name, spaces::ENCRYPTION_NAME);
            // This producer's advisory block size is zero. The OOXML reader
            // above validates the authoritative EncryptionInfo profile.
            assert_eq!(
                &primary[consumed..],
                &[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0]
            );
        } else {
            assert!(bytes.starts_with(b"PK"));
        }
        let p = open(&name, &bytes);
        if name == "word/word-document-protection.docx" {
            assert!(
                p.document()
                    .unwrap()
                    .settings()
                    .unwrap()
                    .unwrap()
                    .is_protected()
            );
            protected += 1;
        }
        assert_eq!(std::fs::read(root().join(name)).unwrap(), bytes);
    }
    assert_eq!((encrypted, protected), (1, 1));
}
