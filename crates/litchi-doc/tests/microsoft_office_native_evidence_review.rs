//! Evidence-only audits: page setup and equation pictures do not qualify other owners.
use litchi_doc::{OpenOptions, Package, Password};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, io::Cursor};
fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office")
}
fn files() -> Vec<String> {
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root().join("manifest.json")).unwrap()).unwrap();
    manifest["files"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|f| {
            let n = f["path"].as_str().unwrap();
            (n.starts_with("word/")
                && matches!(
                    std::path::Path::new(n).extension().and_then(|v| v.to_str()),
                    Some("doc" | "dot")
                ))
            .then(|| n.to_owned())
        })
        .collect()
}
fn open(name: &str, bytes: &[u8]) -> Package {
    if name == "word/word-encrypted.doc" {
        Package::from_bytes_with_options(
            bytes,
            OpenOptions::default().password(Password::new("Office42!".into())),
        )
        .unwrap()
    } else {
        Package::from_bytes(bytes).unwrap()
    }
}
fn hash(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
#[test]
fn native_binary_word_corpus_has_no_saved_selection_or_printer_metadata_records() {
    let names = files();
    assert_eq!(names.len(), 20);
    let mut encrypted = 0;
    for name in names {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        let p = open(&name, &bytes);
        let d = p.document();
        encrypted += usize::from(d.fib().is_encrypted());
        assert!(
            d.saved_selection().unwrap().is_none(),
            "{name}: new native selection evidence"
        );
        assert!(
            d.print_environment().unwrap().is_none(),
            "{name}: new native printer metadata evidence"
        );
        for index in [27, 28, 29, 30] {
            let (_, length) = d.fib().get_table_pointer(index).unwrap();
            assert_eq!(length, 0, "{name}: FIB slot {index}");
        }
        assert!(
            !d.body_paragraphs()
                .collect::<Result<Vec<_>, _>>()
                .unwrap()
                .is_empty()
        );
        assert!(p.to_bytes().unwrap() == bytes);
    }
    assert_eq!(encrypted, 1);
}
#[test]
fn native_binary_word_math_has_emf_fallbacks_without_equation_native_storage() {
    let names = files();
    assert_eq!(names.len(), 20);
    let mut math = false;
    let mut workbook_storages = 0;
    for name in names {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        let p = open(&name, &bytes);
        let mut ole =
            litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
                .unwrap();
        let streams = ole.list_streams();
        for path in &streams {
            assert!(
                !["Equation Native", "MSWordEquation", "Equation.3"]
                    .iter()
                    .any(|n| path.last().unwrap().eq_ignore_ascii_case(n)),
                "{name}: new equation stream {path:?}"
            );
            if path.last().is_some_and(|s| s == "\u{1}CompObj") {
                let names = path.iter().map(String::as_str).collect::<Vec<_>>();
                let data = ole.open_stream(&names).unwrap();
                assert!(data.len() >= 28);
                assert_ne!(&data[12..28], &litchi_doc::equation::EQUATION_3_CLSID);
                if path.first().is_some_and(|s| s == "ObjectPool") {
                    assert_eq!(name, "word/word-visual-objects.doc");
                    workbook_storages += 1;
                }
            }
        }
        if name == "word/word-math-navigation.doc" {
            assert!(
                streams
                    .iter()
                    .all(|p| p.first().is_none_or(|v| v != "ObjectPool"))
            );
            let d = p.document();
            let mut images = BTreeMap::new();
            for para in d.paragraphs().unwrap() {
                for run in para.runs().unwrap() {
                    if let Some(image) = run.image() {
                        let data = d.image_data(image).unwrap();
                        assert_eq!(data.kind(), litchi_odraw::image::Kind::Emf);
                        let payload = data.data().unwrap();
                        images.insert(image.pic_offset(), (payload.len(), hash(payload)));
                    }
                }
            }
            assert_eq!(
                images,
                BTreeMap::from([
                    (
                        500,
                        (
                            12398,
                            "470418143d90103432afb25d7384cbe51198149da0bda81523c4f1fb26a813e4"
                                .into()
                        )
                    ),
                    (
                        18383,
                        (
                            12398,
                            "470418143d90103432afb25d7384cbe51198149da0bda81523c4f1fb26a813e4"
                                .into()
                        )
                    ),
                    (
                        36266,
                        (
                            12470,
                            "cf898ccce10b71d9cd38e04ade7450c5801ecfe4ff944da3b409789f5299962b"
                                .into()
                        )
                    )
                ])
            );
            math = true;
        }
        assert!(p.to_bytes().unwrap() == bytes);
    }
    assert!(math);
    assert_eq!(workbook_storages, 2);
}
