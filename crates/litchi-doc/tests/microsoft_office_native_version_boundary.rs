//! Native-source FIB derivatives audit admission, not native-authored Word 6/95 semantics.
use litchi_doc::{Leniency, OpenOptions, Package};
use std::{collections::BTreeMap, io::Cursor, path::Path};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let data = ole
                .open_stream(&path.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (path, data)
        })
        .collect()
}
fn derivative(bytes: &[u8], version: u16) -> Vec<u8> {
    let mut parts = streams(bytes);
    let word = parts.get_mut(&vec!["WordDocument".into()]).unwrap();
    assert_eq!(&word[2..4], &[0xc1, 0]);
    word[2..4].copy_from_slice(&version.to_le_bytes());
    assert!(parts.contains_key(&vec!["1Table".into()]));
    let mut writer = litchi_cfb::OleWriter::new();
    for (path, bytes) in &parts {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        if path.len() > 1 {
            writer.create_storage(&path[..path.len() - 1]).unwrap();
        }
        writer.create_stream(&path, bytes).unwrap();
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let out = out.into_inner();
    assert!(streams(&out) == parts);
    out
}

#[test]
fn native_doc_derivatives_reject_older_fib_profiles_even_with_modern_table_streams() {
    for name in ["word-text-structure.doc", "word-layout-extensions.dot"] {
        let original = native(name);
        let native = Package::from_bytes(&original).unwrap();
        assert_eq!(native.document().fib().version(), 0x00c1);
        for (version, expected_name) in [
            (0x0021, "Word 1.0"),
            (0x0045, "Word 2.0"),
            (0x0065, "Word 6.0"),
            (0x0067, "Word 95 (7.0)"),
            (0x0068, "Word 95 (7.0)"),
        ] {
            let bytes = derivative(&original, version);
            for leniency in [Leniency::Strict, Leniency::TolerateStylesheetDefects] {
                match Package::from_bytes_with_options(
                    &bytes,
                    OpenOptions::default().leniency(leniency),
                ) {
                    Err(litchi_doc::Error::UnsupportedVersion { nfib, name }) => {
                        assert_eq!(nfib, version);
                        assert_eq!(name, expected_name);
                    },
                    Err(e) => {
                        panic!("{name}, {version:#06x}: expected UnsupportedVersion, got {e:?}")
                    },
                    Ok(_) => panic!("{name}: older profile was admitted"),
                }
            }
            assert!(matches!(
                litchi_doc::source_backed::Package::from_bytes(bytes),
                Err(litchi_doc::source_backed::Error::Refused(
                    litchi_doc::source_backed::Refusal::UnsupportedSource
                ))
            ));
        }
        assert!(native.to_bytes().unwrap() == original);
    }
}

#[test]
fn native_word97_shell_and_bidi_derivatives_remain_in_the_modern_profile() {
    for name in ["word-text-structure.doc", "word-layout-extensions.dot"] {
        let original = native(name);
        let source = Package::from_bytes(&original).unwrap();
        let text = source.document().text().unwrap();
        let source_refusal = match litchi_doc::source_backed::Package::from_bytes(original.clone())
        {
            Err(litchi_doc::source_backed::Error::Refused(reason)) => reason,
            other => panic!("{name}: expected the native rich-body edit boundary, got {other:?}"),
        };
        assert_ne!(
            source_refusal,
            litchi_doc::source_backed::Refusal::UnsupportedSource
        );
        for version in [0x00c0, 0x00c2] {
            let bytes = derivative(&original, version);
            let p = Package::from_bytes(&bytes).unwrap();
            assert_eq!(p.document().text().unwrap(), text);
            assert_eq!(p.document().fib().table_pointer_count(), Some(183));
            assert_eq!(p.document().fib().is_template(), name.ends_with(".dot"));
            assert!(p.to_bytes().unwrap() == bytes);
            assert!(
                matches!(litchi_doc::source_backed::Package::from_bytes(bytes),
                Err(litchi_doc::source_backed::Error::Refused(reason)) if reason == source_refusal)
            );
        }
        assert!(source.to_bytes().unwrap() == original);
    }
}
