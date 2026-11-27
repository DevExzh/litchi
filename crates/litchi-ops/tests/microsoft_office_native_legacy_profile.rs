//! Native-source invalid-profile derivatives; no legacy Word interpretation is claimed.
#![cfg(feature = "doc")]
use litchi_ops::{Code, OpenOptions, SourceMode};
use std::{io::Cursor, path::Path};
#[test]
fn native_doc_profile_derivatives_keep_public_unsupported_format_errors() {
    for name in ["word-text-structure.doc", "word-layout-extensions.dot"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name);
        let original = std::fs::read(&path).unwrap();
        for version in [0x0065u16, 0x0067] {
            let mut source = litchi_cfb::OleFile::open(
                Cursor::new(&original),
                litchi_cfb::OleFileLimits::default(),
            )
            .unwrap();
            let mut writer = litchi_cfb::OleWriter::new();
            for path in source.list_streams() {
                let names = path.iter().map(String::as_str).collect::<Vec<_>>();
                let mut bytes = source.open_stream(&names).unwrap();
                if names == ["WordDocument"] {
                    assert_eq!(&bytes[2..4], &[0xc1, 0]);
                    bytes[2..4].copy_from_slice(&version.to_le_bytes());
                }
                if names.len() > 1 {
                    writer.create_storage(&names[..names.len() - 1]).unwrap();
                }
                writer.create_stream(&names, &bytes).unwrap();
            }
            let mut output = Cursor::new(Vec::new());
            writer.write_to(&mut output).unwrap();
            let bytes = output.into_inner();
            let dir = tempfile::tempdir().unwrap();
            let changed = dir.path().join(name);
            std::fs::write(&changed, &bytes).unwrap();
            for mode in [SourceMode::Auto, SourceMode::Eager] {
                let options = OpenOptions::new().source_mode(mode);
                let error = if mode == SourceMode::Auto {
                    litchi_ops::open_with(&changed, &options)
                } else {
                    litchi_ops::open_bytes_with(bytes.clone(), &options)
                }
                .unwrap_err();
                assert_eq!(error.code(), Code::UnsupportedFormat, "{name}: {error}");
                assert!(
                    matches!(error.leaf(),litchi_ops::error::Source::Doc(litchi_doc::Error::UnsupportedVersion {nfib,..}) if *nfib==version)
                );
            }
            assert!(std::fs::read(changed).unwrap() == bytes);
        }
        assert!(std::fs::read(path).unwrap() == original);
    }
}
