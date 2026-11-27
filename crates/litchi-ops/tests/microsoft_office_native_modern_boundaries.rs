//! Declared public path gaps complement the owning formats' semantic profiles.
#![cfg(all(feature = "xls", feature = "ppt"))]
use litchi_ops::{Code, OpenOptions, SourceMode};
use std::path::Path;
#[test]
fn native_legacy_readers_refuse_declared_modern_feature_paths() {
    for (file, owner, kinds) in [
        (
            "excel/excel-interaction-print.xls",
            "/sheet[1]",
            &[
                ("threadedComment", Code::UnsupportedPath, false),
                ("slicer", Code::UnsupportedCapability, true),
            ][..],
        ),
        (
            "powerpoint/powerpoint-master-motion.ppt",
            "/slide[1]",
            &[
                ("modernComment", Code::UnsupportedCapability, true),
                ("smartart", Code::UnsupportedCapability, true),
            ][..],
        ),
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(file);
        let source = std::fs::read(&path).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let p = litchi_ops::open_with(&path, &OpenOptions::new().source_mode(mode)).unwrap();
            p.get(owner).unwrap();
            let before = p.to_bytes().unwrap();
            for (kind, code, has_gap) in kinds {
                let target = format!("{owner}/{kind}[1]");
                let error = p.get(target.as_str()).unwrap_err();
                assert_eq!(error.code(), *code, "{file} {mode:?} {target}: {error}");
                let element = p.capabilities().element(kind);
                assert_eq!(element.is_some(), *has_gap);
                if let Some(element) = element {
                    assert_eq!(element.gap().unwrap().code(), *code);
                }
            }
            assert!(p.to_bytes().unwrap() == before);
        }
        assert!(std::fs::read(path).unwrap() == source);
    }
}
