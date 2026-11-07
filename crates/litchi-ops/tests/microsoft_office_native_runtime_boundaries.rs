//! Evidence of runtime boundaries; these tests do not qualify rendering or field execution.
#![cfg(all(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "encryption",
    feature = "eval"
))]
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, OpenOptions, Password, SourceMode};
use std::{fmt, path::Path};

#[derive(Default)]
struct RefusingSink {
    calls: usize,
}
impl fmt::Write for RefusingSink {
    fn write_str(&mut self, _: &str) -> fmt::Result {
        self.calls += 1;
        Err(fmt::Error)
    }
}
fn refused_views(file: &File, name: &str, expected: Code) {
    use litchi_ops::capabilities::Support;
    for mode in [Mode::Html, Mode::Svg, Mode::Screenshot, Mode::Pdf] {
        assert_eq!(
            file.capabilities().view(mode).code(),
            Some(expected),
            "{name}: {mode:?}"
        );
        let Support::Unsupported { reason, .. } = file.capabilities().view(mode) else {
            panic!("{name}: unexpected visual renderer")
        };
        assert_eq!(
            reason,
            if expected == Code::DocumentProtected {
                "the document is encrypted; reads need its password"
            } else {
                "visual rendering is outside the unified layer"
            }
        );
        let error = file.view(mode).unwrap_err();
        assert_eq!(error.code(), expected, "{name}: {mode:?}: {error}");
        if expected == Code::UnsupportedCapability {
            assert!(error.message().contains(mode.name()), "{name}: {error}");
        }
        let error = file.view_json(mode).unwrap_err();
        assert_eq!(error.code(), expected, "{name}: JSON {mode:?}: {error}");
        let mut sink = RefusingSink::default();
        let error = file.view_to(mode, &mut sink).unwrap_err();
        assert_eq!(error.code(), expected, "{name}: sink {mode:?}: {error}");
        assert_eq!(
            sink.calls, 0,
            "{name}: unsupported view wrote partial output"
        );
    }
}

#[test]
fn native_office_corpus_visual_views_refuse_with_no_output_in_both_ingress_modes() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut count = 0;
    let mut protected = 0;
    let mut unlocked = 0;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        let path = root.join(name);
        let bytes = std::fs::read(&path).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let options = OpenOptions::new().source_mode(mode);
            let file = if mode == SourceMode::Auto {
                litchi_ops::open_with(&path, &options).unwrap()
            } else {
                litchi_ops::open_bytes_with(bytes.clone(), &options).unwrap()
            };
            if let File::Locked(locked) = &file {
                protected += 1;
                assert!(
                    name.contains("-encrypted."),
                    "{name}: unexpected locked source"
                );
                refused_views(&file, name, Code::DocumentProtected);
                let password = if name.starts_with("word/") {
                    "Office42!"
                } else {
                    "OfficeFixture42!"
                };
                let result = locked.clone().unlock(Password::new(password.to_owned()));
                if name == "excel/excel-encrypted.xlsb" {
                    let error = result.unwrap_err();
                    assert_eq!(error.code(), Code::UnsupportedCapability);
                    assert_eq!(
                        error.message(),
                        "an encrypted XLSB package cannot be decrypted: the XLSB reader has no decryption support"
                    );
                } else {
                    let clear = result.unwrap_or_else(|e| panic!("{name}: {e}"));
                    refused_views(&clear, name, Code::UnsupportedCapability);
                    unlocked += 1;
                }
            } else {
                assert!(
                    !name.contains("-encrypted."),
                    "{name}: encrypted source bypassed Locked"
                );
                let before = file.to_bytes().unwrap();
                refused_views(&file, name, Code::UnsupportedCapability);
                assert!(
                    file.to_bytes().unwrap() == before,
                    "{name}: read changed publication"
                );
            }
            assert!(
                std::fs::read(&path).unwrap() == bytes,
                "{name}: input changed"
            );
            count += 1;
        }
    }
    assert_eq!((count, protected, unlocked), (278, 14, 12));
}

#[test]
fn native_word_field_and_mail_merge_inputs_refuse_worksheet_evaluation() {
    use litchi_ops::calc::Options;
    use litchi_ops::capabilities::Verb;
    for extension in ["doc", "docx", "rtf"] {
        for base in [
            "word-mailmerge-fields",
            "word-math-navigation",
            "word-layout-extensions",
        ] {
            let name = format!("{base}.{extension}");
            let path = Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/microsoft-office/word")
                .join(&name);
            let bytes = std::fs::read(&path).unwrap();
            for mode in [SourceMode::Auto, SourceMode::Eager] {
                let options = OpenOptions::new().source_mode(mode);
                let file = if mode == SourceMode::Auto {
                    litchi_ops::open_with(&path, &options).unwrap()
                } else {
                    litchi_ops::open_bytes_with(bytes.clone(), &options).unwrap()
                };
                let before = file.to_bytes().unwrap();
                let text = file.view(Mode::Text).unwrap();
                assert!(!text.is_empty());
                if base == "word-mailmerge-fields" {
                    if extension == "rtf" {
                        // RTF field results have separate story owners and do
                        // not enter its ordinary paragraph-text projection.
                        let package = litchi_rtf::Package::from_bytes(&before).unwrap();
                        assert!(package.fields().iter().any(|field| {
                            field.instruction.trim_start().starts_with("IF ")
                                && field.result == "MATCH_009"
                        }));
                    } else {
                        assert!(text.contains("MATCH_009"), "{name}: cached IF result");
                    }
                }
                assert_eq!(
                    file.capabilities().verb(Verb::Evaluate).code(),
                    Some(Code::UnsupportedCapability)
                );
                assert!(matches!(
                    file.capabilities().verb(Verb::Evaluate),
                    litchi_ops::capabilities::Support::Unsupported {
                        reason: "this format has no worksheets",
                        ..
                    }
                ));
                let e = file
                    .evaluate("/", &Options::new().recalculate(true))
                    .unwrap_err();
                assert_eq!(e.code(), Code::UnsupportedCapability, "{name}: {e}");
                let e = file.recalculate(&Options::new()).unwrap_err();
                assert_eq!(e.code(), Code::UnsupportedCapability, "{name}: {e}");
                assert_eq!(file.view(Mode::Text).unwrap(), text);
                assert!(file.to_bytes().unwrap() == before);
            }
            assert!(std::fs::read(path).unwrap() == bytes);
        }
    }
}
