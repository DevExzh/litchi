//! Flat (single-XML) OpenDocument through the unified layer (record 0292):
//! `.fodt` and `.fodg` open and read through their families' adapters over
//! the leaves' packaged read projections, and publish their flat bytes
//! unchanged.

#![cfg(any(feature = "odt", feature = "odg"))]
#![cfg_attr(
    not(all(feature = "odt", feature = "odg")),
    allow(dead_code, unused_imports)
)]

use litchi_core::FileFormat;
use litchi_ops::capabilities::Verb;
use litchi_ops::{Code, File, open, open_bytes, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path)
}

fn bytes(path: &str) -> Vec<u8> {
    std::fs::read(fixture(path)).expect("read fixture")
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("path")
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

#[cfg(feature = "odt")]
mod odt {
    use super::*;

    const NESTED: &str = "odf/odt/note-nested.fodt";

    #[test]
    fn opens_reads_and_republishes_a_flat_text_document() {
        let source = bytes(NESTED);
        let file = open_bytes(source.clone()).expect("open .fodt");
        assert_eq!(file.format(), Some(FileFormat::Odt));
        assert!(matches!(file, File::Odt(_)));
        let first = text(&file, "/body/p[1]").expect("paragraph text");
        assert!(first.starts_with("Lorem ipsum"), "{first}");
        assert!(file.query("footnote").unwrap().count() >= 1);
        assert!(file.text().unwrap().contains("Lorem ipsum"));
        // Publication is the flat document, byte for byte.
        assert_eq!(file.to_bytes().unwrap(), source);
        let reopened = open(fixture(NESTED)).expect("open path");
        assert_eq!(reopened.to_bytes().unwrap(), source);
    }

    #[test]
    fn every_corpus_fodt_opens() {
        let mut stack = vec![fixture("")];
        let mut opened = 0;
        while let Some(dir) = stack.pop() {
            for entry in std::fs::read_dir(&dir).unwrap().flatten() {
                let path = entry.path();
                if path.is_dir() {
                    stack.push(path);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension == "fodt")
                {
                    let file = open(&path).unwrap_or_else(|error| {
                        panic!("{}: {error}", path.display());
                    });
                    assert_eq!(file.format(), Some(FileFormat::Odt));
                    file.text().unwrap();
                    opened += 1;
                }
            }
        }
        assert!(opened >= 40, "{opened}");
    }
}

#[cfg(feature = "odg")]
mod odg {
    use super::*;

    const LATHE: &str = "libreoffice-core/xmloff/qa/unit/data/tdf161327_LatheEndAngle.fodg";

    #[test]
    fn opens_reads_and_republishes_a_flat_drawing() {
        let source = bytes(LATHE);
        let file = open_bytes(source.clone()).expect("open .fodg");
        assert_eq!(file.format(), Some(FileFormat::Odg));
        assert!(matches!(file, File::Odg(_)));
        assert!(file.query("page").unwrap().count() >= 1);
        assert!(file.query("shape").unwrap().count() >= 1);
        assert_eq!(file.to_bytes().unwrap(), source);
        let error = file
            .set("/page[1]", props! { "name" => "Renamed" })
            .expect_err("flat drawings are read-only");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let report = file.capabilities();
        assert!(report.verb(Verb::Get).is_supported());
        assert!(!report.verb(Verb::Set).is_supported());
    }

    #[test]
    fn a_drawing_without_a_body_is_a_typed_refusal() {
        let error = open_bytes(bytes("odf/drawing/hatch-angles.fodg")).expect_err("no body");
        assert_ne!(error.code(), Code::InternalError);
    }
}
