//! Flat (single-XML) OpenDocument through the unified layer (record 0292):
//! `.fodt` and `.fodg` open and read through their families' adapters over
//! the leaves' packaged projections; every flat family publishes its flat
//! bytes unchanged, and an edit of the projection republishes into the flat
//! document element by element, keeping every other byte.

#![cfg(any(feature = "odt", feature = "odg", feature = "ods", feature = "odp"))]
#![cfg_attr(
    not(all(feature = "odt", feature = "odg", feature = "ods", feature = "odp")),
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

/// The bytes outside one changed window are unchanged: the common prefix
/// and suffix of `before` and `after` cover all but `changed` source bytes.
fn assert_local(before: &[u8], after: &[u8], changed: usize) {
    let prefix = before
        .iter()
        .zip(after)
        .take_while(|(left, right)| left == right)
        .count();
    let suffix = before
        .iter()
        .rev()
        .zip(after.iter().rev())
        .take_while(|(left, right)| left == right)
        .count();
    assert!(
        prefix + suffix + changed >= before.len(),
        "the edit rewrote more than {changed} bytes (prefix {prefix}, suffix {suffix}, length {})",
        before.len()
    );
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
    fn writes_paragraph_text_into_the_flat_bytes() {
        const TITLED: &str = "odf/odt/title-field-invalidate.fodt";
        let source = bytes(TITLED);
        let file = open_bytes(source.clone()).expect("open");
        let path = file
            .query("p")
            .unwrap()
            .map(Result::unwrap)
            .find(|node| node.view().text().unwrap().as_deref() == Some("body text"))
            .expect("the body paragraph")
            .path()
            .to_string();
        let edited = file
            .set(path.as_str(), props! { "text" => "new body" })
            .expect("flat paragraph text");
        let published = edited.to_bytes().unwrap();
        assert_local(&source, &published, "body text".len());
        // A second edit of the edited handle, without a reopen.
        let chained = edited
            .set(path.as_str(), props! { "text" => "third body" })
            .expect("chained flat edit");
        assert_eq!(
            String::from_utf8(chained.to_bytes().unwrap()).unwrap(),
            String::from_utf8(source.clone())
                .unwrap()
                .replace(">body text<", ">third body<")
        );
        let reopened = open_bytes(published.clone()).expect("reopen");
        assert_eq!(text(&reopened, &path).as_deref(), Some("new body"));
        assert_eq!(
            String::from_utf8(published).unwrap(),
            String::from_utf8(source)
                .unwrap()
                .replace(">body text<", ">new body<")
        );
        // A transaction of two sets, and metadata, publish flat too.
        let committed = reopened
            .edit(|tx| {
                tx.set(path.as_str(), props! { "text" => "one" })?;
                tx.set("/metadata", props! { "title" => "Flat" })?;
                Ok(())
            })
            .expect("flat transaction");
        let reopened = open_bytes(committed.file().to_bytes().unwrap()).expect("reopen");
        assert_eq!(text(&reopened, &path).as_deref(), Some("one"));
        assert_eq!(
            reopened
                .metadata()
                .unwrap()
                .and_then(|metadata| metadata.title),
            Some("Flat".to_owned())
        );
        assert!(
            String::from_utf8(committed.file().to_bytes().unwrap())
                .unwrap()
                .starts_with("<?xml")
        );
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
        let edited = file
            .set("/page[1]", props! { "name" => "Renamed" })
            .expect("flat page rename");
        let published = edited.to_bytes().unwrap();
        assert_local(&source, &published, 64);
        let chained = edited
            .set("/page[1]", props! { "name" => "Again" })
            .expect("chained flat edit");
        assert_local(&source, &chained.to_bytes().unwrap(), 64);
        let reopened = open_bytes(published).expect("reopen");
        assert_eq!(
            reopened
                .get("/page[1]")
                .unwrap()
                .view()
                .format()
                .unwrap()
                .into_iter()
                .find(|prop| prop.key() == "name")
                .map(|prop| prop.to_string())
                .as_deref(),
            Some("Renamed")
        );
        let report = file.capabilities();
        assert!(report.verb(Verb::Set).is_supported());
    }

    #[test]
    fn a_drawing_without_a_body_is_a_typed_refusal() {
        let error = open_bytes(bytes("odf/drawing/hatch-angles.fodg")).expect_err("no body");
        assert_ne!(error.code(), Code::InternalError);
    }
}

#[cfg(feature = "ods")]
mod ods {
    use super::*;

    const GRID: &str = "odf/ods/grid-padding.fods";

    #[test]
    fn writes_cell_values_into_the_flat_bytes() {
        let source = bytes(GRID);
        let file = open_bytes(source.clone()).expect("open .fods");
        let edited = file
            .set("/Padded/A1", props! { "value" => "changed" })
            .expect("flat cell value");
        let published = edited.to_bytes().unwrap();
        let reopened = open_bytes(published.clone()).expect("reopen");
        assert_eq!(text(&reopened, "/Padded/A1").as_deref(), Some("changed"));
        // Only the edited row changed.
        assert_local(&source, &published, 2_048);
        assert!(published.len() < source.len() + 2_048);
    }
}

#[cfg(feature = "odp")]
mod odp {
    use super::*;

    const MATH: &str = "libreoffice-core/sd/qa/unit/data/odp/Math.fodp";

    #[test]
    fn writes_slide_names_into_the_flat_bytes() {
        let source = bytes(MATH);
        let file = open_bytes(source.clone()).expect("open .fodp");
        let edited = file
            .set("/slide[1]", props! { "name" => "Renamed" })
            .expect("flat slide name");
        let published = edited.to_bytes().unwrap();
        assert_local(&source, &published, 16);
        let reopened = open_bytes(published).expect("reopen");
        assert_eq!(
            reopened
                .get("/slide[1]")
                .unwrap()
                .view()
                .format()
                .unwrap()
                .into_iter()
                .find(|prop| prop.key() == "name")
                .map(|prop| prop.to_string())
                .as_deref(),
            Some("Renamed")
        );
    }
}
