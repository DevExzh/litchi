#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! The eager slide catalog is parsed at most once per borrowed
//! [`litchi_pptx::presentation::Presentation`] (change 0637). These tests lock
//! the two properties that make the memo value-identical: repeating any
//! catalog query returns what the first call returned, and the memo never
//! moves a refusal — in particular it must not lend `slide_count`'s
//! package-graph validation to `slide`, which has never performed it.

use litchi_core::TextOutputOptions;
use litchi_opc::PackURI;
use litchi_pptx::Package;
use tempfile::NamedTempFile;

/// An authored deck with `slides` slides, round-tripped through the package
/// writer so the catalog and the slide parts are physically present.
fn authored(slides: usize) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    {
        let presentation = package.presentation_mut().unwrap();
        for index in 0..slides {
            presentation
                .add_slide()
                .unwrap()
                .set_title(&format!("Slide {index}"));
        }
    }
    package.save(output.path()).unwrap();
    Package::open(output.path()).unwrap()
}

/// Rewrite `/ppt/presentation.xml` without asking the typed reader to accept
/// the replacement during the edit itself.
fn with_presentation_blob(package: &mut Package, blob: Vec<u8>) {
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(blob.clone());
            Ok(())
        })
        .unwrap();
}

fn presentation_xml(package: &Package) -> Vec<u8> {
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc()
        .unwrap()
        .get_part(&part_name)
        .unwrap()
        .blob()
        .to_vec()
}

#[test]
fn repeated_catalog_queries_return_the_first_answer() {
    let package = authored(3);
    let presentation = package.presentation().unwrap();

    let first_references = presentation.slide_references().unwrap();
    let first_count = presentation.slide_count().unwrap();
    let first_parts: Vec<String> = presentation
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.part().part().partname().to_string())
        .collect();
    let first_text = presentation.text().unwrap();

    assert_eq!(first_count, 3);
    assert_eq!(first_references.len(), 3);

    // Every repeat, in a different order, sees the memoized catalog.
    for _ in 0..3 {
        assert_eq!(presentation.slide_count().unwrap(), first_count);
        assert_eq!(presentation.slide_references().unwrap(), first_references);
        let parts: Vec<String> = presentation
            .slides()
            .unwrap()
            .iter()
            .map(|slide| slide.part().part().partname().to_string())
            .collect();
        assert_eq!(parts, first_parts);
        assert_eq!(presentation.text().unwrap(), first_text);
        for (index, expected) in first_parts.iter().enumerate() {
            let slide = presentation.slide(index).unwrap().unwrap();
            assert_eq!(&slide.part().part().partname().to_string(), expected);
        }
        assert!(presentation.slide(first_count).unwrap().is_none());
    }

    // A second borrow of the same package, with its own empty memo, agrees.
    let fresh = package.presentation().unwrap();
    assert_eq!(fresh.slide_references().unwrap(), first_references);
    assert_eq!(fresh.slide_count().unwrap(), first_count);
    assert_eq!(fresh.text().unwrap(), first_text);
}

#[test]
fn a_malformed_catalog_refuses_on_every_call() {
    let mut package = authored(2);
    with_presentation_blob(
        &mut package,
        br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst></p:presentation>"#
            .to_vec(),
    );
    let presentation = package.presentation().unwrap();

    let first = presentation.slide_references().unwrap_err().to_string();
    for _ in 0..3 {
        assert_eq!(
            presentation.slide_references().unwrap_err().to_string(),
            first
        );
        assert_eq!(presentation.slide_count().unwrap_err().to_string(), first);
        assert_eq!(
            presentation.slide(0).err().map(|error| error.to_string()),
            Some(first.clone())
        );
        assert_eq!(
            presentation.slides().err().map(|error| error.to_string()),
            Some(first.clone())
        );
        assert_eq!(presentation.text().unwrap_err().to_string(), first);
    }
}

#[test]
fn the_memo_does_not_lend_catalog_validation_to_slide() {
    // Slide 0 resolves; slide 1's relationship does not exist. `slide_count`
    // and `slide_references` validate the whole catalog and refuse; `slide(0)`
    // resolves only the reference it was asked for and has never refused.
    let mut package = authored(2);
    let xml = String::from_utf8(presentation_xml(&package)).unwrap();
    let broken = xml.replacen("rId5", "rIdAbsent", 1);
    assert_ne!(
        broken, xml,
        "authored catalog did not contain the second slide relationship rId5"
    );
    with_presentation_blob(&mut package, broken.into_bytes());

    // Filling the unvalidated memo first must not change what the validated
    // queries answer, and vice versa.
    let presentation = package.presentation().unwrap();
    assert!(presentation.slide(0).unwrap().is_some());
    let refusal = presentation.slide_count().unwrap_err().to_string();
    assert!(presentation.slide(0).unwrap().is_some());
    assert_eq!(presentation.slide_count().unwrap_err().to_string(), refusal);
    assert_eq!(
        presentation.slide_references().unwrap_err().to_string(),
        refusal
    );
    assert_eq!(
        presentation.slides().err().map(|error| error.to_string()),
        Some(refusal.clone())
    );

    let reversed = package.presentation().unwrap();
    assert_eq!(reversed.slide_count().unwrap_err().to_string(), refusal);
    assert!(reversed.slide(0).unwrap().is_some());
}

#[test]
fn write_text_to_validates_the_whole_catalog_before_the_first_byte() {
    let mut package = authored(2);
    let xml = String::from_utf8(presentation_xml(&package)).unwrap();
    with_presentation_blob(
        &mut package,
        xml.replacen("rId5", "rIdAbsent", 1).into_bytes(),
    );

    let presentation = package.presentation().unwrap();
    // Filling the unvalidated memo first must not let partial text escape.
    assert!(presentation.slide(0).unwrap().is_some());
    let mut sink = Vec::new();
    let error = presentation
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap_err();
    assert!(sink.is_empty(), "partial text escaped a refused catalog");
    let repeated = {
        let mut sink = Vec::new();
        let error = presentation
            .write_text_to(&mut sink, TextOutputOptions::default())
            .unwrap_err();
        assert!(sink.is_empty());
        format!("{error:?}")
    };
    assert_eq!(format!("{error:?}"), repeated);
}
