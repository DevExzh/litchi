//! Change 0755: malformed DrawingML text-run markup is refused with a typed
//! error by every reader and opened-transaction verb, never a panic.
//!
//! An empty `a:t` inside an open `a:t` used to pass the scene reader and be
//! recorded by the text-run locator as a span that starts inside the
//! enclosing element's span; `set_shape_text` then sliced the slide with a
//! reversed range and panicked (`panic = "abort"` in release).

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused regression tests use panic-on-fixture-failure assertions"
)]

use std::panic::{AssertUnwindSafe, catch_unwind};

use litchi_opc::PackURI;

use super::ShapeTextReplacement;
use crate::{Error, Package, Result};

const SLIDE: &str = "/ppt/slides/slide1.xml";

fn text_box_package(texts: &[&str]) -> Result<Package> {
    let mut package = Package::new()?;
    let slide = package.presentation_mut()?.add_slide()?;
    for (index, text) in texts.iter().enumerate() {
        let offset = i64::try_from(index).map_err(|_| Error::Invalid("offset".into()))?;
        slide.add_text_box(text, 10, 20 + offset * 500, 300, 400);
    }
    Package::from_vec(package.to_bytes()?)
}

fn slide_xml(package: &Package) -> Result<String> {
    let name = PackURI::new(SLIDE).map_err(Error::Invalid)?;
    std::str::from_utf8(package.opc.get_part(&name)?.blob())
        .map(str::to_owned)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn set_slide_xml(package: &mut Package, xml: Vec<u8>) -> Result<()> {
    let name = PackURI::new(SLIDE).map_err(Error::Invalid)?;
    package.opc.get_part_mut(&name)?.set_blob(xml);
    Ok(())
}

/// The reported reproduction: `<a:t>My` becomes `<a:t>M<a:t/>y`.
fn nested_empty_run_package() -> Result<Package> {
    let mut package = text_box_package(&["My text"])?;
    let xml = slide_xml(&package)?;
    let nested = xml.replacen("<a:t>My", "<a:t>M<a:t/>y", 1);
    assert_ne!(nested, xml, "the fixture edit must apply");
    set_slide_xml(&mut package, nested.into_bytes())?;
    Ok(package)
}

fn is_nested_text_refusal(error: &Error) -> bool {
    matches!(error, Error::Invalid(message) if message.contains("nested DrawingML text elements"))
}

#[test]
fn a_nested_empty_text_run_is_refused_by_set_shape_text_not_a_panic() -> Result<()> {
    let package = nested_empty_run_package()?;
    // Capture validates the package graph and the notes roots, not text runs,
    // so it still opens the deck; every verb that reads the runs refuses.
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    let outcome = catch_unwind(AssertUnwindSafe(|| edit.set_shape_text(0, 0, "x")));
    let error = outcome
        .expect("set_shape_text must not panic")
        .expect_err("a nested text run must be refused");
    assert!(is_nested_text_refusal(&error), "{error:?}");
    let batch = catch_unwind(AssertUnwindSafe(|| {
        edit.set_shape_texts(0, &[ShapeTextReplacement::at(0, "x")])
    }));
    let error = batch
        .expect("set_shape_texts must not panic")
        .expect_err("a nested text run must be refused");
    assert!(is_nested_text_refusal(&error), "{error:?}");
    Ok(())
}

#[test]
fn the_scene_reader_and_semantic_text_refuse_nested_text_runs_alike() -> Result<()> {
    let package = nested_empty_run_package()?;
    let presentation = package.presentation()?;
    let slide = presentation.slide(0)?.expect("slide");
    let scene_error = slide.shapes().expect_err("the scene reader must refuse");
    assert!(is_nested_text_refusal(&scene_error), "{scene_error:?}");
    let text_error = slide.text().expect_err("semantic text must refuse");
    assert!(is_nested_text_refusal(&text_error), "{text_error:?}");

    // A nested start tag was already refused by both, and still is.
    let mut package = text_box_package(&["My text"])?;
    let xml = slide_xml(&package)?.replacen("<a:t>My", "<a:t>M<a:t>z</a:t>y", 1);
    set_slide_xml(&mut package, xml.into_bytes())?;
    let presentation = package.presentation()?;
    let slide = presentation.slide(0)?.expect("slide");
    assert!(is_nested_text_refusal(&slide.shapes().expect_err("scene")));
    assert!(is_nested_text_refusal(&slide.text().expect_err("text")));
    Ok(())
}

#[test]
fn well_formed_empty_text_runs_still_edit() -> Result<()> {
    // Sibling empty runs are legitimate and keep working.
    let mut package = text_box_package(&["My text"])?;
    let xml = slide_xml(&package)?.replacen("</a:t></a:r>", "</a:t></a:r><a:r><a:t/></a:r>", 1);
    set_slide_xml(&mut package, xml.into_bytes())?;
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    assert!(edit.set_shape_text(0, 0, "edited")?);
    let commit = edit.commit()?;
    let scene_text = crate::shape::Scene::read(
        commit
            .snapshot()
            .package
            .get_part(&PackURI::new(SLIDE).map_err(Error::Invalid)?)?
            .blob(),
    )?
    .at(0)
    .map_err(|error| Error::Invalid(error.to_string()))?
    .common()
    .text()
    .map(str::to_owned);
    assert_eq!(scene_text.as_deref(), Some("edited"));
    Ok(())
}

/// Deterministic markup mutations inside the shapes' text bodies.
fn text_body_mutations(xml: &str) -> Vec<String> {
    const SNIPPETS: &[&str] = &[
        "<a:t/>",
        "<a:t>",
        "</a:t>",
        "<a:t>z</a:t>",
        "<a:t xml:space=\"preserve\"/>",
        "<a:r>",
        "</a:r>",
        "<a:r><a:t/></a:r>",
        "<a:br/>",
        "<!---->",
        "&amp;",
        "<![CDATA[c]]>",
        "<q:t/>",
        "<a:p>",
        "</a:p>",
    ];
    let mut variants = Vec::new();
    let mut search = 0usize;
    while let Some(relative) = xml[search..].find("<p:txBody>") {
        let body_start = search + relative;
        let body_end = body_start + xml[body_start..].find("</p:txBody>").expect("closed body");
        for position in body_start..=body_end {
            if !xml.is_char_boundary(position) {
                continue;
            }
            for snippet in SNIPPETS {
                let mut variant = String::with_capacity(xml.len() + snippet.len());
                variant.push_str(&xml[..position]);
                variant.push_str(snippet);
                variant.push_str(&xml[position..]);
                variants.push(variant);
            }
        }
        search = body_end + 1;
    }
    variants
}

#[test]
fn no_text_verb_panics_on_mutated_text_run_markup() -> Result<()> {
    let base = text_box_package(&["alpha", "b", "gamma ray"])?;
    let xml = slide_xml(&base)?;
    let variants = text_body_mutations(&xml);
    assert!(variants.len() > 1_000, "expected a broad mutation set");
    let mut refused = 0usize;
    let mut edited = 0usize;
    for (index, variant) in variants.iter().enumerate() {
        let mut package = Package::from_opc_package(base.opc.clone())?;
        set_slide_xml(&mut package, variant.clone().into_bytes())?;
        let outcome = catch_unwind(AssertUnwindSafe(|| -> (usize, usize) {
            let mut refused = 0usize;
            let mut edited = 0usize;
            if let Ok(presentation) = package.presentation()
                && let Ok(Some(slide)) = presentation.slide(0)
            {
                let _ = slide.shapes();
                let _ = slide.text();
            }
            let Ok(snapshot) = package.opened_presentation() else {
                return (1, 0);
            };
            for shape in 0..3 {
                let mut edit = snapshot.edit();
                match edit.set_shape_text(0, shape, "replacement & <text>") {
                    Ok(_) => match edit.commit() {
                        Ok(_) => edited += 1,
                        Err(_) => refused += 1,
                    },
                    Err(_) => refused += 1,
                }
                let mut batch = snapshot.edit();
                if batch
                    .set_shape_texts(0, &[ShapeTextReplacement::at(shape, "batch")])
                    .is_err()
                {
                    refused += 1;
                }
            }
            (refused, edited)
        }));
        let (variant_refused, variant_edited) =
            outcome.unwrap_or_else(|_| panic!("variant {index} panicked:\n{variant}"));
        refused += variant_refused;
        edited += variant_edited;
    }
    println!(
        "0755-text-run-mutations variants={} refused={refused} edited={edited}",
        variants.len()
    );
    assert!(refused > 0 && edited > 0);
    Ok(())
}

/// A leading UTF-8 byte-order mark shifts every reader position by three
/// bytes (quick-xml skips it without counting it), so text edits of such a
/// slide are refused rather than applied. This pins that the edit is refused
/// or correct, never corrupted and never a panic; the positional fix is a
/// follow-up of change 0755.
#[test]
fn a_byte_order_marked_slide_is_refused_or_edited_correctly_never_corrupted() -> Result<()> {
    let mut package = text_box_package(&["My text"])?;
    let mut xml = vec![0xef, 0xbb, 0xbf];
    xml.extend_from_slice(slide_xml(&package)?.as_bytes());
    set_slide_xml(&mut package, xml)?;
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    let outcome = catch_unwind(AssertUnwindSafe(|| edit.set_shape_text(0, 0, "edited")));
    match outcome.expect("a byte-order mark must not panic") {
        Err(_) => {},
        Ok(_) => {
            let commit = edit.commit()?;
            let blob = commit
                .snapshot()
                .package
                .get_part(&PackURI::new(SLIDE).map_err(Error::Invalid)?)?
                .blob()
                .to_vec();
            let text = crate::shape::Scene::read(&blob)?
                .at(0)
                .map_err(|error| Error::Invalid(error.to_string()))?
                .common()
                .text()
                .map(str::to_owned);
            assert_eq!(text.as_deref(), Some("edited"));
        },
    }
    Ok(())
}
