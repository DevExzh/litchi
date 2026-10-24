#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "shape-text closure tests use panic-driven assertions"
)]

//! Length-changing shape text on `PowerPoint`-authored decks (record 0260).
//!
//! `PowerPoint` writes a `TextSpecialInfoAtom` and a `TextRulerAtom` beside
//! almost every text atom, and keeps title/body placeholder text in the
//! slide list (`OutlineTextRefAtom`, MS-PPT 2.9.78) instead of the drawing.
//! These tests pin that both now accept a text of another length, that the
//! formatting ranges which cannot be resized losslessly still refuse, and
//! that every other shape, slide and stream is preserved.

use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
use litchi_ppt::Package;
use litchi_ppt::package::edit::{Error, Position, ShapeTextReplacement, Target};
use litchi_ppt::slide_order;
use litchi_ppt::text_edit::Refusal;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

fn texts(package: &Package) -> Vec<Vec<String>> {
    let presentation = package.presentation().unwrap();
    presentation
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .unwrap()
                .iter()
                .map(|shape| shape.text().unwrap())
                .collect()
        })
        .collect()
}

fn target(slide: usize, shape: usize) -> Target {
    Target::new(Position::new(slide), Position::new(shape))
}

/// Applies one replacement, then checks readback, preservation of every
/// other shape text, exact patch replay and exact inverse.
fn assert_replaced(path: &str, slide: usize, shape: usize, text: &str) {
    let source_bytes = fixture(path);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let before = texts(&source);
    assert_ne!(
        before[slide][shape], text,
        "{path}: the fixture already holds the text"
    );
    assert_ne!(
        before[slide][shape].encode_utf16().count(),
        text.encode_utf16().count(),
        "{path}: the replacement must change the length"
    );

    let mut edit = source.edit().unwrap();
    edit.set_shape_text(target(slide, shape), text).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());

    let mut expected = before.clone();
    expected[slide][shape] = text.to_owned();
    assert_eq!(texts(commit.package()), expected, "{path}");
    // The opened source snapshot is immutable.
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    // A fresh open of the published bytes reads the same presentation.
    let published = commit.package().to_bytes().unwrap();
    assert_eq!(texts(&Package::from_bytes(&published).unwrap()), expected);
    // Exact replay and exact inverse.
    let replayed = source.apply(commit.patch()).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), published);
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);
}

fn refusal(error: &Error) -> Option<&Refusal> {
    let mut source: Option<&(dyn std::error::Error + 'static)> = Some(error);
    while let Some(current) = source {
        if let Some(slide_order::Error::Text(text)) = current.downcast_ref::<slide_order::Error>()
            && let litchi_ppt::text_edit::Error::Refused(refusal) = text.as_ref()
        {
            return Some(refusal);
        }
        if let Some(litchi_ppt::text_edit::Error::Refused(refusal)) =
            current.downcast_ref::<litchi_ppt::text_edit::Error>()
        {
            return Some(refusal);
        }
        source = current.source();
    }
    None
}

#[test]
fn drawing_text_with_special_info_and_ruler_changes_length() {
    // `TextHeaderAtom`, `TextBytesAtom`, `StyleTextPropAtom`,
    // `TextSpecialInfoAtom` and `TextRulerAtom`: the plain PowerPoint layout.
    assert_replaced(
        "ole/ppt/text_shapes.ppt",
        0,
        2,
        "A rectangle with a longer label",
    );
    assert_replaced("ole/ppt/text-margins.ppt", 0, 0, "T");
    assert_replaced(
        "ole/ppt/ppt_with_png.ppt",
        0,
        1,
        "Here is a PNG image, captioned at greater length:",
    );
}

#[test]
fn outline_placeholder_text_changes_length() {
    // Title placeholders whose ClientTextbox holds only an OutlineTextRefAtom.
    assert_replaced("ole/ppt/SampleShow.ppt", 0, 0, "A new title");
    assert_replaced(
        "ole/ppt/text_shapes.ppt",
        0,
        0,
        "A title placeholder rewritten in the slide list",
    );
    // A multi-paragraph body placeholder in a 36-slide deck.
    assert_replaced(
        "poi/test-data/slideshow/41246-1.ppt",
        7,
        1,
        "First paragraph\rSecond paragraph",
    );
    // Text may also become shorter, down to one character.
    assert_replaced("poi/test-data/slideshow/41246-1.ppt", 0, 0, "J");
}

#[test]
fn one_call_publishes_outline_and_drawing_targets_together() {
    let source = Package::from_bytes(&fixture("ole/ppt/text_shapes.ppt")).unwrap();
    let mut expected = texts(&source);
    expected[0][0] = "Outline title".to_owned();
    expected[0][2] = "Drawing rectangle, longer".to_owned();
    let mut edit = source.edit().unwrap();
    let changed = edit
        .set_shape_texts(&[
            ShapeTextReplacement::new(target(0, 0), "Outline title"),
            ShapeTextReplacement::new(target(0, 2), "Drawing rectangle, longer"),
        ])
        .unwrap();
    assert_eq!(changed, 2);
    let commit = edit.commit().unwrap();
    assert_eq!(texts(commit.package()), expected);
}

#[test]
fn outline_text_composes_with_staged_slide_structure() {
    let source = Package::from_bytes(&fixture("poi/test-data/slideshow/41246-1.ppt")).unwrap();
    let before = texts(&source);
    let mut edit = source.edit().unwrap();
    // Move the last slide first, then retitle the moved slide and the slide
    // that was first, both outline-backed.
    let last = before.len() - 1;
    edit.move_slide(Position::new(last), Position::new(0))
        .unwrap();
    edit.set_shape_text(target(0, 0), "Moved and retitled")
        .unwrap();
    edit.set_shape_text(target(1, 0), "Second now").unwrap();
    edit.remove_slide(Position::new(2)).unwrap();
    let commit = edit.commit().unwrap();

    let mut expected = before.clone();
    let moved = expected.remove(last);
    expected.insert(0, moved);
    expected[0][0] = "Moved and retitled".to_owned();
    expected[1][0] = "Second now".to_owned();
    expected.remove(2);
    assert_eq!(texts(commit.package()), expected);
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn position_dependent_ranges_keep_the_same_length_contract() {
    // datetime.ppt slide 1 shape 1 holds a date field (DateTimeMCAtom) at a
    // character position, which a changed length would invalidate.
    let path = "poi/test-data/slideshow/datetime.ppt";
    let source = Package::from_bytes(&fixture(path)).unwrap();
    let before = texts(&source);
    let (slide, shape) = before
        .iter()
        .enumerate()
        .flat_map(|(slide, shapes)| (0..shapes.len()).map(move |shape| (slide, shape)))
        .find(|&(slide, shape)| {
            let mut edit = source.edit().unwrap();
            let error = edit
                .set_shape_text(target(slide, shape), "a text of some other length entirely")
                .err();
            error.as_ref().and_then(refusal) == Some(&Refusal::DependencyClosure)
        })
        .expect("datetime.ppt has a field-bearing text body");
    // The refusal leaves the edit usable and the same-length edit still works.
    let original = &before[slide][shape];
    let same_length: String = original
        .chars()
        .map(|character| {
            if character.is_ascii_alphabetic() {
                'x'
            } else {
                character
            }
        })
        .collect();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.set_shape_text(target(slide, shape), "different length!")
            .err()
            .as_ref()
            .and_then(refusal),
        Some(&Refusal::DependencyClosure)
    );
    if same_length != *original {
        edit.set_shape_text(target(slide, shape), same_length.as_str())
            .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(texts(commit.package())[slide][shape], same_length);
    }
}

#[test]
fn signed_outline_text_admits_exact_noops_and_refuses_changes() {
    let unsigned = fixture("ole/ppt/SampleShow.ppt");
    let mut source =
        OleFile::open(Cursor::new(unsigned.as_slice()), OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = source.open_stream(&names).unwrap();
        writer.create_stream(&names, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"opaque synthetic signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let signed_bytes = output.into_inner();
    let signed = Package::from_bytes(&signed_bytes).unwrap();
    let title = texts(&signed)[0][0].clone();

    let mut edit = signed.edit().unwrap();
    edit.set_shape_text(target(0, 0), title.as_str()).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().to_bytes().unwrap(), signed_bytes);

    let mut edit = signed.edit().unwrap();
    assert!(edit.set_shape_text(target(0, 0), "Changed").is_err());
    assert!(
        edit.set_shape_texts(&[ShapeTextReplacement::new(target(0, 0), "Changed")])
            .is_err()
    );
}

/// Corpus coverage of a length-changing replacement (record 0260): every
/// top-level shape of every readable, unencrypted PPT fixture under
/// `test-data`, one edit per shape. Before record 0260 no target accepted a
/// changed length; the remaining refusals are shapes without a text body
/// and the position-dependent records the closure keeps.
#[test]
fn corpus_length_change_coverage_is_pinned() {
    const DECKS: &[&str] = &[
        "ole/ppt/SampleShow.ppt",
        "ole/ppt/empty.ppt",
        "ole/ppt/empty_textbox.ppt",
        "ole/ppt/ppt_with_png.ppt",
        "ole/ppt/text-margins.ppt",
        "ole/ppt/text_shapes.ppt",
        "poi/test-data/slideshow/41246-1.ppt",
        "poi/test-data/slideshow/45543.ppt",
        "poi/test-data/slideshow/54880_chinese.ppt",
        "poi/test-data/slideshow/WithComments.ppt",
        "poi/test-data/slideshow/WithLinks.ppt",
        "poi/test-data/slideshow/basic_test_ppt_file.ppt",
        "poi/test-data/slideshow/bug58144-headers-footers-2003.ppt",
        "poi/test-data/slideshow/bug58144-headers-footers-2007.ppt",
        "poi/test-data/slideshow/datetime.ppt",
        "poi/test-data/slideshow/headers_footers.ppt",
        "poi/test-data/slideshow/headers_footers_2007.ppt",
        "poi/test-data/slideshow/incorrect_slide_order.ppt",
        "poi/test-data/slideshow/ole2-embedding-2003.ppt",
        "poi/test-data/slideshow/ppt_with_embeded.ppt",
        "poi/test-data/slideshow/sound.ppt",
        "poi/test-data/slideshow/testPPT_oleWorkbook.ppt",
        "poi/test-data/slideshow/text-margins.ppt",
    ];
    let (mut changed, mut closure, mut no_text_atom, mut no_textbox) = (0, 0, 0, 0);
    for deck in DECKS {
        let package = Package::from_bytes(&fixture(deck)).unwrap();
        let shapes = texts(&package);
        for (slide, shapes) in shapes.iter().enumerate() {
            for shape in 0..shapes.len() {
                let mut edit = package.edit().unwrap();
                match edit.set_shape_text(target(slide, shape), "Changed text of another length!") {
                    Ok(()) => {
                        let commit = edit.commit().unwrap();
                        assert_eq!(
                            texts(commit.package())[slide][shape],
                            "Changed text of another length!",
                            "{deck} {slide}/{shape}"
                        );
                        changed += 1;
                    },
                    Err(error) => match refusal(&error) {
                        Some(Refusal::DependencyClosure) => closure += 1,
                        Some(Refusal::NoTextAtom) => no_text_atom += 1,
                        Some(Refusal::NoTextbox) => no_textbox += 1,
                        other => panic!("{deck} {slide}/{shape}: unexpected {other:?}"),
                    },
                }
            }
        }
    }
    assert_eq!(
        (changed, closure, no_text_atom, no_textbox),
        (129, 47, 13, 25)
    );
}
