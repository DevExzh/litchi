//! Save-fidelity regressions for `MutablePresentation`.
//!
//! Every fixture here previously lost slide text when an unmodified
//! presentation was opened and saved: the writer emitted only the constructs
//! the slide model understands and dropped everything else.

use litchi_odp::{MutablePresentation, Presentation, Shape};

/// Bytes of a fixture stored under `test-data/odf/odp`.
macro_rules! fixture {
    ($name:literal) => {
        include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/odf/odp/",
            $name
        ))
        .to_vec()
    };
}

/// Concatenate every slide's visible text, in slide order.
fn deck_text(bytes: Vec<u8>) -> String {
    let presentation = Presentation::from_bytes(bytes).expect("open presentation");
    presentation
        .slides()
        .expect("read slides")
        .iter()
        .map(|slide| slide.all_text())
        .collect::<Vec<_>>()
        .join("\n")
}

/// Open, save, and reopen without touching the model.
fn round_trip(bytes: Vec<u8>) -> Vec<u8> {
    let presentation = Presentation::from_bytes(bytes).expect("open presentation");
    let mutable = MutablePresentation::from_presentation(presentation).expect("make mutable");
    mutable.to_bytes().expect("save presentation")
}

/// The saved `content.xml` of a presentation package.
fn saved_content_xml(bytes: Vec<u8>) -> String {
    let package = litchi_odp::Package::from_bytes(bytes).expect("open package");
    String::from_utf8(package.get_file("content.xml").expect("content.xml")).expect("utf-8")
}

#[test]
fn table_shape_text_survives_a_save() {
    let source = fixture!("cellspan.odp");
    let before = deck_text(source.clone());
    assert!(
        before.contains("0,0"),
        "fixture lost its table text: {before}"
    );

    let saved = round_trip(source);
    assert_eq!(deck_text(saved.clone()), before);

    let xml = saved_content_xml(saved);
    assert!(xml.contains("<table:table"), "table markup was dropped");
    assert!(
        xml.contains("table:number-columns-spanned=\"2\""),
        "cell spans were dropped"
    );
}

#[test]
fn text_in_image_survives_a_save() {
    let source = fixture!("text-in-image.odp");
    let before = deck_text(source.clone());
    assert!(!before.trim().is_empty(), "fixture has no text to preserve");
    assert_eq!(deck_text(round_trip(source)), before);
}

#[test]
fn outline_text_survives_a_save() {
    let source = fixture!("tdf102223.odp");
    let before = deck_text(source.clone());
    assert!(!before.trim().is_empty(), "fixture has no text to preserve");
    assert_eq!(deck_text(round_trip(source)), before);
}

#[test]
fn placeholder_text_survives_a_save() {
    let source = fixture!("tdf105502.odp");
    let before = deck_text(source.clone());
    assert!(!before.trim().is_empty(), "fixture has no text to preserve");
    assert_eq!(deck_text(round_trip(source)), before);
}

#[test]
fn empty_style_name_reference_opens_and_round_trips() {
    // `style:data-style-name` is an ODF `styleNameRef`, which explicitly allows
    // the empty string to mean "no referenced style".
    let source = fixture!("tdf169979.odp");
    let before = deck_text(source.clone());
    assert_eq!(deck_text(round_trip(source)), before);
}

#[test]
fn font_declarations_and_automatic_styles_survive_a_save() {
    let source = fixture!("tdf102223.odp");
    let original = saved_content_xml(source.clone());
    let saved = saved_content_xml(round_trip(source));
    assert!(
        original.contains("<office:font-face-decls>"),
        "fixture has no font declarations"
    );
    assert!(
        saved.contains("<office:font-face-decls>"),
        "font declarations were dropped"
    );
    assert!(
        saved.contains("style:family=\"drawing-page\""),
        "automatic drawing-page styles were dropped"
    );
}

#[test]
fn editing_one_slide_leaves_other_slides_verbatim() {
    let source = fixture!("tdf102223.odp");
    let presentation = Presentation::from_bytes(source).expect("open presentation");
    let mut mutable = MutablePresentation::from_presentation(presentation).expect("make mutable");
    assert!(
        !mutable.slides().is_empty(),
        "fixture needs at least one slide"
    );
    let untouched_text = mutable
        .slides()
        .iter()
        .skip(1)
        .map(|slide| slide.all_text())
        .collect::<Vec<_>>()
        .join("\n");

    let mut shape = Shape::new();
    shape.text = "Appended by the test".to_string();
    mutable.add_shape(0, shape).expect("add shape");
    let saved = mutable.to_bytes().expect("save presentation");

    let reopened = Presentation::from_bytes(saved).expect("reopen presentation");
    let slides = reopened.slides().expect("read slides");
    assert!(
        slides[0].all_text().contains("Appended by the test"),
        "the edit was not written"
    );
    let after = slides
        .iter()
        .skip(1)
        .map(|slide| slide.all_text())
        .collect::<Vec<_>>()
        .join("\n");
    assert_eq!(after, untouched_text, "untouched slides changed");
}
