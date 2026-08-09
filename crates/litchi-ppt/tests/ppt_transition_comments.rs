#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Integration tests for legacy PPT slide transitions, slide timings,
//! comments, and custom slide shows using Apache POI test fixtures.

use litchi_ppt::{AdvanceMode, Package, Presentation, TransitionSpeed, TransitionType};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

fn with_presentation<T>(name: &str, f: impl FnOnce(&Presentation) -> T) -> T {
    let mut package = Package::open(fixture(name)).expect("open POI fixture");
    let presentation = package.presentation().expect("parse presentation");
    f(&presentation)
}

#[test]
fn with_comments_ppt_exposes_slide_comment() {
    with_presentation("WithComments.ppt", |presentation| {
        let slides = presentation.slides().expect("slides");
        assert_eq!(slides.len(), 1);

        let comments = slides[0].comments().expect("parse comments");
        assert_eq!(comments.len(), 1);

        let comment = &comments[0];
        assert_eq!(comment.index, 1);
        assert_eq!(comment.author, "Administrator");
        assert_eq!(comment.initials, "A");
        assert!(comment.text.contains("This is a test comment"));
        assert_eq!(comment.year, 2008);
        assert_eq!(comment.month, 8);
        assert_eq!(comment.day, 4);
    });
}

#[test]
fn presentation_comments_aggregates_commented_slides_only() {
    with_presentation("45543.ppt", |presentation| {
        let all = presentation.comments().expect("parse comments");
        // Only slides 1 and 2 carry comments in this 11-slide deck.
        assert_eq!(all.len(), 2);

        assert_eq!(all[0].slide_number, 1);
        assert_eq!(all[0].comments.len(), 1);
        assert_eq!(all[0].comments[0].author, "XPVMWARE01");
        assert_eq!(all[0].comments[0].text, "testdoc");

        assert_eq!(all[1].slide_number, 2);
        assert_eq!(all[1].comments.len(), 1);
        assert_eq!(all[1].comments[0].text, "test phrase");
    });
}

#[test]
fn bug45543_ppt_exposes_random_transition() {
    with_presentation("45543.ppt", |presentation| {
        let slides = presentation.slides().expect("slides");
        let transition = slides[0]
            .transition()
            .expect("parse transition")
            .expect("slide 1 has a transition");

        assert_eq!(transition.transition_type, TransitionType::Random);
        assert_eq!(transition.speed, TransitionSpeed::Slow);
        assert_eq!(transition.advance_mode, AdvanceMode::OnClick);
        assert_eq!(transition.advance_time_ms, None);
        assert!(transition.sound.is_none());
        assert!(transition.has_effect());
    });
}

#[test]
fn with_links_ppt_exposes_slide_timing_and_advance_mode() {
    with_presentation("WithLinks.ppt", |presentation| {
        let slides = presentation.slides().expect("slides");
        assert_eq!(slides.len(), 2);

        let transition = slides[0]
            .transition()
            .expect("parse transition")
            .expect("slide 1 has an SSSlideInfoAtom");
        // Raw atom: slideTime=1024ms, effect=0, flags=0x11 (manual advance + sound).
        assert_eq!(transition.transition_type, TransitionType::None);
        assert_eq!(transition.speed, TransitionSpeed::Medium);
        assert_eq!(transition.advance_mode, AdvanceMode::Both);
        assert_eq!(transition.advance_time_ms, Some(1024));

        let timing = slides[0].timing().expect("slide 1 has timing");
        assert_eq!(timing.advance_time_ms, 1024);
        assert!(timing.advance_on_click);
        assert!(!timing.hidden);
    });
}

#[test]
fn slides_without_slide_info_atom_return_no_transition_or_timing() {
    with_presentation("41246-1.ppt", |presentation| {
        let slides = presentation.slides().expect("slides");
        // Slide 2 of this deck has no SSSlideInfoAtom.
        let slide = &slides[1];
        assert!(slide.transition().expect("parse transition").is_none());
        assert!(slide.timing().is_none());
    });
}

#[test]
fn presentations_without_named_shows_return_empty_custom_shows() {
    with_presentation("WithComments.ppt", |presentation| {
        assert!(presentation.custom_shows().is_empty());
    });
}

#[test]
fn uncommented_presentation_returns_no_comments() {
    with_presentation("WithLinks.ppt", |presentation| {
        assert!(presentation.comments().expect("parse comments").is_empty());
    });
}
