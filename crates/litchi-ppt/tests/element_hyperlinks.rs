#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "hyperlink edit tests use panic-driven assertions"
)]

//! Shape click hyperlinks through the canonical primary edit (record 0323).

use litchi_ppt::Package;
use litchi_ppt::hyperlink::{Click, InteractionJump, Target};
use litchi_ppt::package::edit::{Position, Target as ShapeTarget};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

const DECK: &str = "ole/ppt/SampleShow.ppt";

/// The click hyperlink of the top-level shape at `shape` on `slide`.
fn click(package: &Package, slide: usize, shape: usize) -> Option<Click> {
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let ids = slides
        .iter()
        .map(|slide| slide.slide_id())
        .collect::<Vec<_>>();
    let selected = &slides[slide];
    let shapes = selected.shapes().unwrap();
    let id = shapes[shape].id();
    let hyperlinks = presentation.hyperlinks().unwrap();
    selected
        .shape_interactions()
        .unwrap()
        .into_iter()
        .find(|entry| entry.shape_id == id)
        .and_then(|entry| {
            entry
                .interactions
                .iter()
                .find_map(|interaction| Click::resolve(interaction, &hyperlinks, &ids))
        })
}

fn texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap().to_owned())
        .collect()
}

#[test]
fn url_slide_and_jump_links_round_trip() {
    let source_bytes = fixture(DECK);
    let source = Package::from_bytes(&source_bytes).unwrap();
    assert_eq!(click(&source, 0, 0), None);
    let url = Click {
        target: Target::Url("https://example.com/a?b=c".into()),
        tooltip: Some("Example".into()),
    };
    let slide = Click {
        target: Target::Slide(Position::new(1)),
        tooltip: None,
    };
    let jump = Click {
        target: Target::Jump(InteractionJump::LastSlide),
        tooltip: None,
    };
    let mut edit = source.edit().unwrap();
    let target = |shape| ShapeTarget::new(Position::new(0), Position::new(shape));
    edit.set_shape_hyperlink(target(0), Some(&url)).unwrap();
    edit.set_shape_hyperlink(target(1), Some(&slide)).unwrap();
    edit.set_shape_hyperlink(
        ShapeTarget::new(Position::new(1), Position::new(0)),
        Some(&jump),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(click(&published, 0, 0), Some(url.clone()));
    assert_eq!(click(&published, 0, 1), Some(slide));
    assert_eq!(click(&published, 1, 0), Some(jump));
    assert_eq!(texts(&published), texts(&source));
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );

    // Equal links are exact no-ops; removal drops the click action.
    let mut edit = published.edit().unwrap();
    edit.set_shape_hyperlink(target(0), Some(&url)).unwrap();
    assert!(edit.commit().unwrap().patch().is_noop());
    let mut edit = published.edit().unwrap();
    edit.set_shape_hyperlink(target(0), None).unwrap();
    let removed = edit.commit().unwrap();
    assert_eq!(click(removed.package(), 0, 0), None);
    assert!(click(removed.package(), 0, 1).is_some());

    // Replacing a link keeps the old definition and adds a new one.
    let mut edit = published.edit().unwrap();
    let other = Click {
        target: Target::Url("https://example.org/".into()),
        tooltip: None,
    };
    edit.set_shape_hyperlink(target(0), Some(&other)).unwrap();
    let replaced = edit.commit().unwrap();
    assert_eq!(click(replaced.package(), 0, 0), Some(other));
    assert_eq!(
        replaced
            .package()
            .presentation()
            .unwrap()
            .hyperlinks()
            .unwrap()
            .hyperlinks
            .len(),
        published
            .presentation()
            .unwrap()
            .hyperlinks()
            .unwrap()
            .hyperlinks
            .len()
            + 1
    );
}

#[test]
fn invalid_links_are_refused() {
    let source = Package::from_bytes(&fixture(DECK)).unwrap();
    let target = ShapeTarget::new(Position::new(0), Position::new(0));
    for click in [
        Click {
            target: Target::Url(String::new()),
            tooltip: None,
        },
        Click {
            target: Target::Slide(Position::new(9)),
            tooltip: None,
        },
        Click {
            target: Target::Jump(InteractionJump::NextSlide),
            tooltip: Some("tip".into()),
        },
    ] {
        let mut edit = source.edit().unwrap();
        assert!(edit.set_shape_hyperlink(target, Some(&click)).is_err());
    }
}

#[test]
fn elements_exobjlist_hyperlink_add_reopens_with_chart_and_content_intact() {
    let source_bytes = fixture("unified-ops/0294/elements.ppt");
    let source = Package::from_bytes(&source_bytes).unwrap();
    let source_presentation = source.presentation().unwrap();
    let source_texts = texts(&source);
    let source_hyperlinks = source_presentation.hyperlinks().unwrap();
    let source_charts = source_presentation.charts().unwrap();
    assert_eq!(source_charts.seen(), 1);
    assert_eq!(source_charts.len(), 1);
    assert_eq!(source_charts.failures().count(), 0);

    let added = Click {
        target: Target::Url("https://example.test/added".into()),
        tooltip: None,
    };
    let mut edit = source.edit().unwrap();
    edit.set_shape_hyperlink(
        ShapeTarget::new(Position::new(0), Position::new(0)),
        Some(&added),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    let published = Package::from_bytes(&output).unwrap();

    assert_eq!(click(&published, 0, 0), Some(added));
    assert_eq!(texts(&published), source_texts);

    let published_presentation = published.presentation().unwrap();
    let published_hyperlinks = published_presentation.hyperlinks().unwrap();
    for hyperlink in source_hyperlinks.hyperlinks {
        assert!(published_hyperlinks.hyperlinks.contains(&hyperlink));
    }
    assert!(
        published_hyperlinks
            .hyperlinks
            .iter()
            .any(|hyperlink| hyperlink.target.as_deref() == Some("https://example.test/added"))
    );

    let published_charts = published_presentation.charts().unwrap();
    assert_eq!(published_charts.seen(), source_charts.seen());
    assert_eq!(published_charts.len(), source_charts.len());
    assert_eq!(
        published_charts.failures().count(),
        source_charts.failures().count()
    );
    let source_chart = source_charts.get(0).unwrap();
    let published_chart = published_charts.get(0).unwrap();
    assert_eq!(published_chart.kind(), source_chart.kind());
    assert_eq!(
        published_chart.info().program(),
        source_chart.info().program()
    );
    assert_eq!(published_chart.info().frame(), source_chart.info().frame());
    assert_eq!(
        published_chart.info().object_id(),
        source_chart.info().object_id()
    );
    assert_eq!(
        published_chart.info().persist_id(),
        source_chart.info().persist_id()
    );

    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );
}
