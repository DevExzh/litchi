//! Slide-list text resolved into the shapes that reference it (L-PPT-R1,
//! finding F-10 of `docs/unified-ops/0010-m1-unified-read-plan.md`).
//!
//! `PowerPoint` keeps the text of title and body placeholders in the
//! `SlideListWithTextContainer`; the placeholder's text box holds only an
//! `OutlineTextRefAtom` (MS-PPT 2.9.78). Before L-PPT-R1 those placeholders
//! read as empty. The expected strings below are the slide-list bodies the
//! references select; where Apache POI's own test suite asserts the same
//! text (`poi-scratchpad/src/test/java/org/apache/poi/hslf/...`), the test
//! names the POI assertion it agrees with.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_ppt::shapes::PlaceholderType;
use litchi_ppt::{Package, Presentation};
use std::path::PathBuf;

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path)
}

fn with_presentation(path: &str, password: Option<&str>, check: impl FnOnce(&Presentation)) {
    let package = match password {
        None => Package::open(fixture(path)).unwrap(),
        Some(password) => Package::open_with_options(
            fixture(path),
            litchi_ppt::OpenOptions::default().password(password),
        )
        .unwrap(),
    };
    let presentation = package.presentation().unwrap();
    check(&presentation);
}

/// Every placeholder of every slide, as `(type, text)` in drawing order.
fn placeholders(presentation: &Presentation) -> Vec<Vec<(PlaceholderType, String)>> {
    presentation
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .unwrap()
                .iter()
                .filter_map(|shape| {
                    shape
                        .as_placeholder()
                        .map(|placeholder| (placeholder.placeholder_type(), shape.text().unwrap()))
                })
                .collect()
        })
        .collect()
}

fn owned(slides: &[&[(PlaceholderType, &str)]]) -> Vec<Vec<(PlaceholderType, String)>> {
    slides
        .iter()
        .map(|slide| {
            slide
                .iter()
                .map(|(kind, text)| (*kind, (*text).to_owned()))
                .collect()
        })
        .collect()
}

/// Reading the outline text adds no repair to any affected fixture.
fn assert_clean(presentation: &Presentation) {
    let report = presentation.collect_tolerance_report().unwrap();
    assert!(
        report.defects().iter().all(|defect| defect.defect
            != litchi_ppt::Defect::Shape(litchi_ppt::ShapeDefect::UnresolvedOutlineTextRef)),
        "{:?}",
        report.defects()
    );
}

/// `extract_text_fast` and `Slide::text` agree slide by slide.
fn assert_fast_text_matches(presentation: &Presentation) {
    let fast = presentation.extract_text_fast().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(fast.len(), slides.len());
    for ((number, text), slide) in fast.iter().zip(&slides) {
        assert_eq!(*number, slide.slide_number());
        assert_eq!(text, slide.text().unwrap());
    }
}

use PlaceholderType::{Body, CenterTitle, SubTitle, Title};

/// POI `TestSheetText.testSheetOne`/`testSheetTwo` and
/// `TestTextRun.testGetText` assert these raw texts for the slide-list text
/// of the two slides; `TestSlideOrdering.testTitles` asserts the titles.
#[test]
fn basic_test_placeholders_read_their_slide_list_text() {
    with_presentation(
        "poi/test-data/slideshow/basic_test_ppt_file.ppt",
        None,
        |presentation| {
            assert_eq!(
                placeholders(presentation),
                owned(&[
                    &[
                        (CenterTitle, "This is a test title"),
                        (SubTitle, "This is a test subtitle\rThis is on page 1"),
                    ],
                    &[
                        (Title, "This is the title on page 2"),
                        (
                            Body,
                            "This is page two\rIt has several blocks of text\rNone of them have formatting",
                        ),
                    ],
                ])
            );
            let slides = presentation.slides().unwrap();
            let bodies = slides[0].outline_text_bodies();
            assert_eq!(bodies.len(), 2);
            assert_eq!(bodies[0].text_type(), litchi_ppt::TextType::CenterTitle);
            assert_eq!(bodies[0].text_header_instance(), 0);
            assert_eq!(bodies[0].text().unwrap(), "This is a test title");
            assert_eq!(bodies[1].text_type(), litchi_ppt::TextType::CenterBody);
            assert_eq!(bodies[1].text_header_instance(), 1);
            assert_eq!(
                presentation.slide_directory().entries()[0].outline_text_bodies(),
                bodies
            );
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        },
    );
}

/// POI `TestSlideOrdering.testTitles` asserts the titles "Slide 1",
/// "Slide 2", "Slide 3" in presentation order.
#[test]
fn reordered_slides_resolve_their_own_slide_list_text() {
    with_presentation(
        "poi/test-data/slideshow/incorrect_slide_order.ppt",
        None,
        |presentation| {
            assert_eq!(
                placeholders(presentation),
                owned(&[
                    &[(CenterTitle, "Slide 1"), (SubTitle, "First slide I added")],
                    &[(Title, "Slide 2"), (Body, "Third slide I added")],
                    &[(Title, "Slide 3"), (Body, "Second slide I added")],
                ])
            );
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        },
    );
}

/// POI `TestHyperlink.testTextRunHyperlinks` asserts these bodies (with
/// `\r` shown as `\n`): the hyperlinks of slide-list text stay inside the
/// resolved text.
#[test]
fn hyperlinked_slide_list_text_resolves() {
    with_presentation(
        "poi/test-data/slideshow/WithLinks.ppt",
        None,
        |presentation| {
            assert_eq!(
                placeholders(presentation),
                owned(&[
                    &[
                        (Title, "I am the title of the first page"),
                        (
                            SubTitle,
                            "This page has two links:\rhttp://jakarta.apache.org/poi/\r\rhttp://slashdot.org/\r\rIn addition, its notes has one link",
                        ),
                    ],
                    &[
                        (Title, "I am page two"),
                        (SubTitle, "I have the one link:\rJakarta HSSF"),
                    ],
                ])
            );
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        },
    );
}

#[test]
fn single_reference_fixtures_resolve() {
    let cases: [(&str, &[&[(PlaceholderType, &str)]]); 4] = [
        (
            "ole/ppt/empty_textbox.ppt",
            &[
                &[
                    (CenterTitle, "Simple PPT attachment"),
                    (SubTitle, "Bla bla bla"),
                ],
                // The body's slide-list text has no characters.
                &[(Title, "First Slide"), (Body, "")],
            ],
        ),
        (
            "poi/test-data/slideshow/WithComments.ppt",
            &[&[
                (CenterTitle, "Test Slide"),
                (SubTitle, "With a comment on it"),
            ]],
        ),
        (
            "poi/test-data/slideshow/bug58144-headers-footers-2003.ppt",
            &[&[(CenterTitle, "Test")]],
        ),
        (
            "poi/test-data/slideshow/ole2-embedding-2003.ppt",
            &[&[(CenterTitle, "OLE Embedding")]],
        ),
    ];
    for (path, expected) in cases {
        with_presentation(path, None, |presentation| {
            assert_eq!(placeholders(presentation), owned(expected), "{path}");
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        });
    }
}

/// POI `TestTextShape.read` asserts "Title Placeholder" for the title
/// placeholder of this deck. The slide's text now includes the title beside
/// the shapes that hold their text inline; before L-PPT-R1 the title was
/// missing from both the shape and the slide text.
#[test]
fn slide_text_includes_resolved_placeholders_beside_inline_text() {
    with_presentation("ole/ppt/text_shapes.ppt", None, |presentation| {
        assert_eq!(
            placeholders(presentation),
            owned(&[&[(Title, "Title Placeholder")]])
        );
        let slides = presentation.slides().unwrap();
        assert_eq!(
            slides[0].text().unwrap(),
            "Title Placeholder\nText in a TextBox\nRectangle\nEllipse\nOctagon"
        );
        assert_fast_text_matches(presentation);
    });

    with_presentation(
        "poi/test-data/slideshow/41246-1.ppt",
        None,
        |presentation| {
            let slides = presentation.slides().unwrap();
            // Slide 8 has a code text box; its title and body were dropped
            // from the slide text before.
            assert_eq!(
                slides[7].text().unwrap(),
                "How can I performance profile my Java code?\n[*]java -prof MyClass\rExample A simple class\npublic class Tomer {\r\tpublic static void main(String[] args) {\r\t\tSystem.out.println(\"hi there\");\r\t}\r}"
            );
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        },
    );
}

/// Every placeholder text of a deck whose title and body text all live in
/// the slide list equals the slide's slide-list text, read through the
/// independent record decoder behind `SlideDirectoryEntry::list_text`.
#[test]
fn every_resolved_placeholder_matches_the_independent_slide_list_decoder() {
    with_presentation(
        "poi/test-data/slideshow/41246-1.ppt",
        None,
        |presentation| {
            let entries = presentation.slide_directory().entries();
            let slides = placeholders(presentation);
            assert_eq!(entries.len(), slides.len());
            let mut resolved = 0;
            for (entry, slide) in entries.iter().zip(&slides) {
                let mut expected = entry
                    .list_text()
                    .split('\n')
                    .map(str::trim)
                    .filter(|text| !text.is_empty())
                    .map(str::to_owned)
                    .collect::<Vec<_>>();
                let mut actual = slide
                    .iter()
                    .map(|(_, text)| text.clone())
                    .filter(|text| !text.is_empty())
                    .collect::<Vec<_>>();
                resolved += actual.len();
                expected.sort();
                actual.sort();
                assert_eq!(actual, expected, "slide {}", entry.slide_id());
            }
            // 36 slides; 69 placeholders hold slide-list characters.
            assert_eq!(slides.len(), 36);
            assert_eq!(resolved, 69);
            assert_eq!(
                slides[0],
                [
                    (Title, "JBOSS Profiler".to_owned()),
                    (SubTitle, "Performance Profiling".to_owned()),
                ]
            );
        },
    );
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_decks_resolve_slide_list_text() {
    let protected = [
        (
            "Password_Protected-hello.ppt",
            "It only has the one slide\rIt does have notes\rThe password is \u{201c}hello\u{201d}",
        ),
        (
            "Password_Protected-56-hello.ppt",
            "It only has the one slide\rIt does have notes\rThe password is \u{201c}hello\u{201d}",
        ),
        (
            "Password_Protected-np-hello.ppt",
            "It only has the one slide\rIt does have notes\rProperties are not protected\rThe password is \u{201c}hello\u{201d}",
        ),
    ];
    for (name, subtitle) in protected {
        let path = format!("poi/test-data/slideshow/{name}");
        with_presentation(&path, Some("hello"), |presentation| {
            assert_eq!(
                placeholders(presentation),
                owned(&[&[
                    (CenterTitle, "This is a protected document"),
                    (SubTitle, subtitle),
                ]]),
                "{name}"
            );
            assert_clean(presentation);
            assert_fast_text_matches(presentation);
        });
    }
}
