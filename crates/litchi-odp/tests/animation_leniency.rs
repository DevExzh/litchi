//! LENIENCY_AUDIT F9.10: an animation element the reader cannot map — an
//! unknown `anim:` element or a foreign-namespace element inside an
//! animation node — is skipped with its subtree, the way LibreOffice's
//! `AnimationNodeContext` skips tokens it does not know, instead of failing
//! every slide. Every ODP open path records each skip with its slide; a
//! conforming deck records nothing; a known element in a disallowed position
//! and malformed XML stay hard errors.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "../../litchi-odf-common/tests/support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{ContentDefect, Defect, DefectLocation, ToleranceReport};
use litchi_odp::model::animation::Kind;
use litchi_odp::{Presentation, SourceBackedPresentation};

type Recorded = (Defect, u16, Option<DefectLocation>);

fn recorded(report: &ToleranceReport) -> Vec<Recorded> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context, tolerated.location))
        .collect()
}

fn skipped(context: u16, slide: u32) -> Recorded {
    (
        Defect::Content(ContentDefect::UnknownAnimationElement),
        context,
        Some(DefectLocation::Slide { index: slide }),
    )
}

fn conforming() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/impress-master-layouts.odp"),
    )
    .unwrap()
}

/// Insert `markup` just before the closing tag of page `page` (zero-based).
fn insert_before_page_end(xml: &str, page: usize, markup: &str) -> String {
    let (position, _) = xml
        .match_indices("</draw:page>")
        .nth(page)
        .expect("fixture has the page");
    format!("{}{markup}{}", &xml[..position], &xml[position..])
}

const NESTED: &str = concat!(
    r#"<anim:par><anim:seq>"#,
    r#"<anim:animatePhysics smil:dur="1s"><anim:set smil:attributeName="visibility" smil:to="visible"/></anim:animatePhysics>"#,
    r#"<anim:par><anim:set smil:attributeName="visibility" smil:to="visible"/><vendor:effect xmlns:vendor="urn:example:vendor"><vendor:inner/></vendor:effect></anim:par>"#,
    r#"</anim:seq></anim:par>"#,
);
const PAGE_ROOT: &str = r#"<anim:future-root/><anim:future-tree><anim:par/></anim:future-tree>"#;

fn with_unknown_animations() -> Vec<u8> {
    odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        let xml = insert_before_page_end(xml, 2, NESTED);
        insert_before_page_end(&xml, 4, PAGE_ROOT)
    })
}

fn expected() -> Vec<Recorded> {
    vec![skipped(0, 2), skipped(1, 2), skipped(0, 4), skipped(0, 4)]
}

fn assert_every_open_path(bytes: &[u8], defects: &[Recorded], label: &str) {
    let owned = Presentation::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(owned.slides().unwrap().len(), 5, "{label}");
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );

    // The report is the same whether it or the slides are requested first.
    let report_first = Presentation::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(report_first.tolerance_report()),
        defects,
        "{label}: report first"
    );

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Presentation::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(source.slides().unwrap().len(), 5, "{label}: source-backed");
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
}

#[test]
fn unknown_animation_elements_are_skipped_and_recorded_per_slide() {
    let bytes = with_unknown_animations();
    assert_every_open_path(&bytes, &expected(), "unknown animations");

    let slides = Presentation::from_bytes(bytes).unwrap().slides().unwrap();
    // Slide 2 keeps its timing tree minus the two unmappable elements.
    let root = &slides[2].animations()[0];
    assert_eq!(root.kind(), Kind::Parallel);
    let sequence = &root.children()[0];
    assert_eq!(sequence.kind(), Kind::Sequence);
    assert_eq!(sequence.children().len(), 1);
    let inner = &sequence.children()[0];
    assert_eq!(inner.kind(), Kind::Parallel);
    assert_eq!(
        inner
            .children()
            .iter()
            .map(|node| node.kind())
            .collect::<Vec<_>>(),
        [Kind::Set]
    );
    // Slide 4's page-root unknowns are dropped entirely; the other slides
    // are untouched.
    assert!(slides[4].animations().is_empty());
    assert!(slides[0].animations().is_empty());
}

#[test]
fn the_same_deck_without_the_deviation_repairs_nothing() {
    assert_every_open_path(&conforming(), &[], "original");
    let known_only = odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        insert_before_page_end(
            xml,
            2,
            r#"<anim:par><anim:seq><anim:set smil:attributeName="visibility" smil:to="visible"/></anim:seq></anim:par>"#,
        )
    });
    assert_every_open_path(&known_only, &[], "known animations");
}

#[test]
fn disallowed_known_children_and_broken_xml_stay_hard_errors() {
    let misplaced = odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        insert_before_page_end(
            xml,
            1,
            r#"<anim:set smil:attributeName="visibility"><anim:par/></anim:set>"#,
        )
    });
    assert!(
        Presentation::from_bytes(misplaced)
            .unwrap()
            .slides()
            .is_err()
    );

    let unterminated = odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        insert_before_page_end(xml, 1, r#"<anim:future-tree><anim:par/>"#)
    });
    assert!(
        Presentation::from_bytes(unterminated)
            .and_then(|presentation| presentation.slides())
            .is_err()
    );
}
