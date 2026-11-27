//! A chart, a picture or an equation object inserted and removed in one
//! edit commits, leaves nothing in the saved package, and its patch replays
//! and inverts exactly: every part the edit added goes with the
//! relationships that reached it, with the parts those parts relate.

use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::{Edit, ElementPath, ElementStep, NewChart, NewPicture};
use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};

fn paragraph() -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(0))])
}

fn base() -> Package {
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("Text");
    builder.build().unwrap()
}

fn save(package: &Package) -> Vec<u8> {
    package.clone().to_bytes().unwrap()
}

/// Run `stage` on an edit of a fresh document and check the commit leaves
/// the document as it was, and that the patch replays and inverts exactly.
fn leaves_no_trace(stage: impl FnOnce(&mut Edit)) {
    let base = base();
    let before = save(&base);
    let mut edit = base.edit().unwrap();
    stage(&mut edit);
    let (after, patch) = edit.commit().expect("the commit succeeds").into_parts();
    let saved = save(&after);
    assert_eq!(saved, before, "the saved package changed");
    assert_eq!(save(&base.apply(&patch).unwrap().into_package()), saved);
    assert_eq!(
        save(&after.apply(&patch.inverse()).unwrap().into_package()),
        before
    );
}

fn chart(kind: Kind) -> NewChart {
    NewChart::new(
        Spec::new(Variant::new(kind))
            .with_categories(["Q1", "Q2"])
            .with_series(Series::new("East", vec![1.0, 2.0])),
    )
}

#[test]
fn a_column_chart_inserted_and_removed_leaves_no_trace() {
    leaves_no_trace(|edit| {
        edit.insert_chart_at(&paragraph(), None, &chart(Kind::Column))
            .unwrap();
        edit.remove_chart_at(&paragraph(), Position::new(0))
            .unwrap();
    });
}

#[test]
fn a_waterfall_chart_inserted_and_removed_leaves_no_trace() {
    leaves_no_trace(|edit| {
        edit.insert_chart_at(&paragraph(), None, &chart(Kind::Waterfall))
            .unwrap();
        edit.remove_chart_at(&paragraph(), Position::new(0))
            .unwrap();
    });
}

#[test]
fn a_picture_inserted_and_removed_leaves_no_trace() {
    // A 1x1 PNG.
    let png = [
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, b'I', b'H', b'D', b'R', 0, 0,
        0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89,
    ];
    leaves_no_trace(|edit| {
        edit.insert_picture_at(&paragraph(), None, &NewPicture::new(png.to_vec()))
            .unwrap();
        edit.remove_picture_at(&paragraph(), Position::new(0))
            .unwrap();
    });
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_object_inserted_and_removed_leaves_no_trace() {
    use litchi_docx::document::{EquationContent, EquationSource};
    let content = EquationContent::new(
        EquationSource::Latex("x^{2}"),
        None,
        12.0,
        &litchi_equation::Limits::DEFAULT,
    )
    .unwrap();
    leaves_no_trace(|edit| {
        edit.insert_equation_object_at(None, &paragraph(), None, &content)
            .unwrap();
        edit.remove_equation_object_at(None, &paragraph(), Position::new(0))
            .unwrap();
    });
}

#[test]
fn a_kept_chart_keeps_the_parts_it_relates() {
    // Two charts inserted, one removed: the other keeps its workbook.
    let base = base();
    let mut edit = base.edit().unwrap();
    edit.insert_chart_at(&paragraph(), None, &chart(Kind::Waterfall))
        .unwrap();
    edit.insert_chart_at(&paragraph(), None, &chart(Kind::Column))
        .unwrap();
    edit.remove_chart_at(&paragraph(), Position::new(0))
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let reopened = Package::from_bytes(&save(&after)).unwrap();
    let charts = reopened.document().unwrap().paragraphs().unwrap()[0]
        .charts()
        .unwrap();
    assert_eq!(charts.len(), 1);
    assert!(
        reopened
            .chart_xml(charts[0].relationship_id())
            .unwrap()
            .is_some()
    );
    assert_eq!(
        save(&after.apply(&patch.inverse()).unwrap().into_package()),
        save(&base)
    );
}
