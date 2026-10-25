//! Rich text formatting and table-cell edits read back through the scene
//! and the shared DrawingML text-body reader.

use litchi_core::Position;
use litchi_drawingml::text::Alignment;

use litchi_drawingml::geom::Preset;

use super::{CellFill, Rgb, ShapeFill, TextAssignment, TextProperties, TextTarget};
use crate::{Error, Package, Result};

fn deck() -> Result<Package> {
    let mut builder = crate::builder::Builder::new()?;
    builder
        .presentation()?
        .add_slide()?
        .add_text_box("alpha", 0, 0, 914_400, 914_400);
    let package = builder.build()?;
    let mut edit = package.edit()?;
    edit.add_table(
        0usize,
        &[
            vec!["a1".to_owned(), "b1".to_owned()],
            vec!["a2".to_owned(), "b2".to_owned()],
        ],
        (0, 1_000_000, 2_000_000, 600_000),
    )?;
    Ok(edit.commit()?.into_package())
}

fn slide_xml(package: &Package) -> Result<Vec<u8>> {
    let mut edit = package.edit()?;
    let part = edit.slides()?[0].part_name().clone();
    Ok(package.opc()?.get_part(&part)?.blob().to_vec())
}

fn at(index: usize) -> Position {
    Position::new(index)
}

fn run(paragraph: usize, run: usize) -> TextTarget {
    TextTarget::Run {
        shape: at(0),
        paragraph: at(paragraph),
        run: at(run),
    }
}

fn apply(package: &Package, edits: &[TextAssignment<'_>]) -> Result<(Package, usize)> {
    let mut edit = package.edit()?;
    let changed = edit.set_text_properties(0usize, edits)?;
    Ok((edit.commit()?.into_package(), changed))
}

fn contains(haystack: &[u8], needle: &str) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle.as_bytes())
}

#[test]
fn run_formatting_round_trips() -> Result<()> {
    let package = deck()?;
    let properties = TextProperties::default()
        .with_bold(Some(true))
        .with_italic(Some(true))
        .with_size(Some(2_400))
        .with_color(Some(Rgb([0x1F, 0x4E, 0x79])))
        .with_font(Some("Georgia"));
    let (after, changed) = apply(&package, &[TextAssignment::new(run(0, 0), properties)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    for needle in [
        "b=\"1\"",
        "i=\"1\"",
        "sz=\"2400\"",
        "<a:srgbClr val=\"1F4E79\"/>",
        "<a:latin typeface=\"Georgia\"/>",
        "<a:ea typeface=\"Georgia\"/>",
        "<a:cs typeface=\"Georgia\"/>",
    ] {
        assert!(contains(&xml, needle), "{needle}");
    }
    // Clearing restores the absence of local properties.
    let clear = TextProperties::default()
        .with_bold(None)
        .with_size(None)
        .with_color(None)
        .with_font(None);
    let (cleared, _) = apply(&after, &[TextAssignment::new(run(0, 0), clear)])?;
    let xml = slide_xml(&cleared)?;
    assert!(!contains(&xml, "b=\"1\""));
    assert!(!contains(&xml, "srgbClr val=\"1F4E79\""));
    assert!(!contains(&xml, "Georgia"));
    assert!(contains(&xml, "i=\"1\""));
    Ok(())
}

#[test]
fn paragraph_text_and_alignment() -> Result<()> {
    let package = deck()?;
    let target = TextTarget::Paragraph {
        shape: at(0),
        paragraph: at(0),
    };
    let properties = TextProperties::default()
        .with_text("one\ntwo")
        .with_align(Some(Alignment::Center));
    let (after, changed) = apply(&package, &[TextAssignment::new(target, properties)])?;
    assert_eq!(changed, 1);
    let scene_xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&scene_xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(scene.at(0)?.text(), Some("one\ntwo"));
    assert!(contains(&scene_xml, "algn=\"ctr\""));
    // An exact repeat changes nothing.
    let (_, unchanged) = apply(&after, &[TextAssignment::new(target, properties)])?;
    assert_eq!(unchanged, 0);
    Ok(())
}

#[test]
fn run_text_keeps_properties() -> Result<()> {
    let package = deck()?;
    let bold = TextProperties::default().with_bold(Some(true));
    let (bolded, _) = apply(&package, &[TextAssignment::new(run(0, 0), bold)])?;
    let text = TextProperties::default().with_text("x < y & z");
    let (after, _) = apply(&bolded, &[TextAssignment::new(run(0, 0), text)])?;
    let xml = slide_xml(&after)?;
    assert!(contains(&xml, "x &lt; y &amp; z"));
    assert!(contains(&xml, "b=\"1\""));
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(scene.at(0)?.text(), Some("x < y & z"));
    Ok(())
}

#[test]
fn cell_text_and_fill() -> Result<()> {
    let package = deck()?;
    let target = TextTarget::Cell {
        shape: at(1),
        row: at(1),
        column: at(0),
    };
    let properties = TextProperties::default()
        .with_text("changed")
        .with_fill(Some(CellFill::Rgb(Rgb([0xFF, 0xEE, 0x00]))))
        .with_bold(Some(true));
    let (after, changed) = apply(&package, &[TextAssignment::new(target, properties)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    let cell = table.cell(1, 0)?.expect("cell");
    assert_eq!(cell.text()?, "changed");
    assert_eq!(cell.fill()?, Some([0xFF, 0xEE, 0x00]));
    assert_eq!(table.cell(0, 0)?.expect("cell").fill()?, None);
    let clear = TextProperties::default().with_fill(None);
    let (cleared, _) = apply(&after, &[TextAssignment::new(target, clear)])?;
    let xml = slide_xml(&cleared)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    assert_eq!(table.cell(1, 0)?.expect("cell").fill()?, None);
    Ok(())
}

#[test]
fn refusals_leave_the_edit_unchanged() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    let bad = [
        TextAssignment::new(run(0, 9), TextProperties::default().with_bold(Some(true))),
        TextAssignment::new(run(0, 0), TextProperties::default().with_size(Some(5))),
        TextAssignment::new(run(0, 0), TextProperties::default().with_text("a\nb")),
        TextAssignment::new(run(0, 0), TextProperties::default()),
        TextAssignment::new(
            TextTarget::Shape { shape: at(0) },
            TextProperties::default().with_fill(Some(CellFill::Rgb(Rgb([0, 0, 0])))),
        ),
        TextAssignment::new(
            TextTarget::Cell {
                shape: at(0),
                row: at(0),
                column: at(0),
            },
            TextProperties::default().with_bold(Some(true)),
        ),
    ];
    for item in bad {
        assert!(
            edit.set_text_properties(0usize, &[item]).is_err(),
            "{item:?}"
        );
    }
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(matches!(
        package.edit()?.set_text_properties(
            0usize,
            &[TextAssignment::new(
                run(3, 0),
                TextProperties::default().with_bold(Some(true))
            )]
        ),
        Err(Error::IndexOutOfBounds { .. })
    ));
    Ok(())
}

#[test]
fn source_owner_matches_eager_owner() -> Result<()> {
    use std::sync::Arc;
    let bytes = deck()?.to_bytes()?;
    let properties = TextProperties::default()
        .with_bold(Some(true))
        .with_align(Some(Alignment::Right));
    let edits = [
        TextAssignment::new(run(0, 0), properties),
        TextAssignment::new(
            TextTarget::Paragraph {
                shape: at(0),
                paragraph: at(0),
            },
            properties,
        ),
    ];
    let eager = Package::from_bytes(&bytes)?;
    let (eager, _) = apply(&eager, &edits)?;
    let source =
        crate::source::Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let mut edit = source.edit()?;
    assert_eq!(edit.set_text_properties(Position::new(0), &edits)?, 2);
    let (after, patch) = edit.commit()?.into_parts();
    assert!(!patch.is_noop());
    let reopened = Package::from_bytes(&after.to_bytes()?)?;
    assert_eq!(slide_xml(&reopened)?, slide_xml(&eager)?);
    assert_eq!(
        after.apply(&patch.inverse())?.into_package().to_bytes()?,
        source.to_bytes()?
    );
    Ok(())
}

fn shape_target() -> TextTarget {
    TextTarget::Shape { shape: at(0) }
}

#[test]
fn shape_fill_and_geometry_round_trip() -> Result<()> {
    let package = deck()?;
    let before = slide_xml(&package)?;
    assert!(contains(&before, "<a:prstGeom prst=\"rect\">"));
    let properties = TextProperties::default()
        .with_shape_fill(Some(ShapeFill::Solid(Rgb([0x1E, 0x3A, 0x5F]))))
        .with_geometry(Preset::Ellipse)
        .with_bold(Some(true));
    let (after, changed) = apply(&package, &[TextAssignment::new(shape_target(), properties)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    assert!(contains(
        &xml,
        "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>"
    ));
    assert!(contains(
        &xml,
        "<a:solidFill><a:srgbClr val=\"1E3A5F\"/></a:solidFill></p:spPr>"
    ));
    assert!(
        !contains(&xml, "<a:noFill/></p:spPr>"),
        "the old fill is replaced"
    );
    assert!(contains(&xml, "b=\"1\""));
    // Schema order: geometry before fill, both after the transform.
    let find = |needle: &[u8]| xml.windows(needle.len()).position(|w| w == needle);
    let xfrm = find(b"</a:xfrm>").expect("transform");
    let geometry = find(b"<a:prstGeom").expect("geometry");
    let fill = find(b"<a:solidFill><a:srgbClr val=\"1E3A5F\"").expect("fill");
    assert!(xfrm < geometry && geometry < fill);

    // The same assignments again change nothing.
    let (_, unchanged) = apply(&after, &[TextAssignment::new(shape_target(), properties)])?;
    assert_eq!(unchanged, 0);

    // `None` writes `a:noFill`; removal deletes the local fill choice.
    let none = TextProperties::default().with_shape_fill(Some(ShapeFill::None));
    let (hidden, _) = apply(&after, &[TextAssignment::new(shape_target(), none)])?;
    assert!(contains(&slide_xml(&hidden)?, "<a:noFill/></p:spPr>"));
    let cleared = TextProperties::default().with_shape_fill(None);
    let (cleared, _) = apply(&hidden, &[TextAssignment::new(shape_target(), cleared)])?;
    let xml = slide_xml(&cleared)?;
    assert!(!contains(&xml, "<a:noFill/></p:spPr>"));
    assert!(contains(&xml, "</a:prstGeom></p:spPr>"));
    Ok(())
}

#[test]
fn shape_properties_refuse_other_targets() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    let fill = TextProperties::default().with_shape_fill(Some(ShapeFill::None));
    let geometry = TextProperties::default().with_geometry(Preset::Star5);
    for item in [
        TextAssignment::new(run(0, 0), fill),
        TextAssignment::new(
            TextTarget::Paragraph {
                shape: at(0),
                paragraph: at(0),
            },
            geometry,
        ),
        // The table frame has no shape properties of its own.
        TextAssignment::new(TextTarget::Shape { shape: at(1) }, geometry),
    ] {
        assert!(
            edit.set_text_properties(0usize, &[item]).is_err(),
            "{item:?}"
        );
    }
    assert!(!edit.commit()?.diagnostics().changed());
    Ok(())
}

#[test]
fn source_owner_writes_shape_properties_like_the_eager_owner() -> Result<()> {
    use std::sync::Arc;
    let bytes = deck()?.to_bytes()?;
    let edits = [TextAssignment::new(
        shape_target(),
        TextProperties::default()
            .with_shape_fill(Some(ShapeFill::Solid(Rgb([0xFF, 0x00, 0x00]))))
            .with_geometry(Preset::RoundRect),
    )];
    let eager = Package::from_bytes(&bytes)?;
    let (eager, _) = apply(&eager, &edits)?;
    let source =
        crate::source::Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let mut edit = source.edit()?;
    assert_eq!(edit.set_text_properties(Position::new(0), &edits)?, 1);
    let (after, patch) = edit.commit()?.into_parts();
    let reopened = Package::from_bytes(&after.to_bytes()?)?;
    assert_eq!(slide_xml(&reopened)?, slide_xml(&eager)?);
    assert_eq!(
        after.apply(&patch.inverse())?.into_package().to_bytes()?,
        source.to_bytes()?
    );
    Ok(())
}

#[test]
fn run_spans_split_runs_and_format_only_the_span() -> Result<()> {
    let package = deck()?;
    let span = |start, end| TextTarget::RunSpan {
        shape: at(0),
        paragraph: at(0),
        run: at(0),
        start,
        end,
    };
    // "alpha" → "al" + bold "pha"; the parts keep the run's markup.
    let bold = TextProperties::default().with_bold(Some(true));
    let (after, changed) = apply(&package, &[TextAssignment::new(span(2, 5), bold)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(scene.at(0)?.text(), Some("alpha"));
    assert!(
        contains(&xml, ">al</a:t></a:r><a:r><a:rPr"),
        "{}",
        String::from_utf8_lossy(&xml)
    );
    assert!(contains(&xml, "b=\"1\""));
    // A middle span with replacement text: three runs.
    let text = TextProperties::default()
        .with_text("L & P")
        .with_italic(Some(true));
    let (after, _) = apply(&package, &[TextAssignment::new(span(1, 4), text)])?;
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(scene.at(0)?.text(), Some("aL & Pa"));
    assert!(contains(&xml, "L &amp; P"));
    // Edits apply in order: later spans address the runs a split made.
    let (after, changed) = apply(
        &package,
        &[
            TextAssignment::new(span(3, 4), bold),
            TextAssignment::new(
                span(0, 1),
                TextProperties::default().with_italic(Some(true)),
            ),
        ],
    )?;
    assert_eq!(changed, 2);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(scene.at(0)?.text(), Some("alpha"));
    // Empty, out-of-range and off-boundary spans are refused.
    for (start, end) in [(2, 2), (4, 9)] {
        assert!(apply(&package, &[TextAssignment::new(span(start, end), bold)]).is_err());
    }
    Ok(())
}

/// Record 0279: a field's cached result takes run text; one without an
/// `a:t` gains one after its properties.
#[test]
fn field_results_take_run_text() -> Result<()> {
    let text = TextProperties::default().with_text("7");
    let field =
        br#"<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>&lt;#&gt;</a:t></a:fld>"#;
    let rewritten = super::text_format::rewrite_run(field, &text, Some("7"))?;
    assert_eq!(
        rewritten,
        br#"<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>"#
    );
    let bare = br#"<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/></a:fld>"#;
    let rewritten = super::text_format::rewrite_run(bare, &text, Some("7"))?;
    assert_eq!(
        rewritten,
        br#"<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>"#
    );
    assert!(super::text_format::rewrite_run(b"<a:br/>", &text, Some("7")).is_err());
    Ok(())
}

/// Record 0285: table-cell runs take run text and spans as shape runs do,
/// and the cell's text body reads with the positions they address.
#[test]
fn cell_runs_and_spans_edit_one_run() -> Result<()> {
    let package = deck()?;
    let cell = |run| TextTarget::CellRun {
        shape: at(1),
        row: at(1),
        column: at(1),
        paragraph: at(0),
        run: at(run),
    };
    let text = TextProperties::default().with_text("B2");
    let (after, changed) = apply(&package, &[TextAssignment::new(cell(0), text)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    assert_eq!(table.cell(1, 1)?.expect("cell").text()?, "B2");
    assert_eq!(table.cell(0, 1)?.expect("cell").text()?, "b1");
    let body = table.cell(1, 1)?.expect("cell").text_body()?.expect("body");
    assert_eq!(body.paragraphs[0].runs[0].text, "B2");
    // A span splits the cell's run and formats only itself.
    let span = TextTarget::CellRunSpan {
        shape: at(1),
        row: at(0),
        column: at(0),
        paragraph: at(0),
        run: at(0),
        start: 1,
        end: 2,
    };
    let bold = TextProperties::default()
        .with_text("ONE")
        .with_bold(Some(true));
    let (after, _) = apply(&package, &[TextAssignment::new(span, bold)])?;
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    let body = table.cell(0, 0)?.expect("cell").text_body()?.expect("body");
    let runs: Vec<&str> = body.paragraphs[0]
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect();
    assert_eq!(runs, ["a", "ONE"]);
    // Out-of-range runs, line breaks and a non-table shape are refused.
    assert!(apply(&package, &[TextAssignment::new(cell(4), text)]).is_err());
    let broken = TextProperties::default().with_text("a\nb");
    assert!(apply(&package, &[TextAssignment::new(cell(0), broken)]).is_err());
    let not_a_table = TextTarget::CellRun {
        shape: at(0),
        row: at(0),
        column: at(0),
        paragraph: at(0),
        run: at(0),
    };
    assert!(apply(&package, &[TextAssignment::new(not_a_table, text)]).is_err());
    Ok(())
}

/// Record 0285: a notes page's shapes take the same run edits as a slide's.
#[test]
fn notes_pages_take_run_edits() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    assert!(edit.set_notes_text(0usize, "Speaker draft")?);
    let package = edit.commit()?.into_package();
    let notes = package.notes()?.expect("notes");
    let note = notes.slides().first().expect("note").clone();
    let scene = crate::shape::Scene::read(note.xml(), crate::shape::Limits::DEFAULT)?;
    let (index, _) = (0..scene.len())
        .map(|index| (index, scene.at(index)))
        .find(|(_, shape)| {
            shape
                .as_ref()
                .is_ok_and(|shape| shape.text() == Some("Speaker draft"))
        })
        .expect("the notes body");
    let span = TextTarget::RunSpan {
        shape: at(index),
        paragraph: at(0),
        run: at(0),
        start: 8,
        end: 13,
    };
    let mut edit = package.edit()?;
    let changed = edit.set_notes_text_properties(
        0usize,
        &[TextAssignment::new(
            span,
            TextProperties::default()
                .with_text("final")
                .with_italic(Some(true)),
        )],
    )?;
    assert_eq!(changed, 1);
    let (after, patch) = edit.commit()?.into_parts();
    let notes_text = |package: &Package| -> Result<Option<String>> {
        let notes = package.notes()?.expect("notes");
        let note = notes.slides().first().expect("note");
        let scene = crate::shape::Scene::read(note.xml(), crate::shape::Limits::DEFAULT)?;
        Ok(scene.at(index)?.text().map(str::to_owned))
    };
    assert_eq!(notes_text(&after)?.as_deref(), Some("Speaker final"));
    assert_eq!(
        notes_text(&after.apply(&patch.inverse())?.into_package())?.as_deref(),
        Some("Speaker draft")
    );
    // A slide without notes refuses.
    let mut edit = deck()?.edit()?;
    assert!(
        edit.set_notes_text_properties(
            0usize,
            &[TextAssignment::new(
                span,
                TextProperties::default().with_bold(Some(true))
            )]
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn cell_theme_fill_anchor_and_insets_round_trip() -> Result<()> {
    use super::CellMargin;
    use litchi_drawingml::color::Scheme;
    use litchi_drawingml::text::Anchor;
    let package = deck()?;
    let target = TextTarget::Cell {
        shape: at(1),
        row: at(0),
        column: at(1),
    };
    let properties = TextProperties::default()
        .with_fill(Some(CellFill::Scheme(Scheme::Accent1)))
        .with_cell_anchor(Some(Anchor::Center))
        .with_cell_margin(CellMargin::Left, Some(91_440))
        .with_cell_margin(CellMargin::Bottom, Some(45_720));
    let (after, changed) = apply(&package, &[TextAssignment::new(target, properties)])?;
    assert_eq!(changed, 1);
    let read = |package: &Package| -> Result<crate::table::CellProperties> {
        let xml = slide_xml(package)?;
        let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
        let frame = scene.at(1)?.xml()?.to_vec();
        let table = crate::table::Table::from_graphic_frame(&frame)?;
        table.cell(0, 1)?.expect("cell").properties()
    };
    let cell = read(&after)?;
    assert_eq!(cell.fill, Some(CellFill::Scheme(Scheme::Accent1)));
    assert_eq!(cell.anchor, Some(Anchor::Center));
    assert_eq!(cell.margins, [Some(91_440), None, None, Some(45_720)]);
    assert!(contains(
        &slide_xml(&after)?,
        "<a:schemeClr val=\"accent1\"/>"
    ));
    // An exact repeat changes nothing; `none` writes an explicit no-fill.
    let (_, unchanged) = apply(&after, &[TextAssignment::new(target, properties)])?;
    assert_eq!(unchanged, 0);
    let none = TextProperties::default()
        .with_fill(Some(CellFill::None))
        .with_cell_anchor(None)
        .with_cell_margin(CellMargin::Left, None);
    let (cleared, _) = apply(&after, &[TextAssignment::new(target, none)])?;
    let cell = read(&cleared)?;
    assert_eq!(cell.fill, Some(CellFill::None));
    assert_eq!(cell.anchor, None);
    assert_eq!(cell.margins, [None, None, None, Some(45_720)]);
    // Cell-only properties refuse other targets.
    let mut edit = package.edit()?;
    assert!(
        edit.set_text_properties(
            0usize,
            &[TextAssignment::new(
                run(0, 0),
                TextProperties::default().with_cell_anchor(Some(Anchor::Top))
            )]
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn underline_and_strike_round_trip() -> Result<()> {
    use litchi_drawingml::text::{Strike, Underline};
    let package = deck()?;
    let properties = TextProperties::default()
        .with_underline(Some(Underline::Double))
        .with_strike(Some(Strike::Single));
    let (after, _) = apply(&package, &[TextAssignment::new(run(0, 0), properties)])?;
    let xml = slide_xml(&after)?;
    assert!(contains(&xml, "u=\"dbl\""));
    assert!(contains(&xml, "strike=\"sngStrike\""));
    let clear = TextProperties::default()
        .with_underline(None)
        .with_strike(None);
    let (cleared, _) = apply(&after, &[TextAssignment::new(run(0, 0), clear)])?;
    assert_eq!(slide_xml(&cleared)?, slide_xml(&package)?);
    Ok(())
}

#[test]
fn row_cell_counts_match_the_row_reader() -> Result<()> {
    let package = deck()?;
    let xml = slide_xml(&package)?;
    let scene = crate::shape::owned::Scene::read(xml, crate::shape::Limits::DEFAULT)?;
    let handle = scene
        .shape(1_usize)
        .map_err(|_| Error::Invalid("selector".into()))?
        .expect("table frame");
    let table = handle.table()?;
    let counts = table.row_cell_counts()?;
    let rows: Vec<usize> = table
        .rows()?
        .into_iter()
        .map(|row| row.cell_count())
        .collect();
    assert_eq!(counts, rows);
    assert_eq!(counts, [2, 2]);
    Ok(())
}

#[test]
fn cell_contents_match_the_cell_readers() -> Result<()> {
    let package = deck()?;
    let xml = slide_xml(&package)?;
    let scene = crate::shape::owned::Scene::read(xml, crate::shape::Limits::DEFAULT)?;
    let handle = scene
        .shape(1_usize)
        .map_err(|_| Error::Invalid("selector".into()))?
        .expect("table frame");
    let table = handle.table()?;
    let mut expected = Vec::new();
    for row in table.rows()? {
        let mut cells = Vec::new();
        for cell in row.cells()? {
            cells.push((cell.properties()?, cell.text_body()?));
        }
        expected.push(cells);
    }
    assert_eq!(table.cell_contents()?, expected);
    assert!(expected[0][0].1.is_some());
    Ok(())
}

#[test]
fn chunked_text_edits_equal_one_call_per_chunk() -> Result<()> {
    let package = deck()?;
    let bold = TextProperties::default().with_bold(Some(true));
    let italic = TextProperties::default().with_italic(Some(true));
    let first = [TextAssignment::new(run(0, 0), bold)];
    let second = [
        TextAssignment::new(run(0, 0), italic),
        TextAssignment::new(
            TextTarget::Cell {
                shape: at(1),
                row: at(0),
                column: at(0),
            },
            bold,
        ),
    ];
    let mut one = package.edit()?;
    one.set_text_properties(0usize, &first)?;
    one.set_text_properties(0usize, &second)?;
    let expected = slide_xml(&one.commit()?.into_package())?;
    let mut chunked = package.edit()?;
    assert_eq!(
        chunked.set_text_property_chunks(0usize, &[&first[..], &second[..]])?,
        3
    );
    assert_eq!(slide_xml(&chunked.commit()?.into_package())?, expected);
    let mut empty = package.edit()?;
    assert_eq!(empty.set_text_property_chunks(0usize, &[])?, 0);
    let too_many = vec![TextAssignment::new(run(0, 0), bold); 257];
    assert!(
        empty
            .set_text_property_chunks(0usize, &[&too_many])
            .is_err()
    );
    Ok(())
}

#[test]
fn batched_paragraph_and_cell_edits_equal_one_edit_at_a_time() -> Result<()> {
    // Record 0346: paragraph, run and cell edits of one shape are staged on
    // their own spans and spliced once; the result is the sequential one.
    let package = deck()?;
    let paragraph = |index: usize| TextTarget::Paragraph {
        shape: at(0),
        paragraph: at(index),
    };
    let cell = |row: usize, column: usize| TextTarget::CellParagraph {
        shape: at(1),
        row: at(row),
        column: at(column),
        paragraph: at(0),
    };
    let edits = [
        TextAssignment::new(run(0, 0), TextProperties::default().with_bold(Some(true))),
        TextAssignment::new(
            paragraph(0),
            TextProperties::default().with_align(Some(Alignment::Center)),
        ),
        // A whole-shape edit after the paragraph edits sees them.
        TextAssignment::new(
            TextTarget::Shape { shape: at(0) },
            TextProperties::default().with_italic(Some(true)),
        ),
        TextAssignment::new(run(0, 0), TextProperties::default().with_size(Some(2_000))),
        TextAssignment::new(cell(0, 0), TextProperties::default().with_bold(Some(true))),
        TextAssignment::new(cell(1, 1), TextProperties::default().with_text("changed")),
        TextAssignment::new(
            cell(0, 0),
            TextProperties::default().with_italic(Some(true)),
        ),
    ];
    let (batched, changed) = apply(&package, &edits)?;
    assert_eq!(changed, edits.len());
    let mut sequential = package;
    for edit in &edits {
        sequential = apply(&sequential, std::slice::from_ref(edit))?.0;
    }
    assert_eq!(slide_xml(&batched)?, slide_xml(&sequential)?);
    Ok(())
}

#[test]
fn per_edit_changes_equal_one_edit_at_a_time() -> Result<()> {
    // Record 0348: the per-edit changes of a chunked batch are the changes
    // its edits report applied one after the other, on both owners.
    let package = deck()?;
    let cell = |row: usize, column: usize| TextTarget::Cell {
        shape: at(1),
        row: at(row),
        column: at(column),
    };
    let cell_paragraph = |row: usize, column: usize| TextTarget::CellParagraph {
        shape: at(1),
        row: at(row),
        column: at(column),
        paragraph: at(0),
    };
    let bold = TextProperties::default().with_bold(Some(true));
    let fill = TextProperties::default().with_fill(Some(CellFill::Rgb(Rgb([0xFF, 0xEE, 0xCC]))));
    let edits = [
        TextAssignment::new(cell(0, 0), bold),
        // The same change again: no bytes change.
        TextAssignment::new(cell(0, 0), bold),
        TextAssignment::new(cell(0, 1), fill),
        TextAssignment::new(
            cell_paragraph(1, 0),
            TextProperties::default().with_align(Some(Alignment::Center)),
        ),
        TextAssignment::new(
            cell_paragraph(1, 1),
            TextProperties::default().with_text("b2"),
        ),
        TextAssignment::new(run(0, 0), TextProperties::default().with_italic(Some(true))),
    ];
    let mut expected = Vec::new();
    let mut sequential = deck()?;
    for edit in &edits {
        let (after, changed) = apply(&sequential, std::slice::from_ref(edit))?;
        expected.push(changed == 1);
        sequential = after;
    }
    assert_eq!(expected, [true, false, true, true, false, true]);
    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(2).collect();
    let mut batched = package.edit()?;
    assert_eq!(
        batched.set_text_property_chunks_per_edit(0usize, &chunks)?,
        expected
    );
    assert_eq!(
        slide_xml(&batched.commit()?.into_package())?,
        slide_xml(&sequential)?
    );
    let mut empty = package.edit()?;
    assert!(
        empty
            .set_text_property_chunks_per_edit(0usize, &[])?
            .is_empty()
    );

    let source = crate::source::Package::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(deck()?.to_bytes()?),
    ))?;
    let mut edit = source.edit()?;
    assert_eq!(
        edit.set_text_property_chunks_per_edit(at(0), &chunks)?,
        expected
    );
    let mut counted = source.edit()?;
    assert_eq!(counted.set_text_property_chunks(at(0), &chunks)?, 4);
    Ok(())
}

#[test]
fn notes_chunks_equal_one_call_per_chunk() -> Result<()> {
    // Record 0348: one notes-part pass for every chunk.
    let package = deck()?;
    let mut edit = package.edit()?;
    assert!(edit.set_notes_text(0usize, "one")?);
    let package = edit.commit()?.into_package();
    let notes = package.notes()?.expect("notes");
    let note = notes.slides().first().expect("note").clone();
    let scene = crate::shape::Scene::read(note.xml(), crate::shape::Limits::DEFAULT)?;
    let body = (0..scene.len())
        .find(|index| {
            scene
                .at(*index)
                .is_ok_and(|shape| shape.text().is_some_and(|text| text.starts_with("one")))
        })
        .expect("the notes body");
    let paragraph = |index: usize| TextTarget::Paragraph {
        shape: at(body),
        paragraph: at(index),
    };
    let notes_run = TextTarget::Run {
        shape: at(body),
        paragraph: at(0),
        run: at(0),
    };
    let edits = [
        TextAssignment::new(
            paragraph(0),
            TextProperties::default().with_align(Some(Alignment::Right)),
        ),
        TextAssignment::new(notes_run, TextProperties::default().with_bold(Some(true))),
        TextAssignment::new(
            paragraph(0),
            TextProperties::default().with_italic(Some(true)),
        ),
        TextAssignment::new(notes_run, TextProperties::default().with_bold(Some(true))),
        TextAssignment::new(paragraph(0), TextProperties::default().with_text("one two")),
        // Unchanged after the edit before it.
        TextAssignment::new(paragraph(0), TextProperties::default().with_text("one two")),
    ];
    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(2).collect();
    let notes_xml = |package: &Package| -> Result<Vec<u8>> {
        let notes = package.notes()?.expect("notes");
        Ok(notes.slides().first().expect("note").xml().to_vec())
    };
    let mut one = package.edit()?;
    let mut changed = 0;
    for chunk in &chunks {
        changed += one.set_notes_text_properties(0usize, chunk)?;
    }
    let expected = notes_xml(&one.commit()?.into_package())?;
    let mut chunked = package.edit()?;
    assert_eq!(
        chunked.set_notes_text_property_chunks(0usize, &chunks)?,
        changed
    );
    assert_eq!(notes_xml(&chunked.commit()?.into_package())?, expected);
    // Each edit's change, as one call per edit applied in order reports it.
    let mut sequential = package.edit()?;
    let mut changes = Vec::new();
    for edit in &edits {
        changes
            .push(sequential.set_notes_text_properties(0usize, std::slice::from_ref(edit))? == 1);
    }
    assert_eq!(changes.last(), Some(&false));
    let mut reported = package.edit()?;
    assert_eq!(
        reported.set_notes_text_property_chunks_per_edit(0usize, &chunks)?,
        changes
    );
    // A slide without notes refuses.
    let mut edit = deck()?.edit()?;
    assert!(
        edit.set_notes_text_property_chunks(0usize, &chunks)
            .is_err()
    );
    Ok(())
}

#[test]
fn text_indexes_of_the_corpus_equal_the_readers() -> Result<()> {
    // Record 0348: every table of the tracked PPTX corpus.
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut pending = vec![root];
    let (mut tables, mut scanned, mut bodies) = (0usize, 0usize, 0usize);
    while let Some(directory) = pending.pop() {
        for entry in std::fs::read_dir(&directory)? {
            let path = entry?.path();
            if path.is_dir() {
                pending.push(path);
                continue;
            }
            if path.extension().is_none_or(|extension| extension != "pptx") {
                continue;
            }
            let Ok(package) = Package::from_bytes(&std::fs::read(&path)?) else {
                continue;
            };
            let Ok(mut edit) = package.edit() else {
                continue;
            };
            let opc = package.opc()?;
            let Ok(slides) = edit.slides() else {
                continue;
            };
            for slide in slides {
                let xml = opc.get_part(slide.part_name())?.blob().to_vec();
                let Ok(scene) =
                    crate::shape::owned::Scene::read(xml, crate::shape::Limits::DEFAULT)
                else {
                    continue;
                };
                for position in 0..scene.len() {
                    let Ok(Some(handle)) = scene.shape(position) else {
                        continue;
                    };
                    let common = handle.view().common();
                    if let Ok(Some(body)) = common.text_body() {
                        bodies += 1;
                        let expected = crate::table::CellTextIndex::from_body(&body)?;
                        let index = common.text_index()?.expect("a body");
                        assert_eq!(common.read_text_index()?.as_ref(), Some(&index));
                        assert_eq!(index.paragraph_count(), body.paragraphs.len());
                        for (paragraph, item) in body.paragraphs.iter().enumerate() {
                            let runs = item.runs.iter().filter(|run| run.is_text()).count();
                            for ordinal in 0..=runs {
                                assert_eq!(
                                    index.text_run(paragraph, ordinal),
                                    expected.text_run(0, 0, paragraph, ordinal),
                                    "{} {position}",
                                    path.display()
                                );
                            }
                        }
                    }
                    let Ok(table) = handle.table() else {
                        continue;
                    };
                    tables += 1;
                    scanned += usize::from(table.scan_cell_text_index()?.is_some());
                    assert_eq!(handle.cell_text_index()?, table.cell_text_index()?);
                    assert_eq!(
                        table.cell_text_index()?,
                        crate::table::CellTextIndex::from_cells(table)?,
                        "{} {position}",
                        path.display()
                    );
                }
            }
        }
    }
    assert!(
        tables >= 5 && scanned >= 5 && bodies >= 100,
        "{tables} {scanned} {bodies}"
    );
    Ok(())
}

#[test]
fn a_litchi_table_frame_is_indexed_in_one_pass() -> Result<()> {
    // Record 0348: a frame declaring its namespaces on its root (as
    // `add_table` writes it) takes the one-pass frame index.
    let xml = slide_xml(&deck()?)?;
    let start = memchr::memmem::find(&xml, b"<p:graphicFrame").expect("a frame");
    let end = start
        + memchr::memmem::find(&xml[start..], b"</p:graphicFrame>").expect("its end")
        + b"</p:graphicFrame>".len();
    let index = crate::table::shape::scan_frame_cell_text_index(&xml[start..end], &xml, start)?
        .expect("the one-pass index");
    let scene = crate::shape::owned::Scene::read(xml, crate::shape::Limits::DEFAULT)?;
    let handle = scene
        .shape(1_usize)
        .map_err(|_| Error::Invalid("selector".into()))?
        .expect("table frame");
    assert_eq!(index, handle.table()?.cell_text_index()?);
    assert_eq!(index, handle.cell_text_index()?);
    assert_eq!(index.paragraph_count(1, 1), Some(1));
    Ok(())
}
