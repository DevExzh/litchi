//! Rich text formatting and table-cell edits read back through the scene
//! and the shared DrawingML text-body reader.

use litchi_core::Position;
use litchi_drawingml::text::Alignment;

use litchi_drawingml::geom::Preset;

use super::{CellFill, Rgb, ShapeFill, TextEdit, TextProperties, TextTarget};
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
    let part = package.edit()?.slides()[0].part_name().clone();
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

fn apply(package: &Package, edits: &[TextEdit<'_>]) -> Result<(Package, usize)> {
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
    let (after, changed) = apply(&package, &[TextEdit::new(run(0, 0), properties)])?;
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
    let (cleared, _) = apply(&after, &[TextEdit::new(run(0, 0), clear)])?;
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
    let (after, changed) = apply(&package, &[TextEdit::new(target, properties)])?;
    assert_eq!(changed, 1);
    let scene_xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&scene_xml)?;
    assert_eq!(scene.at(0)?.text(), Some("one\ntwo"));
    assert!(contains(&scene_xml, "algn=\"ctr\""));
    // An exact repeat changes nothing.
    let (_, unchanged) = apply(&after, &[TextEdit::new(target, properties)])?;
    assert_eq!(unchanged, 0);
    Ok(())
}

#[test]
fn run_text_keeps_properties() -> Result<()> {
    let package = deck()?;
    let bold = TextProperties::default().with_bold(Some(true));
    let (bolded, _) = apply(&package, &[TextEdit::new(run(0, 0), bold)])?;
    let text = TextProperties::default().with_text("x < y & z");
    let (after, _) = apply(&bolded, &[TextEdit::new(run(0, 0), text)])?;
    let xml = slide_xml(&after)?;
    assert!(contains(&xml, "x &lt; y &amp; z"));
    assert!(contains(&xml, "b=\"1\""));
    let scene = crate::shape::Scene::read(&xml)?;
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
    let (after, changed) = apply(&package, &[TextEdit::new(target, properties)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    let cell = table.cell(1, 0)?.expect("cell");
    assert_eq!(cell.text()?, "changed");
    assert_eq!(cell.fill()?, Some([0xFF, 0xEE, 0x00]));
    assert_eq!(table.cell(0, 0)?.expect("cell").fill()?, None);
    let clear = TextProperties::default().with_fill(None);
    let (cleared, _) = apply(&after, &[TextEdit::new(target, clear)])?;
    let xml = slide_xml(&cleared)?;
    let scene = crate::shape::Scene::read(&xml)?;
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
        TextEdit::new(run(0, 9), TextProperties::default().with_bold(Some(true))),
        TextEdit::new(run(0, 0), TextProperties::default().with_size(Some(5))),
        TextEdit::new(run(0, 0), TextProperties::default().with_text("a\nb")),
        TextEdit::new(run(0, 0), TextProperties::default()),
        TextEdit::new(
            TextTarget::Shape { shape: at(0) },
            TextProperties::default().with_fill(Some(CellFill::Rgb(Rgb([0, 0, 0])))),
        ),
        TextEdit::new(
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
            &[TextEdit::new(
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
        TextEdit::new(run(0, 0), properties),
        TextEdit::new(
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
        after.apply(&patch.inverse())?.to_bytes()?,
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
    let (after, changed) = apply(&package, &[TextEdit::new(shape_target(), properties)])?;
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
    let (_, unchanged) = apply(&after, &[TextEdit::new(shape_target(), properties)])?;
    assert_eq!(unchanged, 0);

    // `None` writes `a:noFill`; removal deletes the local fill choice.
    let none = TextProperties::default().with_shape_fill(Some(ShapeFill::None));
    let (hidden, _) = apply(&after, &[TextEdit::new(shape_target(), none)])?;
    assert!(contains(&slide_xml(&hidden)?, "<a:noFill/></p:spPr>"));
    let cleared = TextProperties::default().with_shape_fill(None);
    let (cleared, _) = apply(&hidden, &[TextEdit::new(shape_target(), cleared)])?;
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
        TextEdit::new(run(0, 0), fill),
        TextEdit::new(
            TextTarget::Paragraph {
                shape: at(0),
                paragraph: at(0),
            },
            geometry,
        ),
        // The table frame has no shape properties of its own.
        TextEdit::new(TextTarget::Shape { shape: at(1) }, geometry),
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
    let edits = [TextEdit::new(
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
        after.apply(&patch.inverse())?.to_bytes()?,
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
    let (after, changed) = apply(&package, &[TextEdit::new(span(2, 5), bold)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
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
    let (after, _) = apply(&package, &[TextEdit::new(span(1, 4), text)])?;
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
    assert_eq!(scene.at(0)?.text(), Some("aL & Pa"));
    assert!(contains(&xml, "L &amp; P"));
    // Edits apply in order: later spans address the runs a split made.
    let (after, changed) = apply(
        &package,
        &[
            TextEdit::new(span(3, 4), bold),
            TextEdit::new(
                span(0, 1),
                TextProperties::default().with_italic(Some(true)),
            ),
        ],
    )?;
    assert_eq!(changed, 2);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
    assert_eq!(scene.at(0)?.text(), Some("alpha"));
    // Empty, out-of-range and off-boundary spans are refused.
    for (start, end) in [(2, 2), (4, 9)] {
        assert!(apply(&package, &[TextEdit::new(span(start, end), bold)]).is_err());
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
    let (after, changed) = apply(&package, &[TextEdit::new(cell(0), text)])?;
    assert_eq!(changed, 1);
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
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
    let (after, _) = apply(&package, &[TextEdit::new(span, bold)])?;
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml)?;
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
    assert!(apply(&package, &[TextEdit::new(cell(4), text)]).is_err());
    let broken = TextProperties::default().with_text("a\nb");
    assert!(apply(&package, &[TextEdit::new(cell(0), broken)]).is_err());
    let not_a_table = TextTarget::CellRun {
        shape: at(0),
        row: at(0),
        column: at(0),
        paragraph: at(0),
        run: at(0),
    };
    assert!(apply(&package, &[TextEdit::new(not_a_table, text)]).is_err());
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
    let scene = crate::shape::Scene::read(note.xml())?;
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
        &[TextEdit::new(
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
        let scene = crate::shape::Scene::read(note.xml())?;
        Ok(scene.at(index)?.text().map(str::to_owned))
    };
    assert_eq!(notes_text(&after)?.as_deref(), Some("Speaker final"));
    assert_eq!(
        notes_text(&after.apply(&patch.inverse())?)?.as_deref(),
        Some("Speaker draft")
    );
    // A slide without notes refuses.
    let mut edit = deck()?.edit()?;
    assert!(
        edit.set_notes_text_properties(
            0usize,
            &[TextEdit::new(
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
    let (after, changed) = apply(&package, &[TextEdit::new(target, properties)])?;
    assert_eq!(changed, 1);
    let read = |package: &Package| -> Result<crate::table::CellProperties> {
        let xml = slide_xml(package)?;
        let scene = crate::shape::Scene::read(&xml)?;
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
    let (_, unchanged) = apply(&after, &[TextEdit::new(target, properties)])?;
    assert_eq!(unchanged, 0);
    let none = TextProperties::default()
        .with_fill(Some(CellFill::None))
        .with_cell_anchor(None)
        .with_cell_margin(CellMargin::Left, None);
    let (cleared, _) = apply(&after, &[TextEdit::new(target, none)])?;
    let cell = read(&cleared)?;
    assert_eq!(cell.fill, Some(CellFill::None));
    assert_eq!(cell.anchor, None);
    assert_eq!(cell.margins, [None, None, None, Some(45_720)]);
    // Cell-only properties refuse other targets.
    let mut edit = package.edit()?;
    assert!(
        edit.set_text_properties(
            0usize,
            &[TextEdit::new(
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
    let (after, _) = apply(&package, &[TextEdit::new(run(0, 0), properties)])?;
    let xml = slide_xml(&after)?;
    assert!(contains(&xml, "u=\"dbl\""));
    assert!(contains(&xml, "strike=\"sngStrike\""));
    let clear = TextProperties::default()
        .with_underline(None)
        .with_strike(None);
    let (cleared, _) = apply(&after, &[TextEdit::new(run(0, 0), clear)])?;
    assert_eq!(slide_xml(&cleared)?, slide_xml(&package)?);
    Ok(())
}
