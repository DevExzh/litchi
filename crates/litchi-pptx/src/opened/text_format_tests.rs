//! Rich text formatting and table-cell edits read back through the scene
//! and the shared DrawingML text-body reader.

use litchi_core::Position;
use litchi_drawingml::text::Alignment;

use super::{Rgb, TextEdit, TextProperties, TextTarget};
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
        .with_fill(Some(Rgb([0xFF, 0xEE, 0x00])))
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
            TextProperties::default().with_fill(Some(Rgb([0, 0, 0]))),
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
