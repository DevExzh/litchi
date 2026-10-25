//! Paragraph and run insertion and removal in shape, table-cell and
//! notes-page text bodies (record 0333).

use litchi_core::Position;
use litchi_drawingml::text::Alignment;

use super::{Rgb, TextBody, TextProperties};
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
    edit.set_notes_text(0usize, "note one")?;
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

const SHAPE: TextBody = TextBody::Shape {
    shape: Position::new(0),
};

fn text(xml: &[u8], shape: usize) -> Result<String> {
    let scene = crate::shape::Scene::read(xml, crate::shape::Limits::DEFAULT)?;
    Ok(scene.at(shape)?.text().unwrap_or_default().to_owned())
}

fn contains(haystack: &[u8], needle: &str) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle.as_bytes())
}

/// Whether `after` is `before` with `inserted` spliced in once.
fn only_inserted(before: &[u8], after: &[u8], inserted: &str) -> bool {
    let after = String::from_utf8_lossy(after);
    after.len() == before.len() + inserted.len()
        && after.contains(inserted)
        && after.replacen(inserted, "", 1).as_bytes() == before
}

/// The paragraph texts of a table cell's body.
fn cell_paragraphs(xml: &[u8], row: usize, column: usize) -> Result<Vec<String>> {
    let scene = crate::shape::Scene::read(xml, crate::shape::Limits::DEFAULT)?;
    let frame = scene.at(1)?.xml()?.to_vec();
    let table = crate::table::Table::from_graphic_frame(&frame)?;
    let body = table
        .cell(row, column)?
        .expect("cell")
        .text_body()?
        .expect("cell body");
    Ok(body
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()).collect())
        .collect())
}

#[test]
fn paragraphs_insert_before_after_and_at_the_end() -> Result<()> {
    let package = deck()?;
    let before = slide_xml(&package)?;
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        SHAPE,
        at(1),
        TextProperties::default().with_text("omega"),
    )?;
    edit.insert_paragraph(
        0usize,
        SHAPE,
        at(0),
        TextProperties::default()
            .with_text("first")
            .with_bold(Some(true))
            .with_align(Some(Alignment::Center)),
    )?;
    let after = edit.commit()?.into_package();
    let xml = slide_xml(&after)?;
    assert_eq!(text(&xml, 0)?, "first\nalpha\nomega");
    assert!(contains(&xml, "algn=\"ctr\""));
    assert!(contains(&xml, "b=\"1\""));
    // The source deck is unchanged and the inserted paragraphs are the
    // only new bytes.
    assert_eq!(slide_xml(&package)?, before);
    let mut one = package.edit()?;
    one.insert_paragraph(
        0usize,
        SHAPE,
        at(1),
        TextProperties::default().with_text("x & y"),
    )?;
    assert!(only_inserted(
        &before,
        &slide_xml(&one.commit()?.into_package())?,
        "<a:p><a:r><a:t>x &amp; y</a:t></a:r></a:p>"
    ));
    Ok(())
}

#[test]
fn an_empty_paragraph_carries_its_formatting_as_end_properties() -> Result<()> {
    let package = deck()?;
    let before = slide_xml(&package)?;
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        SHAPE,
        at(1),
        TextProperties::default().with_size(Some(800)),
    )?;
    let xml = slide_xml(&edit.commit()?.into_package())?;
    assert!(only_inserted(
        &before,
        &xml,
        "<a:p><a:endParaRPr sz=\"800\"/></a:p>"
    ));
    assert_eq!(text(&xml, 0)?, "alpha\n");
    Ok(())
}

#[test]
fn paragraph_removal_keeps_the_last_paragraph() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        SHAPE,
        at(1),
        TextProperties::default().with_text("beta"),
    )?;
    edit.remove_paragraph(0usize, SHAPE, at(0))?;
    let after = edit.commit()?.into_package();
    assert_eq!(text(&slide_xml(&after)?, 0)?, "beta");
    let mut edit = after.edit()?;
    assert!(matches!(
        edit.remove_paragraph(0usize, SHAPE, at(0)),
        Err(Error::Refused(_))
    ));
    assert!(matches!(
        edit.remove_paragraph(0usize, SHAPE, at(3)),
        Err(Error::IndexOutOfBounds { index: 3, len: 1 })
    ));
    Ok(())
}

#[test]
fn runs_insert_and_remove_around_existing_runs() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    edit.insert_run(
        0usize,
        SHAPE,
        at(0),
        at(1),
        TextProperties::default()
            .with_text(" omega")
            .with_italic(Some(true))
            .with_color(Some(Rgb([0xC0, 0, 0]))),
    )?;
    edit.insert_run(
        0usize,
        SHAPE,
        at(0),
        at(0),
        TextProperties::default().with_text("pre "),
    )?;
    let after = edit.commit()?.into_package();
    let xml = slide_xml(&after)?;
    assert_eq!(text(&xml, 0)?, "pre alpha omega");
    assert!(contains(
        &xml,
        "<a:r><a:rPr i=\"1\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></a:rPr><a:t> omega</a:t></a:r>"
    ));
    let mut edit = after.edit()?;
    edit.remove_run(0usize, SHAPE, at(0), at(1))?;
    assert_eq!(
        text(&slide_xml(&edit.commit()?.into_package())?, 0)?,
        "pre  omega"
    );
    let mut edit = after.edit()?;
    assert!(matches!(
        edit.insert_run(
            0usize,
            SHAPE,
            at(0),
            at(4),
            TextProperties::default().with_text("x")
        ),
        Err(Error::IndexOutOfBounds { index: 4, len: 3 })
    ));
    assert!(
        edit.insert_run(
            0usize,
            SHAPE,
            at(0),
            at(0),
            TextProperties::default().with_text("a\nb")
        )
        .is_err()
    );
    assert!(
        edit.insert_run(
            0usize,
            SHAPE,
            at(0),
            at(0),
            TextProperties::default().with_align(Some(Alignment::Right))
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn a_run_joins_an_empty_paragraph_before_its_end_properties() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        SHAPE,
        at(1),
        TextProperties::default().with_size(Some(1_200)),
    )?;
    edit.insert_run(
        0usize,
        SHAPE,
        at(1),
        at(0),
        TextProperties::default().with_text("late"),
    )?;
    let xml = slide_xml(&edit.commit()?.into_package())?;
    assert!(contains(
        &xml,
        "<a:p><a:r><a:t>late</a:t></a:r><a:endParaRPr sz=\"1200\"/></a:p>"
    ));
    assert_eq!(text(&xml, 0)?, "alpha\nlate");
    Ok(())
}

#[test]
fn table_cell_bodies_take_paragraphs_and_runs() -> Result<()> {
    let package = deck()?;
    let cell = TextBody::Cell {
        shape: at(1),
        row: at(1),
        column: at(0),
    };
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        cell,
        at(1),
        TextProperties::default().with_text("second"),
    )?;
    edit.insert_run(
        0usize,
        cell,
        at(0),
        at(1),
        TextProperties::default().with_text("!"),
    )?;
    let after = edit.commit()?.into_package();
    let xml = slide_xml(&after)?;
    assert_eq!(cell_paragraphs(&xml, 1, 0)?, ["a2!", "second"]);
    let mut edit = after.edit()?;
    edit.remove_paragraph(0usize, cell, at(0))?;
    let xml = slide_xml(&edit.commit()?.into_package())?;
    assert_eq!(cell_paragraphs(&xml, 1, 0)?, ["second"]);
    // A cell body on a text shape, or a shape body on a table, is refused.
    let mut edit = package.edit()?;
    assert!(
        edit.insert_paragraph(
            0usize,
            TextBody::Cell {
                shape: at(0),
                row: at(0),
                column: at(0)
            },
            at(0),
            TextProperties::default()
        )
        .is_err()
    );
    assert!(
        edit.insert_paragraph(
            0usize,
            TextBody::Shape { shape: at(1) },
            at(0),
            TextProperties::default()
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn notes_bodies_take_paragraphs() -> Result<()> {
    let package = deck()?;
    let notes = package
        .presentation()?
        .slides()?
        .into_iter()
        .next()
        .expect("slide")
        .notes()?
        .expect("notes page");
    let body = notes
        .scene()
        .roots()
        .find(|handle| {
            handle
                .view()
                .placeholder()
                .is_some_and(|placeholder| placeholder.kind() == Some("body"))
        })
        .expect("notes body")
        .position()
        .get();
    let target = TextBody::Notes {
        shape: Position::new(body),
    };
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        0usize,
        target,
        at(1),
        TextProperties::default().with_text("note two"),
    )?;
    let after = edit.commit()?.into_package();
    let notes = after
        .presentation()?
        .slides()?
        .into_iter()
        .next()
        .expect("slide")
        .notes()?
        .expect("notes page");
    assert_eq!(notes.text(), Some("note one\nnote two"));
    Ok(())
}

#[test]
fn source_edits_compose_text_structure_and_refuse_notes() -> Result<()> {
    let bytes = deck()?.to_bytes()?;
    let package = crate::source::Package::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(bytes),
    ))?;
    let mut edit = package.edit()?;
    edit.insert_paragraph(
        at(0),
        SHAPE,
        at(1),
        TextProperties::default().with_text("beta"),
    )?;
    edit.insert_run(
        at(0),
        SHAPE,
        at(1),
        at(1),
        TextProperties::default().with_text("!"),
    )?;
    assert!(matches!(
        edit.insert_paragraph(
            at(0),
            TextBody::Notes { shape: at(0) },
            at(0),
            TextProperties::default()
        ),
        Err(Error::Refused(_))
    ));
    let (target, patch) = edit.commit()?.into_parts();
    assert!(!patch.is_noop());
    let reopened = Package::from_bytes(&target.to_bytes()?)?;
    assert_eq!(text(&slide_xml(&reopened)?, 0)?, "alpha\nbeta!");
    Ok(())
}
