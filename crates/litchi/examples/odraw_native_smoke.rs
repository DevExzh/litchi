//! Generate legacy OfficeArt artifacts for native Microsoft Office smoke tests.
//!
//! The output directory defaults to `target/office-smoke` and can be changed
//! with `LITCHI_OFFICE_SMOKE_DIR`.

use std::{error::Error, fs, path::PathBuf};

use litchi::doc::writer::{DocDrawingShape, DocShapeKind, DocWriter, FloatingPosition};
use litchi::ppt::writer::{
    FillStyle, LineStyleConfig, PptWriter, ShapeColor, ShapeStyle, ShapeType, Table,
};
use litchi::xls::writer::shape::{Anchor, Behavior, Rect};
use litchi::xls::writer::{
    XlsShapeColor, XlsShapeFill, XlsShapeGroupChild, XlsShapeGroupWrite, XlsShapeKind,
    XlsShapeLine, XlsShapeText, XlsShapeWrite, XlsWriter,
};

fn main() -> Result<(), Box<dyn Error>> {
    let output = std::env::var_os("LITCHI_OFFICE_SMOKE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/office-smoke"));
    fs::create_dir_all(&output)?;

    let doc = output.join("odraw-smoke.doc");
    let doc_mode =
        std::env::var("LITCHI_OFFICE_SMOKE_DOC_MODE").unwrap_or_else(|_| "full".to_string());
    write_doc(&doc, &doc_mode)?;
    let ppt = output.join("odraw-smoke.ppt");
    write_ppt(&ppt)?;
    let xls = output.join("odraw-smoke.xls");
    let xls_mode =
        std::env::var("LITCHI_OFFICE_SMOKE_XLS_MODE").unwrap_or_else(|_| "full".to_string());
    write_xls(&xls, &xls_mode)?;

    println!("DOC={}", doc.canonicalize()?.display());
    println!("PPT={}", ppt.canonicalize()?.display());
    println!("XLS={}", xls.canonicalize()?.display());
    Ok(())
}

fn write_doc(path: &std::path::Path, mode: &str) -> Result<(), Box<dyn Error>> {
    let mut doc = DocWriter::new();
    doc.add_paragraph("Litchi OfficeArt native smoke test")?;

    if mode == "basic" {
        doc.save(path)?;
        return Ok(());
    }

    doc.insert_floating_shape(
        DocDrawingShape::new(DocShapeKind::Rectangle, 2880, 1440)?
            .with_fill(0x1F, 0x4E, 0x78)
            .with_line(0xFF, 0xFF, 0xFF),
        FloatingPosition::new(900, 1200),
    )?;

    if mode == "shape" {
        doc.save(path)?;
        return Ok(());
    }

    doc.insert_floating_text_box(
        DocDrawingShape::new(DocShapeKind::RoundRectangle, 3600, 1440)?
            .with_fill(0xE2, 0xF0, 0xD9)
            .with_line(0x54, 0x8B, 0x2F),
        FloatingPosition::new(4200, 2400),
        "Typed DOC textbox",
    )?;
    doc.add_paragraph("The document remains editable in Microsoft Word.")?;
    doc.save(path)?;
    Ok(())
}

fn write_ppt(path: &std::path::Path) -> Result<(), Box<dyn Error>> {
    let mut ppt = PptWriter::new_widescreen();
    let slide = ppt.add_slide()?;
    ppt.add_textbox(slide, 50, 25, 650, 40, "Litchi OfficeArt native smoke test")?;
    let rectangle = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::rgb(31, 78, 120)))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::WHITE,
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 80, 120, 220, 120, rectangle)?;
    let ellipse = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::rgb(226, 240, 217)))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(84, 139, 47),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 380, 120, 220, 120, ellipse)?;
    ppt.add_textbox(
        slide,
        80,
        285,
        520,
        60,
        "Typed shapes, anchors, colors, and text",
    )?;
    let mut table = Table::new(2, 2)?;
    table.set_cell_text(0, 0, "Typed")?;
    table.set_cell_text(0, 1, "table")?;
    table.set_cell_text(1, 0, "safe")?;
    table.set_cell_text(1, 1, "anchor")?;
    ppt.add_table(slide, 80, 380, table)?;
    ppt.save(path)?;
    Ok(())
}

fn write_xls(path: &std::path::Path, mode: &str) -> Result<(), Box<dyn Error>> {
    let mut xls = XlsWriter::new();
    let sheet = xls.add_worksheet("OfficeArt")?;
    xls.write_string(sheet, 0, 0, "Litchi OfficeArt native smoke test")?;
    xls.write_number(sheet, 1, 0, 42.0)?;

    if mode == "basic" {
        xls.save(path)?;
        return Ok(());
    }

    let mut rectangle = XlsShapeWrite::new(XlsShapeKind::Rectangle, xls_anchor(1, 2, 5, 8)?);
    rectangle.fill = XlsShapeFill::Solid(XlsShapeColor::rgb(0x1F, 0x4E, 0x78));
    rectangle.line = XlsShapeLine::None;
    xls.add_shape(sheet, rectangle)?;

    if mode == "primitive" {
        xls.save(path)?;
        return Ok(());
    }

    let mut textbox = XlsShapeWrite::new(XlsShapeKind::TextBox, xls_anchor(6, 2, 11, 8)?);
    textbox.text = Some(XlsShapeText::new("Typed XLS textbox 世界"));
    textbox.fill = XlsShapeFill::Solid(XlsShapeColor::rgb(0xE2, 0xF0, 0xD9));
    xls.add_shape(sheet, textbox)?;

    if mode == "textbox" {
        xls.save(path)?;
        return Ok(());
    }

    let mut group = XlsShapeGroupWrite::new(xls_anchor(1, 10, 11, 20)?);
    group.coordinates = Rect::new(0, 0, 2000, 1000)?;
    let mut child = XlsShapeGroupChild::new(XlsShapeKind::Ellipse, Rect::new(0, 0, 900, 900)?);
    child.fill = XlsShapeFill::Solid(XlsShapeColor::rgb(0xFF, 0xC0, 0x00));
    group.children.push(child);
    let mut label = XlsShapeGroupChild::new(XlsShapeKind::TextBox, Rect::new(950, 100, 2000, 900)?);
    label.text = Some(XlsShapeText::new("Grouped"));
    group.children.push(label);
    xls.add_shape_group(sheet, group)?;

    xls.save(path)?;
    Ok(())
}

fn xls_anchor(
    first_col: u16,
    first_row: u32,
    last_col: u16,
    last_row: u32,
) -> litchi::xls::XlsResult<Anchor> {
    Anchor::cells(
        first_row,
        first_col,
        last_row,
        last_col,
        Behavior::MoveAndSize,
    )
}
