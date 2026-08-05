//! Snapshot/edit XLSX example with typed feature models.
//!
//! The workbook mutation surface intentionally covers the currently supported
//! lossless cell, formula, merge, row, and column edits. Styles, conditional
//! formatting, page setup, and worksheet drawings are constructed below as
//! typed models so this example remains a compile-checked guide to those
//! domains; their worksheet package authoring operations are not exposed by
//! the current facade yet.

use litchi::drawing::geom::Preset;
use litchi::ooxml::xlsx::conditional_formatting::{IconSet, Kind, Operator};
use litchi::ooxml::xlsx::shapes::{Anchor, CellMarker, EditAs, Emu};
use litchi::ooxml::xlsx::style::format::{CellFont, CellFormat};
use litchi::ooxml::xlsx::style::stylesheet::{Scheme, Underline};
use litchi::ooxml::xlsx::writer::shape::ShapeSpec;
use litchi::ooxml::xlsx::{Fit, Formula, Orientation, Paper, Setup};
use litchi::ooxml::xlsx::{Number, Value, Workbook};
use std::path::PathBuf;

fn marker(column: u32, row: u32) -> CellMarker {
    CellMarker {
        column,
        column_offset: Emu(0),
        row,
        row_offset: Emu(0),
    }
}

fn anchor(from: (u32, u32), to: (u32, u32)) -> Anchor {
    Anchor::TwoCell {
        from: marker(from.0, from.1),
        to: marker(to.0, to.1),
        edit_as: EditAs::TwoCell,
    }
}

fn typed_feature_models() {
    let _header_format = CellFormat {
        font: Some(CellFont {
            bold: true,
            underline: Some(Underline::Single),
            scheme: Some(Scheme::Minor),
            ..CellFont::default()
        }),
        ..CellFormat::default()
    };
    let _conditional_kinds = [Kind::ColorScale, Kind::IconSet, Kind::CellIs];
    let _conditional_operator = Operator::GreaterThan;
    let _icon_set = IconSet::ThreeTrafficLights1;
    let _page_setup = Setup {
        orientation: Some(Orientation::Landscape),
        paper: Some(Paper::A4),
        fit_to_width: Some(Fit::ONE),
        fit_to_height: Some(Fit::NONE),
        ..Setup::default()
    };
    let _shapes = [
        ShapeSpec::text_box(
            "Summary",
            anchor((0, 8), (4, 14)),
            Preset::RoundRect,
            "Typed worksheet drawing",
        ),
        ShapeSpec::shape(
            "Status",
            anchor((5, 8), (9, 14)),
            Preset::Ellipse,
            "Authoring boundary",
        ),
    ];
}

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("xlsx_comprehensive_features.xlsx"));

    // Immutable snapshots make the source state cheap to share; one edit
    // publishes all worksheet changes atomically at commit time.
    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("missing initial worksheet tab")?
        .rename("Sales")?;

    {
        let mut sheet = edit.sheet("Sales")?.ok_or("missing Sales worksheet")?;
        sheet.set("A1", "Product")?;
        sheet.set("B1", "Quantity")?;
        sheet.set("C1", "Price")?;
        sheet.set("D1", "Total")?;

        sheet.set("A2", Value::text("Widget A"))?;
        sheet.set("B2", Value::Number(Number::new("10")?))?;
        sheet.set("C2", Value::Number(Number::new("25.50")?))?;
        sheet.set("D2", Formula::new("B2*C2")?)?;
        sheet.set("A3", Value::text("Widget B"))?;
        sheet.set("B3", Value::Number(Number::new("5")?))?;
        sheet.set("C3", Value::Number(Number::new("42.00")?))?;
        sheet.set("D3", Formula::new("B3*C3")?)?;

        sheet.merge("A5:D5")?.set("A5", "Sales summary")?;
        sheet.column("A")?.width(18.0)?;
        sheet.column("B")?.width(12.0)?;
        sheet.column("C")?.width(12.0)?;
        sheet.column("D")?.width(14.0)?;
        sheet.row(1)?.height(22.0)?.thick_bottom();
    }

    let mut links = edit.add("Links")?;
    links.set("A1", "External link data")?;
    links.set("B1", "https://example.com")?;
    links.set("A2", "Snapshot/edit example")?;

    let mut analysis = edit.add("Analysis")?;
    analysis.set("A1", "Value")?;
    for (row, value) in (2..=6).zip([10, 25, 50, 75, 100]) {
        analysis.set((row, 1), Value::Number(Number::new(value.to_string())?))?;
    }
    analysis.set("B1", "Total")?;
    analysis.set("B2", Formula::new("SUM(A2:A6)")?)?;
    analysis.activate();

    typed_feature_models();

    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output)?;
    println!("saved {}", output.display());
    Ok(())
}
