//! Typed worksheet editing plus typed style, page, conditional-formatting,
//! and SpreadsheetDrawing model construction.
//!
//! ```bash
//! cargo run -p litchi --example xlsx_typed_page_and_shapes --features ooxml -- output.xlsx
//! ```

use litchi::drawing::geom::Preset;
use litchi::ooxml::xlsx::conditional_formatting::{IconSet, Kind, Operator};
use litchi::ooxml::xlsx::shapes::{Anchor, CellMarker, EditAs, Emu};
use litchi::ooxml::xlsx::style::format::{CellFont, CellFormat};
use litchi::ooxml::xlsx::style::stylesheet::{Scheme, Underline};
use litchi::ooxml::xlsx::writer::shape::ShapeSpec;
use litchi::ooxml::xlsx::{Fit, Orientation, Paper, Setup};
use litchi::ooxml::xlsx::{Formula, Rgb, Workbook};
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

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("xlsx_typed_page_and_shapes.xlsx"));

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Typed API")?;

    {
        let mut sheet = edit.sheet("Typed API")?.ok_or("missing worksheet")?;
        sheet.set("A1", "Typed Excel properties")?;
        for (row, value) in (2..=6).zip([10_i32, 25, 50, 75, 100]) {
            sheet.set((row, 1), value)?;
            sheet.set((row, 2), value)?;
            sheet.set((row, 3), value)?;
        }
        sheet.set("D2", Formula::new("SUM(A2:A6)")?)?;
    }

    // These typed models are owned by the standalone XLSX crate. They are
    // intentionally kept as values here until the corresponding worksheet
    // package authoring operations are exposed by that facade.
    let _title_format = CellFormat {
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
            "Closed preset domain\nNo string tokens",
        ),
        ShapeSpec::shape(
            "Status",
            anchor((5, 8), (9, 14)),
            Preset::Ellipse,
            "Typed ellipse",
        ),
    ];
    let _accent = Rgb::new(0x2F, 0x75, 0xB5);

    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output)?;
    println!("saved {}", output.display());
    Ok(())
}
