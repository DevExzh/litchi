//! Typed worksheet styles, conditional formatting, page setup, and shapes.
//!
//! ```bash
//! cargo run -p litchi --example xlsx_typed_page_and_shapes --features ooxml -- output.xlsx
//! ```

use litchi::ooxml::xlsx::writer::{ConditionalFormatType, IconSet, Operator, ShapeSpec};
use litchi::ooxml::xlsx::{
    CellFont, CellFormat, CellMarker, EditAs, Emu, Preset, Rgb, Scheme, ShapeAnchor, Underline,
    Workbook,
    page_setup::{Fit, Orientation, Paper, Setup},
};
use std::path::PathBuf;

fn marker(column: u32, row: u32) -> CellMarker {
    CellMarker {
        column,
        column_offset: Emu(0),
        row,
        row_offset: Emu(0),
    }
}

fn anchor(from: (u32, u32), to: (u32, u32)) -> ShapeAnchor {
    ShapeAnchor::TwoCell {
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

    let mut workbook = Workbook::create()?;
    let sheet = workbook.worksheet_mut(0)?;
    sheet.set_name("Typed API".to_string());
    sheet.set_cell_value(1, 1, "Typed Excel properties");
    sheet.set_cell_format(
        1,
        1,
        CellFormat {
            font: Some(CellFont {
                bold: true,
                underline: Some(Underline::Single),
                scheme: Some(Scheme::Minor),
                ..CellFont::default()
            }),
            ..CellFormat::default()
        },
    );
    sheet.set_tab_color(Rgb::new(0x2F, 0x75, 0xB5));
    for (row, value) in (2..=6).zip([10_i64, 25, 50, 75, 100]) {
        sheet.set_cell_value(row, 1, value);
        sheet.set_cell_value(row, 2, value);
        sheet.set_cell_value(row, 3, value);
    }
    sheet.add_conditional_formatting(
        "A2:A6",
        ConditionalFormatType::ColorScale {
            min_color: Rgb::new(0xF8, 0x69, 0x6B),
            max_color: Rgb::new(0x63, 0xBE, 0x7B),
            mid_color: Some(Rgb::new(0xFF, 0xEB, 0x84)),
        },
        1,
        None,
    );
    sheet.add_conditional_formatting(
        "B2:B6",
        ConditionalFormatType::IconSet {
            icon_set: IconSet::ThreeTrafficLights1,
            show_value: true,
        },
        2,
        None,
    );
    sheet.add_conditional_formatting(
        "C2:C6",
        ConditionalFormatType::CellIs {
            operator: Operator::GreaterThan,
            formula: "50".to_string(),
        },
        3,
        None,
    );
    sheet.set_page(Setup {
        orientation: Some(Orientation::Landscape),
        paper: Some(Paper::A4),
        ..Setup::default()
    });
    sheet.set_fit(Fit::ONE, Fit::NONE);

    sheet.add_shape(ShapeSpec::text_box(
        "Summary",
        anchor((0, 8), (4, 14)),
        Preset::RoundRect,
        "Closed preset domain\nNo string tokens",
    ))?;
    sheet.add_shape(ShapeSpec::shape(
        "Status",
        anchor((5, 8), (9, 14)),
        Preset::Ellipse,
        "Typed ellipse",
    ))?;

    workbook.save(&output)?;
    println!("saved {}", output.display());
    Ok(())
}
