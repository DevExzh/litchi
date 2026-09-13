use std::error::Error;
use std::io;

use litchi_xlsx::{Rect, Workbook, WorksheetKind};

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(io::ErrorKind::InvalidInput, "usage: open <workbook.xlsx>")
    })?;
    let workbook = Workbook::open(path)?;
    for sheet in workbook.sheets() {
        println!(
            "{}: {} ({:?}, {:?}, active={})",
            sheet.position(),
            sheet.name(),
            sheet.kind(),
            sheet.visibility(),
            sheet.is_active()
        );
        if sheet.kind() == WorksheetKind::Worksheet {
            if let Some(defaults) = sheet.defaults()? {
                println!(
                    "  defaults: base_width={}, stored_base_width={:?}, width={:?}, height={}, descent={:?}, custom_height={}, hidden={}, thick_top={}, thick_bottom={}, row_outline={}, column_outline={}",
                    defaults.base_width(),
                    defaults.stored_base_width(),
                    defaults.width().map(litchi_xlsx::layout::Width::get),
                    defaults.height().get(),
                    defaults.descent().map(litchi_xlsx::layout::Descent::get),
                    defaults.custom_height(),
                    defaults.hidden(),
                    defaults.thick_top(),
                    defaults.thick_bottom(),
                    defaults.row_outline().get(),
                    defaults.column_outline().get(),
                );
            }
            let extents = sheet.extents()?;
            println!(
                "  bounds: declared={:?}, used={:?}, stored={:?}",
                extents.declared().map(Rect::a1),
                extents.used().map(Rect::a1),
                extents.stored().map(Rect::a1),
            );
            for range in sheet.merges()? {
                println!("  merged: {range}");
            }
            for (address, cell) in sheet.cells(Rect::ALL)? {
                println!(
                    "  ({}, {}): {cell:?}",
                    address.row().get(),
                    address.column().get()
                );
            }
            for row in sheet.rows()? {
                println!(
                    "  row {} (index {}): hidden={}, height={:?}, descent={:?}, style={:?}, custom_height={}, outline={}, collapsed={}, thick_top={}, thick_bottom={}, phonetic={}, custom_format={}",
                    row.index().get() + 1,
                    row.index().get(),
                    row.hidden(),
                    row.height().map(litchi_xlsx::Height::get),
                    row.descent().map(litchi_xlsx::layout::Descent::get),
                    sheet.row_style(row.index())?,
                    row.custom_height(),
                    row.outline().get(),
                    row.collapsed(),
                    row.thick_top(),
                    row.thick_bottom(),
                    row.phonetic(),
                    row.custom_format(),
                );
            }
            for column in sheet.columns()? {
                println!(
                    "  column {} (index {}): hidden={}, width={:?}, style={:?}, outline={}, collapsed={}, best_fit={}, phonetic={}",
                    column.index().a1(),
                    column.index().get(),
                    column.hidden(),
                    column.width().map(litchi_xlsx::Width::get),
                    sheet.column_style(column.index())?,
                    column.outline().get(),
                    column.collapsed(),
                    column.best_fit(),
                    column.phonetic(),
                );
            }
        }
    }
    Ok(())
}
