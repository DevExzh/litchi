use std::error::Error;
use std::io;

use litchi_xlsx::{Rect, SheetKind, Workbook};

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(io::ErrorKind::InvalidInput, "usage: open <workbook.xlsx>")
    })?;
    let workbook = Workbook::open(path)?;
    for sheet in workbook.sheets() {
        println!(
            "{}: {} ({:?})",
            sheet.position(),
            sheet.name(),
            sheet.kind()
        );
        if sheet.kind() == SheetKind::Worksheet {
            let extents = sheet.extents()?;
            println!(
                "  bounds: declared={:?}, used={:?}, stored={:?}",
                extents.declared().map(Rect::a1),
                extents.used().map(Rect::a1),
                extents.stored().map(Rect::a1),
            );
            for (address, cell) in sheet.cells(Rect::ALL)? {
                println!(
                    "  ({}, {}): {cell:?}",
                    address.row().get(),
                    address.column().get()
                );
            }
            for row in sheet.rows()?.filter(|row| row.hidden()) {
                println!("  hidden row: {}", row.index().get());
            }
            for column in sheet.columns()?.filter(|column| column.hidden()) {
                println!("  hidden column: {}", column.index().get());
            }
        }
    }
    Ok(())
}
