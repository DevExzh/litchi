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
            for (address, cell) in sheet.cells(Rect::ALL)? {
                println!(
                    "  ({}, {}): {cell:?}",
                    address.row().get(),
                    address.column().get()
                );
            }
        }
    }
    Ok(())
}
