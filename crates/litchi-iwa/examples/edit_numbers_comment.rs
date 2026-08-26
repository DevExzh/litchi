//! Update or delete an existing Numbers cell comment.

use std::{env, fs::File};

use litchi_numbers::{Package, SheetSelector, TableSelector, table::CellPosition};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_comment <input.numbers> <output.numbers> <sheet-index> <table-index> <row> <column> <text|--clear>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet = SheetSelector::index(
        arguments
            .next()
            .ok_or("missing zero-based sheet index")?
            .parse::<usize>()?,
    );
    let table = TableSelector::index(
        arguments
            .next()
            .ok_or("missing zero-based table index")?
            .parse::<usize>()?,
    );
    let row = arguments
        .next()
        .ok_or("missing zero-based row")?
        .parse::<usize>()?;
    let column = arguments
        .next()
        .ok_or("missing zero-based column")?
        .parse::<usize>()?;
    let replacement = arguments.collect::<Vec<_>>().join(" ");
    if replacement.is_empty() {
        return Err("missing comment text or --clear".into());
    }

    let position = CellPosition::try_from_usize(row, column)?;
    let package = Package::open(&input)?;

    let commit = if replacement == "--clear" {
        package.clear_table_cell_comment(sheet, table, position)?
    } else {
        package.set_table_cell_comment(sheet, table, position, replacement)?
    };
    let mut destination = File::create(&output)?;
    commit.package().write_to(&mut destination)?;

    let verified = Package::open(&output)?.table_cell_comment(sheet, table, position)?;
    println!("comment={verified:?}");
    Ok(())
}
