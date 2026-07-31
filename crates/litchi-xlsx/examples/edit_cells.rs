use std::error::Error;
use std::io;

use litchi_xlsx::{Formula, Workbook};

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let (workbook, path) = match arguments.as_slice() {
        [output] => (Workbook::new()?, output),
        [input, output] => (Workbook::open(input)?, output),
        _ => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "usage: cargo run -p litchi-xlsx --example edit_cells -- [INPUT.xlsx] OUTPUT.xlsx",
            )
            .into());
        },
    };

    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit.sheet("Sheet1")?.ok_or_else(|| {
            io::Error::new(io::ErrorKind::NotFound, "baseline worksheet is missing")
        })?;
        sheet
            .set("A1", "Litchi cell CRUD")?
            .set("B2", 42_i32)?
            .set("C3", Formula::new("B2*2")?)?
            .set("D4", "temporary")?
            .clear("D4")?
            .set("E5", "removed")?
            .remove("E5")?;
    }
    let commit = edit.commit()?;
    std::fs::write(path, commit.workbook().to_bytes()?)?;
    println!("committed {} semantic changes", commit.patch().len());
    Ok(())
}
