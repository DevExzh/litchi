use std::error::Error;
use std::io;

use litchi_xlsx::{Formula, Workbook};

fn main() -> Result<(), Box<dyn Error>> {
    let output = std::env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example add_sheet -- OUTPUT.xlsx",
        )
    })?;

    let source = Workbook::new()?;
    let mut edit = source.edit()?;
    edit.add("Summary")?
        .set("A1", "Created atomically")?
        .set("B2", Formula::new("1+1")?)?
        .column(2u32)?
        .hide();
    edit.add("Active Data")?.set("A1", 42_i32)?.activate();

    let committed = edit.commit()?;
    committed.workbook().save(output)?;
    println!("committed {} semantic changes", committed.patch().len());
    Ok(())
}
