use std::error::Error;
use std::io;

use litchi_xlsx::{Formula, Workbook};

fn main() -> Result<(), Box<dyn Error>> {
    let output = std::env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example insert_sheet -- OUTPUT.xlsx",
        )
    })?;

    let source = Workbook::new()?;
    let mut edit = source.edit()?;
    edit.add_before("Inputs", "Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?
        .set("A1", "Revenue")?
        .set("B1", 120_i32)?;
    edit.add_after("Results", "Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?
        .set("A1", "Projected")?
        .set("B1", Formula::new("Inputs!B1*1.1")?)?
        .activate();
    edit.add("Archive")?.set("A1", "Tail insertion")?;

    let committed = edit.commit()?;
    committed.workbook().save(output)?;
    println!("committed {} semantic changes", committed.patch().len());
    Ok(())
}
