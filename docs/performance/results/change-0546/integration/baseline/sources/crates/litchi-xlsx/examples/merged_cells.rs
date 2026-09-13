use std::error::Error;
use std::io;

use litchi_xlsx::{Rect, Workbook, cell};

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let [output] = arguments.as_slice() else {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example merged_cells -- OUTPUT.xlsx",
        )
        .into());
    };

    let workbook = Workbook::new()?;
    let mut create = workbook.edit()?;
    create
        .sheet("Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?
        .set("A1", "Merged title")?
        .merge("A1:C2")?
        .set("E1", "Temporary merge")?
        .merge("E1:F2")?;
    let created = create.commit()?;

    let mut revise = created.workbook().edit()?;
    revise
        .sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?
        .unmerge("F2")?
        .set("F2", "Unmerged follower")?
        .set("A4", "Merged footer")?
        .merge("A4:C4")?;
    let revised = revise.commit()?;

    let sheet = revised
        .workbook()
        .sheet("Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?;
    let ranges = sheet.merges()?.map(Rect::a1).collect::<Vec<_>>();
    if ranges != ["A1:C2", "A4:C4"] {
        return Err(io::Error::other("committed merged ranges differ").into());
    }
    if !matches!(sheet.cell("B2")?, cell::View::Covered(_))
        || !matches!(sheet.cell("F2")?, cell::View::Stored(_))
    {
        return Err(io::Error::other("covered/unmerged cell views differ").into());
    }

    revised.workbook().save(output)?;
    println!(
        "saved two merged ranges after one unmerge in {} semantic changes",
        created.patch().len().saturating_add(revised.patch().len())
    );
    Ok(())
}
