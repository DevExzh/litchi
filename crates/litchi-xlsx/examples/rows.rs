use std::env;
use std::error::Error;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let output = env::args().nth(1).ok_or("usage: rows <output.xlsx>")?;
    let source = Workbook::new()?;
    let mut edit = source.edit()?;
    let mut sheet = edit.sheet("Sheet1")?.ok_or("missing Sheet1")?;
    sheet
        .set("A1", "Visible above")?
        .set("A2", "Hidden by Litchi")?
        .set("A3", "Visible below")?;
    sheet.row(1)?.hide();

    let committed = edit.commit()?;
    let sheet = committed
        .workbook()
        .sheet("Sheet1")?
        .ok_or("missing committed Sheet1")?;
    let row = sheet.row(1)?;
    assert!(row.stored() && row.hidden());
    committed.workbook().save(&output)?;
    println!(
        "saved {} semantic changes to {output}",
        committed.patch().len()
    );
    Ok(())
}
