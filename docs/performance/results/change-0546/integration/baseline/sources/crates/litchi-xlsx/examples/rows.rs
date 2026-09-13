use std::env;
use std::error::Error;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let output = env::args().nth(1).ok_or("usage: rows <output.xlsx>")?;
    let source = Workbook::new()?;
    let mut edit = source.edit()?;
    let mut sheet = edit.sheet("Sheet1")?.ok_or("missing Sheet1")?;
    sheet
        .set("A1", "Default height")?
        .set("A2", "Tall by Litchi")?
        .set("A3", "Hidden by Litchi")?
        .set("A4", "Outlined by Litchi")?;
    sheet.row(1)?.height(30)?;
    sheet.row(2)?.hide();
    sheet.row(3)?.outline(1)?.collapse();

    let committed = edit.commit()?;
    let sheet = committed
        .workbook()
        .sheet("Sheet1")?
        .ok_or("missing committed Sheet1")?;
    let tall = sheet.row(1)?;
    assert_eq!(tall.height().map(litchi_xlsx::Height::get), Some(30.0));
    assert!(sheet.row(2)?.hidden());
    let outlined = sheet.row(3)?;
    assert_eq!(outlined.outline().get(), 1);
    assert!(outlined.collapsed());
    committed.workbook().save(&output)?;
    println!(
        "saved {} semantic changes to {output}",
        committed.patch().len()
    );
    Ok(())
}
