use std::env;
use std::error::Error;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let output = env::args().nth(1).ok_or("usage: columns <output.xlsx>")?;
    let source = Workbook::new()?;
    let mut edit = source.edit()?;
    let mut sheet = edit.sheet("Sheet1")?.ok_or("missing Sheet1")?;
    sheet
        .set("A1", "Default width")?
        .set("B1", "Wide by Litchi")?
        .set("C1", "Hidden by Litchi")?
        .set("D1", "Outlined by Litchi")?;
    sheet.column("B")?.width(24)?;
    sheet.column("C")?.hide();
    sheet.column("D")?.outline(1)?.collapse();

    let committed = edit.commit()?;
    let sheet = committed
        .workbook()
        .sheet("Sheet1")?
        .ok_or("missing committed Sheet1")?;
    let wide = sheet.column("B")?;
    assert_eq!(wide.width().map(litchi_xlsx::Width::get), Some(24.0));
    assert!(sheet.column("C")?.hidden());
    let outlined = sheet.column("D")?;
    assert_eq!(outlined.outline().get(), 1);
    assert!(outlined.collapsed());
    committed.workbook().save(&output)?;
    println!(
        "saved {} semantic changes to {output}",
        committed.patch().len()
    );
    Ok(())
}
