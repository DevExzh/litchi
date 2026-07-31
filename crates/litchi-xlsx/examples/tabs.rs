use std::env;
use std::error::Error;
use std::io;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = env::args().skip(1);
    let input = args
        .next()
        .ok_or("usage: tabs <input.xlsx> <output.xlsx> <sheet-name> <show|hide|very-hide>")?;
    let output = args
        .next()
        .ok_or("usage: tabs <input.xlsx> <output.xlsx> <sheet-name> <show|hide|very-hide>")?;
    let name = args
        .next()
        .ok_or("usage: tabs <input.xlsx> <output.xlsx> <sheet-name> <show|hide|very-hide>")?;
    let operation = args
        .next()
        .ok_or("usage: tabs <input.xlsx> <output.xlsx> <sheet-name> <show|hide|very-hide>")?;

    let source = Workbook::open(&input)?;
    let mut edit = source.edit()?;
    let mut tab = edit.tab(name.as_str())?.ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::NotFound,
            format!("workbook has no sheet named '{name}'"),
        )
    })?;
    match operation.as_str() {
        "show" => {
            tab.show();
        },
        "hide" => {
            tab.hide();
        },
        "very-hide" => {
            tab.very_hide();
        },
        other => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("unknown tab operation '{other}'"),
            )
            .into());
        },
    }

    let committed = edit.commit()?;
    let tab = committed
        .workbook()
        .sheet(name.as_str())?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "edited tab disappeared"))?;
    committed.workbook().save(&output)?;
    println!(
        "saved {} semantic change(s); {name} is {:?}",
        committed.patch().len(),
        tab.visibility()
    );
    Ok(())
}
