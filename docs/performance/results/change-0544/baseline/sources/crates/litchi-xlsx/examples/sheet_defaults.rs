use std::error::Error;
use std::io;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let [input, output] = arguments.as_slice() else {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example sheet_defaults -- INPUT.xlsx OUTPUT.xlsx",
        )
        .into());
    };

    let workbook = Workbook::open(input)?;
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet(0usize)?
            .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?;
        {
            let mut defaults = sheet.defaults();
            // Supplying the height also makes this valid for producer files
            // that omit `sheetFormatPr` entirely.
            defaults.height(24)?.width(14)?.descent(0.2)?;
        }
        sheet.row(1)?.height(32)?.descent(0.3)?;
    }

    let committed = edit.commit()?;
    let sheet = committed
        .workbook()
        .sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?;
    let defaults = sheet
        .defaults()?
        .ok_or_else(|| io::Error::other("worksheet defaults were not committed"))?;
    if defaults.height().get() != 24.0
        || defaults.width().map(litchi_xlsx::layout::Width::get) != Some(14.0)
        || defaults.descent().map(litchi_xlsx::layout::Descent::get) != Some(0.2)
    {
        return Err(io::Error::other("committed worksheet defaults differ").into());
    }
    let row = sheet.row(1)?;
    if row.height().map(litchi_xlsx::Height::get) != Some(32.0)
        || row.descent().map(litchi_xlsx::layout::Descent::get) != Some(0.3)
    {
        return Err(io::Error::other("committed row layout differs").into());
    }

    committed.workbook().save(output)?;
    println!(
        "saved worksheet defaults and row 2 layout in {} semantic changes",
        committed.patch().len()
    );
    Ok(())
}
