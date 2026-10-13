use std::error::Error;
use std::io;

use litchi_xlsx::LocalStyle;

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let [input, output] = arguments.as_slice() else {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example grid_styles -- INPUT.xlsx OUTPUT.xlsx",
        )
        .into());
    };

    let workbook = litchi_xlsx::Package::open(input)?;
    let sheet = workbook
        .workbook_ref()
        .sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?;
    let style = sheet.style("A1")?.ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::NotFound,
            "A1 has no stored cell or resolvable shared style",
        )
    })?;

    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet(0usize)?
            .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?;
        sheet.row(3)?.style(&style)?;
        sheet.column("D")?.width(12)?.style(&style)?;
    }

    let committed = edit.commit()?;
    let sheet = committed
        .package()
        .workbook_ref()
        .sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?;
    let row_style = match sheet.row_style(3)? {
        Some(LocalStyle::Shared(row_style)) => row_style,
        _ => return Err(io::Error::other("row style was not committed").into()),
    };
    let column_style = match sheet.column_style("D")? {
        Some(LocalStyle::Shared(column_style)) => column_style,
        _ => return Err(io::Error::other("column style was not committed").into()),
    };
    if !row_style.same(&style) || !column_style.same(&style) {
        return Err(io::Error::other("committed grid style changed identity").into());
    }

    committed.package().save(output)?;
    println!(
        "applied A1's shared style to row 4 and width-12 column D in {} semantic changes",
        committed.patch().operation_count()
    );
    Ok(())
}
