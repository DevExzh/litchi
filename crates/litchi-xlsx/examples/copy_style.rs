use std::error::Error;
use std::io;

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let [input, output] = arguments.as_slice() else {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example copy_style -- INPUT.xlsx OUTPUT.xlsx",
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
    edit.sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?
        .set("C1", 42_i32)?
        .style("C1", &style)?;
    let committed = edit.commit()?;
    committed.package().save(output)?;
    println!(
        "copied A1's shared style to C1; source fan-out was {} stored cell(s)",
        style.fan_out()?
    );
    Ok(())
}
