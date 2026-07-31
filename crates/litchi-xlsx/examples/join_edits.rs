use std::error::Error;
use std::io;

use litchi_xlsx::{Cell, Value, Workbook};

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args_os().skip(1).collect::<Vec<_>>();
    let output = match arguments.as_slice() {
        [] => None,
        [output] => Some(output),
        _ => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "usage: cargo run -p litchi-xlsx --example join_edits -- [OUTPUT.xlsx]",
            )
            .into());
        },
    };
    let workbook = Workbook::new()?;

    let mut labels = workbook.edit()?;
    labels
        .sheet("Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?
        .set("A1", "Revenue")?;

    let mut values = workbook.edit()?;
    values
        .sheet(0usize)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "first sheet is missing"))?
        .set("B1", 42_i32)?;

    labels.join(values)?;
    let committed = labels.commit()?;
    let sheet = committed
        .workbook()
        .sheet("Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?;
    if !matches!(sheet.cell("A1")?, Some(Cell::Value(Value::Text(_)))) {
        return Err(io::Error::new(io::ErrorKind::InvalidData, "A1 was not joined").into());
    }
    if !matches!(
        sheet.cell("B1")?,
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ) {
        return Err(io::Error::new(io::ErrorKind::InvalidData, "B1 was not joined").into());
    }
    if let Some(output) = output {
        committed.workbook().save(output)?;
    }
    Ok(())
}
