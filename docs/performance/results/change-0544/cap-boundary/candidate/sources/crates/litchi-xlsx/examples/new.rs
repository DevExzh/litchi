use std::error::Error;
use std::io;

use litchi_xlsx::Workbook;

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os()
        .nth(1)
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "usage: new <workbook.xlsx>"))?;
    let workbook = Workbook::new()?;
    std::fs::write(path, workbook.to_bytes()?)?;
    Ok(())
}
