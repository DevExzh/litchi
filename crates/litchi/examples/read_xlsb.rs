use litchi::sheet::WorkbookTrait;
use litchi::xlsb::Workbook;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open("target/example.xlsb")?;
    let wb = Workbook::new(file)?;
    println!("Sheets: {}", wb.worksheet_count());
    for name in wb.worksheet_names() {
        println!("- {}", name);
    }
    Ok(())
}
