use litchi::ooxml::xlsb::XlsbWorkbook;
use litchi::sheet::WorkbookTrait;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open("target/example.xlsb")?;
    let wb = XlsbWorkbook::new(file)?;
    println!("Sheets: {}", wb.worksheet_count());
    for name in wb.worksheet_names() {
        println!("- {}", name);
    }
    Ok(())
}
