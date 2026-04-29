use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut wb = XlsbWorkbookWriter::new();

    // Sheet1: per spec example
    let mut sheet1 = MutableXlsbWorksheet::new("Sheet1");
    // B4: "Number" => row=3, col=1
    sheet1.set_cell(3, 1, "Number");
    // B5: 1 => row=4, col=1
    sheet1.set_cell(4, 1, 1.0);
    // B6: "Formula" => row=5, col=1
    sheet1.set_cell(5, 1, "Formula");
    // B7: SQRT(B5*2) -> write cached numeric result as constant for now
    sheet1.set_cell(6, 1, 1.4142135623730951_f64);
    wb.add_worksheet(sheet1);

    // Sheet2 and Sheet3 empty
    wb.add_worksheet(MutableXlsbWorksheet::new("Sheet2"));
    wb.add_worksheet(MutableXlsbWorksheet::new("Sheet3"));

    let file = File::create("target/example.xlsb")?;
    wb.save(file)?;
    println!("Wrote target/example.xlsb");
    Ok(())
}
