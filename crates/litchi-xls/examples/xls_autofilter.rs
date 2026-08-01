use litchi_xls::XlsWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("AutoFilter")?;

    writer.write_string(sheet, 0, 0, "Item")?;
    writer.write_string(sheet, 0, 1, "Category")?;
    writer.write_string(sheet, 0, 2, "Value")?;

    writer.write_string(sheet, 1, 0, "A")?;
    writer.write_string(sheet, 1, 1, "Fruit")?;
    writer.write_number(sheet, 1, 2, 10.0)?;

    writer.write_string(sheet, 2, 0, "B")?;
    writer.write_string(sheet, 2, 1, "Fruit")?;
    writer.write_number(sheet, 2, 2, 20.0)?;

    writer.write_string(sheet, 3, 0, "C")?;
    writer.write_string(sheet, 3, 1, "Vegetable")?;
    writer.write_number(sheet, 3, 2, 30.0)?;

    writer.write_string(sheet, 4, 0, "D")?;
    writer.write_string(sheet, 4, 1, "Vegetable")?;
    writer.write_number(sheet, 4, 2, 40.0)?;

    writer.set_auto_filter(sheet, 0, 4, 0, 2)?;

    writer.save("xls_autofilter.xls")?;
    Ok(())
}
