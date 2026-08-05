use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Protected")?;

    // Some sample content
    writer.write_string(sheet, 0, 0, "This sheet is protected")?;
    writer.write_string(sheet, 1, 0, "Try editing cells in Excel")?;

    // Protect workbook structure and windows with a simple password.
    writer.protect_workbook(Some("secret"), true, true);

    // Protect the sheet, including objects and scenarios, with its own password.
    writer.protect_sheet(sheet, Some("sheetpw"), true, true)?;

    writer.save("xls_protection.xls")?;
    Ok(())
}
