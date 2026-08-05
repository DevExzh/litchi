use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Hyperlinks")?;

    // Header row
    writer.write_string(sheet, 0, 0, "Description")?;
    writer.write_string(sheet, 0, 1, "Link")?;

    // Web hyperlink
    writer.write_string(sheet, 1, 0, "Rust website")?;
    writer.write_string(sheet, 1, 1, "https://www.rust-lang.org")?;
    writer.set_hyperlink(sheet, 1, 1, "https://www.rust-lang.org")?;

    // Mail hyperlink
    writer.write_string(sheet, 2, 0, "Send email")?;
    writer.write_string(sheet, 2, 1, "mailto:example@example.com")?;
    writer.set_hyperlink(sheet, 2, 1, "mailto:example@example.com")?;

    // Internal hyperlink back to A1 on the same sheet
    writer.write_string(sheet, 3, 0, "Jump to A1")?;
    writer.write_string(sheet, 3, 1, "Go to A1")?;
    writer.set_hyperlink(sheet, 3, 1, "internal:Hyperlinks!A1")?;

    writer.save("xls_hyperlinks.xls")?;
    Ok(())
}
