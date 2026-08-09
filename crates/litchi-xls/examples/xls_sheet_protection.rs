//! Sheet Protection — XLS Writer + Round-Trip Read Example
//!
//! Demonstrates writing sheet protection (PROTECT, OBJECTPROTECT, SCENPROTECT,
//! PASSWORD records) and then reading them back to verify the round-trip.
//!
//! Run with: `cargo run --example xls_sheet_protection`
//!
//! The file is saved to `output/xls_sheet_protection.xls`.

use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output_dir = std::path::Path::new("output");
    std::fs::create_dir_all(output_dir)?;
    let output_path = output_dir.join("xls_sheet_protection.xls");

    let mut w = Writer::new();

    // ================================================================
    // Sheet 1 — Protected with password + object/scenario protection
    // ================================================================
    let s1 = w.add_worksheet("Protected")?;

    w.write_string(s1, 0, 0, "This sheet is protected.")?;
    w.write_string(s1, 1, 0, "Password: secret")?;
    w.write_string(s1, 2, 0, "Objects and scenarios are also protected.")?;

    w.write_string(s1, 4, 0, "Item")?;
    w.write_string(s1, 4, 1, "Value")?;
    for i in 0..5 {
        let row = 5 + i as u32;
        w.write_string(s1, row, 0, &format!("Row {}", i + 1))?;
        w.write_number(s1, row, 1, f64::from(i + 1) * 100.0)?;
    }

    // Protect: password "secret", protect objects, protect scenarios.
    w.protect_sheet(s1, Some("secret"), true, true)?;

    w.set_column_width(s1, 0, 40.0)?;
    w.set_column_width(s1, 1, 12.0)?;

    println!("[Sheet 1] Protected — password 'secret', objects+scenarios protected");

    // ================================================================
    // Sheet 2 — Protected without password
    // ================================================================
    let s2 = w.add_worksheet("No Password")?;

    w.write_string(s2, 0, 0, "This sheet is protected without a password.")?;
    w.write_string(s2, 1, 0, "Users can unprotect it freely.")?;

    w.write_string(s2, 3, 0, "Name")?;
    w.write_string(s2, 3, 1, "Score")?;
    let names = ["Alice", "Bob", "Carol", "Dave"];
    for (i, name) in names.iter().enumerate() {
        let row = 4 + i as u32;
        w.write_string(s2, row, 0, name)?;
        w.write_number(s2, row, 1, 80.0 + i as f64 * 5.0)?;
    }

    // Protect without password, no object/scenario protection.
    w.protect_sheet(s2, None, false, false)?;

    w.set_column_width(s2, 0, 40.0)?;
    w.set_column_width(s2, 1, 12.0)?;

    println!("[Sheet 2] No Password — protected but no password set");

    // ================================================================
    // Sheet 3 — Unprotected (for comparison)
    // ================================================================
    let s3 = w.add_worksheet("Unprotected")?;

    w.write_string(s3, 0, 0, "This sheet has no protection.")?;
    w.write_string(s3, 1, 0, "All cells are editable.")?;
    w.write_number(s3, 3, 0, 42.0)?;

    w.set_column_width(s3, 0, 40.0)?;

    println!("[Sheet 3] Unprotected — no protection records written");

    // ================================================================
    // Save
    // ================================================================
    w.save(&output_path)?;
    println!("\nSaved to: {}", output_path.display());

    // ================================================================
    // Round-trip: read back and verify protection state
    // ================================================================
    println!("\n=== Round-trip verification ===\n");
    round_trip_verify(&output_path)?;

    Ok(())
}

/// Read the generated XLS back and print the parsed protection state
/// for each sheet.
fn round_trip_verify(path: &std::path::Path) -> Result<(), Box<dyn std::error::Error>> {
    use litchi_core::sheet::WorkbookTrait;
    use litchi_xls::Workbook;
    use std::io::Cursor;

    let data = std::fs::read(path)?;
    let wb = Workbook::new(Cursor::new(data))?;

    println!("Worksheets: {:?}", wb.worksheet_names());

    for (idx, name) in wb.worksheet_names().iter().enumerate() {
        println!("\n--- Sheet {idx}: \"{name}\" ---");

        let xls_sheet = wb.xls_worksheet(idx)?;

        let prot = xls_sheet.protection();
        println!("  sheet_protected:    {}", prot.is_protected());
        println!("  objects_protected:  {}", prot.objects_protected());
        println!("  scenarios_protected:{}", prot.scenarios_protected());
        println!(
            "  password_hash:      0x{:04X}",
            prot.password()
                .map_or(0, litchi_xls::protection::PasswordVerifier::raw)
        );
        println!("  has_password:       {}", prot.has_password());
    }

    println!("\nRound-trip complete.");
    Ok(())
}
