use litchi_cfb::OleFile;
use std::fs::File;
use std::io::BufReader;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open("comprehensive_test.doc")?;
    let reader = BufReader::new(file);
    let mut ole = OleFile::open(reader)?;

    let table_data = ole.open_stream(&["1Table"])?;

    // Section table is at offset 601 (from FIB)
    let section_table_offset = 601;

    println!("=== Section Table (at offset {}) ===", section_table_offset);
    let section_data = &table_data[section_table_offset..section_table_offset + 20];

    // Parse PLCF structure:
    // 2 CPs (8 bytes)
    // 1 SED (12 bytes)

    let cp0 = u32::from_le_bytes([
        section_data[0],
        section_data[1],
        section_data[2],
        section_data[3],
    ]);
    let cp1 = u32::from_le_bytes([
        section_data[4],
        section_data[5],
        section_data[6],
        section_data[7],
    ]);

    println!("CP[0] = {} (start)", cp0);
    println!("CP[1] = {} (end)", cp1);

    // SED structure (12 bytes)
    let fn_val = u16::from_le_bytes([section_data[8], section_data[9]]);
    let fc_sepx = u32::from_le_bytes([
        section_data[10],
        section_data[11],
        section_data[12],
        section_data[13],
    ]);
    let fn_mpr = u16::from_le_bytes([section_data[14], section_data[15]]);
    let fc_mpr = u32::from_le_bytes([
        section_data[16],
        section_data[17],
        section_data[18],
        section_data[19],
    ]);

    println!("\nSED (Section Descriptor):");
    println!("  fn:     {} (0x{:04X})", fn_val, fn_val);
    println!(
        "  fcSepx: {} (0x{:04X}) <- CRITICAL: must point to SEPX in WordDocument stream",
        fc_sepx, fc_sepx
    );
    println!("  fnMpr:  {} (0x{:04X})", fn_mpr, fn_mpr);
    println!("  fcMpr:  {} (0x{:04X})", fc_mpr, fc_mpr);

    // Now check WordDocument stream at that offset
    let wd_data = ole.open_stream(&["WordDocument"])?;

    println!("\n=== WordDocument stream at SEPX offset {} ===", fc_sepx);
    if (fc_sepx as usize) < wd_data.len() {
        let sepx_data =
            &wd_data[fc_sepx as usize..std::cmp::min((fc_sepx + 10) as usize, wd_data.len())];
        print!("SEPX data: ");
        for byte in sepx_data {
            print!("{:02X} ", byte);
        }
        println!();

        // SEPX structure: 2-byte size + grpprl
        let sepx_size = u16::from_le_bytes([sepx_data[0], sepx_data[1]]);
        println!("SEPX size: {} bytes", sepx_size);
    } else {
        println!(
            "ERROR: fcSepx {} is beyond WordDocument stream size {}",
            fc_sepx,
            wd_data.len()
        );
    }

    Ok(())
}
