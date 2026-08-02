use litchi_cfb::writer::OleWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = OleWriter::new();

    // Create WordDocument stream (512 bytes - should get 1 sector)
    let word_doc = vec![0xEC, 0xA5]; // FIB magic + padding
    let word_doc_padded = {
        let mut v = word_doc.clone();
        v.resize(4096, 0);
        v
    };
    writer.create_stream(&["WordDocument"], &word_doc_padded)?;

    // Create 1Table stream (512 bytes - should get 1 sector)
    let table = vec![0x01, 0x02];
    let table_padded = {
        let mut v = table.clone();
        v.resize(4096, 0);
        v
    };
    writer.create_stream(&["1Table"], &table_padded)?;

    // Save
    writer.save("debug_fat_test.doc")?;

    println!("Created debug_fat_test.doc");

    // Now read the FAT from the file
    use std::fs::File;
    use std::io::Read;

    let mut file = File::open("debug_fat_test.doc")?;
    let mut data = Vec::new();
    file.read_to_end(&mut data)?;

    println!("\nFAT entries from header (first 20):");
    for i in 0..20 {
        let offset = 76 + i * 4;
        let entry = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]);

        let status = match entry {
            0xFFFFFFFF => "END".to_string(),
            0xFFFFFFFE => "FREE".to_string(),
            0xFFFFFFFD => "FATSECT".to_string(),
            0xFFFFFFFC => "DIFSECT".to_string(),
            _ => format!("-> {}", entry),
        };

        println!("  Sector {:2}: {}", i, status);
    }

    Ok(())
}
