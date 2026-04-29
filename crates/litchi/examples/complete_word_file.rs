//! Create a complete Word file with all required OLE streams

use std::fs::File;
use std::io::Read;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load all extracted streams
    let mut word_doc = Vec::new();
    let mut table = Vec::new();
    let mut summary = Vec::new();
    let mut doc_summary = Vec::new();
    let mut compobj = Vec::new();

    // Create WordDocument
    let mut fib = vec![0u8; 512];
    fib[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib[2..4].copy_from_slice(&193u16.to_le_bytes());
    fib[4..6].copy_from_slice(&1813u16.to_le_bytes());
    fib[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
    fib[8..10].copy_from_slice(&0u16.to_le_bytes());
    fib[10..12].copy_from_slice(&0x52F0u16.to_le_bytes());
    fib[12..14].copy_from_slice(&0x00BFu16.to_le_bytes());
    fib[14..18].copy_from_slice(&0u32.to_le_bytes());
    fib[18] = 1;
    fib[19] = 16;
    fib[20..22].copy_from_slice(&0u16.to_le_bytes());
    fib[22..24].copy_from_slice(&0u16.to_le_bytes());
    fib[24..28].copy_from_slice(&2048u32.to_le_bytes());
    fib[28..32].copy_from_slice(&2060u32.to_le_bytes());

    fib[32..34].copy_from_slice(&14u16.to_le_bytes());
    let rgw = [
        0x62, 0x6a, 0x62, 0x6a, 0x62, 0xb5, 0x62, 0xb5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 4, 8, 22, 0,
    ];
    fib[34..62].copy_from_slice(&rgw);

    fib[62..64].copy_from_slice(&22u16.to_le_bytes());
    fib[64..68].copy_from_slice(&3634u32.to_le_bytes());
    fib[68..72].copy_from_slice(&0x5ff7df00u32.to_le_bytes());
    fib[72..76].copy_from_slice(&0x5ff7df00u32.to_le_bytes());
    fib[76..80].copy_from_slice(&12u32.to_le_bytes());

    fib[152..154].copy_from_slice(&183u16.to_le_bytes());
    let fcllcb = 154;
    fib[fcllcb..fcllcb + 4].copy_from_slice(&0u32.to_le_bytes());
    fib[fcllcb + 4..fcllcb + 8].copy_from_slice(&3788u32.to_le_bytes());
    fib[fcllcb + 8..fcllcb + 12].copy_from_slice(&0u32.to_le_bytes());
    fib[fcllcb + 12..fcllcb + 16].copy_from_slice(&3788u32.to_le_bytes());

    word_doc.extend_from_slice(&fib);
    word_doc.resize(2048, 0);
    word_doc.extend_from_slice(b"Hello World\r");
    word_doc.resize(4096, 0);

    // Load 1Table
    File::open("/tmp/reference_stylesheet.bin")?.read_to_end(&mut table)?;
    table.resize(4096, 0);

    // Load property streams
    File::open("/tmp/ref_SummaryInformation.bin")?.read_to_end(&mut summary)?;
    File::open("/tmp/ref_DocumentSummaryInformation.bin")?.read_to_end(&mut doc_summary)?;
    File::open("/tmp/ref_CompObj.bin")?.read_to_end(&mut compobj)?;

    println!("Loaded all streams:");
    println!("  WordDocument: {} bytes", word_doc.len());
    println!("  1Table: {} bytes", table.len());
    println!("  SummaryInformation: {} bytes", summary.len());
    println!("  DocumentSummaryInformation: {} bytes", doc_summary.len());
    println!("  CompObj: {} bytes", compobj.len());

    // Create OLE file with all streams
    // CRITICAL: WordDocument MUST be added first to get sector 0!
    let mut ole = litchi::ole::writer::OleWriter::new();
    ole.create_stream(&["WordDocument"], &word_doc)?;
    ole.create_stream(&["1Table"], &table)?;
    ole.create_stream(&["SummaryInformation"], &summary)?;
    ole.create_stream(&["DocumentSummaryInformation"], &doc_summary)?;
    ole.create_stream(&["CompObj"], &compobj)?;

    let mut file = File::create("complete_word_file.doc")?;
    ole.write_to(&mut file)?;

    println!("\n✅ Created complete_word_file.doc with all 5 streams!");
    Ok(())
}
