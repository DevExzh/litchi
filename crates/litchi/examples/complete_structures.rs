//! Use COMPLETE reference WordDocument, 1Table, and properties
//! Only modify the text bytes at offset 2048-2078

use std::fs::File;
use std::io::Read;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Use COMPLETE reference WordDocument (includes FIB + all formatting structures)
    let mut word_doc = Vec::new();
    File::open("/tmp/ref_WordDocument.bin")?.read_to_end(&mut word_doc)?;

    println!(
        "Loaded complete reference WordDocument ({} bytes)",
        word_doc.len()
    );
    println!("This includes FIB + all formatting structures at correct offsets");

    // Only modify the text at offset 2048
    let text = b"Hello World\r";
    word_doc[2048..2048 + text.len()].copy_from_slice(text);

    println!(
        "Modified text at offset 2048 to: {:?}",
        std::str::from_utf8(text)
    );

    // Use COMPLETE reference 1Table
    let mut table = Vec::new();
    File::open("/tmp/reference_1table.bin")?.read_to_end(&mut table)?;
    println!("Using complete reference 1Table ({} bytes)", table.len());

    // Load property streams
    let mut summary = Vec::new();
    File::open("/tmp/ref_SummaryInformation.bin")?.read_to_end(&mut summary)?;

    let mut doc_summary = Vec::new();
    File::open("/tmp/ref_DocumentSummaryInformation.bin")?.read_to_end(&mut doc_summary)?;

    let mut compobj = Vec::new();
    File::open("/tmp/ref_CompObj.bin")?.read_to_end(&mut compobj)?;

    println!("\nCreating file with COMPLETE reference structures...");

    let mut ole = litchi_cfb::writer::OleWriter::new();
    // CRITICAL: Allocation order determines sector numbers!
    // WordDocument MUST be first to get sector 0
    // Directory entry order is determined by BST, not creation order
    ole.create_stream(&["WordDocument"], &word_doc)?;
    ole.create_stream(&["1Table"], &table)?;
    ole.create_stream(&["SummaryInformation"], &summary)?;
    ole.create_stream(&["DocumentSummaryInformation"], &doc_summary)?;
    ole.create_stream(&["CompObj"], &compobj)?;

    let mut file = File::create("complete_structures.doc")?;
    ole.write_to(&mut file)?;

    println!("✅ Created complete_structures.doc");
    println!("\nThis file is IDENTICAL to reference except:");
    println!("  ✅ Complete reference WordDocument (with all formatting structures)");
    println!("  ✅ Complete reference 1Table (with all offsets intact)");
    println!("  ✅ Complete reference property streams");
    println!("  📝 Only text at offset 2048 modified to 'Hello World'");
    println!("\nThis should open without ANY conversion dialog!");
    Ok(())
}
