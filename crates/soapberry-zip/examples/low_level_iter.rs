//! Iterate ZIP entries using the lower-level [`soapberry_zip::ZipArchive`] API.
//!
//! Demonstrates [`ZipArchive::from_slice`] and the central-directory iterator,
//! printing each entry's path, compression method, CRC32, and sizes.
//!
//! # Run
//!
//! ```sh
//! cargo run -p soapberry-zip --example low_level_iter
//! cargo run -p soapberry-zip --example low_level_iter -- path/to/file.docx
//! ```

use soapberry_zip::ZipArchive;
use std::env;
use std::fs;
use std::path::PathBuf;

fn default_path() -> PathBuf {
    PathBuf::from("test-data/ooxml/docx/documentProperties.docx")
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(default_path);

    println!("Reading archive: {}", path.display());
    let data = fs::read(&path)?;

    // ZipArchive::from_slice returns a ZipSliceArchive over the borrowed bytes.
    let archive = ZipArchive::from_slice(&data)?;
    println!("Entries hint (from EOCD): {}", archive.entries_hint());
    println!("Directory offset: {}", archive.directory_offset());
    println!("End offset:       {}", archive.end_offset());
    println!();

    println!(
        "{:<60}  {:<14}  {:>10}  {:>10}  {:<10}",
        "name", "method", "comp", "uncomp", "crc32"
    );
    println!("{}", "-".repeat(110));

    let mut total = 0usize;
    let mut total_compressed: u64 = 0;
    let mut total_uncompressed: u64 = 0;

    for entry_result in archive.entries() {
        let entry = entry_result?;
        let raw_path = entry.file_path();
        let display_name = match raw_path.try_normalize() {
            Ok(p) => p.as_ref().to_string(),
            Err(_) => String::from_utf8_lossy(raw_path.as_ref()).to_string(),
        };
        let method = entry.compression_method();
        let comp = entry.compressed_size_hint();
        let uncomp = entry.uncompressed_size_hint();

        println!(
            "{:<60}  {:<14}  {:>10}  {:>10}  {:08x}",
            display_name,
            format!("{:?}", method),
            comp,
            uncomp,
            entry.crc32(),
        );

        total += 1;
        total_compressed += comp;
        total_uncompressed += uncomp;
    }

    println!();
    println!(
        "Total: {} entries, {} compressed bytes, {} uncompressed bytes",
        total, total_compressed, total_uncompressed
    );

    Ok(())
}
