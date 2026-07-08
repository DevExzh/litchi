//! Extract a single entry from a ZIP archive and print its content.
//!
//! Demonstrates [`soapberry_zip::office::ArchiveReader::read`] which transparently
//! handles both stored and Deflate-compressed entries.
//!
//! # Run
//!
//! ```sh
//! cargo run -p soapberry-zip --example extract_entry
//! cargo run -p soapberry-zip --example extract_entry -- path/to/file.docx word/document.xml
//! ```

use soapberry_zip::office::ArchiveReader;
use std::env;
use std::fs;
use std::path::PathBuf;

const MAX_PRINT_BYTES: usize = 500;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let path = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("test-data/ooxml/docx/documentProperties.docx"));
    let entry_name = args
        .next()
        .unwrap_or_else(|| "word/document.xml".to_string());

    println!("Archive: {}", path.display());
    println!("Entry:   {}", entry_name);
    println!();

    let data = fs::read(&path)?;
    let archive = ArchiveReader::new(&data)?;

    if !archive.contains(&entry_name) {
        eprintln!("Entry not found. Available entries:");
        let mut names: Vec<&str> = archive.file_names().collect();
        names.sort_unstable();
        for name in names {
            eprintln!("  {}", name);
        }
        return Err(format!("entry {entry_name:?} not found in archive").into());
    }

    let bytes = archive.read(&entry_name)?;
    println!("Decompressed size: {} bytes", bytes.len());
    println!("--- first {} bytes (UTF-8 lossy) ---", MAX_PRINT_BYTES);

    let preview_len = bytes.len().min(MAX_PRINT_BYTES);
    let preview = String::from_utf8_lossy(&bytes[..preview_len]);
    print!("{}", preview);
    if bytes.len() > MAX_PRINT_BYTES {
        println!();
        println!(
            "... ({} more bytes truncated)",
            bytes.len() - MAX_PRINT_BYTES
        );
    } else {
        println!();
    }

    Ok(())
}
