//! List all entries in a ZIP archive (e.g. an OOXML document).
//!
//! Demonstrates the high-level [`soapberry_zip::office::ArchiveReader`] API.
//! It opens a ZIP file, prints every entry's name and, when available, its
//! compressed and uncompressed sizes (recovered via the lower-level slice
//! archive iterator since the high-level reader hides them).
//!
//! # Run
//!
//! ```sh
//! cargo run -p soapberry-zip --example list_archive
//! cargo run -p soapberry-zip --example list_archive -- path/to/file.docx
//! ```

use soapberry_zip::ZipArchive;
use soapberry_zip::office::ArchiveReader;
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
    println!("File size: {} bytes", data.len());

    // High-level API: indexed by name, files-only (directories filtered).
    let archive = ArchiveReader::new(&data)?;
    println!("Entries (high-level, files only): {}", archive.len());

    // Lower-level slice archive lets us print compressed and uncompressed
    // sizes per entry, including directory entries.
    let slice_archive = ZipArchive::from_slice(&data)?;
    println!();
    println!("{:<60}  {:>12}  {:>12}", "name", "comp size", "uncomp size");
    println!("{}", "-".repeat(90));
    for entry_result in slice_archive.entries() {
        let entry = entry_result?;
        let path = entry.file_path();
        let display_name = match path.try_normalize() {
            Ok(p) => p.as_ref().to_string(),
            Err(_) => String::from_utf8_lossy(path.as_ref()).to_string(),
        };
        let kind = if entry.is_dir() { " (dir)" } else { "" };
        println!(
            "{:<60}  {:>12}  {:>12}{}",
            display_name,
            entry.compressed_size_hint(),
            entry.uncompressed_size_hint(),
            kind,
        );
    }

    Ok(())
}
