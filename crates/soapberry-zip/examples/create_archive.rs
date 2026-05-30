//! Create a tiny ODF-like ZIP package, write it to a tempfile, then re-open it.
//!
//! Demonstrates [`soapberry_zip::office::StreamingArchiveWriter`] for writing
//! and [`soapberry_zip::office::ArchiveReader`] for round-trip verification.
//!
//! # Run
//!
//! ```sh
//! cargo run -p soapberry-zip --example create_archive
//! ```

use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::fs;

const MIMETYPE_VALUE: &[u8] = b"application/vnd.oasis.opendocument.text";
const CONTENT_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
    <office:body>
        <office:text>
            <text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">Hello, soapberry-zip!</text:p>
        </office:text>
    </office:body>
</office:document-content>
"#;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Build the archive in memory.
    let mut writer = StreamingArchiveWriter::new();

    // ODF requires "mimetype" be the first entry, stored uncompressed.
    writer.write_stored("mimetype", MIMETYPE_VALUE)?;
    // Everything else can be deflated.
    writer.write_deflated("content.xml", CONTENT_XML)?;

    let bytes = writer.finish_to_bytes()?;
    println!("Built archive ({} bytes in memory).", bytes.len());

    // Persist to a tempfile so we exercise the real on-disk round-trip.
    let tempfile = tempfile::Builder::new()
        .prefix("soapberry-example-")
        .suffix(".odt")
        .tempfile()?;
    let temp_path = tempfile.path().to_path_buf();
    fs::write(&temp_path, &bytes)?;
    println!("Wrote archive to: {}", temp_path.display());

    // Re-open from disk and verify contents.
    let on_disk = fs::read(&temp_path)?;
    let archive = ArchiveReader::new(&on_disk)?;
    println!("Re-opened archive: {} files", archive.len());

    let mut names: Vec<&str> = archive.file_names().collect();
    names.sort_unstable();
    for name in &names {
        println!("  - {}", name);
    }

    // Verify each entry round-trips byte-for-byte.
    let mimetype = archive.read("mimetype")?;
    assert_eq!(mimetype, MIMETYPE_VALUE, "mimetype round-trip mismatch");
    println!("mimetype OK ({} bytes, stored)", mimetype.len());

    let content = archive.read("content.xml")?;
    assert_eq!(content, CONTENT_XML, "content.xml round-trip mismatch");
    println!("content.xml OK ({} bytes, deflated)", content.len());

    println!("All entries verified.");
    // tempfile is deleted when it goes out of scope.
    Ok(())
}
