//! Extract a single part from an OPC package by partname.
//!
//! Demonstrates parsing a `PackURI` and using `OpcPackage::get_part` to locate
//! the part inside the package, then either prints the bytes to stdout or
//! writes them to a destination file.
//!
//! # Run
//!
//! ```bash
//! # Print /word/document.xml of the default test file to stdout:
//! cargo run -p litchi-opc --example extract_part
//!
//! # Specify package + partname:
//! cargo run -p litchi-opc --example extract_part -- \
//!     test-data/ooxml/docx/documentProperties.docx /word/document.xml
//!
//! # Save to a file instead of stdout:
//! cargo run -p litchi-opc --example extract_part -- \
//!     test-data/ooxml/docx/documentProperties.docx /word/document.xml /tmp/document.xml
//! ```

use litchi_opc::{OpcPackage, PackURI};
use std::io::Write;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let pkg_path: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("test-data/ooxml/docx/documentProperties.docx"));
    let partname_str = args
        .next()
        .unwrap_or_else(|| "/word/document.xml".to_string());
    let out_path: Option<PathBuf> = args.next().map(PathBuf::from);

    // Demonstrate PackURI parsing and inspection.
    let partname = PackURI::new(partname_str.clone())
        .map_err(|e| format!("invalid PackURI {partname_str:?}: {e}"))?;

    eprintln!("PackURI parsing:");
    eprintln!("  full:       {}", partname);
    eprintln!("  base_uri:   {}", partname.base_uri());
    eprintln!("  filename:   {}", partname.filename());
    eprintln!("  ext:        {}", partname.ext());
    eprintln!("  membername: {}", partname.membername());
    if let Some(idx) = partname.idx() {
        eprintln!("  idx:        {idx}");
    }
    eprintln!();

    let pkg = OpcPackage::open(&pkg_path)?;
    eprintln!(
        "Opened {} ({} parts). Looking up {}...",
        pkg_path.display(),
        pkg.part_count(),
        partname,
    );

    let part = pkg.get_part(&partname)?;
    let blob = part.blob();
    eprintln!(
        "Found part: content_type={}, size={} bytes",
        part.content_type(),
        blob.len()
    );

    match out_path {
        Some(path) => {
            std::fs::write(&path, blob)?;
            eprintln!("Wrote {} bytes to {}", blob.len(), path.display());
        },
        None => {
            // Stream the blob to stdout. Use a locked handle for efficiency.
            let stdout = std::io::stdout();
            let mut handle = stdout.lock();
            handle.write_all(blob)?;
            handle.flush()?;
        },
    }

    Ok(())
}
