//! Build a tiny OPC package from scratch with `PackageWriter`.
//!
//! This example assembles an `OpcPackage` containing a single XML part
//! (`/custom/data.xml`), wires up a package-level relationship pointing to
//! that part, writes the package to a temporary file via `PackageWriter`,
//! and finally re-opens it with `OpcPackage::open` to verify the round-trip.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-opc --example write_package
//! ```

use litchi_opc::{OpcPackage, PackURI, PackageWriter, XmlPart};
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // --- 1. Build the package in memory --------------------------------
    let mut pkg = OpcPackage::new();

    let partname = PackURI::new("/custom/data.xml")
        .map_err(|e| format!("invalid PackURI: {e}"))?;
    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<root xmlns="urn:litchi-opc:example">
  <hello>world</hello>
  <pi value="3.14159"/>
</root>"#
        .to_vec();
    let content_type = "application/vnd.litchi-opc.example+xml".to_string();

    let part = XmlPart::new(partname.clone(), content_type.clone(), xml.clone());
    pkg.add_part(Box::new(part));

    // Wire up a package-level relationship pointing at the new part.
    let reltype = "http://schemas.litchi.example/2024/relationships/customData";
    let r_id = pkg.relate_to(partname.as_str(), reltype);
    println!("Created package relationship {r_id} -> {partname} (type {reltype})");

    // Add an external relationship for good measure.
    let ext_r_id = pkg.relate_to_external(
        "https://example.com/spec",
        "http://schemas.litchi.example/2024/relationships/spec",
    );
    println!("Created external relationship {ext_r_id} -> https://example.com/spec");

    // --- 2. Save to a temp file ----------------------------------------
    // Use std::env::temp_dir() + a timestamp suffix to avoid pulling
    // `tempfile` in just for the example.
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let mut out_path: PathBuf = std::env::temp_dir();
    out_path.push(format!("litchi_opc_write_demo_{nanos}.opc"));

    PackageWriter::write(&out_path, &pkg)?;
    println!("\nWrote package to {}", out_path.display());

    let on_disk_size = std::fs::metadata(&out_path)?.len();
    println!("Serialized package size: {on_disk_size} bytes");

    // --- 3. Read the package back to verify the round trip --------------
    let reopened = OpcPackage::open(&out_path)?;
    println!(
        "\nRe-opened package: {} parts, {} package relationships",
        reopened.part_count(),
        reopened.rels().len()
    );

    let round_tripped = reopened.get_part(&partname)?;
    println!(
        "  /custom/data.xml -> content_type={}, size={}",
        round_tripped.content_type(),
        round_tripped.blob().len()
    );
    assert_eq!(round_tripped.content_type(), content_type);
    assert_eq!(round_tripped.blob(), xml.as_slice());

    let internal_rel = reopened.rels().iter().find(|r| !r.is_external());
    if let Some(rel) = internal_rel {
        println!(
            "  internal rel: {} -> {} (type {})",
            rel.r_id(),
            rel.target_ref(),
            rel.reltype()
        );
    }
    let external_rel = reopened.rels().iter().find(|r| r.is_external());
    if let Some(rel) = external_rel {
        println!(
            "  external rel: {} -> {} (type {})",
            rel.r_id(),
            rel.target_ref(),
            rel.reltype()
        );
    }

    // --- 4. Clean up ----------------------------------------------------
    if let Err(e) = std::fs::remove_file(&out_path) {
        eprintln!(
            "warning: could not remove temp file {}: {}",
            out_path.display(),
            e
        );
    }

    println!("\nRound-trip verified successfully.");
    Ok(())
}
