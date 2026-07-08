//! Write a small CFB / OLE2 container with a single user stream.
//!
//! This example uses [`OleWriter`] to construct a minimal CFB file in
//! memory, save it to a temporary file, and then re-open it with
//! [`OleFile`] to round-trip-verify the contents.
//!
//! Gated on the `write` feature.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-cfb --features write --example write_ole [-- <output-path>]
//! ```
//!
//! If no output path is supplied the file is written to the system temp
//! directory.
//!
//! [`OleWriter`]: litchi_cfb::OleWriter
//! [`OleFile`]: litchi_cfb::OleFile

#[cfg(feature = "write")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::fs::File;
    use std::path::PathBuf;

    use litchi_cfb::{OleFile, OleWriter, is_ole_file};

    let out_path: PathBuf = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| std::env::temp_dir().join("litchi-cfb-demo.ole"));

    println!("=== litchi-cfb: write_ole ===");
    println!("Output file: {}", out_path.display());

    // Build a tiny container: one top-level storage with one stream inside,
    // plus a top-level stream alongside it.
    let mut writer = OleWriter::new();
    writer.create_stream(&["Greeting"], b"Hello from litchi-cfb!\n")?;
    writer.create_storage(&["Demo"])?;
    writer.create_stream(
        &["Demo", "Payload"],
        b"Nested stream contents written by the example.",
    )?;

    writer.save(&out_path)?;
    println!(
        "Wrote CFB file: {} bytes",
        std::fs::metadata(&out_path)?.len()
    );

    // Round-trip verify: re-open and list.
    // `is_ole_file` checks both the magic *and* a minimum size, so read a
    // sufficiently large prefix (>= 1536 bytes) rather than just 8 bytes.
    let mut head = vec![0u8; 4096];
    let n = {
        use std::io::Read;
        let mut f = File::open(&out_path)?;
        f.read(&mut head)?
    };
    head.truncate(n);
    assert!(is_ole_file(&head), "output is not a valid CFB file");

    let mut ole = OleFile::open(File::open(&out_path)?)?;
    println!("\nStreams in re-opened file:");
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = ole.open_stream(&refs)?;
        println!("  /{} ({} bytes)", path.join("/"), data.len());
    }

    println!("\nDone.");
    Ok(())
}

#[cfg(not(feature = "write"))]
fn main() {
    eprintln!(
        "This example requires the `write` feature. Re-run with:\n  \
         cargo run -p litchi-cfb --features write --example write_ole"
    );
    std::process::exit(1);
}
