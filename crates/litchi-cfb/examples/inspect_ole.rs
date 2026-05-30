//! Inspect a Microsoft Compound File Binary (CFB / OLE2) container.
//!
//! This example verifies the file signature with [`is_ole_file`], opens it
//! with [`OleFile::open`], walks the directory tree printing each entry's
//! name, type, and size, and prints summary properties from
//! [`OleMetadata`] when available.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-cfb --example inspect_ole -- <path-to-ole-file>
//! ```
//!
//! If no path is supplied, defaults to a sample `.doc` file from the
//! workspace's `test-data/ole/doc/` directory.
//!
//! [`is_ole_file`]: litchi_cfb::is_ole_file
//! [`OleFile::open`]: litchi_cfb::OleFile::open
//! [`OleMetadata`]: litchi_cfb::OleMetadata

use std::fs::File;
use std::io::Read;
use std::path::PathBuf;

use litchi_cfb::{DirectoryEntry, OleFile, is_ole_file};

type ExampleResult<T> = Result<T, Box<dyn std::error::Error>>;

const DEFAULT_SAMPLE: &str = "test-data/ole/doc/Lists.doc";

fn main() -> ExampleResult<()> {
    let path: PathBuf = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(DEFAULT_SAMPLE));

    println!("=== litchi-cfb: inspect_ole ===");
    println!("Target file: {}", path.display());

    // 1. Read a prefix of the file and verify the CFB signature.
    // `is_ole_file` requires at least MINIMAL_OLEFILE_SIZE (1536) bytes,
    // so read a generously-sized prefix rather than just the magic bytes.
    let mut probe_buf = vec![0u8; 4096];
    let n = {
        let mut probe = File::open(&path)?;
        probe.read(&mut probe_buf)?
    };
    probe_buf.truncate(n);
    if !is_ole_file(&probe_buf) {
        return Err(format!(
            "Not a CFB/OLE2 file (magic mismatch): {}",
            path.display()
        )
        .into());
    }
    println!("Signature OK: D0 CF 11 E0 A1 B1 1A E1");

    // 2. Open the CFB container.
    let file = File::open(&path)?;
    let mut ole = OleFile::open(file)?;
    println!("File size:   {} bytes", ole.file_size());
    if let Some(name) = ole.get_root_name() {
        println!("Root entry:  {name}");
    }

    // 3. Walk every directory level and print each entry.
    println!("\n--- Directory tree ---");
    print_dir(&ole, &[], 0)?;

    // 4. Stream paths (flattened).
    let streams = ole.list_streams();
    println!("\n--- Streams ({}) ---", streams.len());
    for path in &streams {
        println!("  /{}", path.join("/"));
    }

    // 5. Metadata, if SummaryInformation is present.
    println!("\n--- Metadata ---");
    match ole.get_metadata() {
        Ok(meta) => print_metadata(&meta),
        Err(e) => println!("  (unable to read metadata: {e})"),
    }

    Ok(())
}

/// Recursively print directory entries. `path` is the slice of names that
/// addresses the current storage relative to the root.
fn print_dir<R: Read + std::io::Seek>(
    ole: &OleFile<R>,
    path: &[&str],
    depth: usize,
) -> ExampleResult<()> {
    let entries = ole.list_directory_entries(path)?;
    for entry in entries {
        let indent = "  ".repeat(depth);
        let kind = describe_type(entry.entry_type);
        println!(
            "{indent}- {name:<32} [{kind}] size={size}",
            name = entry.name,
            kind = kind,
            size = entry.size,
        );

        // Recurse into nested storages. Build a new path slice with the
        // child name appended.
        if is_storage(entry) {
            let mut next: Vec<&str> = path.to_vec();
            next.push(&entry.name);
            print_dir(ole, &next, depth + 1)?;
        }
    }
    Ok(())
}

fn is_storage(entry: &DirectoryEntry) -> bool {
    // STGTY_STORAGE = 1, STGTY_ROOT = 5
    entry.entry_type == 1 || entry.entry_type == 5
}

fn describe_type(t: u8) -> &'static str {
    match t {
        0 => "empty",
        1 => "storage",
        2 => "stream",
        3 => "lockbytes",
        4 => "property",
        5 => "root",
        _ => "unknown",
    }
}

fn print_metadata(meta: &litchi_cfb::OleMetadata) {
    let mut printed = false;
    let mut row = |label: &str, value: Option<&str>| {
        if let Some(v) = value
            && !v.is_empty()
        {
            println!("  {label:<22} {v}");
            printed = true;
        }
    };

    row("Title:", meta.title.as_deref());
    row("Subject:", meta.subject.as_deref());
    row("Author:", meta.author.as_deref());
    row("Keywords:", meta.keywords.as_deref());
    row("Comments:", meta.comments.as_deref());
    row("Template:", meta.template.as_deref());
    row("Last saved by:", meta.last_saved_by.as_deref());
    row("Revision:", meta.revision_number.as_deref());
    row("Application:", meta.creating_application.as_deref());
    row("Category:", meta.category.as_deref());
    row("Manager:", meta.manager.as_deref());
    row("Company:", meta.company.as_deref());

    if let Some(t) = meta.create_time {
        println!("  {:<22} {t}", "Created:");
        printed = true;
    }
    if let Some(t) = meta.last_saved_time {
        println!("  {:<22} {t}", "Last saved:");
        printed = true;
    }
    if let Some(t) = meta.last_printed_time {
        println!("  {:<22} {t}", "Last printed:");
        printed = true;
    }
    if let Some(d) = meta.edit_time {
        println!("  {:<22} {d}", "Edit time:");
        printed = true;
    }
    if let Some(p) = meta.num_pages {
        println!("  {:<22} {p}", "Pages:");
        printed = true;
    }
    if let Some(w) = meta.num_words {
        println!("  {:<22} {w}", "Words:");
        printed = true;
    }
    if let Some(c) = meta.num_chars {
        println!("  {:<22} {c}", "Characters:");
        printed = true;
    }
    if let Some(cp) = meta.codepage {
        println!("  {:<22} {cp}", "Codepage:");
        printed = true;
    }

    if !printed {
        println!("  (no SummaryInformation properties)");
    }
}
