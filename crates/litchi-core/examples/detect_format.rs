//! Demonstrates signature-based file format detection in `litchi-core`.
//!
//! Reads up to the first 512 bytes of a file (or every file in a directory)
//! and runs the public detection helpers from `litchi_core::detection` to
//! classify the contents.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-core --example detect_format -- <path>
//! cargo run -p litchi-core --example detect_format -- test-data/ooxml/docx
//! ```

use litchi_core::FileFormat;
use litchi_core::detection::simd_utils::check_office_signatures;
use std::fs;
use std::io::Read;
use std::path::Path;

const HEAD_BYTES: usize = 512;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let target = match args.next() {
        Some(p) => p,
        None => {
            eprintln!("usage: detect_format <file-or-directory>");
            std::process::exit(2);
        },
    };

    let path = Path::new(&target);
    if path.is_dir() {
        walk_dir(path)?;
    } else if path.is_file() {
        let label = describe(path)?;
        println!("{}: {}", path.display(), label);
    } else {
        eprintln!("path does not exist or is not accessible: {}", target);
        std::process::exit(1);
    }

    Ok(())
}

/// Walk the immediate children of `dir` (non-recursive) and detect each file.
fn walk_dir(dir: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut entries: Vec<_> = fs::read_dir(dir)?.filter_map(Result::ok).collect();
    entries.sort_by_key(|e| e.file_name());

    println!("Scanning {} ...", dir.display());
    for entry in entries {
        let p = entry.path();
        if !p.is_file() {
            continue;
        }
        let name = p.file_name().unwrap().to_string_lossy().into_owned();
        match describe(&p) {
            Ok(label) => println!("  {:<40}  {}", name, label),
            Err(e) => println!("  {:<40}  <error: {}>", name, e),
        }
    }
    Ok(())
}

/// Read the first `HEAD_BYTES` bytes of `path` and classify them.
fn describe(path: &Path) -> Result<String, Box<dyn std::error::Error>> {
    let mut file = fs::File::open(path)?;
    let mut buf = vec![0u8; HEAD_BYTES];
    let read = read_up_to(&mut file, &mut buf)?;
    buf.truncate(read);

    Ok(format_label(&buf, path))
}

/// Read until either the buffer is full or EOF.
fn read_up_to<R: Read>(reader: &mut R, buf: &mut [u8]) -> std::io::Result<usize> {
    let mut filled = 0;
    while filled < buf.len() {
        match reader.read(&mut buf[filled..])? {
            0 => break,
            n => filled += n,
        }
    }
    Ok(filled)
}

/// Run the available signature checks against `bytes` and return a friendly
/// label.
///
/// `litchi-core` exposes signature-level helpers; the umbrella `litchi`
/// crate layers ZIP-content inspection on top to disambiguate
/// .docx vs .xlsx vs .pptx vs iWork. We surface the underlying signature
/// here so callers can see exactly what the byte-level detection tells us.
fn format_label(bytes: &[u8], path: &Path) -> String {
    // RTF: matches `{\rtf` prefix (always available regardless of feature
    // flags — the helper itself has no `cfg` gate).
    if let Some(fmt) = litchi_core::detection::rtf::detect_rtf_format(bytes) {
        return describe_format(fmt);
    }

    // Lower-level signature mask for the remaining categories.
    let mask = check_office_signatures(bytes);
    if mask.is_ole2() {
        return format!(
            "OLE2 container (legacy Office: .doc/.xls/.ppt) — {}",
            suffix(path)
        );
    }
    if mask.is_zip() {
        return format!("ZIP container (OOXML/ODF/iWork) — {}", suffix(path));
    }
    if mask.is_rtf() {
        // Defensive — should already be caught above.
        return describe_format(FileFormat::Rtf);
    }

    "unknown / unrecognised signature".to_string()
}

fn describe_format(fmt: FileFormat) -> String {
    format!("{:?}", fmt)
}

fn suffix(path: &Path) -> String {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|e| format!("extension: .{}", e))
        .unwrap_or_else(|| "no extension".to_string())
}
