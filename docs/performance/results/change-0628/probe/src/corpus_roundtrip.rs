//! Change 0628 probe A: open and republish every OOXML fixture in the
//! repository corpus and print a SHA-256 of the published bytes.
//!
//! Change 0628 alters only which of several equally matching relationships
//! `Relationships::get_or_add` / `get_or_add_ext_rel` reuses. Neither method is
//! reached by open or by publication, so every digest below must be identical
//! on the before and the after leg. Any difference would mean the change
//! touched a path it was not supposed to touch.
//!
//! Usage: corpus_roundtrip <test-data-root>
use sha2::{Digest, Sha256};
use std::path::{Path, PathBuf};

const EXTS: &[&str] = &[
    "docx", "docm", "dotx", "dotm", "xlsx", "xlsm", "xltx", "xltm", "xlsb", "pptx", "pptm", "potx",
    "potm", "ppsx", "ppsm", "thmx",
];

fn walk(root: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    let mut names: Vec<PathBuf> = entries.filter_map(|e| e.ok()).map(|e| e.path()).collect();
    names.sort();
    for path in names {
        if path.is_dir() {
            walk(&path, out);
        } else if path
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| EXTS.contains(&e.to_ascii_lowercase().as_str()))
        {
            out.push(path);
        }
    }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn main() {
    let root = PathBuf::from(std::env::args().nth(1).expect("test-data root"));
    let mut files = Vec::new();
    walk(&root, &mut files);
    let mut opened = 0usize;
    let mut published = 0usize;
    let mut refused_open = 0usize;
    let mut refused_save = 0usize;
    for path in &files {
        let relative = path.strip_prefix(&root).unwrap_or(path).display().to_string();
        let bytes = match std::fs::read(path) {
            Ok(bytes) => bytes,
            Err(error) => {
                println!("{relative}\tREAD-ERROR\t{error}");
                continue;
            },
        };
        let package = match litchi_opc::OpcPackage::from_bytes(&bytes) {
            Ok(package) => package,
            Err(error) => {
                refused_open += 1;
                println!("{relative}\tOPEN-REFUSED\t{error}");
                continue;
            },
        };
        opened += 1;
        let parts = package.part_count();
        let package_rels = package.rels().len();
        let part_rels: usize = package.iter_parts().map(|part| part.rels().len()).sum();
        match litchi_opc::PackageWriter::to_bytes(&package) {
            Ok(output) => {
                published += 1;
                let mut hasher = Sha256::new();
                hasher.update(&output);
                println!(
                    "{relative}\tOK\tparts={parts}\tpkg_rels={package_rels}\tpart_rels={part_rels}\tout_bytes={}\tsha256={}",
                    output.len(),
                    hex(&hasher.finalize())
                );
            },
            Err(error) => {
                refused_save += 1;
                println!("{relative}\tSAVE-REFUSED\tparts={parts}\tpkg_rels={package_rels}\tpart_rels={part_rels}\t{error}");
            },
        }
    }
    println!();
    println!("files={} opened={opened} published={published} open_refused={refused_open} save_refused={refused_save}", files.len());
}
