//! Change 0647 probe E: what a one-object edit through DOCX's ordinary route
//! does to the relationships members of an opened package.
//!
//! The route is the documented one change 0638 registered as
//! `docx_ordinary_save_*`: `Package::open(path)` →
//! `document_mut().add_paragraph_with_text(..)` → `to_stream(sink)`.
//!
//! For every `.docx`-family fixture this prints, per relationships member the
//! source archive carries, whether the published member is byte-identical to
//! the source member (`kept`) or was republished from a regenerated
//! serialization (`moved`), plus a SHA-256 of the whole publication. Running it
//! on both legs and differencing answers two questions at once: whether the
//! route's published bytes move (they must not), and whether any member that
//! was `moved` before the change becomes `kept` after it, which is the only way
//! this change can reach a format's ordinary save.
//!
//! Usage: docx_route <test-data-root>
mod corpus;

use litchi_opc::phys_pkg::PhysPkgReader;
use sha2::{Digest, Sha256};
use std::path::{Path, PathBuf};

const DOCX_EXTS: &[&str] = &["docx", "docm", "dotx", "dotm"];

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
            .is_some_and(|e| DOCX_EXTS.contains(&e.to_ascii_lowercase().as_str()))
        {
            out.push(path);
        }
    }
}

/// Every `.rels` member of an archive, as (name, bytes), in name order.
fn rels_members(archive: &[u8]) -> Vec<(String, Vec<u8>)> {
    let Ok(reader) = PhysPkgReader::new(archive) else {
        return Vec::new();
    };
    let Ok(names) = reader.member_names() else {
        return Vec::new();
    };
    let mut rows: Vec<(String, Vec<u8>)> = names
        .into_iter()
        .filter(|name| name.ends_with(".rels"))
        .filter_map(|name| reader.read_member(&name).ok().map(|bytes| (name, bytes)))
        .collect();
    rows.sort();
    rows
}

fn main() {
    let root = PathBuf::from(std::env::args().nth(1).expect("test-data root"));
    let mut files = Vec::new();
    walk(&root, &mut files);

    let mut opened = 0usize;
    let mut edited = 0usize;
    let mut published = 0usize;
    let mut members_kept = 0usize;
    let mut members_moved = 0usize;
    let mut members_absent = 0usize;

    for path in &files {
        let relative = path
            .strip_prefix(&root)
            .unwrap_or(path)
            .display()
            .to_string();
        let Ok(source) = std::fs::read(path) else {
            println!("{relative}\tREAD-ERROR");
            continue;
        };
        let mut package = match litchi_docx::Package::open(path) {
            Ok(package) => package,
            Err(error) => {
                println!("{relative}\tOPEN-REFUSED\t{error}");
                continue;
            },
        };
        opened += 1;
        match package.document_mut() {
            Ok(document) => {
                document.add_paragraph_with_text("litchi change 0647");
                edited += 1;
            },
            Err(error) => {
                // A typed editor refusal is an outcome, not a failure: the
                // save then publishes the unedited opened package, which is
                // change 0593's `noop` scenario.
                println!("{relative}\tEDIT-REFUSED\t{error}");
            },
        }
        let mut output = Vec::new();
        if let Err(error) = package.to_stream(&mut output) {
            println!("{relative}\tSAVE-REFUSED\t{error}");
            continue;
        }
        published += 1;

        let mut hasher = Sha256::new();
        hasher.update(&output);
        let digest = corpus::hex(&hasher.finalize());
        let source_members = rels_members(&source);
        let output_members: std::collections::BTreeMap<String, Vec<u8>> =
            rels_members(&output).into_iter().collect();
        let mut kept = 0usize;
        let mut moved = 0usize;
        let mut absent = 0usize;
        for (name, bytes) in &source_members {
            match output_members.get(name) {
                Some(published_bytes) if published_bytes == bytes => {
                    kept += 1;
                    members_kept += 1;
                },
                Some(_) => {
                    moved += 1;
                    members_moved += 1;
                    println!("{relative}\t{name}\tmoved");
                },
                None => {
                    absent += 1;
                    members_absent += 1;
                    println!("{relative}\t{name}\tabsent");
                },
            }
        }
        println!(
            "#FIXTURE\t{relative}\tsource_rels={}\tkept={kept}\tmoved={moved}\tabsent={absent}\tout_bytes={}\tsha256={digest}",
            source_members.len(),
            output.len()
        );
    }

    println!("#TOTAL files={}", files.len());
    println!("#TOTAL opened={opened}");
    println!("#TOTAL edited={edited}");
    println!("#TOTAL published={published}");
    println!("#TOTAL rels_members_kept={members_kept}");
    println!("#TOTAL rels_members_moved={members_moved}");
    println!("#TOTAL rels_members_absent={members_absent}");
}
