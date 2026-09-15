//! Change 0628 probe B: for every OOXML fixture, reuse each distinct
//! (type, target, mode) relationship triple on the part that owns it and print
//! the identifier that reuse selected.
//!
//! On the before leg the selection is taken from `HashMap` visit order, so any
//! part owning two matching relationships prints a different identifier from
//! run to run. On the after leg every line is stable and equal to the smallest
//! matching identifier in byte order.
//!
//! Usage: reuse_selection <test-data-root> <repeat-count>
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use std::collections::{BTreeMap, BTreeSet};
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

fn main() {
    let root = PathBuf::from(std::env::args().nth(1).expect("test-data root"));
    let repeats: usize = std::env::args()
        .nth(2)
        .and_then(|value| value.parse().ok())
        .unwrap_or(16);
    let mut files = Vec::new();
    walk(&root, &mut files);
    let mut unstable = 0usize;
    let mut probed = 0usize;
    for path in &files {
        let relative = path.strip_prefix(&root).unwrap_or(path).display().to_string();
        let Ok(bytes) = std::fs::read(path) else {
            continue;
        };
        let Ok(package) = OpcPackage::from_bytes(&bytes) else {
            continue;
        };
        // Collect every (owner, type, target, mode) triple worth probing.
        let mut triples: BTreeSet<(String, String, String, bool)> = BTreeSet::new();
        for part in package.iter_parts() {
            let owner = part.partname().as_str().to_string();
            for relationship in part.rels().iter() {
                triples.insert((
                    owner.clone(),
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.target_mode() == TargetMode::External,
                ));
            }
        }
        drop(package);
        let mut selections: BTreeMap<(String, String, String, bool), BTreeSet<String>> =
            BTreeMap::new();
        for _ in 0..repeats {
            let Ok(mut package) = OpcPackage::from_bytes(&bytes) else {
                continue;
            };
            for triple in &triples {
                let Ok(partname) = PackURI::new(triple.0.as_str()) else {
                    continue;
                };
                let Ok(part) = package.get_part_mut(&partname) else {
                    continue;
                };
                let selected = if triple.3 {
                    part.relate_to_ext(triple.2.as_str(), triple.1.as_str())
                } else {
                    part.relate_to(triple.2.as_str(), triple.1.as_str())
                };
                selections.entry(triple.clone()).or_default().insert(selected);
            }
        }
        for (triple, ids) in selections {
            probed += 1;
            if ids.len() > 1 {
                unstable += 1;
                let joined: Vec<&str> = ids.iter().map(String::as_str).collect();
                println!(
                    "UNSTABLE\t{relative}\towner={}\tmode={}\ttarget={}\tselected={:?}\ttype={}",
                    triple.0,
                    if triple.3 { "External" } else { "Internal" },
                    triple.2,
                    joined,
                    triple.1
                );
            }
        }
    }
    println!();
    println!("files={} triples_probed={probed} unstable_triples={unstable} repeats={repeats}", files.len());
}
