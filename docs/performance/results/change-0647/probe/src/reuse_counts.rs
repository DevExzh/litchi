//! Change 0647 probe C: deterministic counts for one no-op reusing
//! `get_or_add` followed by an owned-source publication.
//!
//! Counted per fixture, with a counting global allocator:
//!
//! * `open`   — `OpcPackage::from_vec` (the owned-source route, which is the
//!   only one that binds change 0593's open-time captures);
//! * `seam`   — taking the mutable seam alone, which is the baseline both legs
//!   share;
//! * `reuse`  — the `get_or_add(type, target)` call on a pair the collection
//!   already carries;
//! * `save`   — `PackageWriter::to_bytes`;
//! * `published` — the published byte count and its SHA-256.
//!
//! The owner is the package's own `_rels/.rels` when `--package` is given and
//! the named part otherwise. `canonical_rels_bytes` reports the size of the
//! canonical `.rels` serialization for that owner: the bytes the publication
//! plan builds and then discards when the capture has been invalidated.
//!
//! Usage: reuse_counts [--package|--part /partname] <fixture>...
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicU64, Ordering};

struct CountingAlloc;
static ALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
unsafe impl GlobalAlloc for CountingAlloc {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
        unsafe { System.realloc(pointer, layout, new_size) }
    }
}
#[global_allocator]
static GLOBAL: CountingAlloc = CountingAlloc;

fn snapshot() -> (u64, u64) {
    (
        ALLOCS.load(Ordering::Relaxed),
        ALLOC_BYTES.load(Ordering::Relaxed),
    )
}

use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use sha2::{Digest, Sha256};

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mut part: Option<String> = None;
    let mut fixtures: Vec<String> = Vec::new();
    let mut index = 0usize;
    while index < args.len() {
        match args[index].as_str() {
            "--package" => index += 1,
            "--part" => {
                part = args.get(index + 1).cloned();
                index += 2;
            },
            other => {
                fixtures.push(other.to_string());
                index += 1;
            },
        }
    }

    for fixture in &fixtures {
        let bytes = std::fs::read(fixture).expect("fixture");

        // Which (type, target) pair to reuse: the smallest internal pair of
        // the chosen owner, so the selection is a function of the package.
        let probe = OpcPackage::from_vec(bytes.clone()).expect("open");
        let uri = part.as_ref().map(|name| PackURI::new(name).expect("partname"));
        let (reltype, target, owner_label, canonical_len) = {
            let rels = match &uri {
                None => probe.rels(),
                Some(uri) => probe.get_part(uri).expect("part").rels(),
            };
            let mut pairs: Vec<(String, String)> = rels
                .iter()
                .filter(|rel| !rel.is_external())
                .map(|rel| (rel.reltype().to_string(), rel.target_ref().to_string()))
                .collect();
            pairs.sort();
            let (reltype, target) = pairs.first().cloned().expect("an internal relationship");
            // `to_xml` is the public twin of the crate-private
            // `try_to_xml_bytes`; `rel.rs:973` asserts they agree byte for byte.
            let canonical_len = rels.to_xml().len();
            let label = uri
                .as_ref()
                .map(|uri| uri.as_str().to_string())
                .unwrap_or_else(|| "/".to_string());
            (reltype, target, label, canonical_len)
        };
        drop(probe);

        let before_open = snapshot();
        let mut package = OpcPackage::from_vec(bytes.clone()).expect("open");
        let after_open = snapshot();

        let before_seam = snapshot();
        match &uri {
            None => {
                let _seam = package.rels_mut();
            },
            Some(uri) => {
                let _seam = package.get_part_mut(uri).expect("part");
            },
        }
        let after_seam = snapshot();

        let before_reuse = snapshot();
        let r_id = match &uri {
            None => package
                .rels_mut()
                .get_or_add(&reltype, &target)
                .r_id()
                .to_string(),
            Some(uri) => package
                .get_part_mut(uri)
                .expect("part")
                .rels_mut()
                .get_or_add(&reltype, &target)
                .r_id()
                .to_string(),
        };
        let after_reuse = snapshot();

        let before_save = snapshot();
        let output = PackageWriter::to_bytes(&package).expect("save");
        let after_save = snapshot();

        let mut hasher = Sha256::new();
        hasher.update(&output);
        println!(
            "{fixture}\towner={owner_label}\tr_id={r_id}\tcanonical_rels_bytes={canonical_len}\n  \
             open   allocs={} bytes={}\n  \
             seam   allocs={} bytes={}\n  \
             reuse  allocs={} bytes={}\n  \
             save   allocs={} bytes={}\n  \
             published bytes={} sha256={}",
            after_open.0 - before_open.0,
            after_open.1 - before_open.1,
            after_seam.0 - before_seam.0,
            after_seam.1 - before_seam.1,
            after_reuse.0 - before_reuse.0,
            after_reuse.1 - before_reuse.1,
            after_save.0 - before_save.0,
            after_save.1 - before_save.1,
            output.len(),
            hex(&hasher.finalize())
        );
    }
}
