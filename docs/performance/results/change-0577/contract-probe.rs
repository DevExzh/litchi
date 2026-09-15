//! Change 0577 contract probe.
//!
//! Opens each generated package through both OOXML ingress modes and reports
//! the open's verdict. Every fixture isolates one open-time decision that the
//! relationship-part reads are responsible for. Nothing under `crates/` is
//! touched; this only calls the public API.

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};

/// A positional source that counts the calls the library makes.
struct Counting {
    bytes: Vec<u8>,
    calls: AtomicUsize,
}

impl Counting {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            calls: AtomicUsize::new(0),
        }
    }
    fn calls(&self) -> usize {
        self.calls.load(Ordering::SeqCst)
    }
}

impl ReadAt for Counting {
    fn read_at(&self, offset: u64, buf: &mut [u8]) -> std::io::Result<usize> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        let start = usize::try_from(offset)
            .unwrap_or(usize::MAX)
            .min(self.bytes.len());
        let end = start.saturating_add(buf.len()).min(self.bytes.len());
        let n = end - start;
        buf[..n].copy_from_slice(&self.bytes[start..end]);
        Ok(n)
    }
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }
    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x0577, 1))
    }
}

fn verdict<T, E: std::fmt::Display>(r: &Result<T, E>) -> String {
    match r {
        Ok(_) => "Ok".to_string(),
        Err(e) => format!("Err({e})"),
    }
}

fn main() {
    let dir = std::env::args().nth(1).expect("usage: contract <fixture-dir>");
    let mut entries: Vec<_> = std::fs::read_dir(&dir)
        .expect("fixture dir")
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|x| x == "docx"))
        .collect();
    entries.sort();

    let mut rows = Vec::new();
    for path in entries {
        let name = path.file_stem().unwrap().to_string_lossy().to_string();
        let bytes = std::fs::read(&path).expect("fixture bytes");

        // Source-backed ingress (the path change 0572 measured).
        let counting = Arc::new(Counting::new(bytes.clone()));
        let counter = Arc::clone(&counting);
        let dynamic: Arc<dyn ReadAt> = counting;
        let sb = litchi_opc::SourceBackedPackage::from_read_at(dynamic);
        let sb_open_calls = counter.calls();
        let sb_verdict = verdict(&sb);

        // Materialized ingress (the eager path).
        let eager = litchi_opc::OpcPackage::from_bytes(&bytes);
        let eager_verdict = verdict(&eager);

        // Non-part members reported by a successful source-backed open.
        let non_parts: Vec<String> = sb
            .as_ref()
            .ok()
            .map(|p| {
                p.non_part_members()
                    .iter()
                    .map(|m| format!("{}", m.name()))
                    .collect()
            })
            .unwrap_or_default();

        rows.push(serde_json::json!({
            "fixture": name,
            "source_backed_open": sb_verdict,
            "materialized_open": eager_verdict,
            "source_backed_open_requests": sb_open_calls,
            "non_part_members": non_parts,
        }));
    }
    println!("{}", serde_json::to_string_pretty(&rows).unwrap());
}
