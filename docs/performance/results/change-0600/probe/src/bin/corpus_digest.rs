//! Corpus digest: open every OOXML fixture through `SourceBackedPackage` and
//! print one deterministic line per package and per Part, covering the catalog
//! (Part names, content types, relationship ids/types/targets), the payload
//! bytes of every Part, and the exact rendering of every refusal. Run it on
//! both legs and diff the outputs.
use std::fmt::Write as _;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{PackURI, SourceBackedPackage};

struct Bytes(Vec<u8>);

impl ReadAt for Bytes {
    fn len(&self) -> std::io::Result<u64> { Ok(self.0.len() as u64) }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        if start >= self.0.len() { return Ok(0); }
        let take = output.len().min(self.0.len() - start);
        output[..take].copy_from_slice(&self.0[start..start + take]);
        Ok(take)
    }
    fn version(&self) -> std::io::Result<SourceVersion> { Ok(SourceVersion::new(0x600, 1)) }
}

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn corpus(root: &Path) -> Vec<PathBuf> {
    let mut stack = vec![root.to_path_buf()];
    let mut found = Vec::new();
    while let Some(directory) = stack.pop() {
        let Ok(entries) = fs::read_dir(&directory) else { continue };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() { stack.push(path); continue; }
            if path.extension().and_then(|e| e.to_str()).is_some_and(|e| matches!(
                e.to_ascii_lowercase().as_str(),
                "docx"|"docm"|"dotx"|"xlsx"|"xlsm"|"xlsb"|"xltx"|"pptx"|"pptm"|"potx"|"ppsx"|"thmx"|"zip")) {
                found.push(path);
            }
        }
    }
    found.sort();
    found
}

fn main() {
    let root = std::env::args().nth(1).expect("corpus root");
    let mut out = String::new();
    for path in corpus(Path::new(&root)) {
        let relative = path.strip_prefix(&root).unwrap_or(&path).display().to_string();
        let Ok(bytes) = fs::read(&path) else {
            let _ = writeln!(out, "{relative}\tUNREADABLE");
            continue;
        };
        let source: Arc<dyn ReadAt> = Arc::new(Bytes(bytes));
        let package = match SourceBackedPackage::from_read_at(source) {
            Ok(package) => package,
            Err(error) => { let _ = writeln!(out, "{relative}\tOPEN-ERR\t{error}"); continue; },
        };
        let mut parts: Vec<(String, String, Vec<(String, String, String)>)> = package
            .iter_parts()
            .map(|part| {
                // `Relationships::iter()` walks a `HashMap` whose order is
                // randomized per process, so sort before digesting.
                let mut rels: Vec<(String, String, String)> = part
                    .rels()
                    .iter()
                    .map(|rel| (rel.r_id().to_string(), rel.reltype().to_string(), rel.target_ref().to_string()))
                    .collect();
                rels.sort();
                (part.partname().to_string(), part.content_type().to_string(), rels)
            })
            .collect();
        parts.sort();
        let _ = writeln!(out, "{relative}\tparts={}", parts.len());
        for (name, content_type, rels) in &parts {
            let uri = PackURI::new(name.as_str());
            let payload = match uri.as_ref().map(|uri| package.part(uri)) {
                Ok(Ok(view)) => match view.data() {
                    Ok(data) => format!("ok:{}:{:016x}", data.as_bytes().len(), fnv1a(data.as_bytes())),
                    Err(error) => format!("data-err:{error}"),
                },
                Ok(Err(error)) => format!("part-err:{error}"),
                Err(error) => format!("uri-err:{error}"),
            };
            // Stream the same Part: the monitored path must agree byte for byte.
            let streamed = match uri.as_ref().map(|uri| package.part(uri)) {
                Ok(Ok(view)) => {
                    let mut sink = Vec::new();
                    match view.stream_to(&mut sink) {
                        Ok(n) => format!("ok:{n}:{:016x}", fnv1a(&sink)),
                        Err(error) => format!("stream-err:{error}"),
                    }
                },
                Ok(Err(error)) => format!("part-err:{error}"),
                Err(error) => format!("uri-err:{error}"),
            };
            // A cold read taken AFTER the stream on the same package: the
            // bounded monitor scope must not change what it returns.
            let after_stream = match uri.as_ref().map(|uri| package.part(uri)) {
                Ok(Ok(view)) => match view.data() {
                    Ok(data) => format!("ok:{}:{:016x}", data.as_bytes().len(), fnv1a(data.as_bytes())),
                    Err(error) => format!("data-err:{error}"),
                },
                Ok(Err(error)) => format!("part-err:{error}"),
                Err(error) => format!("uri-err:{error}"),
            };
            let mut rel_text = String::new();
            for (id, reltype, target) in rels {
                let _ = write!(rel_text, "{id}|{reltype}|{target};");
            }
            let _ = writeln!(
                out,
                "{relative}\t{name}\t{content_type}\t{payload}\t{streamed}\t{after_stream}\trels={}\t{rel_text}",
                rels.len()
            );
        }
        // Absent and non-canonical Part spellings: the refusals must be identical.
        for probe in ["/word/absent.xml", "/xl/../xl/workbook.xml", "/"] {
            let rendered = match PackURI::new(probe) {
                Ok(uri) => match package.part(&uri) {
                    Ok(view) => match view.data() {
                        Ok(data) => format!("ok:{}", data.as_bytes().len()),
                        Err(error) => format!("data-err:{error}"),
                    },
                    Err(error) => format!("part-err:{error}"),
                },
                Err(error) => format!("uri-err:{error}"),
            };
            let _ = writeln!(out, "{relative}\tPROBE\t{probe}\t{rendered}");
        }
    }
    print!("{out}");
}
