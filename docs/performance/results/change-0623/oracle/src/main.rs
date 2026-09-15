//! Change 0623's open-differential oracle.
//!
//! For every container named on the input list this prints one deterministic
//! block: the open's verdict, the requests and bytes the open asked of the
//! source, the package relationships, every admitted Part with its content
//! type and relationships, every non-part member, and the decoded bytes of
//! every Part. Two legs' reports are compared with `cmp`.
use std::{
    fmt::Write as _,
    fs, io,
    sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}},
};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

struct Counting {
    inner: Vec<u8>,
    requests: AtomicU64,
    bytes: AtomicU64,
    versions: AtomicU64,
    log: Mutex<Vec<(u64, u64, u64)>>,
}

impl Counting {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: bytes,
            requests: AtomicU64::new(0),
            bytes: AtomicU64::new(0),
            versions: AtomicU64::new(0),
            log: Mutex::new(Vec::new()),
        }
    }
    fn take(&self) -> (u64, u64, u64, Vec<(u64, u64, u64)>) {
        (
            self.requests.swap(0, Ordering::SeqCst),
            self.bytes.swap(0, Ordering::SeqCst),
            self.versions.swap(0, Ordering::SeqCst),
            std::mem::take(&mut *self.log.lock().unwrap()),
        )
    }
}

impl ReadAt for Counting {
    fn len(&self) -> io::Result<u64> {
        Ok(self.inner.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        self.requests.fetch_add(1, Ordering::SeqCst);
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let take = if start >= self.inner.len() {
            0
        } else {
            output.len().min(self.inner.len() - start)
        };
        if take > 0 {
            output[..take].copy_from_slice(&self.inner[start..start + take]);
        }
        self.bytes.fetch_add(take as u64, Ordering::SeqCst);
        self.log.lock().unwrap().push((offset, output.len() as u64, take as u64));
        Ok(take)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::SeqCst);
        Ok(SourceVersion::new(0x5a1, 1))
    }
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut table = [0u32; 256];
    for (i, slot) in table.iter_mut().enumerate() {
        let mut c = i as u32;
        for _ in 0..8 {
            c = if c & 1 != 0 { 0xEDB8_8320 ^ (c >> 1) } else { c >> 1 };
        }
        *slot = c;
    }
    let mut crc = 0xFFFF_FFFFu32;
    for b in bytes {
        crc = table[((crc ^ u32::from(*b)) & 0xFF) as usize] ^ (crc >> 8);
    }
    crc ^ 0xFFFF_FFFF
}

fn open(src: Arc<Counting>) -> Result<SourceBackedPackage, litchi_opc::OpcError> {
    let dynamic: Arc<dyn ReadAt> = src;
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        dynamic,
        ReadLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
    )
}

fn main() {
    let mut args = std::env::args().skip(1);
    let list = args.next().expect("usage: oracle <file-list> <report> [--log]");
    let report_path = args.next().expect("usage: oracle <file-list> <report> [--log]");
    let want_log = args.next().as_deref() == Some("--log");
    let listing = fs::read_to_string(&list).expect("file list");
    let mut names: Vec<&str> = listing.lines().filter(|l| !l.is_empty()).collect();
    names.sort_unstable();
    let mut out = String::with_capacity(1 << 22);
    for name in names {
        let Ok(bytes) = fs::read(name) else {
            let _ = writeln!(out, "=== {name} unreadable");
            continue;
        };
        let _ = writeln!(out, "=== {name} len={}", bytes.len());
        let src = Arc::new(Counting::new(bytes));
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| open(Arc::clone(&src))));
        let package = match result {
            Err(_) => {
                let _ = writeln!(out, "open PANIC");
                continue;
            },
            Ok(Err(error)) => {
                let (req, by, ver, log) = src.take();
                let _ = writeln!(out, "open err({error:?})");
                let _ = writeln!(out, "open-cost requests={req} bytes={by} versions={ver}");
                if want_log {
                    for (offset, len, got) in log {
                        let _ = writeln!(out, "open-read {offset} {len} {got}");
                    }
                }
                continue;
            },
            Ok(Ok(package)) => package,
        };
        let (req, by, ver, log) = src.take();
        let _ = writeln!(out, "open ok");
        let _ = writeln!(out, "open-cost requests={req} bytes={by} versions={ver}");
        if want_log {
            for (offset, len, got) in log {
                let _ = writeln!(out, "open-read {offset} {len} {got}");
            }
        }
        let mut lines: Vec<String> = Vec::new();
        for rel in package.rels().iter() {
            lines.push(format!(
                "pkgrel {} {} {} {:?} ext={}",
                rel.r_id(),
                rel.reltype(),
                rel.target_ref(),
                rel.target_mode(),
                rel.is_external()
            ));
        }
        for member in package.non_part_members() {
            lines.push(format!("nonpart {} {}", member.name(), member.reason().as_str()));
        }
        let partnames: Vec<String> = package
            .iter_parts()
            .map(|part| part.partname().to_string())
            .collect();
        for part in package.iter_parts() {
            lines.push(format!(
                "part {} ct={} rels={}",
                part.partname(),
                part.content_type(),
                part.rels().len()
            ));
            for rel in part.rels().iter() {
                lines.push(format!(
                    "partrel {} {} {} {} {:?} ext={}",
                    part.partname(),
                    rel.r_id(),
                    rel.reltype(),
                    rel.target_ref(),
                    rel.target_mode(),
                    rel.is_external()
                ));
            }
        }
        lines.sort();
        for line in lines {
            let _ = writeln!(out, "{line}");
        }
        let mut sorted = partnames.clone();
        sorted.sort();
        for name in &sorted {
            let uri = match PackURI::new(name.as_str()) {
                Ok(uri) => uri,
                Err(error) => {
                    let _ = writeln!(out, "data {name} uri-err({error:?})");
                    continue;
                },
            };
            match package.part(&uri).and_then(|part| part.data()) {
                Ok(data) => {
                    let payload = data.as_bytes();
                    let _ = writeln!(
                        out,
                        "data {name} ok(len={} crc={:08x})",
                        payload.len(),
                        crc32(payload)
                    );
                },
                Err(error) => {
                    let _ = writeln!(out, "data {name} err({error:?})");
                },
            }
        }
        let (req, by, ver, _) = src.take();
        let _ = writeln!(out, "read-cost requests={req} bytes={by} versions={ver}");
    }
    fs::write(&report_path, out).expect("report");
}
