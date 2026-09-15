use std::{fs, io, alloc::{GlobalAlloc, Layout, System}, sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}}};
struct Counting_Alloc;
static ALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
unsafe impl GlobalAlloc for Counting_Alloc {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 { ALLOCS.fetch_add(1, Ordering::Relaxed); ALLOC_BYTES.fetch_add(l.size() as u64, Ordering::Relaxed); unsafe { System.alloc(l) } }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) { unsafe { System.dealloc(p, l) } }
    unsafe fn realloc(&self, p: *mut u8, l: Layout, n: usize) -> *mut u8 { ALLOCS.fetch_add(1, Ordering::Relaxed); ALLOC_BYTES.fetch_add(n as u64, Ordering::Relaxed); unsafe { System.realloc(p, l, n) } }
}
#[global_allocator]
static GA: Counting_Alloc = Counting_Alloc;
fn allocs_snapshot() -> (u64, u64) { (ALLOCS.load(Ordering::Relaxed), ALLOC_BYTES.load(Ordering::Relaxed)) }
static VERSION_CALLS: AtomicU64 = AtomicU64::new(0);
use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

#[derive(Clone, Copy, Debug)]
struct Req { offset: u64, len: u64, returned: u64 }

struct Counting { inner: Vec<u8>, log: Mutex<Vec<Req>>, calls: AtomicU64 }
impl Counting {
    fn new(bytes: Vec<u8>) -> Self { Self { inner: bytes, log: Mutex::new(Vec::new()), calls: AtomicU64::new(0) } }
    fn take(&self) -> Vec<Req> { std::mem::take(&mut *self.log.lock().unwrap()) }
}
impl ReadAt for Counting {
    fn len(&self) -> io::Result<u64> { Ok(self.inner.len() as u64) }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() { return Ok(0); }
        self.calls.fetch_add(1, Ordering::SeqCst);
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let take = if start >= self.inner.len() { 0 } else { output.len().min(self.inner.len() - start) };
        if take > 0 { output[..take].copy_from_slice(&self.inner[start..start + take]); }
        self.log.lock().unwrap().push(Req { offset, len: output.len() as u64, returned: take as u64 });
        Ok(take)
    }
    fn version(&self) -> io::Result<SourceVersion> { VERSION_CALLS.fetch_add(1, Ordering::SeqCst); Ok(SourceVersion::new(0x5a1, 1)) }
}

// Minimal central-directory walk for request classification (independent of soapberry-zip).
struct Geometry { cd_offset: u64, cd_size: u64, eocd_offset: u64, members: Vec<(u64, String, u64, bool)> }
fn geometry(bytes: &[u8]) -> Geometry {
    let n = bytes.len();
    let mut eocd = None;
    for i in (0..n.saturating_sub(21)).rev() { if bytes[i..i+4] == [0x50,0x4b,0x05,0x06] { eocd = Some(i); break; } }
    let e = eocd.expect("eocd");
    let u16at = |p: usize| u16::from_le_bytes([bytes[p], bytes[p+1]]) as u64;
    let u32at = |p: usize| u32::from_le_bytes([bytes[p], bytes[p+1], bytes[p+2], bytes[p+3]]) as u64;
    let cd_size = u32at(e + 12); let cd_offset = u32at(e + 16);
    let mut members = Vec::new(); let mut p = cd_offset as usize;
    while (p as u64) < cd_offset + cd_size {
        assert_eq!(&bytes[p..p+4], &[0x50,0x4b,0x01,0x02]);
        let flags = u16at(p + 8); let csize = u32at(p + 20);
        let nlen = u16at(p + 28) as usize; let xlen = u16at(p + 30) as usize; let clen = u16at(p + 32) as usize;
        let lho = u32at(p + 42);
        let name = String::from_utf8_lossy(&bytes[p+46..p+46+nlen]).to_string();
        members.push((lho, name, csize, flags & 8 != 0));
        p += 46 + nlen + xlen + clen;
    }
    members.sort_by_key(|m| m.0);
    Geometry { cd_offset, cd_size, eocd_offset: e as u64, members }
}
fn classify(g: &Geometry, r: &Req) -> &'static str {
    if r.offset == g.eocd_offset || (r.offset + r.len >= g.eocd_offset && r.offset >= g.cd_offset + g.cd_size) { return "eocd/tail"; }
    if r.offset == g.cd_offset && r.len == 46 { return "cd-probe(46)"; }
    if r.offset >= g.cd_offset && r.offset < g.cd_offset + g.cd_size { return "central-dir"; }
    if let Some(m) = g.members.iter().find(|m| m.0 == r.offset) {
        let _ = m;
        // Change 0623 extension: a read that begins at one member's local
        // header and reaches past the next member's local header covers a run
        // of members, not one member's span. Every other class is 0611's, so a
        // leg that issues no run read reproduces 0611's counts byte for byte.
        let members_covered = g.members.iter().filter(|m| m.0 >= r.offset && m.0 + 30 <= r.offset + r.len).count();
        if members_covered > 1 { return "run-span"; }
        return if r.len == 30 { "local-fixed(30)" } else if r.len <= 640 { "local-window(<=640)" } else { "member-span" };
    }
    if r.len == 16 || r.len == 24 { return "descriptor"; }
    "payload"
}
static LAST: Mutex<(u64, u64, u64)> = Mutex::new((0, 0, 0));
fn summarize(g: &Geometry, label: &str, reqs: &[Req]) {
    let (a, b) = allocs_snapshot(); let v = VERSION_CALLS.load(Ordering::SeqCst);
    let mut last = LAST.lock().unwrap();
    let (da, db, dv) = (a - last.0, b - last.1, v - last.2);
    *last = (a, b, v); drop(last);
    println!("  [{label}] allocs={da} alloc_bytes={db} version_calls={dv}");
    let mut counts: Vec<(&str, u64, u64)> = Vec::new();
    for r in reqs { let c = classify(g, r); match counts.iter_mut().find(|e| e.0 == c) { Some(e) => { e.1 += 1; e.2 += r.returned; } None => counts.push((c, 1, r.returned)) } }
    let total: u64 = reqs.iter().map(|r| r.returned).sum();
    println!("  {label}: requests={} bytes={} :: {}", reqs.len(), total,
        counts.iter().map(|(c, n, b)| format!("{c}={n}/{b}B")).collect::<Vec<_>>().join(" "));
}
fn open(src: Arc<Counting>) -> SourceBackedPackage {
    let dynamic: Arc<dyn ReadAt> = src;
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        dynamic, ReadLimits::default(), SourceCacheLimits::default(), SourceReadPolicy::exact()).expect("open")
}
fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let path = &args[0]; let part = &args[1];
    let bytes = fs::read(path).expect("fixture");
    let g = geometry(&bytes);
    println!("fixture={path} bytes={} members={} cd_offset={} cd_size={} eocd={}", bytes.len(), g.members.len(), g.cd_offset, g.cd_size, g.eocd_offset);
    let src = Arc::new(Counting::new(bytes.clone()));
    let pkg = open(Arc::clone(&src));
    summarize(&g, "open", &src.take());
    let uri = PackURI::new(part.as_str()).expect("uri");
    let data = pkg.part(&uri).expect("part").data().expect("data");
    summarize(&g, &format!("first read of {part} ({} B decoded)", data.as_bytes().len()), &src.take());
    let data = pkg.part(&uri).expect("part").data().expect("data");
    let _ = data;
    summarize(&g, "second read of same part", &src.take());
    { let (a0, b0) = allocs_snapshot(); let d = flate2::read::DeflateDecoder::new(&b""[..]); let (a1, b1) = allocs_snapshot(); drop(d);
      println!("  [flate2 DeflateDecoder::new alone] allocs={} alloc_bytes={}", a1 - a0, b1 - b0); }
    if std::env::var("PROBE_PHASES").as_deref() == Ok("open") { return; }
    // all parts, physical order, fresh package
    let names: Vec<String> = pkg.iter_parts().map(|p| p.partname().to_string()).collect();
    let mut ordered: Vec<(u64, String)> = names.iter().map(|n| {
        let member = n.trim_start_matches('/');
        let off = g.members.iter().find(|m| m.1 == member).map(|m| m.0).unwrap_or(u64::MAX);
        (off, n.clone())
    }).collect();
    ordered.sort();
    let src2 = Arc::new(Counting::new(bytes.clone()));
    let pkg2 = open(Arc::clone(&src2)); src2.take();
    for (_, n) in &ordered { let _ = pkg2.part(&PackURI::new(n.as_str()).unwrap()).unwrap().data().unwrap(); }
    summarize(&g, &format!("all {} parts, physical order (after open)", ordered.len()), &src2.take());
    let src3 = Arc::new(Counting::new(bytes.clone()));
    let pkg3 = open(Arc::clone(&src3)); src3.take();
    for (_, n) in ordered.iter().rev() { let _ = pkg3.part(&PackURI::new(n.as_str()).unwrap()).unwrap().data().unwrap(); }
    summarize(&g, &format!("all {} parts, reverse physical order (after open)", ordered.len()), &src3.take());
    // streaming read of the one part (with_verified_decoded_reader path)
    let src4 = Arc::new(Counting::new(bytes));
    let pkg4 = open(Arc::clone(&src4)); src4.take();
    let mut sink = Vec::new();
    let n = pkg4.part(&uri).unwrap().stream_to(&mut sink).unwrap();
    summarize(&g, &format!("stream_to of {part} ({n} B)"), &src4.take());
}
