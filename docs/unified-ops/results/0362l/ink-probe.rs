//! Matched durable Ink replay probe; `before` adapts only Commit extraction.
use litchi_core::{
    Position,
    patch::{BlobLimits, Patch as CorePatch, PatchLimits, Reversible},
};
use litchi_docx::{
    Package,
    ink::{BaseProfile, Destination, Draft, Style},
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use sha2::{Digest, Sha256};
use std::hint::black_box;
use std::io::{self, Cursor, Read, Write};
use std::path::Path;
use std::time::Instant;

const READ_CHUNK: usize = 4096;
const SINK_WINDOW: usize = 8192;

/// A reader with a fixed read window.  It is used only while preparing the
/// fixture, so the timed operation never includes input construction.  The
/// counters make the preflight prove that the replay path accepts a bounded
/// streaming reader rather than relying on one bulk read.
struct ChunkedReader<'a> {
    input: Cursor<&'a [u8]>,
    reads: usize,
    bytes: usize,
}

impl<'a> ChunkedReader<'a> {
    fn new(input: &'a [u8]) -> Self {
        Self {
            input: Cursor::new(input),
            reads: 0,
            bytes: 0,
        }
    }
}

impl Read for ChunkedReader<'_> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let end = buffer.len().min(READ_CHUNK);
        let count = self.input.read(&mut buffer[..end])?;
        self.reads += usize::from(count != 0);
        self.bytes += count;
        Ok(count)
    }
}

/// A fixed-state sink for publication preflight.  It retains only the hash
/// state and counters; output bytes never accumulate in a `Vec`.
struct DigestSink {
    digest: Sha256,
    bytes: usize,
    writes: usize,
    max_write: usize,
}

impl DigestSink {
    fn new() -> Self {
        Self {
            digest: Sha256::new(),
            bytes: 0,
            writes: 0,
            max_write: 0,
        }
    }

    fn finish(self) -> StreamEvidence {
        let hash = self.digest.finalize();
        StreamEvidence {
            bytes: self.bytes,
            writes: self.writes,
            max_write: self.max_write,
            sha256: hex_digest(hash.as_slice()),
        }
    }
}

impl Write for DigestSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(SINK_WINDOW);
        self.digest.update(&bytes[..count]);
        self.bytes += count;
        self.writes += 1;
        self.max_write = self.max_write.max(count);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone)]
struct StreamEvidence {
    bytes: usize,
    writes: usize,
    max_write: usize,
    sha256: String,
}

fn open_bounded(bytes: &[u8]) -> (Package, usize, usize) {
    let mut reader = ChunkedReader::new(bytes);
    let package = Package::from_reader(&mut reader).unwrap();
    (package, reader.reads, reader.bytes)
}

fn stream(package: &Package) -> StreamEvidence {
    let mut sink = DigestSink::new();
    package.write_to(&mut sink).unwrap();
    sink.finish()
}

#[cfg(feature = "allocations")]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub static CURRENT: AtomicUsize = AtomicUsize::new(0);
    pub static PEAK: AtomicUsize = AtomicUsize::new(0);
    pub static BYTES: AtomicUsize = AtomicUsize::new(0);
    pub static CALLS: AtomicUsize = AtomicUsize::new(0);

    pub struct Counting;

    // SAFETY: every operation forwards to the system allocator with the
    // caller's layout and pointer; the atomics only observe sizes.
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: the system allocator receives the original layout.
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                add(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: the pointer and layout came from this allocator.
            unsafe { System.dealloc(pointer, layout) };
            CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
            // SAFETY: the pointer and old layout came from this allocator.
            let moved = unsafe { System.realloc(pointer, layout, size) };
            if !moved.is_null() {
                let live = if size >= layout.size() {
                    CURRENT.fetch_add(size - layout.size(), Ordering::Relaxed) + size
                        - layout.size()
                } else {
                    CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed)
                        - (layout.size() - size)
                };
                PEAK.fetch_max(live, Ordering::Relaxed);
                BYTES.fetch_add(size, Ordering::Relaxed);
                CALLS.fetch_add(1, Ordering::Relaxed);
            }
            moved
        }
    }

    fn add(size: usize) {
        let live = CURRENT.fetch_add(size, Ordering::Relaxed) + size;
        PEAK.fetch_max(live, Ordering::Relaxed);
        BYTES.fetch_add(size, Ordering::Relaxed);
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    pub fn start() -> usize {
        let live = CURRENT.load(Ordering::Relaxed);
        PEAK.store(live, Ordering::Relaxed);
        BYTES.store(0, Ordering::Relaxed);
        CALLS.store(0, Ordering::Relaxed);
        live
    }

    pub fn result(base: usize) -> (usize, usize, usize, usize) {
        (
            PEAK.load(Ordering::Relaxed).saturating_sub(base),
            CURRENT.load(Ordering::Relaxed).saturating_sub(base),
            BYTES.load(Ordering::Relaxed),
            CALLS.load(Ordering::Relaxed),
        )
    }
}

#[cfg(feature = "allocations")]
#[global_allocator]
static ALLOCATOR: allocation::Counting = allocation::Counting;

fn patch_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(16, 64 * 1024 * 1024, 128 * 1024 * 1024),
        32 * 1024 * 1024,
        64,
        32,
        1024 * 1024,
        64 * 1024 * 1024,
    )
}

fn fixture() -> Vec<u8> {
    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut opc = OpcPackage::new();
    let mut document = format!(r#"<w:document xmlns:w="{WORD}" xmlns:r="{REL}"><w:body>"#);
    for index in 0..1000 {
        document.push_str(&format!(
            "<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"
        ));
    }
    document.push_str("</w:body></w:document>");
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
        document.into_bytes(),
    )));
    opc.relate_to("word/document.xml", &format!("{REL}/officeDocument"));
    PackageWriter::to_bytes(&opc).unwrap()
}

fn encode(package: &Package) -> Vec<u8> {
    let mut out = Cursor::new(Vec::new());
    package.write_to(&mut out).unwrap();
    out.into_inner()
}

#[cfg(feature = "before")]
fn apply(package: &Package, patch: &CorePatch<Reversible>) -> Package {
    package.apply_durable_ink(patch).unwrap()
}
#[cfg(not(feature = "before"))]
fn apply(package: &Package, patch: &CorePatch<Reversible>) -> Package {
    package.apply_durable_ink(patch).unwrap().into_package()
}

struct Fixture {
    original: Vec<u8>,
    base: Package,
    target: Package,
    forward: CorePatch<Reversible>,
    inverse: CorePatch<Reversible>,
    noop: CorePatch<Reversible>,
}
impl Fixture {
    fn new() -> Self {
        let original = fixture();
        let (base, _, _) = open_bounded(&original);
        let mut edit = base.edit_ink().unwrap();
        edit.insert(
            Destination::main(Position::new(500)),
            Draft::default().finish().unwrap(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
        let authored = edit.commit().unwrap();
        let wire = authored
            .patch()
            .to_durable(patch_limits())
            .unwrap()
            .to_deterministic_json()
            .unwrap();
        let forward =
            CorePatch::<Reversible>::from_deterministic_json(&wire, patch_limits()).unwrap();
        let inverse = forward.inverse();
        let target_bytes = encode(authored.package());
        let (target, _, _) = open_bounded(&target_bytes);
        let noop = base
            .edit_ink()
            .unwrap()
            .commit()
            .unwrap()
            .patch()
            .to_durable(patch_limits())
            .unwrap();
        Self {
            original,
            base,
            target,
            forward,
            inverse,
            noop,
        }
    }
    fn route(&self, name: &str) -> (&Package, &CorePatch<Reversible>, usize) {
        match name {
            "forward" | "forward_write" => (&self.base, &self.forward, 1),
            "inverse" | "inverse_write" => (&self.target, &self.inverse, 0),
            "noop" | "noop_write" => (&self.base, &self.noop, 0),
            _ => panic!("unknown route {name}"),
        }
    }
}
fn hex_digest(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn digest(bytes: &[u8]) -> String {
    hex_digest(Sha256::digest(bytes).as_slice())
}
fn validate_outputs(out: &Path) {
    std::fs::create_dir_all(out).unwrap();
    let fixture = Fixture::new();
    let forward = apply(&fixture.base, &fixture.forward);
    let inverse = apply(&fixture.target, &fixture.inverse);
    let noop = apply(&fixture.base, &fixture.noop);
    let (matching_base, base_reads, base_bytes) = open_bounded(&fixture.original);
    let target_bytes = encode(&fixture.target);
    let (matching_target, target_reads, target_read_bytes) = open_bounded(&target_bytes);
    assert!(base_reads > 0 && target_reads > 0);
    assert_eq!(base_bytes, fixture.original.len());
    assert_eq!(target_read_bytes, target_bytes.len());
    let matching_forward = apply(&matching_base, &fixture.forward);
    let matching_inverse = apply(&matching_target, &fixture.inverse);
    let matching_noop = apply(&matching_base, &fixture.noop);
    assert_eq!(forward.ink().unwrap().annotations().len(), 1);
    assert_eq!(inverse.ink().unwrap().annotations().len(), 0);
    assert_eq!(encode(&forward), encode(&fixture.target));
    assert_eq!(encode(&inverse), encode(&fixture.base));
    assert_eq!(encode(&noop), fixture.original);
    assert_eq!(encode(&fixture.base), fixture.original);
    assert_eq!(encode(&matching_forward), target_bytes);
    assert_eq!(encode(&matching_inverse), fixture.original);
    assert_eq!(encode(&matching_noop), fixture.original);
    let streamed = [
        ("forward", &forward),
        ("inverse", &inverse),
        ("noop", &noop),
    ]
    .into_iter()
    .map(|(name, package)| {
        let encoded = encode(package);
        let evidence = stream(package);
        assert_eq!(evidence.bytes, encoded.len());
        assert_eq!(evidence.sha256, digest(&encoded));
        assert!(evidence.writes > 0);
        assert!(evidence.max_write <= SINK_WINDOW);
        (name, encoded.len(), evidence)
    })
    .collect::<Vec<_>>();
    for (name, bytes) in [
        ("source.docx", fixture.original),
        ("forward.docx", encode(&forward)),
        ("inverse.docx", encode(&inverse)),
        ("noop.docx", encode(&noop)),
        (
            "forward.json",
            fixture.forward.to_deterministic_json().unwrap(),
        ),
        (
            "inverse.json",
            fixture.inverse.to_deterministic_json().unwrap(),
        ),
        ("noop.json", fixture.noop.to_deterministic_json().unwrap()),
    ] {
        std::fs::write(out.join(name), bytes).unwrap();
    }
    let mut streaming =
        String::from("{\"schema\":\"litchi.docx-durable-ink-stream.v1\",\"read_chunk\":");
    streaming.push_str(&READ_CHUNK.to_string());
    streaming.push_str(",\"sink_window\":");
    streaming.push_str(&SINK_WINDOW.to_string());
    streaming.push_str(&format!(
        ",\"base_reads\":{base_reads},\"base_bytes\":{base_bytes},\"target_reads\":{target_reads},\"target_bytes\":{target_read_bytes},\"routes\":["
    ));
    for (index, (name, bytes, evidence)) in streamed.iter().enumerate() {
        if index != 0 {
            streaming.push(',');
        }
        streaming.push_str(&format!(
            "{{\"route\":\"{name}\",\"bytes\":{bytes},\"writes\":{},\"max_write\":{},\"sha256\":\"{}\"}}",
            evidence.writes, evidence.max_write, evidence.sha256
        ));
    }
    streaming.push_str("]}\n");
    std::fs::write(out.join("streaming.json"), streaming).unwrap();
}
fn main() {
    let mut args = std::env::args().skip(1);
    let route = args.next().unwrap();
    if route == "validate_outputs" {
        validate_outputs(Path::new(&args.next().unwrap()));
        return;
    }
    let iterations: usize = args.next().unwrap().parse().unwrap();
    let fixture = Fixture::new();
    let (source, patch, expected_annotations) = fixture.route(&route);
    let expected = apply(source, patch);
    assert_eq!(
        expected.ink().unwrap().annotations().len(),
        expected_annotations
    );
    let output = encode(&expected);
    let output_sha256 = digest(&output);
    let output_bytes = output.len();
    drop(expected);
    drop(output);
    let write = route.ends_with("_write");
    #[cfg(feature = "allocations")]
    let before = allocation::start();
    let started = Instant::now();
    let mut checksum = 0u64;
    for _ in 0..iterations {
        let package = apply(black_box(source), black_box(patch));
        let count = package.ink().unwrap().annotations().len();
        assert_eq!(count, expected_annotations);
        checksum = checksum.wrapping_add(count as u64);
        if write {
            let bytes = encode(black_box(&package));
            checksum = checksum.wrapping_add(bytes.len() as u64);
            black_box(bytes);
        }
        black_box(package);
    }
    let ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let (peak, retained, allocated, allocations) = allocation::result(before);
    #[cfg(not(feature = "allocations"))]
    let (peak, retained, allocated, allocations) = (0, 0, 0, 0);
    println!(
        "{{\"schema\":\"litchi.docx-durable-ink.v1\",\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{ns},\"checksum\":{checksum},\"output_bytes\":{output_bytes},\"output_sha256\":\"{output_sha256}\",\"peak\":{peak},\"retained\":{retained},\"allocated\":{allocated},\"allocations\":{allocations}}}"
    );
}
