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
use std::io::Cursor;
use std::path::Path;
use std::time::Instant;

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
        let base = Package::from_reader(Cursor::new(&original)).unwrap();
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
        let target = Package::from_reader(Cursor::new(encode(authored.package()))).unwrap();
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
fn digest(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
fn validate_outputs(out: &Path) {
    std::fs::create_dir_all(out).unwrap();
    let fixture = Fixture::new();
    let forward = apply(&fixture.base, &fixture.forward);
    let inverse = apply(&fixture.target, &fixture.inverse);
    let noop = apply(&fixture.base, &fixture.noop);
    assert_eq!(forward.ink().unwrap().annotations().len(), 1);
    assert_eq!(inverse.ink().unwrap().annotations().len(), 0);
    assert_eq!(encode(&forward), encode(&fixture.target));
    assert_eq!(encode(&inverse), encode(&fixture.base));
    assert_eq!(encode(&noop), fixture.original);
    assert_eq!(encode(&fixture.base), fixture.original);
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
    ] {
        std::fs::write(out.join(name), bytes).unwrap();
    }
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
