//! Change 0680 source probe: current DOCX paragraph-index ownership.
//!
//! The probe keeps package construction and opening outside the measured loop.
//! `*_document_count` creates a fresh semantic document view for every
//! iteration; `*_one_view_count` reuses one view.  The difference prices the
//! work that is repeated by the current per-view `OnceLock` layout.  The
//! counting allocator reports deterministic allocation deltas; no latency or
//! speedup claim is inferred from these numbers.
//!
//! Usage: `probe0680 <operation> <paragraphs> <repetitions>`

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io::Cursor;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::OwnedSource;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

static ALLOCATIONS: AtomicU64 = AtomicU64::new(0);
static ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);

struct CountingAllocator;

// SAFETY: every operation delegates to the system allocator. Relaxed counters
// are observational only and never affect returned pointers.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn document_xml(paragraphs: usize) -> Vec<u8> {
    let mut xml = format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body>"#);
    for index in 0..paragraphs {
        xml.push_str(&format!(
            "<w:p><w:r><w:t>litchi-0680-{index:05}</w:t></w:r></w:p>"
        ));
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn package_bytes(paragraphs: usize) -> (Vec<u8>, usize) {
    let xml = document_xml(paragraphs);
    let xml_len = xml.len();
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").expect("document URI"),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        xml,
    )));
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    (
        PackageWriter::to_bytes(&package).expect("package bytes"),
        xml_len,
    )
}

fn main() {
    let mut args = std::env::args().skip(1);
    let operation = args.next().expect("operation");
    let paragraphs: usize = args.next().expect("paragraph count").parse().unwrap();
    let repetitions: usize = args.next().expect("repetitions").parse().unwrap();
    let (bytes, xml_len) = package_bytes(paragraphs);
    let archive_len = bytes.len();
    let mut source_diagnostics = None;

    match operation.as_str() {
        "eager_document" | "eager_document_count" | "eager_one_view_count" => {
            let package = litchi_docx::Package::from_reader(Cursor::new(bytes)).unwrap();
            if operation == "eager_one_view_count" {
                let document = package.document().unwrap();
                let before = counters();
                for _ in 0..repetitions {
                    assert_eq!(document.paragraph_count().unwrap(), paragraphs);
                }
                let after = counters();
                print_sampled(before, after);
            } else {
                let before = counters();
                for _ in 0..repetitions {
                    let document = package.document().unwrap();
                    if operation == "eager_document_count" {
                        assert_eq!(document.paragraph_count().unwrap(), paragraphs);
                    } else {
                        black_box(document);
                    }
                }
                let after = counters();
                print_sampled(before, after);
            }
        },
        "source_document" | "source_document_count" | "source_one_view_count" => {
            let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(
                OwnedSource::new(bytes),
            ))
            .unwrap();
            if operation == "source_one_view_count" {
                let document = package.document().unwrap();
                let before = counters();
                for _ in 0..repetitions {
                    assert_eq!(document.paragraph_count().unwrap(), paragraphs);
                }
                let after = counters();
                print_sampled(before, after);
            } else {
                let before = counters();
                for _ in 0..repetitions {
                    let document = package.document().unwrap();
                    if operation == "source_document_count" {
                        assert_eq!(document.paragraph_count().unwrap(), paragraphs);
                    } else {
                        black_box(document);
                    }
                }
                let after = counters();
                print_sampled(before, after);
            }
            let diagnostics = package.cache_diagnostics();
            source_diagnostics = Some((
                diagnostics.cold_loads,
                diagnostics.hits,
                diagnostics.successful_loads,
                diagnostics.retained_entries,
                diagnostics.retained_bytes,
            ));
        },
        other => panic!("unknown operation {other}"),
    }

    if let Some((cold_loads, hits, successful_loads, retained_entries, retained_bytes)) =
        source_diagnostics
    {
        println!(
            "source_cache cold_loads={cold_loads} hits={hits} successful_loads={successful_loads} retained_entries={retained_entries} retained_bytes={retained_bytes}"
        );
    }
    println!(
        "operation={operation} paragraphs={paragraphs} xml_bytes={xml_len} reps={repetitions} archive_bytes={archive_len}"
    );
}

fn counters() -> (u64, u64) {
    (
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
    )
}

fn print_sampled(before: (u64, u64), after: (u64, u64)) {
    println!(
        "sampled allocations={} allocated_bytes={}",
        after.0.saturating_sub(before.0),
        after.1.saturating_sub(before.1),
    );
}
