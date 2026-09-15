//! Change 0592 isolation probe: DOCX main-document read operations.
//!
//! One invocation repeats a single operation `reps` times over a fixture that
//! is built once, outside the loop. Two invocations at `N` and `N + M` are
//! differenced (callgrind totals for instructions, the counting allocator for
//! allocations) so that process startup, fixture construction and the package
//! open are removed from the per-operation figure.
//!
//! Usage: `probe0592 <operation> <paragraphs> <reps>`

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

// SAFETY: every method forwards to the system allocator unchanged; the
// counters are plain relaxed atomics and never affect the returned pointers.
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

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// The same paragraph payload the perf-baseline semantic DOCX corpus writes.
fn paragraph_text(index: usize) -> String {
    format!("litchi-perf-baseline-docx-semantic-v1-source-{index:05}")
}

fn document_xml(paragraphs: usize) -> Vec<u8> {
    let mut xml = format!(r#"<w:document xmlns:w="{W}"><w:body>"#);
    for index in 0..paragraphs {
        xml.push_str(&format!(
            "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            paragraph_text(index)
        ));
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn package_bytes(paragraphs: usize) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_xml(paragraphs),
    )));
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

struct DiscardSink(u64);

impl std::io::Write for DiscardSink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.0 = self.0.wrapping_add(buf.len() as u64);
        Ok(buf.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn main() {
    let mut args = std::env::args().skip(1);
    let operation = args.next().expect("operation");
    let paragraphs: usize = args.next().expect("paragraphs").parse().unwrap();
    let reps: usize = args.next().expect("reps").parse().unwrap();

    let bytes = package_bytes(paragraphs);
    let selected = paragraphs / 2;
    let eager = litchi_docx::Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    let source =
        litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
            .unwrap();

    if operation == "eager_prepared_paragraph_count" {
        // Mirror the perf-baseline filesystem runner's `PreparedDocx::eager`,
        // which extracts the whole text once, outside the timer, before the
        // measured call. On the before leg that preparation also builds and
        // discards a paragraph index, warming the path the measured call then
        // re-enters; on the after leg nothing warms it.
        black_box(eager.document().unwrap().text().unwrap());
    }

    for _ in 0..reps {
        match operation.as_str() {
            // `document()` + `paragraph_count()` after that preparation.
            "eager_prepared_paragraph_count" => {
                let document = eager.document().unwrap();
                black_box(document.paragraph_count().unwrap());
            },
            // `document()` alone: the MCE pass plus (before 0592) the index.
            "eager_document" => {
                black_box(eager.document().unwrap());
            },
            // One full-text extraction through a fresh main-document view.
            "eager_text" => {
                let document = eager.document().unwrap();
                black_box(document.text().unwrap());
            },
            // One streaming text export through a fresh main-document view.
            "eager_write_text" => {
                let document = eager.document().unwrap();
                let mut sink = DiscardSink(0);
                black_box(
                    document
                        .write_text_to(&mut sink, litchi_core::TextOutputOptions::default())
                        .unwrap(),
                );
                black_box(sink.0);
            },
            // First paragraph query on a fresh view: pays for the index.
            "eager_paragraph_first" => {
                let document = eager.document().unwrap();
                black_box(document.paragraph(selected).unwrap().unwrap());
            },
            // First query plus seven repeats on the same view.
            "eager_paragraph_repeat8" => {
                let document = eager.document().unwrap();
                for _ in 0..8 {
                    black_box(document.paragraph(selected).unwrap().unwrap());
                }
            },
            "eager_paragraph_count" => {
                let document = eager.document().unwrap();
                black_box(document.paragraph_count().unwrap());
            },
            "eager_paragraphs" => {
                let document = eager.document().unwrap();
                black_box(document.paragraphs().unwrap());
            },
            "eager_tables" => {
                let document = eager.document().unwrap();
                black_box(document.tables().unwrap());
            },
            "source_document" => {
                black_box(source.document().unwrap());
            },
            "source_text" => {
                let document = source.document().unwrap();
                black_box(document.extract_text().unwrap());
            },
            "source_paragraph_first" => {
                let document = source.document().unwrap();
                black_box(document.paragraph(selected).unwrap().unwrap());
            },
            "source_paragraph_repeat8" => {
                let document = source.document().unwrap();
                for _ in 0..8 {
                    black_box(document.paragraph(selected).unwrap().unwrap());
                }
            },
            "source_paragraph_count" => {
                let document = source.document().unwrap();
                black_box(document.paragraph_count().unwrap());
            },
            other => panic!("unknown operation {other}"),
        }
    }

    println!(
        "operation={operation} paragraphs={paragraphs} reps={reps} archive_bytes={} allocations={} allocated_bytes={}",
        bytes.len(),
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
    );
}
