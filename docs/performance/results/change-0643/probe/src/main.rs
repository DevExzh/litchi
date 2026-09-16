//! Change 0643 isolation probe: DOCX main-document read operations.
//!
//! Derived from the change-0592 probe. Two additions:
//!
//! * the fixture may be a real `.docx` on disk, opened through the same
//!   `detect_format_smart_with_limits` + `Package::from_opc_package` pair the
//!   perf-baseline filesystem runner's `PreparedDocx::eager` uses, so the
//!   probe can be pointed at the very corpus file the harness generated;
//! * three `*_prepared_*` operations differ only in what the *untimed*
//!   preparation does, which is what change 0592's open regression turns on.
//!
//! Counting mode: `probe0643 <operation> <fixture> <reps>` repeats one
//! operation `reps` times after the preparation and prints the allocation
//! counters. `reps` may be 0, so an isolation pair at 0 and 1 prices the
//! *first* call (including whatever it pays for being first), while a pair at
//! 4 and 20 prices the steady-state call. Instructions come from callgrind
//! over the same pair.
//!
//! Timing mode: `probe0643 timed <operation> <fixture> <samples>` prints one
//! nanosecond figure per line for `samples` repeats of the operation after the
//! preparation.
//!
//! `<fixture>` is either an integer (a synthetic package with that many
//! paragraphs, exactly the change-0592 fixture) or a path to a `.docx`.

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io::Cursor;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

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

/// Resolve the fixture argument: an integer builds the synthetic package, a
/// path reads the file.
fn fixture_bytes(fixture: &str) -> (Vec<u8>, usize, bool) {
    if let Ok(paragraphs) = fixture.parse::<usize>() {
        (package_bytes(paragraphs), paragraphs, false)
    } else {
        let bytes = std::fs::read(fixture).expect("fixture file");
        (bytes, 0, true)
    }
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

/// Open the eager package exactly as `PreparedDocx::eager` does when the
/// fixture is a real file, and as the change-0592 probe did otherwise.
fn eager_package(bytes: &[u8], from_file: bool) -> litchi_docx::Package {
    if from_file {
        let detected = litchi::detection_smart::detect_format_smart_with_limits(
            bytes.to_vec(),
            litchi_docx::ReadLimits::default(),
        )
        .expect("DOCX eager detector did not identify a supported package");
        match detected {
            litchi::detection_smart::DetectedFormat::Docx(opc) => {
                litchi_docx::Package::from_opc_package(opc).expect("docx package")
            },
            _ => panic!("DOCX eager detector returned a non-DOCX package"),
        }
    } else {
        litchi_docx::Package::from_reader(Cursor::new(bytes.to_vec())).unwrap()
    }
}

fn run_once(
    operation: &str,
    eager: &litchi_docx::Package,
    source: &litchi_docx::source_backed::Package,
    selected: usize,
) {
    match operation {
        // `document()` + `paragraph_count()`: the perf-baseline
        // `docx_file_eager_paragraph_count` timed region, verbatim.
        "eager_prepared_paragraph_count"
        | "eager_prepared_warm_paragraph_count"
        | "eager_noprep_paragraph_count"
        | "eager_paragraph_count" => {
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
        // `text()` on a view built once, so the two text rows differ only in
        // the projection, not in the view construction.
        "eager_view_text" => {
            let document = eager.document().unwrap();
            black_box(document.text().unwrap());
        },
        "eager_paragraph_first" => {
            let document = eager.document().unwrap();
            black_box(document.paragraph(selected).unwrap().unwrap());
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
        // The source-backed sink entry point lives on the package, not on
        // the pinned document view.
        "source_write_text" => {
            let mut sink = DiscardSink(0);
            black_box(
                source
                    .write_text_to(&mut sink, litchi_core::TextOutputOptions::default())
                    .unwrap(),
            );
            black_box(sink.0);
        },
        "source_paragraph_count" => {
            let document = source.document().unwrap();
            black_box(document.paragraph_count().unwrap());
        },
        other => panic!("unknown operation {other}"),
    }
}

/// The untimed preparation each operation runs before its measured repeats.
///
/// `eager_prepared_paragraph_count` mirrors `PreparedDocx::eager`, which
/// extracts the whole text once before the harness starts its clock. Before
/// change 0592 that preparation also built and discarded a paragraph index,
/// warming the scan the measured call then re-enters; after 0592 it does not.
/// `eager_prepared_warm_paragraph_count` is the control that removes exactly
/// that asymmetry: its preparation builds an index on *both* legs.
/// `eager_noprep_paragraph_count` is the control with no preparation at all.
fn prepare(operation: &str, eager: &litchi_docx::Package) {
    match operation {
        "eager_prepared_paragraph_count" => {
            black_box(eager.document().unwrap().text().unwrap());
        },
        "eager_prepared_warm_paragraph_count" => {
            black_box(eager.document().unwrap().text().unwrap());
            black_box(eager.document().unwrap().paragraph_count().unwrap());
        },
        _ => {},
    }
}

fn main() {
    let mut args = std::env::args().skip(1);
    let first = args.next().expect("mode or operation");
    let timed = first == "timed";
    let operation = if timed {
        args.next().expect("operation")
    } else {
        first
    };
    let fixture = args.next().expect("fixture");
    let reps: usize = args.next().expect("reps").parse().unwrap();

    let (bytes, paragraphs, from_file) = fixture_bytes(&fixture);
    let selected = paragraphs / 2;
    let eager = eager_package(&bytes, from_file);
    let source =
        litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
            .unwrap();

    prepare(&operation, &eager);

    if timed {
        let mut samples = Vec::with_capacity(reps);
        for _ in 0..reps {
            let started = Instant::now();
            run_once(&operation, &eager, &source, selected);
            samples.push(started.elapsed().as_nanos() as u64);
        }
        for sample in samples {
            println!("{sample}");
        }
        return;
    }

    for _ in 0..reps {
        run_once(&operation, &eager, &source, selected);
    }

    println!(
        "operation={operation} fixture={fixture} reps={reps} archive_bytes={} allocations={} allocated_bytes={}",
        bytes.len(),
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
    );
}
