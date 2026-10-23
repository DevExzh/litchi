#![allow(
    clippy::unwrap_used,
    reason = "the focused allocator regression uses deterministic fixtures"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{self, Write};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

use litchi_core::{OwnedSource, TextOutputOptions};
use litchi_odp::SourceBackedPresentation;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

struct Counting;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

// SAFETY: every operation forwards unchanged to the system allocator; counters
// only observe the successful allocation sizes.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: forwarding the caller-provided layout unchanged.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            let now = CURRENT.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            PEAK.fetch_max(now, Ordering::Relaxed);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: forwarding the original allocation and layout unchanged.
        unsafe { System.dealloc(pointer, layout) };
        CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        // SAFETY: forwarding the caller-provided allocation and layouts unchanged.
        let moved = unsafe { System.realloc(pointer, layout, size) };
        if !moved.is_null() {
            if size >= layout.size() {
                let delta = size - layout.size();
                let now = CURRENT.fetch_add(delta, Ordering::Relaxed) + delta;
                PEAK.fetch_max(now, Ordering::Relaxed);
            } else {
                CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

struct Sink;

impl Write for Sink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn package(slides: usize) -> Vec<u8> {
    let mut content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:presentation="{PRESENTATION}" xmlns:text="{TEXT}"><office:body><office:presentation>"#
    );
    for index in 0..slides {
        content.push_str(&format!(
            r#"<draw:page><draw:frame presentation:class="title"><draw:text-box><text:p>Title {index}</text:p></draw:text-box></draw:frame><draw:frame><draw:text-box><text:p>Body {index}</text:p></draw:text-box></draw:frame></draw:page>"#
        ));
    }
    content.push_str("</office:presentation></office:body></office:document-content>");

    let mut writer =
        litchi_odp::core::PackageWriter::new(litchi_odp::core::StreamingArchiveLimits::default());
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn write_peak(slides: usize) -> usize {
    let bytes = package(slides);
    let owner = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let start = CURRENT.load(Ordering::Relaxed);
    PEAK.store(start, Ordering::Relaxed);
    owner
        .write_text_to(&mut Sink, TextOutputOptions::default())
        .unwrap();
    PEAK.load(Ordering::Relaxed) - start
}

#[test]
fn source_text_event_buffer_peak_does_not_follow_document_size() {
    let small = write_peak(256);
    let large = write_peak(2048);

    // The sink keeps only one slide and the parser clears its event buffer on
    // every read. The larger document may need more transient parser state,
    // but retaining every event would add roughly one full content.xml here.
    assert!(
        large < 2 * 1024 * 1024,
        "unexpected retained event data: {large}"
    );
    assert!(
        large < small.saturating_mul(8).saturating_add(512 * 1024),
        "peak grew with event count: small={small}, large={large}"
    );
}
