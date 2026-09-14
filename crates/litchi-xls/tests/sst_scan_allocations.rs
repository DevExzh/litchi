//! Allocation evidence for the open-time shared-string scan.
//!
//! The scan indexes every shared string in a workbook so that a later cell
//! lookup can fetch one entry. It used to obtain each `(start, end)` boundary by
//! decoding the string into a `String` and dropping it, which cost two heap
//! allocations per shared string. This test pins the property that replaced
//! that: one open allocates fewer times than the workbook has shared strings.

#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_cfb::SharedOleFile;
use litchi_core::{OwnedSource, ReadAt};
use litchi_xls::SourceBackedWorkbook;

// The counters are per-thread, not process-wide, so a measurement is unaffected
// by whatever the other tests in this binary are doing on their own threads.
// Both are const-initialised and hold no destructor, so reading them from inside
// the allocator cannot allocate or recurse.
thread_local! {
    static ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
    static ALLOCATED_BYTES: Cell<usize> = const { Cell::new(0) };
}

fn record(bytes: usize) {
    let _ = ALLOCATIONS.try_with(|cell| cell.set(cell.get().wrapping_add(1)));
    let _ = ALLOCATED_BYTES.try_with(|cell| cell.set(cell.get().wrapping_add(bytes)));
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: Every operation delegates to the platform allocator after recording
// only counts and sizes; the allocator does not alter pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        record(layout.size());
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc(layout) }
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        record(layout.size());
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc_zeroed(layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        record(new_size.saturating_sub(layout.size()));
        // SAFETY: `pointer` and `layout` are exactly the pair returned by a
        // previous allocation delegated to `System`.
        unsafe { System.realloc(pointer, layout, new_size) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: `pointer` and `layout` are exactly the pair returned by a
        // previous allocation delegated to `System`.
        unsafe { System.dealloc(pointer, layout) }
    }
}

#[derive(Debug, Clone, Copy)]
struct Counts {
    allocations: usize,
    bytes: usize,
}

fn measure<T>(body: impl FnOnce() -> T) -> (T, Counts) {
    let allocations = ALLOCATIONS.with(Cell::get);
    let bytes = ALLOCATED_BYTES.with(Cell::get);
    let value = std::hint::black_box(body());
    let counts = Counts {
        allocations: ALLOCATIONS.with(Cell::get) - allocations,
        bytes: ALLOCATED_BYTES.with(Cell::get) - bytes,
    };
    (value, counts)
}

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// Reads the `SST` record's declared unique string count out of the workbook
/// globals by hand, so the bound below is the fixture's own property rather
/// than a number copied into the test.
fn unique_shared_strings(bytes: Vec<u8>) -> u32 {
    let file = SharedOleFile::open(Arc::new(OwnedSource::new(bytes))).expect("the fixture opens");
    let stream = file
        .open_stream(&["Workbook"])
        .expect("the fixture has a workbook stream");
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x00FC {
            let payload = offset + 4;
            return u32::from_le_bytes([
                stream[payload + 4],
                stream[payload + 5],
                stream[payload + 6],
                stream[payload + 7],
            ]);
        }
        assert_ne!(kind, 0x000A, "the globals substream ended before its SST");
        offset += 4 + length;
    }
    panic!("the fixture has no SST record");
}

/// One source-backed open allocates fewer times than the workbook has unique
/// shared strings.
///
/// Before the measure-only walk landed, the open-time scan decoded every shared
/// string and dropped it, which is two allocations per string in the scan alone,
/// so this bound was exceeded by roughly an order of magnitude.
#[test]
fn opening_a_string_heavy_workbook_allocates_less_than_once_per_shared_string() {
    let path = fixture("poi/test-data/spreadsheet/54016.xls");
    let bytes = std::fs::read(&path).expect("the fixture is readable");
    let unique = unique_shared_strings(bytes.clone());
    assert_eq!(unique, 7_893, "the fixture's SST changed shape");

    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    let (workbook, counts) = measure(|| SourceBackedWorkbook::from_read_at(Arc::clone(&source)));
    let workbook = workbook.expect("the fixture opens");
    assert!(
        !workbook
            .worksheet_names()
            .expect("the fixture lists worksheets")
            .is_empty()
    );

    println!(
        "54016.xls open: allocations={} bytes={} unique_strings={unique}",
        counts.allocations, counts.bytes
    );
    assert!(
        counts.allocations < unique as usize,
        "one open allocated {} times for {unique} shared strings; the scan is materializing them again",
        counts.allocations
    );
}

/// The same property on a workbook whose shared strings are long rather than
/// numerous: 474 strings across 101,125 bytes of `SST`.
#[test]
fn opening_a_long_string_workbook_allocates_less_than_once_per_shared_string() {
    let path = fixture("ole/xls/WithCustomViews.xls");
    let bytes = std::fs::read(&path).expect("the fixture is readable");
    let unique = unique_shared_strings(bytes.clone());
    assert_eq!(unique, 474, "the fixture's SST changed shape");

    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    let (workbook, counts) = measure(|| SourceBackedWorkbook::from_read_at(Arc::clone(&source)));
    let workbook = workbook.expect("the fixture opens");
    assert!(
        !workbook
            .worksheet_names()
            .expect("the fixture lists worksheets")
            .is_empty()
    );

    println!(
        "WithCustomViews.xls open: allocations={} bytes={} unique_strings={unique}",
        counts.allocations, counts.bytes
    );
    assert!(
        counts.allocations < unique as usize,
        "one open allocated {} times for {unique} shared strings; the scan is materializing them again",
        counts.allocations
    );
}

/// Extracting text still decodes shared strings, so the materializing path is
/// intact and is still where decoding is paid for.
#[test]
fn extracting_text_still_decodes_shared_strings() {
    let path = fixture("ole/xls/WithCustomViews.xls");
    let bytes = std::fs::read(&path).expect("the fixture is readable");
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    let workbook = SourceBackedWorkbook::from_read_at(source).expect("the fixture opens");

    let text = workbook.text().expect("the fixture extracts text");
    assert!(
        text.chars().any(char::is_alphabetic),
        "the fixture should still yield decoded shared-string text"
    );
}
