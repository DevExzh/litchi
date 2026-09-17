//! Counting-allocator companion for the change 0672 probe.
//!
//! `xlsx0672_alloc visit|cells FILE SHEET` reports the allocation calls,
//! allocated bytes, peak live-byte delta and retained live-byte delta of one
//! whole-sheet operation on a source-backed worksheet. The allocator shape
//! follows the change 0634 driver's `alloc.rs`.

use std::alloc::{GlobalAlloc, Layout, System};
use std::error::Error;
use std::hint::black_box;
use std::path::Path;
use std::sync::atomic::{AtomicUsize, Ordering};

type BoxError = Box<dyn Error>;

/// The whole-sheet area used by every mode, as in `probe.rs`.
const WHOLE_SHEET: &str = "A1:XFD1048576";

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_BYTES: AtomicUsize = AtomicUsize::new(0);

struct Counting;

impl Counting {
    fn record(size: usize) {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
        PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        Self::record(layout.size());
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE_BYTES.fetch_sub(layout.size(), Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        if new_size > layout.size() {
            ALLOCATED_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed);
            let live = LIVE_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed)
                + (new_size - layout.size());
            PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
        } else {
            LIVE_BYTES.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
        }
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn reset() {
    ALLOCATIONS.store(0, Ordering::Relaxed);
    ALLOCATED_BYTES.store(0, Ordering::Relaxed);
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
}

fn report(mode: &str, path: &Path, sheet: &str, visited: usize, baseline: usize, retained: usize) {
    println!(
        "{mode}\t{}\t{sheet}\tvisited={visited}\tallocations={}\tallocated_bytes={}\tpeak_live_delta={}\tretained_live_delta={}",
        path.display(),
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
        PEAK_BYTES.load(Ordering::Relaxed).saturating_sub(baseline),
        retained.saturating_sub(baseline),
    );
}

fn main() -> Result<(), BoxError> {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    let mode = arguments[0].clone();
    let path = std::path::PathBuf::from(&arguments[1]);
    let sheet = arguments[2].clone();

    // One untimed warm-up so lazily-initialized globals are not attributed.
    {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(&path)?;
        let worksheet = workbook
            .sheet(sheet.as_str())?
            .ok_or("probe worksheet is missing from this workbook")?;
        let visited = worksheet.visit_cells(WHOLE_SHEET, |_address, cell| {
            black_box(cell);
            Ok(())
        })?;
        black_box(visited);
    }

    match mode.as_str() {
        "visit" => {
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(&path)?;
            let worksheet = workbook
                .sheet(sheet.as_str())?
                .ok_or("probe worksheet is missing from this workbook")?;
            let mut sum = 0u64;
            let visited = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
                sum = sum.wrapping_add(u64::from(address.row().get()));
                black_box(cell);
                Ok(())
            })?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, &sheet, visited, baseline, retained);
            black_box(sum);
            black_box(&workbook);
        },
        "cells" => {
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(&path)?;
            let worksheet = workbook
                .sheet(sheet.as_str())?
                .ok_or("probe worksheet is missing from this workbook")?;
            let values = worksheet.cells(WHOLE_SHEET)?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, &sheet, values.len(), baseline, retained);
            black_box(&values);
            black_box(&workbook);
        },
        "visit-warm" | "cells-warm" => {
            // Materialize the worksheet store before the measured window, so
            // only the walk over an already-parsed sheet is attributed.
            let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(&path)?;
            let worksheet = workbook
                .sheet(sheet.as_str())?
                .ok_or("probe worksheet is missing from this workbook")?;
            let _extent = worksheet.stored_extent()?;
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            if mode == "visit-warm" {
                let mut sum = 0u64;
                let visited = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
                    sum = sum.wrapping_add(u64::from(address.row().get()));
                    black_box(cell);
                    Ok(())
                })?;
                let retained = LIVE_BYTES.load(Ordering::Relaxed);
                report(&mode, &path, &sheet, visited, baseline, retained);
                black_box(sum);
            } else {
                let values = worksheet.cells(WHOLE_SHEET)?;
                let retained = LIVE_BYTES.load(Ordering::Relaxed);
                report(&mode, &path, &sheet, values.len(), baseline, retained);
                black_box(&values);
            }
            black_box(&workbook);
        },
        other => return Err(format!("unknown mode {other}").into()),
    }
    Ok(())
}
