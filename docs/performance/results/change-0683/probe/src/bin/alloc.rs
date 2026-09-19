//! Counting allocator companion for the 0683 selected-record probe.

use std::alloc::{GlobalAlloc, Layout, System};
use std::error::Error;
use std::hint::black_box;
use std::path::Path;
use std::sync::atomic::{AtomicUsize, Ordering};

use xlsx0683_selected_record_probe::{PreparedWorksheet, cold_measure, error_callbacks};

static ALLOCATION_CALLS: AtomicUsize = AtomicUsize::new(0);
static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static DEALLOCATION_CALLS: AtomicUsize = AtomicUsize::new(0);
static DEALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);

struct Counting;

impl Counting {
    fn allocation(size: usize) {
        ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
        PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
    }

    fn deallocation(size: usize) {
        DEALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    }

    fn reset() {
        ALLOCATION_CALLS.store(0, Ordering::Relaxed);
        ALLOCATED_BYTES.store(0, Ordering::Relaxed);
        DEALLOCATION_CALLS.store(0, Ordering::Relaxed);
        DEALLOCATED_BYTES.store(0, Ordering::Relaxed);
        PEAK_LIVE_BYTES.store(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // Count only the request that the system allocator accepted. A null
        // allocation leaves live and peak accounting unchanged.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            Counting::allocation(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        Counting::deallocation(layout.size());
        unsafe { System.dealloc(pointer, layout) };
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if !result.is_null() {
            // `allocated_bytes` is requested bytes: each successful realloc
            // contributes its new requested size, while logical live bytes
            // changes only by the size delta.
            ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED_BYTES.fetch_add(new_size, Ordering::Relaxed);
            if new_size >= layout.size() {
                let delta = new_size - layout.size();
                let live = LIVE_BYTES.fetch_add(delta, Ordering::Relaxed) + delta;
                PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
            } else {
                LIVE_BYTES.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
            }
        }
        result
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn signed_delta(after: usize, before: usize) -> i128 {
    i128::try_from(after).unwrap_or(i128::MAX) - i128::try_from(before).unwrap_or(i128::MAX)
}

fn main() -> Result<(), Box<dyn Error>> {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    if arguments.len() != 3 {
        return Err(
            "usage: xlsx0683_alloc OP FILE.xlsx SHEET\nOP is visit-cold, cells-cold, visit-selected, cells-selected, visit-warm, or cells-warm"
                .into(),
        );
    }
    let operation = &arguments[0];
    let path = Path::new(&arguments[1]);
    let sheet = &arguments[2];

    // Warm/selected setup belongs outside the measured operation. The owner is
    // retained through the gauge reads below, so the retained delta includes
    // only the operation's new logical state. Cold setup is intentionally
    // inside its operation.
    let prepared = if matches!(operation.as_str(), "visit-selected" | "cells-selected") {
        Some(PreparedWorksheet::open(path, sheet, false)?)
    } else if matches!(operation.as_str(), "visit-warm" | "cells-warm") {
        Some(PreparedWorksheet::open(path, sheet, true)?)
    } else {
        None
    };
    Counting::reset();
    let baseline_live = LIVE_BYTES.load(Ordering::Relaxed);
    let started = std::time::Instant::now();
    let result = if let Some(prepared) = prepared.as_ref() {
        prepared.measure(operation)
    } else {
        cold_measure(operation, path, sheet)
    };
    let elapsed_ns = started.elapsed().as_nanos();
    let succeeded = result.is_ok();
    let mut retained_cells = None;
    let callbacks = match result {
        Ok(measurement) => {
            retained_cells = measurement.retained_cells;
            measurement.callbacks
        },
        Err(error) => {
            let callbacks = error_callbacks(error.as_ref());
            black_box(error);
            callbacks
        },
    };

    let live_after = LIVE_BYTES.load(Ordering::Relaxed);
    let peak = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    println!(
        "operation={}\tcallbacks={}\tresult={}\tallocation_calls={}\tallocated_bytes={}\tdeallocation_calls={}\tdeallocated_bytes={}\tpeak_live_delta={}\tretained_live_delta={}\telapsed_ns={}",
        operation,
        callbacks,
        if succeeded { "ok" } else { "refused" },
        ALLOCATION_CALLS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
        DEALLOCATION_CALLS.load(Ordering::Relaxed),
        DEALLOCATED_BYTES.load(Ordering::Relaxed),
        signed_delta(peak, baseline_live),
        signed_delta(live_after, baseline_live),
        elapsed_ns,
    );
    black_box(&retained_cells);
    black_box(&prepared);
    Ok(())
}
