//! Allocation-instrumented entry point.
//!
//! Only this binary owns the process-global `GlobalAlloc` wrapper, so the
//! timing and callgrind binary is unperturbed. This is the same isolation the
//! repository's own `litchi-perf-baseline-alloc` target uses.

use std::alloc::{GlobalAlloc, Layout, System};

use cfb_save_probe::alloc_metrics;

struct CountingSystemAllocator;

// SAFETY: every method delegates to `std::alloc::System`; the metric updates
// are side-effect-only and change no pointer, layout or ownership contract.
unsafe impl GlobalAlloc for CountingSystemAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            alloc_metrics::record_allocation(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            alloc_metrics::record_allocation(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        unsafe { System.dealloc(pointer, layout) };
        alloc_metrics::record_deallocation(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if !result.is_null() {
            alloc_metrics::record_reallocation(layout.size(), new_size);
        }
        result
    }
}

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingSystemAllocator = CountingSystemAllocator;

fn main() -> Result<(), cfb_save_probe::BoxError> {
    alloc_metrics::enable();
    cfb_save_probe::run()
}
