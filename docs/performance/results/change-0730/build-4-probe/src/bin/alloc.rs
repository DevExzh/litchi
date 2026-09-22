use std::alloc::{GlobalAlloc, Layout, System};

use ole_format_save_probe_0730::alloc_metrics;

struct CountingSystemAllocator;

// SAFETY: every operation delegates to `std::alloc::System`; metric updates
// do not alter pointers, layouts, or ownership.
unsafe impl GlobalAlloc for CountingSystemAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            alloc_metrics::record_allocation(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            alloc_metrics::record_allocation(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer and layout came from the matching allocator.
        unsafe { System.dealloc(pointer, layout) };
        alloc_metrics::record_deallocation(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the pointer and old layout satisfy the realloc contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if !result.is_null() {
            alloc_metrics::record_reallocation(layout.size(), new_size);
        }
        result
    }
}

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingSystemAllocator = CountingSystemAllocator;

fn main() -> Result<(), ole_format_save_probe_0730::BoxError> {
    alloc_metrics::enable();
    ole_format_save_probe_0730::run(false)
}
