//! Isolated allocator-instrumented benchmark entry point.
//!
//! The shared harness library is forbid-safe.  Only this target owns the
//! process-global `GlobalAlloc` wrapper and its narrowly scoped unsafe calls.

mod allocator {
    use std::alloc::{GlobalAlloc, Layout, System};

    use litchi_perf_baseline::allocation_metrics;

    struct CountingSystemAllocator;

    // SAFETY: Every method delegates to `std::alloc::System`; the metric updates
    // are side-effect-only and do not change the pointer, layout, or ownership
    // contract required by `GlobalAlloc`.
    unsafe impl GlobalAlloc for CountingSystemAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: The caller upholds the `GlobalAlloc` layout contract.
            let pointer = unsafe { System.alloc(layout) };
            if pointer.is_null() {
                allocation_metrics::record_failed_allocation();
            } else {
                allocation_metrics::record_allocation(layout.size());
            }
            pointer
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            // SAFETY: The caller upholds the `GlobalAlloc` layout contract.
            let pointer = unsafe { System.alloc_zeroed(layout) };
            if pointer.is_null() {
                allocation_metrics::record_failed_allocation();
            } else {
                allocation_metrics::record_allocation(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: The caller upholds the `GlobalAlloc` pointer/layout contract.
            unsafe { System.dealloc(pointer, layout) };
            allocation_metrics::record_deallocation(layout.size());
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            // SAFETY: The caller upholds the `GlobalAlloc` pointer/layout contract.
            let result = unsafe { System.realloc(pointer, layout, new_size) };
            if result.is_null() {
                allocation_metrics::record_failed_allocation();
            } else {
                allocation_metrics::record_reallocation(layout.size(), new_size);
            }
            result
        }
    }

    #[global_allocator]
    static GLOBAL_ALLOCATOR: CountingSystemAllocator = CountingSystemAllocator;

    #[cfg(test)]
    mod tests {
        use std::{
            alloc::{GlobalAlloc, Layout},
            sync::Mutex,
        };

        use super::GLOBAL_ALLOCATOR;
        use litchi_perf_baseline::allocation_metrics;

        static TEST_LOCK: Mutex<()> = Mutex::new(());

        #[test]
        fn global_allocator_records_successful_alloc_and_dealloc_with_live_peak() {
            let _lock = TEST_LOCK.lock().unwrap();
            allocation_metrics::enable();
            let before = allocation_metrics::snapshot();
            let layout = Layout::from_size_align(256, 8).unwrap();
            // SAFETY: `layout` is valid and this direct call uses the same
            // allocator whose returned pointer is deallocated below.
            let pointer = unsafe { GLOBAL_ALLOCATOR.alloc(layout) };
            assert!(!pointer.is_null());
            let during = allocation_metrics::snapshot();
            assert!(
                during.allocation_calls >= before.allocation_calls + 1,
                "successful alloc must increment allocation calls"
            );
            assert!(during.allocated_bytes >= before.allocated_bytes + 256);
            assert!(during.live_bytes >= before.live_bytes + 256);
            // SAFETY: `pointer` came from `GLOBAL_ALLOCATOR.alloc(layout)` and
            // has not been freed or otherwise used since that call.
            unsafe { GLOBAL_ALLOCATOR.dealloc(pointer, layout) };
            let after = allocation_metrics::snapshot();
            assert!(
                after.deallocation_calls >= before.deallocation_calls + 1,
                "successful dealloc must increment deallocation calls"
            );
            assert!(after.deallocated_bytes >= before.deallocated_bytes + layout.size() as u64);
            assert!(after.peak_live_bytes >= during.peak_live_bytes);
        }

        #[test]
        fn global_allocator_records_zeroed_alloc_and_zeroes_memory() {
            let _lock = TEST_LOCK.lock().unwrap();
            allocation_metrics::enable();
            let before = allocation_metrics::snapshot();
            let layout = Layout::from_size_align(64, 8).unwrap();
            // SAFETY: `layout` is valid and this direct call uses the same
            // allocator whose returned pointer is deallocated below.
            let pointer = unsafe { GLOBAL_ALLOCATOR.alloc_zeroed(layout) };
            assert!(!pointer.is_null());
            // SAFETY: `pointer` is a valid allocation of `layout.size()` bytes
            // returned by `GLOBAL_ALLOCATOR.alloc_zeroed` and remains owned by
            // this test until the deallocation below.
            let bytes = unsafe { std::slice::from_raw_parts(pointer, layout.size()) };
            assert!(bytes.iter().all(|byte| *byte == 0));
            let during = allocation_metrics::snapshot();
            assert!(
                during.allocation_calls >= before.allocation_calls + 1,
                "successful alloc_zeroed must increment allocation calls"
            );
            assert!(during.allocated_bytes >= before.allocated_bytes + layout.size() as u64);
            // SAFETY: `pointer` came from `GLOBAL_ALLOCATOR.alloc_zeroed(layout)`
            // and has not been freed or otherwise used since that call.
            unsafe { GLOBAL_ALLOCATOR.dealloc(pointer, layout) };
        }

        #[test]
        fn global_allocator_records_realloc_grow_and_shrink() {
            let _lock = TEST_LOCK.lock().unwrap();
            allocation_metrics::enable();
            let before = allocation_metrics::snapshot();
            let old_layout = Layout::from_size_align(16, 8).unwrap();
            // SAFETY: `old_layout` is valid for this direct allocator call.
            let pointer = unsafe { GLOBAL_ALLOCATOR.alloc(old_layout) };
            assert!(!pointer.is_null());
            // SAFETY: `pointer` was returned for `old_layout` by the same
            // allocator, and the old allocation remains owned by this test.
            let grown = unsafe { GLOBAL_ALLOCATOR.realloc(pointer, old_layout, 32) };
            assert!(!grown.is_null());
            let grown_layout = Layout::from_size_align(32, 8).unwrap();
            // SAFETY: `grown` was returned by the preceding realloc and
            // `grown_layout` describes its current allocation.
            let shrunk = unsafe { GLOBAL_ALLOCATOR.realloc(grown, grown_layout, 8) };
            assert!(!shrunk.is_null());
            let final_layout = Layout::from_size_align(8, 8).unwrap();
            // SAFETY: `shrunk` was returned by the preceding realloc and
            // `final_layout` describes its current allocation.
            unsafe { GLOBAL_ALLOCATOR.dealloc(shrunk, final_layout) };
            let after = allocation_metrics::snapshot();
            assert!(
                after.allocation_calls >= before.allocation_calls + 3,
                "alloc plus two successful reallocations must be counted"
            );
            assert!(after.reallocation_calls >= before.reallocation_calls + 2);
            assert!(after.allocated_bytes >= before.allocated_bytes + 16 + 32 + 8);
            assert!(after.deallocated_bytes >= before.deallocated_bytes + 16 + 32 + 8);
        }

        #[test]
        fn global_allocator_records_a_failed_allocation() {
            let _lock = TEST_LOCK.lock().unwrap();
            allocation_metrics::enable();
            let before = allocation_metrics::snapshot();
            let layout = Layout::from_size_align(isize::MAX as usize, 1).unwrap();
            // SAFETY: `layout` satisfies the `GlobalAlloc` layout contract;
            // this test deliberately exercises the allocator's null result.
            let pointer = unsafe { GLOBAL_ALLOCATOR.alloc(layout) };
            assert!(pointer.is_null(), "the maximal layout should be rejected");
            let after = allocation_metrics::snapshot();
            assert!(after.failed_allocation_calls >= before.failed_allocation_calls + 1);
        }

        #[test]
        fn global_allocator_records_a_failed_realloc_without_losing_old_allocation() {
            let _lock = TEST_LOCK.lock().unwrap();
            allocation_metrics::enable();
            let before = allocation_metrics::snapshot();
            let layout = Layout::from_size_align(32, 8).unwrap();
            // SAFETY: `layout` is valid and this direct call uses the same
            // allocator whose returned pointer is deallocated below.
            let pointer = unsafe { GLOBAL_ALLOCATOR.alloc(layout) };
            assert!(!pointer.is_null());
            // SAFETY: `pointer` is a valid allocation of `layout.size()` bytes
            // returned by `GLOBAL_ALLOCATOR.alloc` and remains owned by this
            // test until the deallocation below.
            unsafe { *pointer = 0xa5 };
            let after_alloc = allocation_metrics::snapshot();
            // SAFETY: `pointer` was returned for `layout` by the same
            // allocator. `isize::MAX` is deliberately unallocatable; a null
            // result leaves the original allocation owned by this test.
            let failed = unsafe { GLOBAL_ALLOCATOR.realloc(pointer, layout, isize::MAX as usize) };
            assert!(failed.is_null(), "the maximal realloc should be rejected");
            let after_failed = allocation_metrics::snapshot();
            assert!(
                after_failed.failed_allocation_calls >= after_alloc.failed_allocation_calls + 1
            );
            // SAFETY: a failed realloc preserves the original allocation and
            // its contents, which remain valid until the deallocation below.
            assert_eq!(unsafe { *pointer }, 0xa5);
            // SAFETY: the failed realloc preserves the original allocation
            // returned for `layout`, which is still valid here.
            unsafe { GLOBAL_ALLOCATOR.dealloc(pointer, layout) };
            assert!(
                allocation_metrics::snapshot().deallocated_bytes
                    >= before.deallocated_bytes + layout.size() as u64
            );
        }
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::allocation_metrics::enable();
    litchi_perf_baseline::run()
}
