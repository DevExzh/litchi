#![cfg(feature = "ooxml")]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};

use litchi_crypto::ooxml::{Mode, encrypt};

struct CountingAllocator;

static COUNTING: AtomicBool = AtomicBool::new(false);
static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

// This allocator is scoped to this integration-test binary. The library itself
// remains safe-only; the wrapper lets this test distinguish fixed workspace
// behavior from a hidden allocation in every 4 KiB cryptographic segment.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if COUNTING.load(Ordering::Relaxed) {
            ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        }
        // SAFETY: delegated unchanged to the platform allocator.
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: delegated unchanged to the platform allocator.
        unsafe { System.dealloc(pointer, layout) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

fn encryption_allocations(clear_len: usize) -> usize {
    let clear = vec![0x5a; clear_len];
    ALLOCATIONS.store(0, Ordering::Relaxed);
    COUNTING.store(true, Ordering::Relaxed);
    let result = encrypt(clear, "allocation regression password", Mode::Agile);
    COUNTING.store(false, Ordering::Relaxed);
    result.expect("Agile encryption");
    ALLOCATIONS.load(Ordering::Relaxed)
}

#[test]
fn agile_segment_iv_derivation_does_not_allocate_per_segment() {
    let one_segment = encryption_allocations(4_096);
    let many_segments = encryption_allocations(4_096 * 256 + 37);

    // The operation necessarily allocates package and CFB buffers, but the
    // segment count must not add one heap allocation per 4 KiB block.
    assert!(
        many_segments <= one_segment + 64,
        "allocation count grew from {one_segment} to {many_segments} across 257 segments"
    );
}
