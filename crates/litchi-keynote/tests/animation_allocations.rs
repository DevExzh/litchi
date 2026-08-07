#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_keynote::{AnimationType, Result};

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: Every operation delegates to the platform allocator after recording
// only the allocation count; the allocator does not alter pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: `pointer` and `layout` are exactly the pair returned by a
        // previous allocation delegated to `System`.
        unsafe { System.dealloc(pointer, layout) }
    }
}

#[test]
fn known_animation_identifier_classification_does_not_allocate() -> Result<()> {
    let before = ALLOCATIONS.load(Ordering::Relaxed);
    let parsed = std::hint::black_box(AnimationType::from_identifier("FADE-and-SCALE"));
    let after = ALLOCATIONS.load(Ordering::Relaxed);

    assert_eq!(parsed?, AnimationType::FadeAndScale);
    assert_eq!(after, before);

    let before_unknown = ALLOCATIONS.load(Ordering::Relaxed);
    let unknown = std::hint::black_box(AnimationType::from_identifier("com.example.future-effect"));
    let after_unknown = ALLOCATIONS.load(Ordering::Relaxed);
    assert!(matches!(unknown, Ok(AnimationType::Unknown(_))));
    assert_eq!(after_unknown - before_unknown, 1);

    let oversized = "x".repeat(litchi_keynote::build::MAX_IDENTIFIER_BYTES + 1);
    assert_eq!(
        AnimationType::from_identifier(&oversized),
        Err(litchi_keynote::Error::IdentifierTooLarge)
    );
    Ok(())
}
