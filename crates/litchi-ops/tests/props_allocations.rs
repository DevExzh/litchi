//! Allocation evidence for borrowed property input construction and iteration.

#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use litchi_ops::value::Length;
use litchi_ops::{IntoProps, props};

// Per-thread counters so other tests in this binary cannot disturb a
// measurement. Const-initialised without destructors, so the allocator can
// read them without allocating or recursing.
thread_local! {
    static ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
}

fn record() {
    let _ = ALLOCATIONS.try_with(|cell| cell.set(cell.get().wrapping_add(1)));
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: Every operation delegates to the platform allocator after recording
// only a count; the allocator does not alter pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        record();
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc(layout) }
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        record();
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc_zeroed(layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        record();
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

fn allocations<T>(body: impl FnOnce() -> T) -> usize {
    let before = ALLOCATIONS.with(Cell::get);
    let value = std::hint::black_box(body());
    let after = ALLOCATIONS.with(Cell::get);
    drop(value);
    after - before
}

#[test]
fn borrowed_macro_and_string_inputs_do_not_allocate() {
    let text = String::from("caller-owned text");
    assert_eq!(
        allocations(|| {
            let input = props! { "text" => text.as_str(), "bold" => true, "size" => Length::from_points(12), "count" => 42 };
            for pair in input.into_props().unwrap() {
                std::hint::black_box(pair);
            }
        }),
        0
    );
    assert_eq!(
        allocations(|| {
            for pair in ["text=some=a=b", "bold=true"].into_props().unwrap() {
                std::hint::black_box(pair);
            }
        }),
        0
    );
}
