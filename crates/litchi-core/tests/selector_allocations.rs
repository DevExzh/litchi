//! Allocation evidence for path and selector parsing.
//!
//! The unified-operations performance bar requires that parsing a short path
//! allocates nothing on the success path. Names and values borrow from the
//! input, steps and single predicates live inline, and the error type is
//! `Copy`. This test pins that property with a counting allocator, and bounds
//! the allocations of the common multi-part query forms.

#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use litchi_core::selector::{Axis, CellRange, Path, Query};

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

const ZERO_ALLOCATION_PATHS: &[&str] = &[
    "/",
    "/body/p[3]/r[1]",
    "/body/p[last()]",
    "/slide[1]/shape[2]",
    "/slide[2]/notes",
    "/slide[1]/shape[@name=Title]",
    "/slideMaster[1]/layout[2]",
    "/header[1]/p[1]",
    "/footnote[2]",
    "/comments/comment[3]",
    "/body/p[@paraId=1A2B3C4D]/r[2]",
    "/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]",
    "/Sheet1/A1",
    "/Sheet1/A1:C5",
    "/Sheet1/row[5]",
    "/Sheet1/col[B]",
    "/Balance Sheet/B18:E18",
    "/Sheet1/chart[1]/axis[@role=value]",
    "/numbering/abstractNum[@id=100]/level[1]",
];

const ZERO_ALLOCATION_QUERIES: &[&str] = &[
    "paragraph",
    "*",
    "paragraph[style=Heading1]",
    "row[Salary>5000]",
    "shape[fill=FF0000]",
    "cell[type=Number]",
    "field[fieldType!=page]",
    "shape[width>=10cm]",
    "row[2]",
    "Sheet1!A1",
    "Sheet1!B[value!=0]",
    "'My Data (2024)'!row[Amount>5000]",
    "cell:contains(\"TODO\")",
    "run:contains(TODO)",
    "p:empty",
    "image:no-alt",
    "cell:TBD",
];

#[test]
fn short_paths_parse_without_allocating() {
    for input in ZERO_ALLOCATION_PATHS {
        let count = allocations(|| Path::parse(input).expect("parses"));
        assert_eq!(count, 0, "{input:?} allocated {count} times");
    }
    let count = allocations(|| Path::parse_relative("p[@paraId=1A2B3C4D]").expect("parses"));
    assert_eq!(count, 0);
}

#[test]
fn simple_queries_parse_without_allocating() {
    for input in ZERO_ALLOCATION_QUERIES {
        let count = allocations(|| Query::parse(input).expect("parses"));
        assert_eq!(count, 0, "{input:?} allocated {count} times");
    }
}

#[test]
fn compound_queries_allocate_a_bounded_few_times() {
    // One `Vec` per multi-operand boolean, combinator chain, pseudo-class
    // list, nested selector list, or boxed scope path.
    let cases: &[(&str, usize)] = &[
        ("row[Salary>5000 and Region=EMEA]", 1),
        ("row[Status=open or Priority=high]", 1),
        ("paragraph[style=Normal] > run[font!=Arial]", 1),
        ("cell:has(formula)", 1),
        ("cell:not(:empty)", 1),
        ("paragraph, table", 1),
        ("/Sheet1/cell[type=Number]", 1),
    ];
    for &(input, bound) in cases {
        let count = allocations(|| Query::parse(input).expect("parses"));
        assert!(
            count <= bound,
            "{input:?} allocated {count} times (bound {bound})"
        );
    }
}

#[test]
fn rejection_and_helpers_do_not_allocate() {
    for input in ["/body/p[0]", "/body//p", "p[style=Normal", "slide + shape"] {
        assert_eq!(allocations(|| Path::parse(input).is_err()), 0, "{input:?}");
        assert_eq!(allocations(|| Query::parse(input).is_err()), 0, "{input:?}");
    }
    assert_eq!(allocations(|| Axis::classify("slideMaster")), 0);
    assert_eq!(allocations(|| CellRange::parse("A1:XFD1048576")), 0);
}
