//! Allocation evidence for the XLSX adapter (record 0010 §4.7, §5.3, §9.2,
//! M1 step R4b).
//!
//! On a warm handle, a query over a workbook visits its cells without
//! allocating per cell: a cell node holds one reference-counted grid
//! handle, its label sits in a stack buffer, its value is borrowed from the
//! leaf's parsed grid and its format resolves through the workbook's shared
//! format table. These tests run the same compiled plans over a small and a
//! large generated workbook (40 and 4,000 data rows) with a per-thread
//! counting allocator, the pattern of `query_allocations.rs`: the
//! allocations of one whole walk must be equal on both, and few. A warm
//! `get` of one cell allocates the same in any sheet.
//!
//! The adapter's test entry points are compiled under the `test-support`
//! feature only, so this test runs with `--features test-support,xlsx` (or
//! `--all-features`).

#![cfg(all(feature = "test-support", feature = "xlsx"))]
#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use litchi_ops::format::xlsx::__unified::{compile, get, run};
use litchi_ops::format::xlsx::Xlsx;
use litchi_ops::query::Query;

#[path = "support/xlsx_workbook.rs"]
mod generated;

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

fn allocations<T>(body: impl FnOnce() -> T) -> (usize, T) {
    let before = ALLOCATIONS.with(Cell::get);
    let value = std::hint::black_box(body());
    let after = ALLOCATIONS.with(Cell::get);
    (after - before, value)
}

fn open(rows: usize) -> Xlsx {
    match litchi_ops::open_bytes(generated::workbook(rows)).expect("opens") {
        litchi_ops::File::Xlsx(file) => file,
        other => panic!("opened as {other:?}"),
    }
}

/// The most allocations one warm walk makes, whatever the workbook's size:
/// the walker's own (its stack, child counters and per-alternative state)
/// and per worksheet the adapter's grid handle and row list.
const PER_WALK: usize = 8;

/// A walk that reads a worksheet's tables (a `table` query, or a header-row
/// query binding its columns) adds the leaf's read of each table part,
/// which the leaf parses again on every call (about 110 allocations for
/// this workbook's three-column table; a leaf follow-up, not the adapter's).
const PER_TABLE_WALK: usize = PER_WALK + 128;

/// Selectors covering the adapter's evaluation paths: every kind, values,
/// text, formats, coordinates, header rows and unions; whether the walk
/// reads tables, and whether its matches grow with the workbook.
const CASES: &[(&str, bool, bool)] = &[
    ("cell", false, true),
    ("cell[value>1000]", false, true),
    ("cell[type=SharedString]", false, true),
    ("Data!cell:contains(West)", false, true),
    ("B", false, true),
    ("A1:C10", false, false),
    ("cell[bold=true]", false, false),
    ("row", false, true),
    ("row[Sales>1000]", true, true),
    ("row[Region=West and Status=open]", true, true),
    ("row[col.Sales>=800][last()]", true, false),
    ("column", false, false),
    ("table", true, false),
    ("row[2], cell[value=2000]", false, true),
];

#[test]
fn the_counter_sees_allocations() {
    assert_eq!(allocations(|| Vec::<u8>::with_capacity(16)).0, 1);
}

#[test]
fn warm_walks_allocate_per_walk_never_per_cell() {
    let (small, large) = (open(40), open(4_000));
    for &(selector, tables, scales) in CASES {
        let query = Query::parse(selector).expect(selector);
        let plan = compile(&query).unwrap_or_else(|error| panic!("{selector}: {error}"));
        // The first walk reads and caches the grid, which allocates with
        // its size.
        for file in [&small, &large] {
            run(file, &plan, usize::MAX, |_| {})
                .unwrap_or_else(|error| panic!("{selector}: {error}"));
        }
        let (small_count, small_matches) = allocations(|| run(&small, &plan, usize::MAX, |_| {}));
        let (large_count, large_matches) = allocations(|| run(&large, &plan, usize::MAX, |_| {}));
        let (small_matches, large_matches) = (
            small_matches.expect(selector),
            large_matches.expect(selector),
        );
        assert!(
            !scales || large_matches >= 50 * small_matches.max(1),
            "{selector}: {small_matches} and {large_matches} matches"
        );
        assert_eq!(
            small_count, large_count,
            "{selector}: {small_count} allocations for {small_matches} matches, {large_count} for {large_matches}"
        );
        let bound = if tables { PER_TABLE_WALK } else { PER_WALK };
        assert!(
            large_count <= bound,
            "{selector}: {large_count} allocations for one walk"
        );
    }
}

#[test]
fn a_first_match_costs_the_same_in_any_workbook() {
    let (small, large) = (open(40), open(4_000));
    for (selector, bound) in [
        ("cell[value>1000]", PER_WALK),
        ("cell:contains(South)", PER_WALK),
        ("row[Sales>1000]", PER_TABLE_WALK),
    ] {
        let query = Query::parse(selector).expect(selector);
        let plan = compile(&query).expect(selector);
        for file in [&small, &large] {
            run(file, &plan, 1, |_| {}).expect(selector);
        }
        let (small_count, _) = allocations(|| run(&small, &plan, 1, |_| {}));
        let (large_count, _) = allocations(|| run(&large, &plan, 1, |_| {}));
        assert_eq!(small_count, large_count, "{selector}");
        assert!(large_count <= bound, "{selector}: {large_count}");
    }
}

/// A warm `get` allocates the same in any sheet, and for a cell no more
/// than the leaf's own `sheet` + `cell` read, which returns an owned copy of
/// the record, plus the one box the node keeps that copy in.
#[test]
fn a_warm_get_allocates_the_same_in_any_sheet() {
    let (small, large) = (open(40), open(4_000));
    for path in [
        "/Data/B20",
        "/Data/row[21]",
        "Data!C30",
        "/Data/table[1]",
        "/data/col[B]",
    ] {
        for file in [&small, &large] {
            get(file, path).unwrap_or_else(|error| panic!("{path}: {error}"));
        }
        let (small_count, node) = allocations(|| get(&small, path).map(drop));
        node.expect(path);
        let (large_count, node) = allocations(|| get(&large, path).map(drop));
        node.expect(path);
        assert_eq!(small_count, large_count, "{path}");
    }
    let leaf = litchi_xlsx::SourceBackedWorkbook::from_reader(std::io::Cursor::new(
        generated::workbook(4_000),
    ))
    .expect("opens");
    let read = || {
        let sheet = leaf.sheet("Data").expect("reads").expect("a sheet");
        sheet.cell("B20").expect("reads")
    };
    drop(read());
    let (leaf_count, _) = allocations(read);
    let (unified_count, node) = allocations(|| get(&large, "/Data/B20").map(drop));
    node.expect("resolves");
    assert!(
        unified_count <= leaf_count + 1,
        "get allocates {unified_count}, the leaf's read {leaf_count}"
    );
}
