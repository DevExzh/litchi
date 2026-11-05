//! Allocation evidence for the query walker (record 0010 §5.3, §9.2, M1
//! step R3).
//!
//! The performance bar requires that a query visits nodes without
//! allocating: the walk's stack, child counters and ancestor memo grow with
//! the document's depth, never with the number of nodes it visits, and
//! predicates are compiled once per query, not per node. These tests run the
//! same compiled plans over a small and a large fake document of the same
//! depth and count, with a per-thread counting allocator (the pattern of
//! `litchi-core`'s `selector_allocations` test, record 0005), the allocations
//! of each whole walk: they must be equal, and few.
//!
//! The fake formats are compiled under the `test-support` feature only, so
//! this test runs with `--features test-support` (or `--all-features`).

use super::ALLOCATIONS;
use litchi_ops::query::__fake::{Document, Family, Outcome, compile_for, paths, run};
use litchi_ops::query::Query;
use std::cell::Cell;

fn allocations<T>(body: impl FnOnce() -> T) -> (usize, T) {
    let before = ALLOCATIONS.with(Cell::get);
    let value = std::hint::black_box(body());
    let after = ALLOCATIONS.with(Cell::get);
    (after - before, value)
}

/// The most allocations one walk makes, whatever the document's size. The
/// walk's own: its stack, child counters and per-alternative state (three for
/// every walk), the ancestor memo when the selector has combinators, the
/// `[last()]` ancestor snapshot, and for a header-row query the binding and
/// key lists; plus, for a header-row query, the four lists the fake
/// adapter's binding builds (its sheets, candidate tables, column numbers and
/// the binding's copy of them).
const PER_WALK: usize = 9;

/// Selectors covering every evaluation path: kinds, properties, typed
/// comparisons, text pseudo-classes, `:has`, `:not`, combinators with
/// backtracking, unions, subject positions, `*` and header rows. The flag
/// says whether the selector walks the whole document; `[N]` stops at its
/// match.
const CASES: &[(Family, &str, bool)] = &[
    (Family::Word, "p", true),
    (Family::Word, "r", true),
    (Family::Word, "p[style=Heading1]", true),
    (Family::Word, "p[style=Heading1] > r[bold=true]", true),
    (Family::Word, "body p:contains(revenue)", true),
    (Family::Word, "p:not(:contains(Body))", true),
    (Family::Word, "p:has(r[bold=true])", true),
    (Family::Word, "tbl > tr tc > p", true),
    (Family::Word, "r[bold=true], p[style=Heading1], tc", true),
    (Family::Word, "p[last()]", true),
    (Family::Word, "p[style=Normal][2]", false),
    (Family::Word, "*", true),
    (Family::Word, "[style=Heading1]", true),
    (Family::Presentation, "shape", true),
    (Family::Presentation, "title", true),
    (Family::Presentation, "textbox[width>=10cm]", true),
    (
        Family::Presentation,
        "slide > shape:contains(revenue)",
        true,
    ),
    (Family::Presentation, "shape[fill=red]", true),
    (Family::Presentation, "picture:no-alt", true),
    (Family::Workbook, "cell", true),
    (Family::Workbook, "cell[value>1000]", true),
    (Family::Workbook, "B", true),
    (Family::Workbook, "A1:C10", true),
    (Family::Workbook, "Data!cell:contains(West)", true),
    (Family::Workbook, "row", true),
    (Family::Workbook, "row[Sales>1000]", true),
    (Family::Workbook, "row[Region=West and Status=open]", true),
    (Family::Workbook, "row[not(Region=West)]", true),
];

/// A small and a large document of one family, of equal depth.
fn documents(family: Family) -> [Document; 2] {
    match family {
        Family::Word => [Document::word(50), Document::word(5_000)],
        Family::Presentation => [Document::deck(5, 4), Document::deck(500, 4)],
        Family::Workbook => [Document::workbook(40), Document::workbook(4_000)],
    }
}

#[test]
fn the_counter_sees_allocations() {
    assert_eq!(allocations(|| Vec::<u8>::with_capacity(16)).0, 1);
    assert_eq!(allocations(|| String::from("x")).0, 1);
}

#[test]
fn walks_allocate_per_walk_never_per_visited_node() {
    for family in [Family::Word, Family::Presentation, Family::Workbook] {
        let [small, large] = documents(family);
        for &(_, selector, whole) in CASES.iter().filter(|(case, ..)| *case == family) {
            let query = Query::parse(selector).expect(selector);
            let plan =
                compile_for(family, &query).unwrap_or_else(|error| panic!("{selector}: {error}"));
            let (small_count, small_run) = allocations(|| run(&small, &plan, usize::MAX));
            let (large_count, large_run) = allocations(|| run(&large, &plan, usize::MAX));
            let small_run: Outcome =
                small_run.unwrap_or_else(|error| panic!("{selector}: {error}"));
            let large_run: Outcome =
                large_run.unwrap_or_else(|error| panic!("{selector}: {error}"));
            assert!(
                !whole || large_run.visited >= 10 * small_run.visited.max(1) / 2,
                "{selector}: the large document visits {} nodes, the small {}",
                large_run.visited,
                small_run.visited
            );
            assert_eq!(
                small_count, large_count,
                "{selector}: {small_count} allocations over {} visited nodes, {large_count} over {}",
                small_run.visited, large_run.visited
            );
            assert!(
                large_count <= PER_WALK,
                "{selector}: {large_count} allocations for one walk"
            );
        }
    }
}

#[test]
fn a_first_match_costs_the_same_in_any_document() {
    for family in [Family::Word, Family::Presentation, Family::Workbook] {
        let [small, large] = documents(family);
        for &(_, selector, _) in CASES.iter().filter(|(case, ..)| *case == family) {
            let query = Query::parse(selector).expect(selector);
            let plan = compile_for(family, &query).expect(selector);
            let (small_count, small_run) = allocations(|| run(&small, &plan, 1));
            let (large_count, large_run) = allocations(|| run(&large, &plan, 1));
            let (small_run, large_run) = (small_run.expect(selector), large_run.expect(selector));
            assert!(large_count <= PER_WALK, "{selector}");
            assert_eq!(small_count, large_count, "{selector}");
            // Streaming: the first match is reached after the same visits,
            // unless the selector needs the whole document (`[last()]`).
            if !selector.contains("last()") && small_run.matches == 1 {
                assert_eq!(small_run.visited, large_run.visited, "{selector}");
            }
        }
    }
}

#[test]
fn writing_match_paths_allocates_only_in_the_sink() {
    let document = Document::word(2_000);
    let query = Query::parse("tbl tc > p, p[style=Heading1]").expect("parses");
    let plan = compile_for(Family::Word, &query).expect("compiles");
    let mut sink = String::with_capacity(1 << 20);
    let (count, outcome) = allocations(|| paths(&document, &plan, usize::MAX, &mut sink));
    let outcome = outcome.expect("runs");
    assert_eq!(outcome.matches, 200 + 9);
    assert!(count <= PER_WALK, "{count} allocations");
    assert!(
        sink.starts_with("/body/tbl[1]/tr[1]/tc[1]/p[1]\n"),
        "{}",
        &sink[..40]
    );
    assert!(sink.contains("/body/tbl[1]/tr[3]/tc[3]/p[1]\n/body/p[1]\n/body/p[11]\n"));
}
