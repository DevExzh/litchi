//! Allocation evidence for path resolution (record 0010 §4.1 and §9.2, M1
//! step R2).
//!
//! The performance bar requires that resolving a short path allocates nothing
//! on the success path: names fold into a stack buffer, the alias and rule
//! tables are `static`, the classified step borrows from the parsed path, and
//! "did you mean" suggestions run only on the error path. This test pins that
//! with a counting allocator, as `litchi-core`'s `selector_allocations` test
//! does for parsing, over the crate's test-only fake formats. A `find` miss
//! allocates nothing either: it computes no suggestion.

#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use litchi_ops::Code;
use litchi_ops::path::__fake::{Document, find, get};
use litchi_ops::path::{IntoPath, Path};

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

/// Typical three- to six-step paths, each with at most one predicate per
/// step, so parsing the text allocates nothing either (record 0005).
const TEXT_PATHS: &[(Document, &str, &str)] = &[
    (Document::Word, "/body/p[2]/r[1]", "r"),
    (Document::Word, "/body/tbl[1]/tr[2]/tc[1]/p[1]/r[2]", "r"),
    (
        Document::Word,
        "/BODY/Paragraph[@paraId=1a2b3c4d]/hyperlink[1]",
        "hyperlink",
    ),
    (Document::Word, "/body/table[1]/row[1]/cell[2]/p[1]", "p"),
    (Document::Word, "/body/sdt[@sdtId=7]/p[1]/r[1]", "r"),
    (Document::Word, "/body/p[last()]/r[1]", "r"),
    (Document::Word, "/header[@type=first]/p[1]", "p"),
    (Document::Word, "/footnotes/footnote[2]/p[1]/r[1]", "r"),
    (Document::Word, "/footnote[2]/p[1]/r[1]", "r"),
    (Document::Word, "/comments/comment[@commentId=1]/p[1]", "p"),
    (Document::Word, "/styles/Heading1/tab[1]", "tab"),
    (
        Document::Word,
        "/numbering/abstractNum[@id=1]/level[1]",
        "level",
    ),
    (
        Document::Presentation,
        "/slide[1]/shape[@name=Title 1]",
        "shape",
    ),
    (
        Document::Presentation,
        "/slide[2]/shape[@name=ÄRGER]",
        "shape",
    ),
    (
        Document::Presentation,
        "/slide[1]/table[@name=grid]/tr[1]/tc[2]",
        "tc",
    ),
    (
        Document::Presentation,
        "/slideMaster[1]/layout[1]/shape[1]",
        "shape",
    ),
    (
        Document::Presentation,
        "/slide[1]/ph[ctrTitle]",
        "placeholder",
    ),
    (Document::Workbook, "/Sheet1/A1", "cell"),
    (Document::Workbook, "Sheet1!B2", "cell"),
    (Document::Workbook, "'Data Q1'!A1:C5", "range"),
    (Document::Workbook, "Sheet1!3:3", "row"),
    (Document::Workbook, "Sheet1!B:B", "col"),
    (Document::Workbook, "/sheet[2]/row[2]", "row"),
    (Document::Workbook, "/Chart/B2", "cell"),
    (Document::Workbook, "/STRAßE/B2", "cell"),
    (Document::Workbook, "/namedrange[Total]", "namedrange"),
    (Document::Workbook, "/Sheet1/table[@name=sales]", "table"),
];

/// Paths whose text parse allocates (two stacked brackets keep their
/// predicates in a `Vec`), resolved here from a parsed `Path`.
const STACKED: &[(Document, &str, &str)] = &[
    (
        Document::Presentation,
        "/slide[1]/placeholder[@type=body][@idx=1]",
        "placeholder",
    ),
    (
        Document::Word,
        "/body/p[@paraId=1A2B3C4D][@textId=77AA0001]/r[2]",
        "r",
    ),
];

#[test]
fn text_paths_parse_and_resolve_without_allocating() {
    for &(document, text, kind) in TEXT_PATHS {
        assert_eq!(
            get(document, litchi_ops::path::raw::path(text).unwrap()).expect(text),
            kind,
            "{text}"
        );
        let count =
            allocations(|| get(document, litchi_ops::path::raw::path(text).unwrap()).expect(text));
        assert_eq!(count, 0, "{text:?} allocated {count} times");
        let count =
            allocations(|| find(document, litchi_ops::path::raw::path(text).unwrap()).expect(text));
        assert_eq!(count, 0, "find {text:?} allocated {count} times");
    }
}

#[test]
fn parsed_paths_resolve_without_allocating() {
    for &(document, text, kind) in TEXT_PATHS.iter().chain(STACKED) {
        let parsed = text.into_path().expect(text);
        let path: &Path<'_> = parsed.path();
        assert_eq!(
            get(document, litchi_ops::path::raw::path(path).unwrap()).expect(text),
            kind,
            "{text}"
        );
        let count =
            allocations(|| get(document, litchi_ops::path::raw::path(path).unwrap()).expect(text));
        assert_eq!(count, 0, "{text:?} allocated {count} times");
    }
}

#[test]
fn find_misses_compute_no_suggestion() {
    let misses: &[(Document, &str)] = &[
        (Document::Word, "/body/p[99]/r[1]"),
        (Document::Word, "/body/p[@paraId=DEADBEEF]"),
        (Document::Word, "/styles/Missing"),
        (Document::Workbook, "/Sheet9/A1"),
        (Document::Workbook, "/namedrange[Nope]"),
        (Document::Presentation, "/slide[1]/shape[@name=Nope]"),
    ];
    for &(document, text) in misses {
        assert_eq!(
            find(document, litchi_ops::path::raw::path(text).unwrap()).expect(text),
            None,
            "{text}"
        );
        let count =
            allocations(|| find(document, litchi_ops::path::raw::path(text).unwrap()).expect(text));
        assert_eq!(count, 0, "find {text:?} allocated {count} times");
        // `get` builds the error, message and suggestion on the error path.
        let error = get(document, litchi_ops::path::raw::path(text).unwrap()).unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound, "{text}");
        assert!(error.suggestion().is_some(), "{text}");
        assert!(
            allocations(|| get(document, litchi_ops::path::raw::path(text).unwrap()).unwrap_err())
                > 0,
            "{text}"
        );
    }
}

#[test]
fn lowercase_column_letters_remain_borrowed() {
    let count = allocations(|| get(Document::Workbook, "Sheet1!b:b").expect("column"));
    assert_eq!(count, 0);
}

#[path = "support/query_allocations.rs"]
mod query_allocations;

#[cfg(feature = "docx")]
#[test]
fn capability_construction_and_iteration_allocate_nothing() {
    assert_eq!(
        allocations(|| {
            let caps = litchi_ops::capabilities::Capabilities::of(litchi_core::FileFormat::Docx);
            for element in caps.elements() {
                std::hint::black_box(element.operations());
                for path in element.paths() {
                    std::hint::black_box(path);
                }
                for property in element.properties() {
                    std::hint::black_box(property.value_type());
                }
            }
            std::hint::black_box(caps.gaps().count());
        }),
        0
    );
}
