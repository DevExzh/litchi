//! Compact regressions for the legacy formula-tokenizer path.
//!
//! These cases keep the common lexer boundary observable without repeating the
//! complete function catalog or the structured OpenFormula reference matrix.

#![allow(
    unsafe_code,
    reason = "The integration-only allocator injects one fallible failure."
)]

use litchi_core::Error;
use litchi_ods::codec::formula::{FormulaParser, Token};
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

thread_local! {
    static FAIL_ONE_BYTE: Cell<bool> = const { Cell::new(false) };
    static INJECTION_FIRED: Cell<bool> = const { Cell::new(false) };
}

struct OneByteFailureAllocator;

#[global_allocator]
static GLOBAL: OneByteFailureAllocator = OneByteFailureAllocator;

// SAFETY: All non-injected operations delegate unchanged to the system
// allocator. The one injected null result is scoped to this test thread and is
// consumed once, so it exercises only a fallible `try_reserve` call.
unsafe impl GlobalAlloc for OneByteFailureAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if should_fail(layout.size()) {
            std::ptr::null_mut()
        } else {
            // SAFETY: The caller supplies a valid layout for the allocation.
            unsafe { System.alloc(layout) }
        }
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        if should_fail(layout.size()) {
            std::ptr::null_mut()
        } else {
            // SAFETY: The caller supplies a valid layout for the allocation.
            unsafe { System.alloc_zeroed(layout) }
        }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: The pointer/layout pair comes from a prior allocator call.
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        if should_fail(size) {
            std::ptr::null_mut()
        } else {
            // SAFETY: The pointer/layout pair and new size are supplied by the
            // allocator caller.
            unsafe { System.realloc(pointer, layout, size) }
        }
    }
}

fn should_fail(size: usize) -> bool {
    FAIL_ONE_BYTE.with(|armed| {
        if armed.get() && size == 1 {
            armed.set(false);
            INJECTION_FIRED.with(|fired| fired.set(true));
            true
        } else {
            false
        }
    })
}

fn arm_one_byte_failure() {
    INJECTION_FIRED.with(|fired| fired.set(false));
    FAIL_ONE_BYTE.with(|armed| armed.set(true));
}

fn disarm_one_byte_failure() {
    FAIL_ONE_BYTE.with(|armed| armed.set(false));
}

fn injection_fired() -> bool {
    INJECTION_FIRED.with(Cell::get)
}

fn assert_cell(
    token: &Token,
    expected_sheet: Option<&str>,
    expected_column: &str,
    expected_row: u32,
    expected_column_absolute: bool,
    expected_row_absolute: bool,
) {
    let Token::CellRef(cell) = token else {
        panic!("expected a legacy cell token, got {token:?}");
    };
    assert_eq!(cell.sheet.as_deref(), expected_sheet);
    assert_eq!(cell.column, expected_column);
    assert_eq!(cell.row, expected_row);
    assert_eq!(cell.column_absolute, expected_column_absolute);
    assert_eq!(cell.row_absolute, expected_row_absolute);
}

#[test]
fn repeated_function_and_cell_scans_keep_each_lexical_boundary() {
    let source = "=LOG10 (A1)+BIN2DEC\t(\"101\")+LOG10";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("mixed function and cell-shaped names should parse");
    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 11);
    assert!(matches!(&formula.tokens[0], Token::Function(name) if name == "LOG10"));
    assert!(matches!(formula.tokens[1], Token::LParen));
    assert_cell(&formula.tokens[2], None, "A", 1, false, false);
    assert!(matches!(formula.tokens[3], Token::RParen));
    assert!(matches!(formula.tokens[4], Token::Operator('+')));
    assert!(matches!(&formula.tokens[5], Token::Function(name) if name == "BIN2DEC"));
    assert!(matches!(formula.tokens[6], Token::LParen));
    assert!(matches!(&formula.tokens[7], Token::String(value) if value == "101"));
    assert!(matches!(formula.tokens[8], Token::RParen));
    assert!(matches!(formula.tokens[9], Token::Operator('+')));
    assert_cell(&formula.tokens[10], None, "LOG", 10, false, false);
}

#[test]
fn legacy_sheet_names_spaces_and_lowercase_absolute_columns_are_preserved() {
    let source = "=Sales Data.$a$1+Summary Sheet.b$2+Data.c3";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("space-bearing sheet names and lowercase columns should parse");
    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 5);
    assert_cell(&formula.tokens[0], Some("Sales Data"), "A", 1, true, true);
    assert!(matches!(formula.tokens[1], Token::Operator('+')));
    assert_cell(
        &formula.tokens[2],
        Some("Summary Sheet"),
        "B",
        2,
        false,
        true,
    );
    assert!(matches!(formula.tokens[3], Token::Operator('+')));
    assert_cell(&formula.tokens[4], Some("Data"), "C", 3, false, false);

    // The legacy scanner treats every alphanumeric/space run before a dot as
    // a sheet name. Keep that behavior when the run begins with a cell-shaped
    // prefix; the compact common path must not split these into two tokens.
    for (source, expected_sheet) in [("=A1 Sheet.B2", "A1 Sheet"), ("=A1 .B2", "A1 ")] {
        let formula = FormulaParser::new(source)
            .parse()
            .expect("legacy spaced sheet scanner should preserve the full prefix");
        assert_eq!(formula.text, source);
        assert_eq!(formula.tokens.len(), 1);
        assert_cell(
            &formula.tokens[0],
            Some(expected_sheet),
            "B",
            2,
            false,
            false,
        );
    }
}

#[test]
fn malformed_identifier_suffixes_remain_typed_syntax_errors() {
    for source in ["=LOG10X(A1)", "=BIN2DEC_1(A1)", "=VLOOKUPX (A1)"] {
        let error = FormulaParser::new(source)
            .parse()
            .expect_err("unknown function suffix should not be reclassified as a valid call");
        assert!(
            matches!(error, Error::InvalidFormat(_)),
            "unexpected error for {source:?}: {error}"
        );
    }
}

#[test]
fn speculative_cell_allocation_failure_is_not_swallowed_by_fallback() {
    // Each spelling first retains the original formula, then reserves the
    // one-byte normalized column. The compact form and the two legacy forms
    // exercise both parser paths while keeping failure independent of the
    // allocator ordinal and unable to abort unrelated paths.
    for source in ["=A1", "=.A1", "=$A$1"] {
        arm_one_byte_failure();
        let result = FormulaParser::new(source).parse();
        let fired = injection_fired();
        disarm_one_byte_failure();

        assert!(
            fired,
            "the targeted one-byte allocation was not exercised for {source}"
        );
        let error = result.expect_err("the injected column allocation must fail");
        let Error::Allocation { resource, .. } = error else {
            panic!("speculative cell parse swallowed allocation failure for {source}: {error}");
        };
        assert_eq!(resource, "formula column");
    }
}
