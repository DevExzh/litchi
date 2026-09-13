//! Integration coverage for the strict OpenFormula 1.4 expression grammar.
//!
//! These tests inspect the parser's immutable syntax tree and source spans. They
//! do not evaluate formulas, resolve names, open external IRIs, or make claims
//! about function arity beyond the syntax rules in Part 4 sections 5.1--5.14.

#![allow(
    unsafe_code,
    reason = "The integration-only allocator injects fallible failures."
)]

use litchi_core::Error;
use litchi_ods::codec::formula::expression::{
    Expression, HARD_MAX_EXPRESSION_DEPTH, InfixOperator, Kind, Limits, NameScope, Node,
    PostfixOperator, PrefixOperator,
};
use litchi_ods::codec::formula::reference::{
    Address, Endpoint, EndpointValue, Reference, SheetSelector, Subtable,
};
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;
use std::process::Command;

thread_local! {
    static COUNT_ALLOCATIONS: Cell<bool> = const { Cell::new(false) };
    static ALLOCATION_COUNT: Cell<usize> = const { Cell::new(0) };
    static FAIL_AT: Cell<Option<usize>> = const { Cell::new(None) };
    static INJECTION_FIRED: Cell<bool> = const { Cell::new(false) };
}

struct OneShotFailureAllocator;

#[global_allocator]
static GLOBAL: OneShotFailureAllocator = OneShotFailureAllocator;

// SAFETY: All non-injected calls delegate unchanged to the system allocator.
// The injected null result is scoped to this test thread and consumed once.
unsafe impl GlobalAlloc for OneShotFailureAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if should_fail() {
            std::ptr::null_mut()
        } else {
            // SAFETY: The caller supplies a valid allocation layout.
            unsafe { System.alloc(layout) }
        }
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        if should_fail() {
            std::ptr::null_mut()
        } else {
            // SAFETY: The caller supplies a valid allocation layout.
            unsafe { System.alloc_zeroed(layout) }
        }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: The pointer/layout pair comes from a prior allocator call.
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        if should_fail() {
            std::ptr::null_mut()
        } else {
            // SAFETY: The pointer/layout pair and new size are supplied by the
            // allocator caller.
            unsafe { System.realloc(pointer, layout, size) }
        }
    }
}

fn should_fail() -> bool {
    COUNT_ALLOCATIONS.with(|counting| {
        if !counting.get() {
            return false;
        }
        let index = ALLOCATION_COUNT.with(|count| {
            let index = count.get();
            count.set(index.saturating_add(1));
            index
        });
        FAIL_AT.with(|target| {
            if target.get() == Some(index) {
                target.set(None);
                INJECTION_FIRED.with(|fired| fired.set(true));
                true
            } else {
                false
            }
        })
    })
}

fn begin_allocation_count() {
    ALLOCATION_COUNT.with(|count| count.set(0));
    FAIL_AT.with(|target| target.set(None));
    INJECTION_FIRED.with(|fired| fired.set(false));
    COUNT_ALLOCATIONS.with(|counting| counting.set(true));
}

fn end_allocation_count() -> usize {
    COUNT_ALLOCATIONS.with(|counting| counting.set(false));
    ALLOCATION_COUNT.with(Cell::get)
}

fn arm_allocation_failure(index: usize) {
    ALLOCATION_COUNT.with(|count| count.set(0));
    FAIL_AT.with(|target| target.set(Some(index)));
    INJECTION_FIRED.with(|fired| fired.set(false));
    COUNT_ALLOCATIONS.with(|counting| counting.set(true));
}

fn injection_fired() -> bool {
    INJECTION_FIRED.with(Cell::get)
}

fn assert_number(node: &Node<'_>) {
    assert!(
        matches!(node.kind(), Kind::Number),
        "expected number: {node:?}"
    );
}

fn assert_string(node: &Node<'_>) {
    assert!(
        matches!(node.kind(), Kind::String),
        "expected string: {node:?}"
    );
}

fn assert_reference(node: &Node<'_>) {
    assert!(
        matches!(node.kind(), Kind::Reference(_)),
        "expected reference: {node:?}"
    );
}

fn assert_array_row(node: &Node<'_>) {
    assert!(
        matches!(node.kind(), Kind::ArrayRow),
        "expected array row: {node:?}"
    );
}

fn assert_cell_endpoint(
    endpoint: &Endpoint,
    column: &str,
    row: u32,
    column_absolute: bool,
    row_absolute: bool,
) {
    let EndpointValue::Cell(cell) = &endpoint.value else {
        panic!("expected a cell endpoint, got {:?}", endpoint.value);
    };
    assert_eq!(cell.column.label, column);
    assert_eq!(cell.column.absolute, column_absolute);
    assert_eq!(cell.row.number, row);
    assert_eq!(cell.row.absolute, row_absolute);
}

fn assert_current_cell_reference(
    node: &Node<'_>,
    expected_text: &str,
    column: &str,
    row: u32,
    column_absolute: bool,
    row_absolute: bool,
) {
    assert_reference(node);
    assert_eq!(node.text(), expected_text);
    let reference = node
        .reference()
        .expect("reference node should expose its parsed value");
    match reference {
        Reference::Local(Address::Cell(endpoint)) => {
            assert!(matches!(&endpoint.sheet, SheetSelector::Current));
            assert_cell_endpoint(endpoint, column, row, column_absolute, row_absolute);
        },
        other => panic!("expected a current-sheet cell reference, got {other:?}"),
    }
}

fn parse(source: &str) -> Expression {
    match Expression::parse(source) {
        Ok(expression) => expression,
        Err(error) => panic!("{source:?} should parse: {error}"),
    }
}

fn assert_invalid(source: &str) {
    let error = Expression::parse(source)
        .err()
        .unwrap_or_else(|| panic!("{source:?} should be rejected"));
    assert!(
        matches!(error, Error::InvalidFormat(_)),
        "{source:?} returned an unexpected syntax error: {error}"
    );
}

fn assert_resource_limit(source: &str, limits: &Limits) {
    let error = Expression::parse_with_limits(source, limits)
        .err()
        .unwrap_or_else(|| panic!("{source:?} should exceed its explicit limit"));
    assert!(
        matches!(error, Error::ResourceLimit(_)),
        "{source:?} returned a non-resource error: {error}"
    );
}

#[test]
fn precedence_and_associativity_are_retained_in_the_tree() {
    let expression = parse("=1+2*3");
    assert_eq!(expression.source(), "=1+2*3");
    let root = expression.root();
    assert_eq!(root.kind(), Kind::Infix(InfixOperator::Add));
    let children: Vec<_> = root.children().collect();
    assert_eq!(children.len(), 2);
    assert_number(&children[0]);
    assert_eq!(children[0].text(), "1");
    assert_eq!(children[1].kind(), Kind::Infix(InfixOperator::Multiply));
    assert_eq!(children[1].text(), "2*3");

    // §5.5 makes power left associative and puts prefix unary operators above
    // power. Parentheses must still create an observable tree node.
    let power = parse("=2^3^2");
    let power_root = power.root();
    assert_eq!(power_root.kind(), Kind::Infix(InfixOperator::Power));
    let power_children: Vec<_> = power_root.children().collect();
    assert_eq!(power_children.len(), 2);
    assert_eq!(power_children[0].kind(), Kind::Infix(InfixOperator::Power));
    assert_eq!(power_children[0].text(), "2^3");
    assert_number(&power_children[1]);

    let unary = parse("=-2^2");
    let unary_root = unary.root();
    assert_eq!(unary_root.kind(), Kind::Infix(InfixOperator::Power));
    let unary_children: Vec<_> = unary_root.children().collect();
    assert_eq!(
        unary_children[0].kind(),
        Kind::Prefix(PrefixOperator::Minus)
    );
    assert_number(&unary_children[1]);

    let parenthesized = parse("=-(2^2)");
    let parenthesized_root = parenthesized.root();
    assert_eq!(
        parenthesized_root.kind(),
        Kind::Prefix(PrefixOperator::Minus)
    );
    let prefix_child = parenthesized_root
        .child(0)
        .expect("prefix must have one operand");
    assert!(matches!(prefix_child.kind(), Kind::Parenthesized));
    assert_eq!(prefix_child.text(), "(2^2)");

    let postfix = parse("=50%+1");
    let postfix_root = postfix.root();
    assert_eq!(postfix_root.kind(), Kind::Infix(InfixOperator::Add));
    let postfix_children: Vec<_> = postfix_root.children().collect();
    assert_eq!(
        postfix_children[0].kind(),
        Kind::Postfix(PostfixOperator::Percent)
    );
    assert_number(&postfix_children[1]);
}

#[test]
fn reference_operators_follow_their_distinct_precedence() {
    let expression = parse("=[.A1:.B2]![.A1]~[.B2]");
    let root = expression.root();
    assert_eq!(root.kind(), Kind::Infix(InfixOperator::Union));
    let union_children: Vec<_> = root.children().collect();
    assert_eq!(union_children.len(), 2);
    assert_eq!(
        union_children[0].kind(),
        Kind::Infix(InfixOperator::Intersection)
    );
    assert_eq!(union_children[0].text(), "[.A1:.B2]![.A1]");
    assert_reference(&union_children[1]);
    let intersection_children: Vec<_> = union_children[0].children().collect();
    assert_eq!(intersection_children.len(), 2);
    assert_reference(&intersection_children[0]);
    assert_reference(&intersection_children[1]);
}

#[test]
fn function_calls_have_strict_parenthesis_and_separator_grammar() {
    let zero = parse("=SUM()");
    let zero_root = zero.root();
    assert!(matches!(zero_root.kind(), Kind::Function { name: "SUM" }));
    assert_eq!(zero_root.function_name(), Some("SUM"));
    assert_eq!(zero_root.children().count(), 0);

    let slots = parse("=SUM(;1;)");
    let slot_root = slots.root();
    assert!(matches!(slot_root.kind(), Kind::Function { name: "SUM" }));
    let arguments: Vec<_> = slot_root.children().collect();
    assert_eq!(arguments.len(), 3);
    assert!(matches!(arguments[0].kind(), Kind::Missing));
    assert_number(&arguments[1]);
    assert!(matches!(arguments[2].kind(), Kind::Missing));

    let separated = parse("=SUM(1;2)");
    assert_eq!(separated.root().children().count(), 2);

    let all_missing = parse("=SUM(;)");
    let all_missing_arguments: Vec<_> = all_missing.root().children().collect();
    assert_eq!(all_missing_arguments.len(), 2);
    for argument in &all_missing_arguments {
        assert!(matches!(argument.kind(), Kind::Missing));
    }

    for source in ["=SUM (1)", "=SUM(1,2)", "=SUM(,1)", "=SUM(1,)", "=SUM(1 2)"] {
        assert_invalid(source);
    }

    // A name may be a function name lexically; the call remains syntactically
    // valid even when the evaluator has no implementation for it.
    let cell_shaped_call = parse("=A1()");
    assert!(matches!(
        cell_shaped_call.root().kind(),
        Kind::Function { name: "A1" }
    ));
}

#[test]
fn inline_arrays_preserve_rows_columns_and_nested_expression_nodes() {
    let rectangular = parse("={1;2|3;4}");
    let root = rectangular.root();
    let dimensions = match root.kind() {
        Kind::Array(dimensions) => dimensions,
        kind => panic!("expected array, got {kind:?}"),
    };
    assert_eq!(dimensions.rows(), 2);
    assert_eq!(dimensions.columns(), Some(2));
    assert!(dimensions.is_rectangular());
    let rows: Vec<_> = root.children().collect();
    assert_eq!(rows.len(), 2);
    for row in &rows {
        assert_array_row(row);
        assert_eq!(row.children().count(), 2);
    }

    let nested = parse("={SUM(1;2);\"α🌟\"|3}");
    let nested_rows: Vec<_> = nested.root().children().collect();
    assert_eq!(nested_rows.len(), 2);
    assert_eq!(nested_rows[0].children().count(), 2);
    let first_row: Vec<_> = nested_rows[0].children().collect();
    assert!(matches!(
        first_row[0].kind(),
        Kind::Function { name: "SUM" }
    ));
    assert_string(&first_row[1]);

    // §5.13 describes syntax for nonempty rows; the rectangular restriction is
    // an evaluator acceptance rule. The syntax tree must retain a ragged array
    // instead of silently padding or discarding its last cell.
    let ragged = parse("={1;2|3}");
    let ragged_dimensions = ragged
        .root()
        .array_dimensions()
        .expect("ragged array dimensions");
    assert_eq!(ragged_dimensions.rows(), 2);
    assert_eq!(ragged_dimensions.columns(), None);
    assert_eq!(ragged_dimensions.min_columns(), 1);
    assert_eq!(ragged_dimensions.max_columns(), 2);
    let ragged_rows: Vec<_> = ragged.root().children().collect();
    assert_eq!(ragged_rows.len(), 2);
    assert_eq!(ragged_rows[0].children().count(), 2);
    assert_eq!(ragged_rows[1].children().count(), 1);

    for source in ["={}", "={|1}", "={1;}", "={1||2}"] {
        assert_invalid(source);
    }
}

#[test]
fn names_sheets_sources_and_labels_are_inert_typed_nodes() {
    let name = parse("='Sheet One'.Total");
    let name_root = name.root();
    assert!(matches!(
        name_root.kind(),
        Kind::NamedExpression {
            name: "Total",
            scope: NameScope::SheetLocal {
                sheet: "Sheet One",
                ..
            }
        }
    ));
    assert_eq!(name_root.name(), Some("Total"));

    let forced_name = parse("=$$Global");
    assert!(matches!(
        forced_name.root().kind(),
        Kind::NamedExpression {
            name: "Global",
            scope: NameScope::Simple { forced: true }
        }
    ));

    let external = parse("='file:///tmp/foreign.ods'#'Sheet One'.Total");
    let external_root = external.root();
    assert!(matches!(
        external_root.kind(),
        Kind::NamedExpression {
            name: "Total",
            scope: NameScope::External {
                source: "file:///tmp/foreign.ods",
                sheet: Some("Sheet One"),
                ..
            }
        }
    ));
    assert_eq!(external_root.name(), Some("Total"));

    // A quoted named-expression component is legal only with the forced-name
    // marker. Without `$$`, the same spelling is a malformed sheet/name
    // qualification and must not be reclassified as a label.
    assert_invalid("='Sheet'.'Name'");
    assert_invalid("='file:///tmp/foreign.ods'#'Sheet'.'Name'");

    let quoted_name = parse("='Sheet'.$$'Name'");
    assert!(matches!(
        quoted_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::SheetLocal {
                sheet: "Sheet",
                forced: true,
                ..
            }
        }
    ));
    let quoted_external_name = parse("='file:///tmp/foreign.ods'#'Sheet'.$$'Name'");
    assert!(matches!(
        quoted_external_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::External {
                source: "file:///tmp/foreign.ods",
                sheet: Some("Sheet"),
                forced: true,
                ..
            }
        }
    ));

    let local_reference = parse("=[Sheet:Two.A1]");
    let local_root = local_reference.root();
    assert_reference(&local_root);
    assert!(local_root.reference().is_some());

    let external_reference = parse("=['file:///tmp/foreign.ods'#'Sheet One'.A1]");
    let external_reference_root = external_reference.root();
    let reference = external_reference_root
        .reference()
        .expect("external reference node should expose its parsed reference");
    assert_eq!(
        reference.source().map(|source| source.as_str()),
        Some("file:///tmp/foreign.ods")
    );

    let label = parse("='Column Label'");
    let label_root = label.root();
    assert!(matches!(label_root.kind(), Kind::QuotedLabel));
    assert_eq!(label_root.label(), Some("Column Label"));

    let intersection = parse("='Row'!!'Column'");
    assert!(matches!(
        intersection.root().kind(),
        Kind::AutomaticIntersection
    ));
    assert_eq!(intersection.root().children().count(), 2);
}

#[test]
fn whitespace_between_nonterminating_components_is_preserved_and_typed() {
    let forced_name = parse("=$$ Name");
    assert_eq!(forced_name.source(), "=$$ Name");
    assert_eq!(forced_name.root().text(), "$$ Name");
    assert!(matches!(
        forced_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::Simple { forced: true }
        }
    ));

    let forced_quoted_name = parse("=$$ 'Name'");
    assert_eq!(forced_quoted_name.source(), "=$$ 'Name'");
    assert_eq!(forced_quoted_name.root().text(), "$$ 'Name'");
    assert!(matches!(
        forced_quoted_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::Simple { forced: true }
        }
    ));

    let sheet_name = parse("='Sheet' . Name");
    assert_eq!(sheet_name.source(), "='Sheet' . Name");
    assert_eq!(sheet_name.root().text(), "'Sheet' . Name");
    assert!(matches!(
        sheet_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::SheetLocal {
                sheet: "Sheet",
                forced: false,
                ..
            }
        }
    ));

    let absolute_sheet_name = parse("=$ 'Sheet' . Name");
    assert_eq!(absolute_sheet_name.source(), "=$ 'Sheet' . Name");
    assert_eq!(absolute_sheet_name.root().text(), "$ 'Sheet' . Name");
    assert!(matches!(
        absolute_sheet_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::SheetLocal {
                sheet: "Sheet",
                absolute: true,
                forced: false,
            }
        }
    ));

    let external_name = parse("='file:///book.ods' # Name");
    assert_eq!(external_name.source(), "='file:///book.ods' # Name");
    assert_eq!(external_name.root().text(), "'file:///book.ods' # Name");
    assert!(matches!(
        external_name.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::External {
                source: "file:///book.ods",
                sheet: None,
                forced: false,
                ..
            }
        }
    ));

    let external_absolute_sheet = parse("='file:///book.ods' # $ 'Sheet' . Name");
    assert_eq!(
        external_absolute_sheet.source(),
        "='file:///book.ods' # $ 'Sheet' . Name"
    );
    assert_eq!(
        external_absolute_sheet.root().text(),
        "'file:///book.ods' # $ 'Sheet' . Name"
    );
    assert!(matches!(
        external_absolute_sheet.root().kind(),
        Kind::NamedExpression {
            name: "Name",
            scope: NameScope::External {
                source: "file:///book.ods",
                sheet: Some("Sheet"),
                sheet_absolute: true,
                forced: false,
            }
        }
    ));

    let spaced_cell = parse("=[ . $A$1 ]");
    assert_current_cell_reference(&spaced_cell.root(), "[ . $A$1 ]", "A", 1, true, true);

    let spaced_range = parse("=[ . $A$1 : . $B$2 ]");
    assert_eq!(spaced_range.source(), "=[ . $A$1 : . $B$2 ]");
    assert_eq!(spaced_range.root().text(), "[ . $A$1 : . $B$2 ]");
    match spaced_range.root().reference() {
        Some(Reference::Local(Address::Cells(start, end))) => {
            assert!(matches!(&start.sheet, SheetSelector::Current));
            assert!(matches!(&end.sheet, SheetSelector::Inherited));
            assert_cell_endpoint(start, "A", 1, true, true);
            assert_cell_endpoint(end, "B", 2, true, true);
        },
        other => panic!("expected a spaced current-sheet cell range, got {other:?}"),
    }

    let spaced_external_range = parse("=['file:///book.ods' # . $A$1 : . $B$2 ]");
    assert_eq!(
        spaced_external_range.source(),
        "=['file:///book.ods' # . $A$1 : . $B$2 ]"
    );
    assert_eq!(
        spaced_external_range.root().text(),
        "['file:///book.ods' # . $A$1 : . $B$2 ]"
    );
    match spaced_external_range.root().reference() {
        Some(Reference::Source {
            source,
            address: Address::Cells(start, end),
        }) => {
            assert_eq!(source.as_str(), "file:///book.ods");
            assert!(matches!(&start.sheet, SheetSelector::Current));
            assert!(matches!(&end.sheet, SheetSelector::Inherited));
            assert_cell_endpoint(start, "A", 1, true, true);
            assert_cell_endpoint(end, "B", 2, true, true);
        },
        other => panic!("expected a spaced external cell range, got {other:?}"),
    }

    let spaced_subtables = parse("=[ Sheet . A1 . 'Sub.Table' . B2 ]");
    assert_eq!(
        spaced_subtables.source(),
        "=[ Sheet . A1 . 'Sub.Table' . B2 ]"
    );
    assert_eq!(
        spaced_subtables.root().text(),
        "[ Sheet . A1 . 'Sub.Table' . B2 ]"
    );
    match spaced_subtables.root().reference() {
        Some(Reference::Local(Address::Cell(endpoint))) => {
            let SheetSelector::Explicit(locator) = &endpoint.sheet else {
                panic!("expected an explicit sheet locator: {:?}", endpoint.sheet);
            };
            assert_eq!(locator.sheet.name, "Sheet");
            assert_eq!(locator.subtables.len(), 2);
            assert!(matches!(
                &locator.subtables[0],
                Subtable::Cell(cell) if cell.column.label == "A" && cell.row.number == 1
            ));
            assert!(matches!(
                &locator.subtables[1],
                Subtable::Name(name) if name.name == "Sub.Table" && name.quoted
            ));
            assert_cell_endpoint(endpoint, "B", 2, false, false);
        },
        other => panic!("expected a spaced subtable reference, got {other:?}"),
    }

    let spaced_external_absolute = parse("=['file:///book.ods' # $ 'Sheet' . $A$1 ]");
    assert_eq!(
        spaced_external_absolute.source(),
        "=['file:///book.ods' # $ 'Sheet' . $A$1 ]"
    );
    assert_eq!(
        spaced_external_absolute.root().text(),
        "['file:///book.ods' # $ 'Sheet' . $A$1 ]"
    );
    match spaced_external_absolute.root().reference() {
        Some(Reference::Source {
            source,
            address: Address::Cell(endpoint),
        }) => {
            assert_eq!(source.as_str(), "file:///book.ods");
            let SheetSelector::Explicit(locator) = &endpoint.sheet else {
                panic!(
                    "expected an explicit quoted sheet locator: {:?}",
                    endpoint.sheet
                );
            };
            assert_eq!(locator.sheet.name, "Sheet");
            assert!(locator.sheet.absolute);
            assert!(locator.sheet.quoted);
            assert_cell_endpoint(endpoint, "A", 1, true, true);
        },
        other => panic!("expected a spaced external absolute cell, got {other:?}"),
    }

    for source in [
        "=$ $Name",
        "=Na me",
        "=SUM (1)",
        "=$ Sheet.Name",
        "=[.$ A1]",
        "=[.A$ 1]",
        "=\u{000b}1",
        "=1\u{000c}+2",
        "=[.\u{000b}A1]",
        "=[.A1\u{000c}]",
    ] {
        assert_invalid(source);
    }

    for (source, column, row, column_absolute, row_absolute) in [
        ("=[.A 1]", "A", 1, false, false),
        ("=[.$A $1]", "A", 1, true, true),
        ("=[.A $1]", "A", 1, false, true),
    ] {
        let expression = parse(source);
        assert_current_cell_reference(
            &expression.root(),
            &source[1..],
            column,
            row,
            column_absolute,
            row_absolute,
        );
    }
}

#[test]
fn numbers_strings_and_errors_keep_lexical_boundaries() {
    for source in ["=.5", "=1.25e+3", "=1e-9", "=0"] {
        let expression = parse(source);
        assert_number(&expression.root());
        assert_eq!(expression.root().text(), &source[1..]);
    }

    assert_invalid("=1.");
    assert_invalid("=1e+");

    let huge_lexeme = "9".repeat(4096);
    let huge_source = format!("={huge_lexeme}");
    let huge = parse(&huge_source);
    assert_number(&huge.root());
    assert_eq!(huge.root().text(), huge_lexeme);

    let string_source = "=\"α🌟\"\"quote\"";
    let string = parse(string_source);
    assert_string(&string.root());
    assert_eq!(string.root().text(), "\"α🌟\"\"quote\"");

    for source in ["=\"a\0b\"", "=\"unterminated", "=\"a\"b\""] {
        assert_invalid(source);
    }

    for source in ["=#N/A", "=#DIV/0!", "=#FOO?"] {
        let error = parse(source);
        assert!(matches!(error.root().kind(), Kind::Error));
        assert_eq!(error.root().text(), &source[1..]);
    }
    for source in ["=#BAD", "=#N/AA", "=#DIV/01!"] {
        assert_invalid(source);
    }
}

#[test]
fn exact_formula_source_and_forced_recalculation_spelling_are_retained() {
    let source = "of:== SUM(1; 2) ";
    let expression = parse(source);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
    assert_eq!(expression.root().text(), "SUM(1; 2)");

    let ordinary = parse("of:=SUM(1; 2)");
    assert!(!ordinary.is_force_recalculate());

    let whitespace = parse("=  1 +\t2\n");
    assert_eq!(whitespace.source(), "=  1 +\t2\n");
    assert_eq!(whitespace.root().text(), "1 +\t2");
}

#[test]
fn strict_expression_consumes_the_entire_input() {
    for source in [
        "=1 2", "=1+2 3", "=1)", "=(1", "=1@2", "=1,2", "=TRUE", "=FALSE", "=A1",
    ] {
        assert_invalid(source);
    }
}

#[test]
fn explicit_resource_limits_fail_atomically_at_the_boundary() {
    let source = "=1+2";
    let exact_bytes = Limits::default().with_max_bytes(source.len());
    assert_eq!(parse_with_limits(source, &exact_bytes).source(), source);

    let short_bytes = Limits::default().with_max_bytes(source.len() - 1);
    assert_resource_limit(source, &short_bytes);

    let exact_nodes = Limits::default().with_max_nodes(3);
    assert_eq!(parse_with_limits(source, &exact_nodes).source(), source);
    let short_nodes = Limits::default().with_max_nodes(2);
    assert_resource_limit(source, &short_nodes);

    let array_source = "={1;2|3;4}";
    let exact_cells = Limits::default().with_max_array_cells(4);
    assert_eq!(
        parse_with_limits(array_source, &exact_cells).source(),
        array_source
    );
    let short_cells = Limits::default().with_max_array_cells(3);
    assert_resource_limit(array_source, &short_cells);

    let shallow = Limits::default().with_max_depth(1);
    assert_eq!(parse_with_limits("=1", &shallow).source(), "=1");
    assert_resource_limit("=((1))", &shallow);
}

fn parse_with_limits(source: &str, limits: &Limits) -> Expression {
    match Expression::parse_with_limits(source, limits) {
        Ok(expression) => expression,
        Err(error) => panic!("{source:?} should parse: {error}"),
    }
}

#[test]
fn deep_parentheses_and_long_left_chains_are_bounded_without_recursive_drop() {
    let mut long = String::from("=");
    for index in 0..1024 {
        if index != 0 {
            long.push('+');
        }
        long.push('1');
    }
    let expression = parse(&long);
    let mut current = expression.root();
    let mut infix_count = 0;
    loop {
        let children: Vec<_> = current.children().collect();
        if children.len() != 2 || !matches!(current.kind(), Kind::Infix(InfixOperator::Add)) {
            break;
        }
        infix_count += 1;
        current = children.into_iter().next().expect("left child");
    }
    assert_eq!(infix_count, 1023);

    let hostile = format!("={}", "(".repeat(256) + "1" + &")".repeat(256));
    let limits = Limits::default().with_max_depth(64);
    assert_resource_limit(&hostile, &limits);
}

#[test]
fn hard_depth_cap_is_checked_in_a_subprocess_before_stack_overflow() {
    const CHILD_MARKER: &str = "LITCHI_ODS_EXPRESSION_DEPTH_CHILD";

    if std::env::var_os(CHILD_MARKER).is_some() {
        let depth = HARD_MAX_EXPRESSION_DEPTH + 32;
        let limits = Limits::default().with_max_depth(usize::MAX);
        let sources = [
            format!("={}1{}", "(".repeat(depth), ")".repeat(depth)),
            format!("={}1{}", "F(".repeat(depth), ")".repeat(depth)),
            format!("={}1{}", "{".repeat(depth), "}".repeat(depth)),
            format!("={}1", "-".repeat(depth)),
        ];

        for source in &sources {
            let error = Expression::parse_with_limits(source, &limits)
                .err()
                .unwrap_or_else(|| panic!("depth-hostile source unexpectedly parsed"));
            assert!(
                matches!(&error, Error::ResourceLimit(limit) if limit.resource == litchi_core::Resource::Depth),
                "depth-hostile source returned the wrong error: {error}"
            );
        }
        return;
    }

    let executable = std::env::current_exe().expect("test executable path");
    let status = Command::new(executable)
        .arg("--exact")
        .arg("hard_depth_cap_is_checked_in_a_subprocess_before_stack_overflow")
        .arg("--nocapture")
        .env(CHILD_MARKER, "1")
        .status()
        .expect("spawn depth-cap child test");
    assert!(
        status.success(),
        "depth-cap child test failed with status {status}"
    );
}

#[test]
fn malformed_short_inputs_are_bounded_and_panic_free() {
    const FRAGMENTS: &[&str] = &[
        "",
        "=",
        "of:",
        "of:=",
        "(",
        ")",
        "[",
        "]",
        "{",
        "}",
        "{}",
        "{|",
        ";",
        "|",
        ":",
        "!",
        "~",
        "+",
        "-",
        "^",
        "%",
        "&",
        "<",
        ">",
        "<=",
        "<>",
        "#",
        "#N/",
        "'",
        "\"",
        "\"\"",
        "A",
        "A1",
        ".5",
        "1.",
        "1e+",
        "SUM",
        "SUM(",
        "SUM (",
        "SUM;",
        "α",
        "🌟",
        "\0",
        "'a''b'",
        "[.A1]",
        "['x'#.A1]",
    ];
    let limits = Limits::default()
        .with_max_bytes(128)
        .with_max_nodes(32)
        .with_max_depth(8)
        .with_max_array_cells(16);

    for (index, left) in FRAGMENTS.iter().enumerate() {
        for right in FRAGMENTS.iter().skip(index % 7) {
            let source = format!("{left}{right}");
            let _ = Expression::parse_with_limits(&source, &limits);
        }
    }
}

#[test]
fn every_parser_reservation_propagates_a_single_injected_allocation_failure() {
    let source = "of:==SUM(1;[.A1];{2;3|4})";

    begin_allocation_count();
    let baseline = Expression::parse(source);
    let allocation_count = end_allocation_count();
    assert!(baseline.is_ok(), "valid mixed expression should parse");
    assert!(
        allocation_count > 0,
        "parser should make bounded reservations"
    );

    for index in 0..allocation_count {
        arm_allocation_failure(index);
        let result = Expression::parse(source);
        let _ = end_allocation_count();
        FAIL_AT.with(|target| target.set(None));
        assert!(
            injection_fired(),
            "allocation-failure injection did not fire at attempt {index} of {allocation_count}"
        );
        assert!(
            matches!(&result, Err(Error::Allocation { .. })),
            "attempt {index} was swallowed or changed into another error: {result:?}"
        );
    }
}
