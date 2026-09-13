//! Regression coverage for bounded OpenFormula string-literal parsing.
//!
//! The parser must retain decoded literal values and source text while keeping
//! per-literal capacity proportional to the value. OpenFormula permits control
//! characters other than NUL inside a string, and represents a quote by `""`.

#![allow(
    unsafe_code,
    reason = "The integration-only allocator injects one fallible failure."
)]

use litchi_core::{Error, Resource};
use litchi_ods::codec::formula::{FormulaLimits, FormulaParser, Token};
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

thread_local! {
    static FAIL_LITERAL_SIZE: Cell<bool> = const { Cell::new(false) };
    static INJECTION_FIRED: Cell<bool> = const { Cell::new(false) };
}

struct LiteralFailureAllocator;

#[global_allocator]
static GLOBAL: LiteralFailureAllocator = LiteralFailureAllocator;

// SAFETY: All non-injected operations delegate unchanged to the system
// allocator. The injected null result is scoped to this test thread and is
// consumed once, so it exercises only a targeted literal reservation.
unsafe impl GlobalAlloc for LiteralFailureAllocator {
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
    FAIL_LITERAL_SIZE.with(|armed| {
        if armed.get() && matches!(size, 1 | 2) {
            armed.set(false);
            INJECTION_FIRED.with(|fired| fired.set(true));
            true
        } else {
            false
        }
    })
}

fn arm_literal_failure() {
    INJECTION_FIRED.with(|fired| fired.set(false));
    FAIL_LITERAL_SIZE.with(|armed| armed.set(true));
}

fn disarm_literal_failure() {
    FAIL_LITERAL_SIZE.with(|armed| armed.set(false));
}

fn injection_fired() -> bool {
    INJECTION_FIRED.with(Cell::get)
}

fn string_tokens(formula: &litchi_ods::codec::formula::Formula) -> impl Iterator<Item = &String> {
    formula.tokens.iter().filter_map(|token| match token {
        Token::String(value) => Some(value),
        Token::CellRef(_)
        | Token::RangeRef(_)
        | Token::Reference(_)
        | Token::Function(_)
        | Token::Number(_)
        | Token::Boolean(_)
        | Token::Operator(_)
        | Token::LParen
        | Token::RParen
        | Token::Comma
        | Token::Semicolon => None,
    })
}

#[test]
fn many_short_and_empty_utf8_literals_retain_bounded_capacity() {
    for count in [256, 1024] {
        let mut source = String::from("=");
        for index in 0..count {
            if index != 0 {
                source.push(';');
            }
            match index % 4 {
                0 => source.push_str("\"\""),
                1 => source.push_str("\"x\""),
                2 => source.push_str("\"é\""),
                _ => source.push_str("\"🌟\""),
            }
        }

        let formula = FormulaParser::new(&source)
            .parse()
            .expect("short and empty UTF-8 literals should parse");
        let values: Vec<_> = string_tokens(&formula).collect();
        assert_eq!(values.len(), count);

        let mut decoded_bytes = 0_usize;
        let mut retained_capacity = 0_usize;
        for value in values {
            decoded_bytes += value.len();
            retained_capacity += value.capacity();
            if value.is_empty() {
                assert_eq!(value.capacity(), 0, "empty literal retained a buffer");
            } else {
                assert!(
                    value.capacity() <= value.len().saturating_mul(4).saturating_add(8),
                    "literal capacity {} grew beyond decoded length {}",
                    value.capacity(),
                    value.len()
                );
            }
        }

        // This is deliberately a linear bound. The former reservation of all
        // remaining input for every literal grows quadratically with `count`.
        let linear_bound = decoded_bytes
            .saturating_mul(8)
            .saturating_add(count.saturating_mul(16));
        assert!(
            retained_capacity <= linear_bound,
            "retained literal capacity {retained_capacity} exceeded linear bound {linear_bound}"
        );
    }
}

#[test]
fn doubled_quotes_multibyte_values_and_source_text_are_exact() {
    let source = "  OF:=\"α🌟\"\"quoted\"+\"é\"\"Ω\"+\"\"  ";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("doubled quotes and multibyte literals should parse");

    assert_eq!(formula.text, source);
    let values: Vec<&str> = string_tokens(&formula).map(String::as_str).collect();
    assert_eq!(values, ["α🌟\"quoted", "é\"Ω", ""]);
    assert_eq!(formula.tokens.len(), 5);
    assert!(matches!(formula.tokens[1], Token::Operator('+')));
    assert!(matches!(formula.tokens[3], Token::Operator('+')));
}

#[test]
fn nul_is_rejected_while_non_nul_controls_remain_valid() {
    for source in ["=\"a\0\"\"b\"", "=\"a\"\"\0b\"", "=\"é\0\""] {
        let error = FormulaParser::new(source)
            .parse()
            .expect_err("NUL is excluded by the OpenFormula string grammar");
        assert!(
            matches!(error, Error::InvalidFormat(_)),
            "unexpected NUL error for {source:?}: {error}"
        );
    }

    let source = "=\"line\n\t\u{1}é\"";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("non-NUL controls are valid string content");
    assert_eq!(formula.text, source);
    assert!(matches!(
        &formula.tokens[0],
        Token::String(value) if value == "line\n\t\u{1}é"
    ));
}

#[test]
fn unterminated_and_malformed_quotes_remain_typed_syntax_errors() {
    for source in ["=\"unterminated", "=\"a\"\"b", "=\"a\"b"] {
        let error = FormulaParser::new(source)
            .parse()
            .expect_err("malformed string quote should be rejected");
        assert!(
            matches!(error, Error::InvalidFormat(_)),
            "unexpected quote error for {source:?}: {error}"
        );
    }

    let formula = FormulaParser::new("=\"\"")
        .parse()
        .expect("an empty quoted literal is valid");
    assert!(matches!(&formula.tokens[0], Token::String(value) if value.is_empty()));
}

#[test]
fn byte_and_token_limits_are_exact_for_string_formulas() {
    let source = "=\"é\"";
    let exact = FormulaLimits::default().with_max_bytes(source.len());
    let formula = FormulaParser::new(source)
        .parse_with_limits(&exact)
        .expect("a formula at its byte limit should be accepted");
    assert_eq!(formula.text, source);

    let too_small = FormulaLimits::default().with_max_bytes(source.len() - 1);
    let error = FormulaParser::new(source)
        .parse_with_limits(&too_small)
        .expect_err("one byte over the formula limit should be refused");
    let Error::ResourceLimit(limit) = error else {
        panic!("expected a byte resource limit, got {error:?}");
    };
    assert_eq!(limit.resource, Resource::InputBytes);
    assert_eq!(limit.observed, source.len() as u64);
    assert_eq!(limit.limit, (source.len() - 1) as u64);

    let source = "=\"x\";\"y\"";
    let exact = FormulaLimits::default().with_max_tokens(3);
    let formula = FormulaParser::new(source)
        .parse_with_limits(&exact)
        .expect("three tokens at the token limit should be accepted");
    assert_eq!(formula.tokens.len(), 3);

    let too_small = FormulaLimits::default().with_max_tokens(2);
    let error = FormulaParser::new(source)
        .parse_with_limits(&too_small)
        .expect_err("the third token should exceed the token limit");
    let Error::ResourceLimit(limit) = error else {
        panic!("expected a token resource limit, got {error:?}");
    };
    assert_eq!(limit.resource, Resource::Objects);
    assert_eq!(limit.observed, 3);
    assert_eq!(limit.limit, 2);
}

#[test]
fn literal_allocation_failure_is_typed_and_targeted() {
    // The original formula text allocation is four bytes. The one-byte
    // decoded literal allocation in the fixed parser (or two-byte remaining
    // content reservation in the baseline) is the only targeted request.
    arm_literal_failure();
    let result = FormulaParser::new("=\"x\"").parse();
    let fired = injection_fired();
    disarm_literal_failure();

    assert!(fired, "the targeted literal allocation was not exercised");
    let error = result.expect_err("the injected literal reservation must fail");
    let Error::Allocation { resource, .. } = error else {
        panic!("literal allocation failure was converted to another error: {error}");
    };
    assert_eq!(resource, "formula string literal");
}
