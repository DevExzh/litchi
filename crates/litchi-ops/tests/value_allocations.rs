//! Allocation evidence for `litchi_ops::value`.
//!
//! The unified-operations performance bar requires value coercion to
//! allocate nothing on the success path (goal, "Allocation discipline";
//! record 0010 §7.3). A counting global allocator pins that property for
//! the listed coercion cases (excluding long grouped-number normalization), for the canonical renderings written into a fixed buffer,
//! and for query comparison; it also pins that building a failed parse's
//! error is free, so probing parses in a query cost nothing.

#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;
use std::fmt::{self, Write as _};

use litchi_core::selector::CompareOp;
use litchi_ops::value::{
    Color, ColorContext, Decimal, DotNetDouble, Length, LengthContext, LengthSpec, LengthStyle,
    Rgba, Scalar, Test, Tokens, ValueType, parse_bool, parse_integer, parse_number,
    parse_numeric_text, parse_opacity,
};

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

/// A fixed buffer to write canonical text into without a heap.
struct Buffer {
    bytes: [u8; 256],
    len: usize,
}

impl Buffer {
    const fn new() -> Self {
        Self {
            bytes: [0; 256],
            len: 0,
        }
    }
}

impl fmt::Write for Buffer {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        let end = self.len + text.len();
        self.bytes
            .get_mut(self.len..end)
            .ok_or(fmt::Error)?
            .copy_from_slice(text.as_bytes());
        self.len = end;
        Ok(())
    }
}

const LENGTHS: &[(&str, LengthContext)] = &[
    ("914400", LengthContext::Emu),
    ("-914400", LengthContext::Emu),
    ("9.144e5", LengthContext::Emu),
    ("914400emu", LengthContext::Emu),
    ("2.54cm", LengthContext::Emu),
    ("25.4mm", LengthContext::Emu),
    ("1in", LengthContext::Emu),
    ("6pc", LengthContext::Emu),
    ("72pt", LengthContext::Emu),
    ("96px", LengthContext::Emu),
    ("101.6Q", LengthContext::Emu),
    (" 1 IN ", LengthContext::Emu),
    ("2147483647", LengthContext::Extent),
    ("10cm", LengthContext::Extent),
    ("1.5", LengthContext::LineWidth),
    ("2pt", LengthContext::LineWidth),
    ("11", LengthContext::FontSize),
    ("10.5pt", LengthContext::FontSize),
];

const COLORS: &[(&str, ColorContext)] = &[
    ("FF0000", ColorContext::Drawing),
    ("#f00", ColorContext::Drawing),
    ("#F00A", ColorContext::Drawing),
    ("80FF0000", ColorContext::Drawing),
    ("#FF000080", ColorContext::Drawing),
    ("red", ColorContext::Drawing),
    ("CornflowerBlue", ColorContext::Drawing),
    ("transparent", ColorContext::Drawing),
    ("dkBlue", ColorContext::Drawing),
    ("rgb(255, 0, 0)", ColorContext::Drawing),
    ("rgba(255 0 0 / 50%)", ColorContext::Drawing),
    ("hsl(240deg 100% 50%)", ColorContext::Drawing),
    ("hsla(0, 100%, 50%, 0.5)", ColorContext::Drawing),
    ("accent1", ColorContext::Drawing),
    ("accent1+lumMod75+lumOff25", ColorContext::Drawing),
    ("accent1:lumMod75", ColorContext::Drawing),
    ("dk1+lumMod=75500", ColorContext::Drawing),
    ("windowText", ColorContext::Drawing),
    ("accent1+tint40", ColorContext::Spreadsheet),
    ("lt1+tint-25", ColorContext::Spreadsheet),
    ("FFFF0000", ColorContext::Spreadsheet),
    ("auto", ColorContext::Word),
    ("windowText", ColorContext::Word),
    ("1F4E79", ColorContext::Word),
];

static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &["single", "double", "none"],
    &[("dbl", 1), ("sgl", 0)],
);

#[test]
fn the_counter_sees_allocations() {
    assert_eq!(allocations(|| String::from("x")), 1);
    let count = allocations(|| litchi_ops::Error::from(parse_bool("maybe").unwrap_err()));
    assert!(count > 0, "reporting an error allocates its message");
}

#[test]
fn coercions_allocate_nothing_on_success() {
    for &(text, context) in LENGTHS {
        let count = allocations(|| Length::parse_in(text, context).expect("parses"));
        assert_eq!(count, 0, "length {text:?} allocated {count} times");
    }
    for &(text, context) in COLORS {
        let count = allocations(|| Color::parse(text, context).expect("parses"));
        assert_eq!(count, 0, "color {text:?} allocated {count} times");
    }
    let count = allocations(|| {
        (
            parse_bool(" Yes ").expect("parses"),
            parse_number("-1.5e2").expect("parses"),
            parse_integer(" 42 ").expect("parses"),
            parse_numeric_text("(1,234.5)").expect("parses"),
            parse_opacity("30").expect("parses"),
            UNDERLINE.parse(" SGL ").expect("parses"),
            Rgba::named("rebeccapurple").expect("named"),
            Rgba::preset("medAquamarine").expect("preset"),
        )
    });
    assert_eq!(count, 0, "scalars allocated {count} times");
    for ty in [
        ValueType::Bool,
        ValueType::Length(LengthSpec::LINE_WIDTH),
        ValueType::Color(ColorContext::Drawing),
    ] {
        let text = match ty {
            ValueType::Bool => "on",
            ValueType::Length(_) => "0.75",
            _ => "accent2+shade50",
        };
        let count = allocations(|| ty.coerce(text).expect("coerces"));
        assert_eq!(count, 0, "{ty:?} coercion allocated {count} times");
    }
}

#[test]
fn canonical_text_allocates_nothing() {
    let lengths = [0, 1, 1_800, 914_400, 12_192_000, -63, i64::MAX];
    let styles = [
        LengthStyle::Auto,
        LengthStyle::PointsFirst,
        LengthStyle::Lossy,
        LengthStyle::LineWidth,
        LengthStyle::Points,
    ];
    for emu in lengths {
        for style in styles {
            let count = allocations(|| {
                let mut buffer = Buffer::new();
                write!(buffer, "{}", Length::from_emu(emu).display(style)).expect("fits");
                buffer.len
            });
            assert_eq!(count, 0, "{emu} in {style:?} allocated {count} times");
        }
    }
    for &(text, context) in COLORS {
        let color = Color::parse(text, context).expect("parses");
        let count = allocations(|| {
            let mut buffer = Buffer::new();
            write!(buffer, "{color}").expect("fits");
            buffer.len
        });
        assert_eq!(count, 0, "{text:?} printing allocated {count} times");
    }
    let count = allocations(|| {
        let mut buffer = Buffer::new();
        write!(
            buffer,
            "{} {} {} {}",
            DotNetDouble(1e20),
            DotNetDouble(0.000_123),
            Decimal::from_ratio(265, 240, 4).expect("valid"),
            Scalar::Bool(true)
        )
        .expect("fits");
        buffer.len
    });
    assert_eq!(count, 0, "numbers allocated {count} times");
}

#[test]
fn query_comparison_allocates_nothing() {
    let width = Some(ValueType::Length(LengthSpec::EXTENT));
    let fill = Some(ValueType::Color(ColorContext::Drawing));
    let bold = Some(ValueType::Bool);
    let red = Color::rgb(255, 0, 0);
    let cases: &[(CompareOp, &str, Option<ValueType>, Scalar<'_>)] = &[
        (
            CompareOp::Ge,
            "1in",
            width,
            Scalar::Length(Length::from_emu(914_400)),
        ),
        (
            CompareOp::Lt,
            "10cm",
            width,
            Scalar::Length(Length::from_emu(914_400)),
        ),
        (
            CompareOp::Eq,
            "2.54cm",
            width,
            Scalar::Length(Length::from_emu(914_400)),
        ),
        (
            CompareOp::Ne,
            "720000",
            width,
            Scalar::Length(Length::from_emu(914_400)),
        ),
        (CompareOp::Eq, "red", fill, Scalar::Color(red)),
        (CompareOp::Eq, "#FF0000", fill, Scalar::Color(red)),
        (CompareOp::Contains, "ff00", fill, Scalar::Color(red)),
        (CompareOp::Eq, "yes", bold, Scalar::Bool(true)),
        (CompareOp::Gt, "5,000", None, Scalar::Text("12,500")),
        (CompareOp::Eq, "heading1", None, Scalar::Text("Heading1")),
        (
            CompareOp::Contains,
            "ÜBER",
            None,
            Scalar::Text("Überschrift 1"),
        ),
        (CompareOp::Gt, "abc", None, Scalar::Text("abd")),
        (CompareOp::Le, "1e3", None, Scalar::Number(999.5)),
        (CompareOp::Eq, "7", None, Scalar::Int(7)),
    ];
    for &(op, literal, ty, actual) in cases {
        let count = allocations(|| {
            let test = Test::with_text(op, literal, ty);
            (test.eval(Some(actual)), test.eval(None), test.warning())
        });
        assert_eq!(count, 0, "[{op} {literal}] allocated {count} times");
    }
}

#[test]
fn failed_parses_are_free_until_reported() {
    // Building a failed parse's error allocates nothing; its message and
    // suggestion are computed (and may allocate) only when asked for.
    let count = allocations(|| {
        (
            Length::parse_in("12 apples", LengthContext::Emu).is_err(),
            Length::parse_in("2em", LengthContext::Emu).is_err(),
            Length::parse_in("-1", LengthContext::Extent).is_err(),
            Color::parse("reed", ColorContext::Drawing).is_err(),
            Color::parse("accent1+lumMd75", ColorContext::Drawing).is_err(),
            Color::parse("rgb(300, 0, 0)", ColorContext::Word).is_err(),
            parse_bool("maybe").is_err(),
            parse_number("1,000").is_err(),
            UNDERLINE.parse("wavy").is_err(),
        )
    });
    assert_eq!(count, 0, "failed parses allocated {count} times");
}
