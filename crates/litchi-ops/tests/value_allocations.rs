//! Allocation evidence for values and bounded XLS read projections.
//!
//! The unified-operations performance bar requires value coercion to
//! allocate nothing on the success path (goal, "Allocation discipline";
//! record 0010 §7.3). A counting global allocator pins that property for
//! the listed coercion cases (excluding long grouped-number normalization),
//! for the canonical renderings written into a fixed buffer, and for query
//! comparison; it also pins that building a failed parse's error is free, so
//! probing parses in a query costs nothing. XLS adapter tests additionally
//! track requested bytes so a growing vector cannot hide behind one allocation.

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
    static ALLOCATED_BYTES: Cell<usize> = const { Cell::new(0) };
}

fn record(layout: Layout) {
    let _ = ALLOCATIONS.try_with(|cell| cell.set(cell.get().wrapping_add(1)));
    let _ = ALLOCATED_BYTES.try_with(|cell| cell.set(cell.get().wrapping_add(layout.size())));
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: Every operation delegates to the platform allocator after recording
// counts and requested bytes; the allocator does not alter pointer ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        record(layout);
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc(layout) }
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        record(layout);
        // SAFETY: `layout` is supplied by the standard library allocator API.
        unsafe { System.alloc_zeroed(layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        record(Layout::from_size_align(new_size, layout.align()).unwrap_or(layout));
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
    allocation_measure(body).0
}

fn allocation_measure<T>(body: impl FnOnce() -> T) -> (usize, usize) {
    let before_count = ALLOCATIONS.with(Cell::get);
    let before_bytes = ALLOCATED_BYTES.with(Cell::get);
    let value = std::hint::black_box(body());
    let after_count = ALLOCATIONS.with(Cell::get);
    let after_bytes = ALLOCATED_BYTES.with(Cell::get);
    drop(value);
    (after_count - before_count, after_bytes - before_bytes)
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
    assert_eq!(allocation_measure(|| Vec::<u8>::with_capacity(64)), (1, 64));
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

#[cfg(feature = "xls")]
mod xls_allocations {
    use litchi_core::sheet::{Cell as _, Worksheet as _};
    use litchi_ops::{File, OpenOptions, SourceMode};
    use litchi_xls::writer::Builder;

    use super::allocation_measure;

    fn workbook_bytes(cell_count: usize, sheet_count: usize) -> Vec<u8> {
        assert!(sheet_count > 0, "the fixture needs a first worksheet");
        let mut builder = Builder::new();
        let data = builder.add_worksheet("Data").expect("data worksheet");
        for row in 0..cell_count {
            let row = u32::try_from(row).expect("BIFF row");
            if row == 0 {
                builder
                    .write_string(data, row, 0, "first")
                    .expect("string cell");
            } else {
                builder
                    .write_number(data, row, 0, f64::from(row))
                    .expect("number cell");
            }
        }
        for sheet in 1..sheet_count {
            let name = format!("Sheet{sheet}");
            builder.add_worksheet(&name).expect("catalog worksheet");
        }
        builder
            .build()
            .expect("build XLS fixture")
            .to_bytes()
            .expect("serialize XLS fixture")
    }

    fn eager(bytes: Vec<u8>) -> File {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("open eager XLS fixture")
    }

    fn first_cell(file: &File, limited: bool) -> bool {
        let matches = file.query("cell").expect("compile cell query");
        let node = if limited {
            matches
                .limit(1)
                .next()
                .transpose()
                .expect("limited cell query")
        } else {
            matches.first().expect("first cell query")
        };
        let node = node.expect("the fixture has a cell");
        node.kind() == "cell"
    }

    fn assert_first_cell(file: &File) {
        for limited in [false, true] {
            assert!(first_cell(file, limited));
            let matches = file.query("cell").expect("compile cell query");
            let node = if limited {
                matches
                    .limit(1)
                    .next()
                    .transpose()
                    .expect("limited cell query")
            } else {
                matches.first().expect("first cell query")
            }
            .expect("the fixture has a cell");
            assert_eq!(node.path().to_string(), "/sheet[1]/A1");
            assert_eq!(
                node.view().text().expect("cell text").as_deref(),
                Some("first")
            );
        }
    }

    #[test]
    fn eager_cell_query_first_and_limit_do_not_scale_with_cell_count() {
        let small = eager(workbook_bytes(1, 1));
        let large = eager(workbook_bytes(10_000, 1));
        assert_first_cell(&small);
        assert_first_cell(&large);

        for limited in [false, true] {
            let (small_count, small_bytes) = allocation_measure(|| first_cell(&small, limited));
            let (large_count, large_bytes) = allocation_measure(|| first_cell(&large, limited));
            assert!(
                large_count <= small_count + 2,
                "{limited}: first cell used {small_count} allocations in the one-cell workbook and {large_count} in the 10,000-cell workbook"
            );
            assert!(
                large_bytes <= small_bytes + 128,
                "{limited}: first cell used {small_bytes} bytes in the one-cell workbook and {large_bytes} in the 10,000-cell workbook"
            );
        }
    }

    #[test]
    fn eager_ordinal_get_does_not_scale_with_catalog_length() {
        let small = eager(workbook_bytes(1, 1));
        let large = eager(workbook_bytes(1, 100));
        for file in [&small, &large] {
            let node = file.get("/sheet[1]/A1").expect("first cell");
            assert_eq!(node.path().to_string(), "/sheet[1]/A1");
            assert_eq!(
                node.view().text().expect("cell text").as_deref(),
                Some("first")
            );
        }

        let (small_count, small_bytes) = allocation_measure(|| {
            let node = small.get("/sheet[1]/A1").expect("first cell");
            node.kind()
        });
        let (large_count, large_bytes) = allocation_measure(|| {
            let node = large.get("/sheet[1]/A1").expect("first cell");
            node.kind()
        });
        assert!(
            large_count <= small_count + 2,
            "ordinal get used {small_count} allocations in the one-sheet workbook and {large_count} in the 100-sheet workbook"
        );
        assert!(
            large_bytes <= small_bytes + 128,
            "ordinal get used {small_bytes} bytes in the one-sheet workbook and {large_bytes} in the 100-sheet workbook"
        );
    }

    fn native_first_cell(package: &litchi_xls::Package) -> (u32, u32) {
        let workbook = package.workbook();
        let reader = workbook.reader_ref().expect("eager native reader");
        let worksheet = reader.xls_worksheet(0).expect("first native worksheet");
        let mut cells = worksheet.cells();
        let cell = cells
            .next()
            .expect("native cell")
            .expect("native cell read");
        (cell.row(), cell.column())
    }

    #[test]
    fn native_first_cell_iteration_does_not_materialize_the_cell_vector() {
        let small_bytes = workbook_bytes(1, 1);
        let large_bytes = workbook_bytes(10_000, 1);
        let small = litchi_xls::Package::from_bytes(&small_bytes).expect("small native XLS");
        let large = litchi_xls::Package::from_bytes(&large_bytes).expect("large native XLS");
        // The shared Cell projection exposes one-based row and column values.
        assert_eq!(native_first_cell(&small), (1, 1));
        assert_eq!(native_first_cell(&large), (1, 1));

        let (small_count, small_bytes) = allocation_measure(|| native_first_cell(&small));
        let (large_count, large_bytes) = allocation_measure(|| native_first_cell(&large));
        assert!(
            large_count <= small_count + 1,
            "native first cell used {small_count} allocations in the one-cell workbook and {large_count} in the 10,000-cell workbook"
        );
        assert!(
            large_bytes <= small_bytes + 128,
            "native first cell used {small_bytes} bytes in the one-cell workbook and {large_bytes} in the 10,000-cell workbook"
        );
    }
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_selected_reads_add_constant_allocations_over_native_slide_reads() {
    fn deck(slides: usize, shapes: usize) -> (litchi_ops::File, litchi_ppt::Presentation) {
        let mut builder = litchi_ppt::writer::Builder::new();
        for _ in 0..slides {
            let slide = builder.add_slide().expect("slide");
            for _ in 0..shapes {
                builder
                    .add_textbox(slide, 10, 10, 240, 40, "selected text")
                    .expect("textbox");
            }
        }
        let package = builder.build().expect("package");
        let bytes = package.to_bytes().expect("bytes");
        let native = litchi_ppt::Package::from_bytes(&bytes)
            .expect("native")
            .presentation()
            .expect("presentation");
        (litchi_ops::open_bytes(bytes).expect("PPT"), native)
    }
    let (small, small_native) = deck(1, 1);
    let (large, large_native) = deck(100, 1);
    let path = litchi_ops::path::Path::parse("/slide[1]/shape[1]").expect("path");
    let query = litchi_ops::query::Query::parse("shape").expect("query");
    let plan =
        litchi_ops::query::Plan::compile(&query, litchi_core::FileFormat::Ppt).expect("plan");
    let get = |file: &litchi_ops::File| {
        let node = file.get(&path).expect("shape");
        assert_eq!(
            node.view().text().expect("text").as_deref(),
            Some("selected text")
        );
    };
    let first = |file: &litchi_ops::File| {
        let node = file
            .query(&plan)
            .expect("query")
            .first()
            .expect("first")
            .expect("shape");
        assert_eq!(
            node.view().text().expect("text").as_deref(),
            Some("selected text")
        );
    };
    let node = large.get(&path).expect("shared shape");
    let typed = node.as_ppt().expect("PPT node");
    assert_eq!(
        allocation_measure(|| typed.clone()),
        (0, 0),
        "node clones share native data"
    );
    get(&small);
    get(&large);
    first(&small);
    first(&large);
    let small_get = allocation_measure(|| get(&small));
    let large_get = allocation_measure(|| get(&large));
    let small_first = allocation_measure(|| first(&small));
    let large_first = allocation_measure(|| first(&large));
    let native = |presentation: &litchi_ppt::Presentation| {
        let slide = presentation.slide_at(0).expect("slide").expect("first");
        assert_eq!(
            slide.shapes().expect("shapes")[0].text().expect("text"),
            "selected text"
        );
    };
    // Both legs measure warm reads: the presentation's one-slide shape memo
    // (record 0260) serves the facade and native handles alike.
    native(&small_native);
    native(&large_native);
    let small_base = allocation_measure(|| native(&small_native));
    let large_base = allocation_measure(|| native(&large_native));
    let overhead =
        |total: (usize, usize), base: (usize, usize)| (total.0 - base.0, total.1 - base.1);
    assert_eq!(
        overhead(large_get, large_base),
        overhead(small_get, small_base),
        "facade lookup overhead must not scale with unselected slides"
    );
    let (many_shapes, many_shapes_native) = deck(1, 1_000);
    get(&many_shapes);
    native(&many_shapes_native);
    assert_eq!(
        overhead(
            allocation_measure(|| get(&many_shapes)),
            allocation_measure(|| native(&many_shapes_native)),
        ),
        overhead(small_get, small_base),
        "first-shape lookup must not allocate a candidate catalog of siblings"
    );
    assert_eq!(
        overhead(large_first, large_base),
        overhead(small_first, small_base),
        "facade first-match overhead must not scale with unselected slides"
    );
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_coalesced_text_planning_does_not_copy_the_existing_body() {
    fn measure(body: &str) -> (usize, usize) {
        let mut builder = litchi_ppt::writer::Builder::new();
        let slide = builder.add_slide().expect("slide");
        for text in ["warmup", body] {
            builder
                .add_textbox(slide, 10, 10, 240, 40, text)
                .expect("textbox");
        }
        let file =
            litchi_ops::open_bytes(builder.build().expect("package").to_bytes().expect("bytes"))
                .expect("PPT");
        let mut measured = (0, 0);
        let commit = file
            .edit(|edit| {
                edit.set(
                    "/slide[1]/shape[1]",
                    litchi_ops::props! { "text" => "prepared" },
                )?;
                // Opening, decoding the shared scene and publishing are outside
                // this boundary. Only admission into an existing run is measured.
                measured = allocation_measure(|| {
                    edit.set(
                        "/slide[1]/shape[2]",
                        litchi_ops::props! { "text" => "replacement" },
                    )
                    .expect("stage replacement");
                });
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit");
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(
            commit
                .file()
                .get("/slide[1]/shape[2]")
                .expect("shape")
                .view()
                .text()
                .expect("text")
                .as_deref(),
            Some("replacement")
        );
        measured
    }
    let small = measure("old");
    let large = measure(&"old ".repeat(4096));
    assert_eq!(
        large, small,
        "planning must borrow the existing native text body"
    );
}
