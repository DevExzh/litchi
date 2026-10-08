//! Allocation evidence for the node view (record 0010 §2, decision D1).
//!
//! Serializing a node view must allocate nothing beyond its sink: no
//! `DocumentNode`, no per-node map, no path `String`. These tests build a
//! fake document during setup, pre-size the sink, and count, with a
//! per-thread counting allocator (the pattern of litchi-core's
//! `tests/selector_allocations.rs`, record 0005), the allocations that
//! serializing and reading the view perform.
//!
//! The fake format is compiled under the `test-support` feature only, so this
//! test runs with `--features test-support` (or `--all-features`).

#![cfg(feature = "test-support")]
#![allow(
    unsafe_code,
    reason = "The test-only allocator is the measurement boundary; production crates remain safe"
)]

use std::alloc::{GlobalAlloc, Layout, System};
use std::borrow::Cow;
use std::cell::Cell;
use std::fmt::Write as _;

use litchi_core::Position;
use litchi_ops::node::fake::{self, Document, Element, Kind};
use litchi_ops::node::raw::{Ident, identity_step};
use litchi_ops::node::{PathStep, Preview, Value, View, ViewOptions};
use litchi_ops::value::{Color, Length, Scheme, Transform, TransformKind};

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

const LONG_TEXT: &str = "A paragraph long enough that its preview is cut at fifty characters, \
                         with \"quotes\", a back\\slash, a tab\tand non-ASCII text: café, 数据.";

const ACCENT: &[Transform] = &[
    Transform::new(TransformKind::LumMod, 75_000).unwrap(),
    Transform::new(TransformKind::LumOff, 25_000).unwrap(),
];

/// A Word-like tree: a document, a body, 40 paragraphs (half with a
/// `paraId` identity path, all with a cut preview, length, token and boolean
/// properties) of three runs each (text, font size, colours with transforms,
/// numbers), and a slide-like shape with geometry, so every value kind and
/// both path forms are serialized.
fn document() -> Document {
    let mut body = Element::new(Kind::Body, "/body");
    for paragraph in 0..40 {
        let mut path = vec![PathStep::new("body")];
        path.push(if paragraph % 2 == 0 {
            identity_step("p", "paraId", Ident::Number(0x1A2B_0000 + paragraph))
        } else {
            PathStep::at(
                "p",
                Position::new(usize::try_from(paragraph).expect("small")),
            )
        });
        let mut element = Element::at(Kind::Paragraph, path.clone())
            .text(LONG_TEXT)
            .preview(Preview::Prefix(Cow::Borrowed(LONG_TEXT)))
            .style("Heading1")
            .prop("style", "Heading1")
            .prop("align", Value::Token("justify"))
            .prop("spaceBefore", Length::from_twips(240))
            .prop("keepNext", true)
            .prop("outlineLvl", 1);
        for run in 0..3 {
            let mut run_path = path.clone();
            run_path.push(PathStep::at("r", Position::new(run)));
            element = element.child(
                Element::at(Kind::Run, run_path)
                    .text("run text, é")
                    .prop("font.latin", "Calibri")
                    .prop("size", Length::from_half_points(21))
                    .prop("bold", run % 2 == 0)
                    .prop(
                        "color",
                        if run == 0 {
                            Color::scheme(Scheme::Accent1)
                                .with_transforms(ACCENT)
                                .expect("two transforms fit")
                        } else {
                            Color::rgba(0x1F, 0x4E, 0x79, 0x80)
                        },
                    )
                    .prop("underline", Value::Token("single")),
            );
        }
        body = body.child(element);
    }
    let shape = Element::new(Kind::Shape, "/slide[1]/shape[@id=3]")
        .kind_name("textbox")
        .text("Revenue grew 25%")
        .preview(Preview::Text(Cow::Borrowed("TextBox 2")))
        .prop("name", "TextBox 2")
        .prop("id", 3)
        .prop("x", Length::from_emu(720_000))
        .prop("width", Length::from_emu(3_810_000))
        .prop("fill", Color::AUTO)
        .prop("rotation", Value::number(15.5).expect("finite"));
    Document::new(
        Element::new(Kind::Document, "/")
            .prop("title", "Q4 Report")
            .child(body)
            .child(shape),
    )
}

/// Serializes `view` compactly, pretty-printed and through
/// `serialize_into`, each into a sink sized beforehand, and returns the
/// allocation counts.
fn serialization_allocations(view: &View<'_, '_>) -> [usize; 3] {
    let size = serde_json::to_vec_pretty(view).expect("serializes").len();

    let mut sink = Vec::with_capacity(size);
    let compact = allocations(|| serde_json::to_writer(&mut sink, view).expect("serializes"));
    assert!(
        sink.len() < size,
        "compact output is shorter than pretty output"
    );

    let mut sink = Vec::with_capacity(size);
    let pretty = allocations(|| serde_json::to_writer_pretty(&mut sink, view).expect("serializes"));
    assert_eq!(sink.len(), size);

    let mut sink = Vec::with_capacity(size);
    let typed = allocations(|| {
        view.serialize_into(&mut serde_json::Serializer::new(&mut sink))
            .expect("serializes");
    });
    [compact, pretty, typed]
}

#[test]
fn the_counter_sees_allocations() {
    assert_eq!(allocations(|| Vec::<u8>::with_capacity(16)), 1);
    assert_eq!(allocations(|| String::from("x")), 1);
}

#[test]
fn serializing_a_view_allocates_nothing_beyond_the_sink() {
    let document = document();
    let root = document.root();
    for depth in 0..=3 {
        let view = root.view_with(ViewOptions::default().depth(depth));
        assert_eq!(
            serialization_allocations(&view),
            [0, 0, 0],
            "depth {depth}: [compact, pretty, serialize_into]"
        );
    }
    // The whole tree is serialized at depth 3: root, body, shape, 40
    // paragraphs and 120 runs.
    let view = root.view_with(ViewOptions::default().depth(3));
    let json = serde_json::to_string(&view).expect("serializes");
    assert_eq!(json.matches(r#""path":"#).count(), 1 + 2 + 40 + 120);
}

#[test]
fn a_depth_64_view_allocates_nothing_beyond_the_sink() {
    // Each element stores its full address, so the deepest path has 65
    // steps and the grammar's inline step storage would spill if the
    // serializer built a `Path`.
    let document = Document::new(fake::chain(70));
    let root = document.root();
    let view = root.view_with(ViewOptions::default().depth(litchi_ops::node::MAX_VIEW_DEPTH));
    assert_eq!(serialization_allocations(&view), [0, 0, 0]);
}

/// Writes every member of `view` and its listed descendants as text.
fn describe(view: View<'_, '_>, out: &mut String) {
    write!(out, "\n{} {}", view.path(), view.kind()).expect("writes");
    if let Some(preview) = view.preview().expect("readable preview") {
        write!(out, " [{preview}]").expect("writes");
    }
    if let Some(text) = view.text().expect("readable text") {
        out.push_str(&text);
    }
    if let Some(style) = view.style().expect("readable style") {
        out.push_str(&style);
    }
    if let Some(count) = view.child_count().expect("readable count") {
        write!(out, " ({count})").expect("writes");
    }
    for prop in view.format().expect("readable format") {
        write!(out, " {}={prop}", prop.key()).expect("writes");
    }
    for child in view.children() {
        let child = child.expect("readable child");
        describe(child.view_with(view.child_options()), out);
    }
}

#[test]
fn reading_the_view_allocates_nothing() {
    let document = document();
    let root = document.root();
    let view = root.view_with(ViewOptions::default().depth(3));
    let mut text = String::with_capacity(256 * 1024);
    let count = allocations(|| describe(view, &mut text));
    assert_eq!(count, 0);
    assert!(text.contains("\n/body/p[@paraId=439025664] paragraph [A paragraph long enough"));
    assert!(text.contains("\n/body/p[@paraId=439025664]/r[1] run"));
    assert!(text.contains(" color=accent1+lumMod75+lumOff25"));
    assert!(text.contains(" color=#1F4E7980"));
    assert!(text.contains(" size=10.5pt"));
    assert!(text.contains(" spaceBefore=12pt"));
    assert!(text.contains(" x=2cm width=300pt fill=auto rotation=15.5"));
}
