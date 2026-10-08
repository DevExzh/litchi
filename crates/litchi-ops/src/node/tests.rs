//! Unit tests of the node view against the fake format (record 0010 §2,
//! §7.2, risk 12).
//!
//! - Golden files under `tests/golden/node_view/` hold OfficeCLI-shaped
//!   `DocumentNode` JSON, written by hand from OfficeCLI's node builders
//!   (cited at each fixture) and compared byte for byte with the indented
//!   output, as `officecli get --json` prints a node.
//! - An oracle collects random fake trees into a declarative `DocumentNode`
//!   mirror (OfficeCLI's member order, `WhenWritingNull`) and checks that the
//!   streaming serializer writes the same bytes.
//! - The rest pin depth, the node budget, typed errors, value and path
//!   rendering, and the spec checks.

use std::borrow::Cow;
use std::fmt::Write as _;
use std::path::PathBuf;
use std::{fs, io};

use litchi_core::selector::{ErrorCode, Path, limits};
use litchi_core::{FileFormat, Position};
use serde::ser::{Serialize, SerializeMap, Serializer};

use super::fake::{self, Document, Element, Kind};
use super::spec::{self, GapSpec, PropSpec, Subject};
use super::view::{Format, Source};
use super::{
    Children, DEFAULT_MAX_NODES, Ident, MAX_VIEW_DEPTH, Node, NodePath, PREVIEW_CHARS, PathStep,
    Preview, Select, SerializeError, Value, View, ViewOptions,
};
use crate::error::{Code, Error};
use crate::value::{
    Color, ColorContext, Length, LengthSpec, LengthStyle, Scalar, Scheme, Transform, TransformKind,
    ValueType,
};

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const PARAGRAPH_TEXT: &str =
    "Quarterly report: revenue grew 25% year over year, led by the APAC region.";

/// `get /body/p[1] --depth 1` on a Word heading: OfficeCLI's `ParagraphToNode`
/// (`WordHandler.Navigation.cs` 4000-4352: text, 50-character preview,
/// style, run count, `paraId` then `pPr` keys) and the run builder
/// (2427-2527). Paths use the `@paraId` identity form (1748-1778). The
/// `style` member is the style id where OfficeCLI prints the display name
/// (M1-Δ4).
fn word_paragraph() -> Element {
    Element::new(Kind::Paragraph, "/body/p[@paraId=1A2B3C4D]")
        .positional("/body/p[1]")
        .text(PARAGRAPH_TEXT)
        .preview(Preview::Prefix(Cow::Borrowed(PARAGRAPH_TEXT)))
        .style("Heading1")
        .prop("paraId", "1A2B3C4D")
        .prop("style", "Heading1")
        .prop("styleId", "Heading1")
        .prop("align", Value::Token("center"))
        .prop("spaceBefore", twips(240))
        .prop("keepNext", true)
        .prop("numId", "3")
        .prop("numLevel", "0")
        .prop("outlineLvl", 0)
        .child(
            Element::new(Kind::Run, "/body/p[@paraId=1A2B3C4D]/r[1]")
                .text("Quarterly report: ")
                .prop("font.latin", "Calibri")
                .prop("size", half_points(28))
                .prop("bold", true)
                .prop("color", Color::rgb(0x1F, 0x4E, 0x79)),
        )
        .child(
            Element::new(Kind::Run, "/body/p[@paraId=1A2B3C4D]/r[2]")
                .text("revenue grew 25% year over year, led by the APAC region.")
                .prop("size", half_points(21))
                .prop("italic", false)
                .prop("underline", Value::Token("single")),
        )
}

/// `get /` on a Word document: OfficeCLI's `GetRootNode`
/// (`WordHandler.Navigation.cs` 118-477): type `document`, document
/// properties in `format`, `/body` listed with its child count.
fn word_document() -> Document {
    let body = Element::new(Kind::Body, "/body")
        .child(word_paragraph())
        .child(Element::new(Kind::Paragraph, "/body/p[2]").text("Second"))
        .child(Element::new(Kind::Paragraph, "/body/p[3]").text("Third"));
    Document::new(
        Element::new(Kind::Document, "/")
            .prop("title", "Q4 Report")
            .prop("author", "Finance")
            .child(body),
    )
}

/// `get /slide[1]`: OfficeCLI's slide node (`PowerPointHandler.Query.cs`
/// 1297-1324: title preview, `layout`) and shape builder
/// (`PowerPointHandler.NodeBuilder.cs` 880-1026: type by flavour, preview,
/// `name`, `alt`, `id`, `isTitle`, geometry through `FormatEmu`, fill with
/// transform suffixes; paragraph count as `childCount`). Shapes are not
/// expanded at depth 0 (M1-Δ6).
fn slide() -> Element {
    Element::new(Kind::Slide, "/slide[1]")
        .preview(Preview::Text(Cow::Borrowed("Q4 Report")))
        .prop("layout", "Title Slide")
        .child(
            Element::new(Kind::Shape, "/slide[1]/shape[@id=2]")
                .kind_name("title")
                .positional("/slide[1]/shape[1]")
                .text("Q4 Report")
                .preview(Preview::Prefix(Cow::Borrowed("Q4 Report")))
                .child_count(Some(1))
                .prop("name", "Title 1")
                .prop("id", 2)
                .prop("isTitle", true)
                .prop("x", Length::from_emu(720_000))
                .prop("y", Length::from_emu(540_000))
                .prop("width", Length::from_emu(10_753_200))
                .prop("height", Length::from_emu(1_144_800)),
        )
        .child(
            Element::new(Kind::Shape, "/slide[1]/shape[@id=3]")
                .kind_name("textbox")
                .positional("/slide[1]/shape[2]")
                .text("Revenue grew 25%")
                .preview(Preview::Prefix(Cow::Borrowed("Revenue grew 25%")))
                .child_count(Some(1))
                .prop("name", "TextBox 2")
                .prop("alt", "Revenue callout")
                .prop("id", 3)
                .prop("isTitle", false)
                .prop("x", Length::from_emu(720_000))
                .prop("y", Length::from_emu(1_800_000))
                .prop("width", Length::from_emu(3_810_000))
                .prop("height", Length::from_emu(914_400))
                .prop("fill", accent1(&[75_000], &[25_000]))
                .prop("rotation", Value::number(15.5).expect("finite")),
        )
}

/// `get /` on a deck at depth 0: OfficeCLI's root (`PowerPointHandler.Query.cs`
/// 28-60) with the widescreen slide size, both pt-exact (`FormatEmuPaired`).
fn presentation() -> Document {
    Document::new(
        Element::new(Kind::Presentation, "/")
            .prop("slideWidth", Length::from_emu(12_192_000))
            .prop("slideHeight", Length::from_emu(6_858_000))
            .child(slide()),
    )
    .with_format(FileFormat::Pptx)
}

/// `get /` on a workbook: OfficeCLI's root (`ExcelHandler.Query.cs` 24-60)
/// lists each sheet with its name as preview. Litchi omits each sheet's
/// `childCount`, which would decode every sheet part (M1-Δ1).
fn workbook() -> Document {
    let sheet = |path: &'static str, name: &'static str| {
        Element::new(Kind::Sheet, path)
            .preview(Preview::Text(Cow::Borrowed(name)))
            .child_count(None)
            .prop("name", name)
    };
    Document::new(
        Element::new(Kind::Workbook, "/")
            .child(sheet("/Sheet1", "Sheet1").child(cells().swap_remove(0)))
            .child(sheet("/My Data", "My Data")),
    )
    .with_format(FileFormat::Xlsx)
}

/// Cells: OfficeCLI's `CellToNode` (`ExcelHandler.Helpers.Node.cs` 295-563:
/// A1 preview, `type`, `formula`, `cachedValue`, `evaluated`, font keys) and
/// the missing cell (`ExcelHandler.Query.cs` 965: text `(empty)`).
fn cells() -> Vec<Element> {
    vec![
        Element::new(Kind::Cell, "/Sheet1/B2")
            .text("6")
            .preview(Preview::Text(Cow::Borrowed("B2")))
            .prop("type", Value::Token("Number"))
            .prop("formula", "SUM(A1:A3)")
            .prop("cachedValue", "6")
            .prop("evaluated", true)
            .prop("font.size", Length::from_points(11))
            .prop("font.color", Color::scheme(Scheme::Accent1)),
        Element::new(Kind::Cell, "/Sheet1/A1")
            .text("Région")
            .preview(Preview::Text(Cow::Borrowed("A1")))
            .prop("type", Value::Token("SharedString"))
            .prop("font.color", Color::rgba(255, 0, 0, 0x80)),
        Element::new(Kind::Cell, "/Sheet1/Z99")
            .text("(empty)")
            .preview(Preview::Text(Cow::Borrowed("Z99"))),
    ]
}

fn twips(value: i32) -> Length {
    Length::from_twips(value)
}

fn half_points(value: i32) -> Length {
    Length::from_half_points(value)
}

/// A transform of `kind` with `value` thousandths of a percent.
fn transform(kind: TransformKind, value: i32) -> Transform {
    Transform::new(kind, value).expect("in range")
}

/// `accent1` with `lumMod` then `lumOff` transforms, as DrawingML stores a
/// darker or lighter theme colour.
fn accent1(lum_mod: &[i32], lum_off: &[i32]) -> Color {
    let transforms: Vec<Transform> = lum_mod
        .iter()
        .map(|&value| transform(TransformKind::LumMod, value))
        .chain(
            lum_off
                .iter()
                .map(|&value| transform(TransformKind::LumOff, value)),
        )
        .collect();
    Color::scheme(Scheme::Accent1)
        .with_transforms(&transforms)
        .expect("fits")
}

fn compact(view: &View<'_, '_>) -> String {
    let mut sink = Vec::new();
    view.serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect("serializes");
    String::from_utf8(sink).expect("UTF-8")
}

fn pretty(view: &View<'_, '_>) -> String {
    let mut sink = Vec::new();
    view.serialize_into(&mut serde_json::Serializer::pretty(&mut sink))
        .expect("serializes");
    String::from_utf8(sink).expect("UTF-8")
}

fn check_golden(name: &str, actual: &str) {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden/node_view")
        .join(format!("{name}.json"));
    let actual = format!("{actual}\n");
    if std::env::var_os("LITCHI_BLESS").is_some() {
        fs::write(&path, &actual).expect("bless the golden file");
        return;
    }
    let expected = fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("read {}: {error}", path.display()));
    assert_eq!(
        actual, expected,
        "{name}: the node view differs from the OfficeCLI-shaped golden file \
         (LITCHI_BLESS=1 rewrites it)"
    );
}

/// The children of `node`, which must all be readable.
fn children<'f>(node: &Node<'f>) -> Vec<Node<'f>> {
    node.children()
        .collect::<crate::Result<Vec<_>>>()
        .expect("readable children")
}

// ---------------------------------------------------------------------------
// OfficeCLI golden shapes
// ---------------------------------------------------------------------------

#[test]
fn word_document_root_matches_officecli_shape() {
    let document = word_document();
    check_golden("word_document", &pretty(&document.root().view()));
}

#[test]
fn word_paragraph_matches_officecli_shape() {
    let document = Document::new(word_paragraph());
    check_golden("word_paragraph", &pretty(&document.root().view()));
}

#[test]
fn pptx_nodes_match_officecli_shape() {
    let document = presentation();
    let root = document.root();
    check_golden(
        "pptx_presentation",
        &pretty(&root.view_with(ViewOptions::default().depth(0))),
    );
    let slide = children(&root).remove(0);
    assert_eq!(slide.format(), FileFormat::Pptx);
    check_golden(
        "pptx_slide",
        &pretty(&slide.view_with(ViewOptions::default().depth(1))),
    );
}

#[test]
fn xlsx_nodes_match_officecli_shape() {
    let document = workbook();
    check_golden("xlsx_workbook", &pretty(&document.root().view()));

    let documents: Vec<Document> = cells().into_iter().map(Document::new).collect();
    let roots: Vec<Node<'_>> = documents.iter().map(Document::root).collect();
    let views: Vec<View<'_, '_>> = roots.iter().map(Node::view).collect();
    let json = serde_json::to_string_pretty(&views).expect("serializes");
    check_golden("xlsx_cells", &json);
}

#[test]
fn compact_output_has_officecli_member_order() {
    let document = Document::new(word_paragraph());
    let json = compact(&document.root().view_with(ViewOptions::default().depth(0)));
    assert_eq!(
        json,
        concat!(
            r#"{"path":"/body/p[@paraId=1A2B3C4D]","type":"paragraph","#,
            r#""text":"Quarterly report: revenue grew 25% year over year, led by the APAC region.","#,
            r#""preview":"Quarterly report: revenue grew 25% year over year,...","#,
            r#""style":"Heading1","childCount":2,"format":{"paraId":"1A2B3C4D","#,
            r#""style":"Heading1","styleId":"Heading1","align":"center","#,
            r#""spaceBefore":"12pt","keepNext":true,"numId":"3","numLevel":"0","#,
            r#""outlineLvl":0},"children":[]}"#
        )
    );
}

#[test]
fn absent_members_are_omitted_and_containers_kept() {
    let document = Document::new(Element::new(Kind::Item, "/item[1]").child_count(None));
    assert_eq!(
        compact(&document.root().view()),
        r#"{"path":"/item[1]","type":"item","format":{},"children":[]}"#
    );
    let document = Document::new(Element::new(Kind::Item, "/item[1]").child_count(Some(0)));
    assert_eq!(
        compact(&document.root().view()),
        r#"{"path":"/item[1]","type":"item","childCount":0,"format":{},"children":[]}"#
    );
}

#[test]
fn serialize_trait_matches_serialize_into() {
    let document = word_document();
    let root = document.root();
    let view = root.view_with(ViewOptions::default().depth(3));
    assert_eq!(
        serde_json::to_string(&view).expect("serializes"),
        compact(&view)
    );
    assert_eq!(
        serde_json::to_string_pretty(&view).expect("serializes"),
        pretty(&view)
    );
}

// ---------------------------------------------------------------------------
// The streaming serializer against a collect-then-serialize oracle
// ---------------------------------------------------------------------------

/// OfficeCLI's `DocumentNode` (`Core/DocumentNode.cs` 14-29) as a declarative
/// serde struct, with `WhenWritingNull` as `skip_serializing_if`.
/// `childCount` is optional here only for Litchi's M1-Δ1.
#[derive(serde::Serialize)]
struct OfficeNode {
    path: String,
    #[serde(rename = "type")]
    kind: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    preview: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(rename = "childCount", skip_serializing_if = "Option::is_none")]
    child_count: Option<u32>,
    format: Entries,
    children: Vec<OfficeNode>,
}

/// An insertion-ordered map, as .NET's `Dictionary` enumerates.
struct Entries(Vec<(&'static str, serde_json::Value)>);

impl Serialize for Entries {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (key, value) in &self.0 {
            map.serialize_entry(key, value)?;
        }
        map.end()
    }
}

/// Collects a view through its public accessors.
fn collect(view: View<'_, '_>) -> OfficeNode {
    let format = view
        .format()
        .expect("readable format")
        .map(|prop| {
            let value = serde_json::to_value(&prop).expect("a JSON value");
            (prop.key(), value)
        })
        .collect();
    let children = view
        .children()
        .map(|child| {
            collect(
                child
                    .expect("readable child")
                    .view_with(view.child_options()),
            )
        })
        .collect();
    OfficeNode {
        path: view.path().to_string(),
        kind: view.kind(),
        text: view.text().expect("readable text").map(Cow::into_owned),
        preview: view
            .preview()
            .expect("readable preview")
            .map(|preview| preview.to_string()),
        style: view.style().expect("readable style").map(Cow::into_owned),
        child_count: view.child_count().expect("readable count"),
        format: Entries(format),
        children,
    }
}

/// `SplitMix64`, for reproducible random trees without a new dependency.
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound).expect("small"))
            .expect("below the bound")
    }

    fn chance(&mut self, percent: usize) -> bool {
        self.below(100) < percent
    }

    fn pick<T: Copy>(&mut self, items: &[T]) -> T {
        items[self.below(items.len())]
    }
}

const TEXTS: &[&str] = &[
    "",
    "plain",
    "quote \" and backslash \\",
    "line\nbreak\ttab\u{1}",
    "café",
    "数据表",
    "emoji 😀 ok",
    "a much longer text that runs well past the fifty-character preview limit",
];

const NUMBERS: &[f64] = &[0.5, 1.0, -2.25, 0.001, 123_456.789, 1e21, -0.0];

fn random_element(rng: &mut Rng, path: Vec<PathStep<'static>>, levels: u32) -> Element {
    let mut element = Element::at(Kind::Item, path.clone());
    if rng.chance(70) {
        element = element.text(rng.pick(TEXTS));
    }
    if rng.chance(50) {
        let text = Cow::Borrowed(rng.pick(TEXTS));
        element = element.preview(if rng.chance(50) {
            Preview::Prefix(text)
        } else {
            Preview::Text(text)
        });
    }
    if rng.chance(30) {
        element = element.style("Normal");
    }
    if rng.chance(50) {
        element = element.prop("level", i64::try_from(rng.next() >> 1).expect("fits"));
    }
    if rng.chance(50) {
        element = element.prop("weight", Value::Number(rng.pick(NUMBERS)));
    }
    element = match rng.below(3) {
        0 => element.child_count(None),
        1 => element.child_count(Some(u32::try_from(rng.below(9)).expect("small"))),
        _ => element,
    };
    if levels > 0 {
        for index in 0..rng.below(4) {
            let mut child_path = path.clone();
            child_path.push(if rng.chance(25) {
                PathStep::identity("item", "id", Ident::Number(rng.next() % 1000))
            } else {
                PathStep::at("item", Position::new(index))
            });
            element = element.child(random_element(rng, child_path, levels - 1));
        }
    }
    element
}

#[test]
fn streaming_output_matches_the_document_node_oracle() {
    let mut rng = Rng(0x0010_0001);
    for seed in 0..200 {
        let root = random_element(&mut rng, vec![PathStep::at("item", Position::new(0))], 4);
        let document = Document::new(root);
        let node = document.root();
        for depth in 0..=5 {
            let view = node.view_with(ViewOptions::default().depth(depth));
            let oracle = collect(view);
            assert_eq!(
                compact(&view),
                serde_json::to_string(&oracle).expect("oracle serializes"),
                "seed {seed}, depth {depth}"
            );
            assert_eq!(
                pretty(&view),
                serde_json::to_string_pretty(&oracle).expect("oracle serializes"),
                "seed {seed}, depth {depth}"
            );
        }
    }
}

// ---------------------------------------------------------------------------
// Depth and the node budget
// ---------------------------------------------------------------------------

#[test]
fn options_preserve_requests_and_refuse_excessive_depth() {
    let options = ViewOptions::default();
    assert_eq!(options, ViewOptions::DEFAULT);
    assert_eq!(options.depth, 1);
    assert_eq!(options.max_nodes, DEFAULT_MAX_NODES);
    assert_eq!(options.dense, None);
    assert_eq!(ViewOptions::default().depth(65).depth, 65);
    assert_eq!(ViewOptions::default().depth(u32::MAX).depth, u32::MAX);
    assert_eq!(ViewOptions::default().depth(7).depth, 7);
    assert_eq!(ViewOptions::default().dense(9).dense, Some(9));

    // An excessive request is retained and refused before writing output.
    let document = Document::new(fake::chain(3));
    let root = document.root();
    let options = ViewOptions {
        depth: 1_000,
        ..ViewOptions::default()
    };
    let view = root.view_with(options);
    assert_eq!(view.options().depth, 1_000);
    assert_eq!(view.child_options().depth, 999);
    let mut output = Vec::new();
    assert!(
        matches!(view.serialize_into(&mut serde_json::Serializer::new(&mut output)), Err(SerializeError::View(error)) if error.code() == Code::MaxDepthExceeded)
    );
    assert!(output.is_empty());
    assert!(std::ptr::eq(view.node(), &root));
}

/// The number of node objects in serialized output.
fn node_count(json: &str) -> usize {
    json.matches(r#""path":"#).count()
}

#[test]
fn depth_counts_child_levels() {
    let document = Document::new(fake::chain(5));
    let root = document.root();
    for depth in 0..=7_u32 {
        let json = compact(&root.view_with(ViewOptions::default().depth(depth)));
        let expected = usize::try_from(depth.min(5)).expect("small") + 1;
        assert_eq!(node_count(&json), expected, "depth {depth}");
    }
    // Depth 0 lists no children but still counts them.
    let json = compact(&root.view_with(ViewOptions::default().depth(0)));
    assert!(json.contains(r#""childCount":1,"#), "{json}");
    assert!(json.ends_with(r#""children":[]}"#), "{json}");
    assert_eq!(
        root.view_with(ViewOptions::default().depth(0))
            .children()
            .count(),
        0
    );
    assert_eq!(root.view().children().count(), 1);
}

/// Record 0010 risk 12: a deep tree serializes with an explicit depth of 64,
/// on a thread with a small, fixed stack.
#[test]
fn depth_64_tree_serializes_within_a_small_stack() {
    let worker = std::thread::Builder::new()
        .name("depth-64".to_owned())
        .stack_size(1 << 20)
        .spawn(|| {
            let document = Document::new(fake::chain(200));
            let root = document.root();
            let view = root.view_with(ViewOptions::default().depth(MAX_VIEW_DEPTH));
            assert_eq!(view.options().depth, MAX_VIEW_DEPTH);
            let json = compact(&view);
            assert_eq!(node_count(&json), 65);
            // The 65th node is the deepest listed: its child is counted, not
            // listed.
            let deepest = format!(
                r#"{{"path":"{}","type":"item","childCount":1,"format":{{"level":64}},"children":[]}}"#,
                "/item[1]".repeat(65)
            );
            assert!(json.contains(&deepest), "{}", &json[json.len() - 400..]);
            let pretty_json = pretty(&view);
            assert_eq!(node_count(&pretty_json.replace(": ", ":")), 65);
        })
        .expect("spawn");
    worker.join().expect("the depth-64 walk completes");
}

fn small_tree() -> Document {
    // 1 root + 3 children + 6 grandchildren = 10 nodes.
    let mut root = Element::new(Kind::Item, "/item[1]");
    for child in 0..3 {
        let child_path = vec![
            PathStep::at("item", Position::new(0)),
            PathStep::at("item", Position::new(child)),
        ];
        let mut element = Element::at(Kind::Item, child_path.clone());
        for grandchild in 0..2 {
            let mut path = child_path.clone();
            path.push(PathStep::at("item", Position::new(grandchild)));
            element = element.child(Element::at(Kind::Item, path));
        }
        root = root.child(element);
    }
    Document::new(root)
}

#[test]
fn node_budget_is_output_too_large() {
    let document = small_tree();
    let root = document.root();
    let at = |max_nodes| {
        let view = root.view_with(ViewOptions::default().depth(2).max_nodes(max_nodes));
        let mut sink = Vec::new();
        view.serialize_into(&mut serde_json::Serializer::new(&mut sink))
    };
    assert!(at(10).is_ok(), "exactly the budget fits");
    for max_nodes in [0, 1, 5, 9] {
        let Err(SerializeError::View(error)) = at(max_nodes) else {
            panic!("max_nodes {max_nodes} must fail with a view error");
        };
        assert_eq!(error.code(), Code::OutputTooLarge);
        assert!(error.message().contains(&max_nodes.to_string()));
        assert!(
            error
                .suggestion()
                .is_some_and(|text| text.contains("max_nodes"))
        );
    }
    let view = root.view_with(ViewOptions::default().depth(2).max_nodes(4));
    let error = serde_json::to_string(&view).expect_err("over budget");
    assert!(error.to_string().contains("output_too_large"), "{error}");
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

fn view_error(document: &Document, depth: u32) -> Error {
    let root = document.root();
    let view = root.view_with(ViewOptions::default().depth(depth));
    let mut sink = Vec::new();
    match view.serialize_into(&mut serde_json::Serializer::new(&mut sink)) {
        Err(SerializeError::View(error)) => error,
        other => panic!("expected a view error, got {other:?}"),
    }
}

#[test]
fn read_failures_stay_typed() {
    let text = Document::new(Element::new(Kind::Item, "/item[1]").fail_text(Code::CorruptFile));
    assert_eq!(view_error(&text, 1).code(), Code::CorruptFile);

    let format =
        Document::new(Element::new(Kind::Item, "/item[1]").fail_format(Code::LimitExceeded));
    assert_eq!(view_error(&format, 1).code(), Code::LimitExceeded);

    let child = Document::new(
        Element::new(Kind::Item, "/item[1]")
            .child(Element::new(Kind::Item, "/item[1]/item[1]"))
            .child(Element::new(Kind::Item, "/item[1]/item[2]"))
            .fail_child(1, Code::IoError),
    );
    assert_eq!(view_error(&child, 1).code(), Code::IoError);
    // At depth 0 the children are not read, so nothing fails.
    assert!(
        serde_json::to_string(&child.root().view_with(ViewOptions::default().depth(0))).is_ok()
    );

    // A failing grandchild surfaces through the nested walk, and the plain
    // `Serialize` reports the unified error's text.
    let nested = Document::new(
        Element::new(Kind::Item, "/item[1]")
            .child(Element::new(Kind::Item, "/item[1]/item[1]").fail_text(Code::CorruptFile)),
    );
    assert_eq!(view_error(&nested, 3).code(), Code::CorruptFile);
    let error = serde_json::to_string(&nested.root().view()).expect_err("fails");
    assert!(error.to_string().contains("corrupt_file"), "{error}");
}

/// A sink that accepts `remaining` bytes, then fails.
struct FullSink {
    remaining: usize,
}

impl io::Write for FullSink {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::other("the sink is full"));
        }
        let written = buf.len().min(self.remaining);
        self.remaining -= written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_failures_are_serializer_errors() {
    let document = word_document();
    let root = document.root();
    let view = root.view_with(ViewOptions::default().depth(3));
    let result = view.serialize_into(&mut serde_json::Serializer::new(FullSink { remaining: 40 }));
    let Err(SerializeError::Serializer(error)) = result else {
        panic!("expected a serializer error, got {result:?}");
    };
    assert!(error.is_io());
    let wrapped: SerializeError<serde_json::Error> = SerializeError::Serializer(error);
    assert_eq!(wrapped.to_string(), "node serializer failed");
    assert!(std::error::Error::source(&wrapped).is_none());

    let view_error =
        SerializeError::<serde_json::Error>::View(Error::new(Code::CorruptFile, "bad"));
    assert_eq!(view_error.to_string(), "corrupt_file: bad");
    assert!(std::error::Error::source(&view_error).is_some());
}

#[test]
fn children_are_fused_after_an_error() {
    let document = Document::new(
        Element::new(Kind::Item, "/item[1]")
            .child(Element::new(Kind::Item, "/item[1]/item[1]"))
            .child(Element::new(Kind::Item, "/item[1]/item[2]"))
            .fail_child(0, Code::CorruptFile),
    );
    let root = document.root();
    let mut iter = root.children();
    assert!(matches!(iter.next(), Some(Err(error)) if error.code() == Code::CorruptFile));
    assert!(iter.next().is_none());
    assert!(iter.next().is_none());

    let mut failed = Children::failed(Error::new(Code::IoError, "unreadable"));
    assert!(matches!(failed.next(), Some(Err(error)) if error.code() == Code::IoError));
    assert!(failed.next().is_none());
    assert!(Children::empty().next().is_none());
    assert_eq!(format!("{:?}", Children::empty()), "Children { .. }");
}

// ---------------------------------------------------------------------------
// Accessors
// ---------------------------------------------------------------------------

#[test]
fn node_and_view_accessors_dispatch() {
    let document = Document::new(word_paragraph());
    let node = document.root();
    assert_eq!(node.kind(), "paragraph");
    assert_eq!(node.format(), FileFormat::Docx);
    assert_eq!(node.path().to_string(), "/body/p[@paraId=1A2B3C4D]");
    assert_eq!(node.positional_path().to_string(), "/body/p[1]");
    assert!(format!("{node:?}").starts_with("Node"));

    let view = node.view();
    assert_eq!(view.kind(), "paragraph");
    assert_eq!(view.path(), node.path());
    assert_eq!(view.text().expect("text").as_deref(), Some(PARAGRAPH_TEXT));
    assert_eq!(
        view.preview()
            .expect("preview")
            .map(|preview| preview.to_string()),
        Some("Quarterly report: revenue grew 25% year over year,...".to_owned())
    );
    assert_eq!(view.style().expect("style").as_deref(), Some("Heading1"));
    assert_eq!(view.child_count().expect("count"), Some(2));

    let runs = children(&node);
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[0].kind(), "run");
    assert_eq!(runs[1].path().to_string(), "/body/p[@paraId=1A2B3C4D]/r[2]");
    // A run without an identity form reports the canonical path again.
    assert_eq!(runs[1].positional_path(), runs[1].path());
    let run = runs[0].view();
    assert_eq!(run.preview().expect("preview"), None);
    assert_eq!(run.style().expect("style"), None);
}

#[test]
fn format_follows_the_property_table_and_skips_absent_keys() {
    let document = Document::new(word_paragraph());
    let node = document.root();
    let keys: Vec<_> = node
        .view()
        .format()
        .expect("format")
        .map(|prop| prop.key())
        .collect();
    assert_eq!(
        keys,
        [
            "paraId",
            "style",
            "styleId",
            "align",
            "spaceBefore",
            "keepNext",
            "numId",
            "numLevel",
            "outlineLvl"
        ]
    );
    let format = node.view().format().expect("format");
    assert_eq!(format.size_hint(), (0, Some(10)));
    assert!(format!("{format:?}").starts_with("Format {"));

    // 720 twips is both 36pt and 1.27cm: the key's style writes points,
    // while `Value`'s own display is centimetre-first.
    let spaced =
        Document::new(Element::new(Kind::Paragraph, "/body/p[1]").prop("spaceBefore", twips(720)));
    let spaced = spaced.root();
    let space = spaced
        .view()
        .format()
        .expect("format")
        .find(|prop| prop.key() == "spaceBefore")
        .expect("spaceBefore");
    assert_eq!(space.to_string(), "36pt");
    assert_eq!(space.value().to_string(), "1.27cm");
    assert_eq!(
        serde_json::to_string(&space).expect("serializes"),
        r#""36pt""#
    );
    assert_eq!(
        serde_json::to_string(space.value()).expect("serializes"),
        r#""1.27cm""#
    );
    assert_eq!(space.clone().into_value(), Value::Length(twips(720)));
    assert_eq!(
        space.into_parts(),
        ("spaceBefore", Value::Length(twips(720)))
    );

    assert_eq!(Format::empty().count(), 0);
    static PROPS: [PropSpec; 1] = [PropSpec::new("bold", ValueType::Bool)];
    assert_eq!(
        Format::new(&PROPS, Source::Empty(std::marker::PhantomData)).count(),
        0
    );
}

#[test]
fn prop_values_keep_their_json_types() {
    let document = Document::new(
        Element::new(Kind::Item, "/item[1]")
            .prop("level", 7)
            .prop("weight", Value::number(2.0).expect("finite")),
    );
    assert_eq!(
        compact(&document.root().view()),
        r#"{"path":"/item[1]","type":"item","childCount":0,"format":{"level":7,"weight":2},"children":[]}"#
    );
    let json = |value: Value<'_>| serde_json::to_string(&value).expect("serializes");
    assert_eq!(json(Value::Bool(false)), "false");
    assert_eq!(json(Value::Int(-3)), "-3");
    assert_eq!(json(Value::Number(2.5)), "2.5");
    assert_eq!(json(Value::Number(-4.0)), "-4");
    assert!(serde_json::to_string(&Value::Number(f64::NAN)).is_err());
    assert_eq!(
        json(Value::Number(1e300)),
        serde_json::to_string(&1e300_f64).expect("serializes")
    );
    assert_eq!(json(Value::from("a\"b")), r#""a\"b""#);
    assert_eq!(json(Value::Token("single")), r#""single""#);
    assert_eq!(
        json(Value::Length(Length::from_emu(914_400))),
        r#""2.54cm""#
    );
    assert_eq!(json(Value::Color(Color::AUTO)), r#""auto""#);
}

// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

/// R1's OfficeCLI vectors (`EmuConverter.FormatEmu`, `FormatEmuPtFirst`, and
/// the `0.##pt` forms of spacing and line widths, `FormatLineWidth`), printed
/// by the canonical [`Length`] renderer the node view now uses. The rows
/// agree with R1's renderer except `i64::MAX`: above 2^50 EMU the canonical
/// renderer replicates OfficeCLI's double re-read check, so the value has no
/// exact centimetre or point form there (R1's exact integer arithmetic
/// printed `726249766681478.41pt`).
#[test]
fn lengths_render_through_the_canonical_renderer() {
    let cases: &[(i64, &str, &str, &str)] = &[
        // emu, FormatEmu, FormatEmuPtFirst, FormatLineWidth
        (0, "0cm", "0pt", "0pt"),
        (360_000, "1cm", "1cm", "28.35pt"),
        (914_400, "2.54cm", "72pt", "72pt"),
        (12_192_000, "960pt", "960pt", "960pt"),
        (6_858_000, "19.05cm", "540pt", "540pt"),
        (457_200, "1.27cm", "36pt", "36pt"),
        (3_600, "0.01cm", "0.01cm", "0.28pt"),
        (127, "0.01pt", "0.01pt", "0.01pt"),
        (12_700, "1pt", "1pt", "1pt"),
        (139_700, "11pt", "11pt", "11pt"),
        (133_350, "10.5pt", "10.5pt", "10.5pt"),
        (-720_000, "-2cm", "-2cm", "-56.69pt"),
        (-1_800, "-1800emu", "-1800emu", "-0.14pt"),
        (180_000, "0.5cm", "0.5cm", "14.17pt"),
        (1, "1emu", "1emu", "1emu"),
        (63, "63emu", "63emu", "63emu"),
        (64, "64emu", "64emu", "0.01pt"),
        (-63, "-63emu", "-63emu", "-63emu"),
        (190, "190emu", "190emu", "0.01pt"),
        (191, "191emu", "191emu", "0.02pt"),
        (
            i64::MAX,
            "9223372036854775807emu",
            "9223372036854775807emu",
            "726249766681478.41pt",
        ),
        (
            i64::MIN,
            "-9223372036854775808emu",
            "-9223372036854775808emu",
            "-726249766681478.41pt",
        ),
    ];
    for &(emu, cm_first, pt_first, line) in cases {
        let length = Length::from_emu(emu);
        assert_eq!(
            length.display(LengthStyle::Auto).to_string(),
            cm_first,
            "{emu} cm-first"
        );
        assert_eq!(
            length.display(LengthStyle::PointsFirst).to_string(),
            pt_first,
            "{emu} pt-first"
        );
        assert_eq!(
            length.display(LengthStyle::LineWidth).to_string(),
            line,
            "{emu} line width"
        );
        assert_eq!(length.to_string(), cm_first);
        assert_eq!(Value::Length(length).to_string(), cm_first);
    }
}

/// A key's value type picks the rendering of a length: the same EMU prints
/// centimetre-first as a position and in points as a font size.
#[test]
fn a_keys_value_type_picks_the_length_rendering() {
    let length = Value::Length(Length::from_twips(720));
    let render = |ty| {
        struct Show<'v>(&'v Value<'v>, ValueType);
        impl std::fmt::Display for Show<'_> {
            fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                self.0.fmt_typed(Some(self.1), f)
            }
        }
        Show(&length, ty).to_string()
    };
    assert_eq!(render(ValueType::Length(LengthSpec::FONT_SIZE)), "36pt");
    assert_eq!(render(ValueType::Length(LengthSpec::POSITION)), "1.27cm");
    // A type that is not a length leaves the centimetre-first default.
    assert_eq!(render(ValueType::Bool), "1.27cm");
    assert_eq!(length.to_string(), "1.27cm");
}

/// R1's colour vectors, printed by the canonical [`Color`] renderer. Two R1
/// rules changed with the reconciliation: a transform value that is not a
/// whole percentage prints OfficeCLI's raw `name=value` form, which reads
/// back exactly, instead of a truncated percentage; and a colour prints every
/// transform it holds, so a reader folds an sRGB base's `a:alpha` into the
/// base's alpha channel (OfficeCLI's `#RRGGBBAA` readback) rather than
/// keeping it as a transform.
#[test]
fn colors_render_through_the_canonical_renderer() {
    let cases: Vec<(Color, &str)> = vec![
        (Color::rgb(0x44, 0x72, 0xC4), "#4472C4"),
        (Color::rgb(0, 0, 0), "#000000"),
        (Color::rgba(0xFF, 0, 0, 0x80), "#FF000080"),
        (Color::rgba(1, 2, 3, 0), "#01020300"),
        (Color::AUTO, "auto"),
        (accent1(&[75_000], &[25_000]), "accent1+lumMod75+lumOff25"),
        (accent1(&[75_500], &[]), "accent1+lumMod=75500"),
        (
            Color::scheme(Scheme::FollowedHyperlink),
            "followedHyperlink",
        ),
    ];
    for (color, expected) in &cases {
        assert_eq!(color.to_string(), *expected);
        assert_eq!(Value::Color(*color).to_string(), *expected);
        assert_eq!(Value::Color(*color).as_scalar(), Scalar::Color(*color));
    }
    for (color, expected) in &cases[5..] {
        let parsed = Color::parse(expected, ColorContext::Drawing).expect("reads back");
        assert_eq!(parsed, *color, "{expected}");
    }
    let names: Vec<_> = [
        Scheme::Dark1,
        Scheme::Light1,
        Scheme::Dark2,
        Scheme::Light2,
        Scheme::Accent1,
        Scheme::Accent2,
        Scheme::Accent3,
        Scheme::Accent4,
        Scheme::Accent5,
        Scheme::Accent6,
        Scheme::Hyperlink,
        Scheme::FollowedHyperlink,
        Scheme::Text1,
        Scheme::Text2,
        Scheme::Background1,
        Scheme::Background2,
    ]
    .iter()
    .map(|scheme| scheme.name())
    .collect();
    assert_eq!(
        names.join(","),
        "dark1,light1,dark2,light2,accent1,accent2,accent3,accent4,accent5,accent6,\
         hyperlink,followedHyperlink,text1,text2,background1,background2"
    );
}

#[test]
fn values_borrow_and_own_without_changing() {
    let text1 = Color::scheme(Scheme::Text1)
        .with_transforms(&[transform(TransformKind::LumMod, 50_000)])
        .expect("fits");
    let values = [
        Value::Bool(true),
        Value::Int(i64::MIN),
        Value::Number(-0.25),
        Value::Text(Cow::Owned("owned".to_owned())),
        Value::Token("center"),
        Value::Length(Length::from_emu(42)),
        Value::Color(text1),
    ];
    for value in &values {
        assert_eq!(&value.borrowed(), value);
        assert_eq!(&value.clone().into_owned(), value);
        assert_eq!(value.borrowed().to_string(), value.to_string());
        assert_eq!(Value::from(value.as_scalar()), value.borrowed());
    }
    assert_eq!(values[0].to_string(), "true");
    assert_eq!(values[1].to_string(), "-9223372036854775808");
    assert_eq!(values[2].to_string(), "-0.25");
    assert_eq!(values[6].to_string(), "text1+lumMod50");
    assert_eq!(Value::Number(3.0).to_string(), "3");
    // Numbers print as .NET prints doubles, as the canonical text does.
    assert_eq!(Value::Number(1e21).to_string(), "1E+21");
    assert_eq!(values[3].as_str(), Some("owned"));
    assert_eq!(values[4].as_str(), Some("center"));
    assert_eq!(values[5].as_str(), None);
    assert_eq!(Value::from(5_u32), Value::Int(5));
    assert_eq!(Value::from(-5_i64), Value::Int(-5));
    assert_eq!(
        Value::from(String::from("s")),
        Value::Text(Cow::Borrowed("s"))
    );
    assert_eq!(
        Value::from(Cow::Borrowed("c")),
        Value::Text(Cow::Borrowed("c"))
    );
    assert_eq!(Value::from(Color::AUTO), Value::Color(Color::AUTO));
    assert_eq!(Value::from(Scalar::Token("center")), Value::Token("center"));
}

#[test]
fn previews_cut_characters_not_bytes() {
    let exact = "é".repeat(PREVIEW_CHARS);
    assert_eq!(Preview::Prefix(Cow::Borrowed(&exact)).to_string(), exact);
    let long = "é".repeat(PREVIEW_CHARS + 1);
    assert_eq!(
        Preview::Prefix(Cow::Borrowed(&long)).to_string(),
        format!("{exact}...")
    );
    let emoji = "😀".repeat(60);
    assert_eq!(
        Preview::Prefix(Cow::Borrowed(&emoji)).to_string(),
        format!("{}...", "😀".repeat(50))
    );
    assert_eq!(Preview::Text(Cow::Borrowed(&long)).to_string(), long);
    let owned = Preview::Prefix(Cow::Owned(long.clone()));
    assert_eq!(owned.borrowed(), owned);
}

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

#[test]
fn node_paths_render_and_round_trip_through_the_grammar() {
    let cases: Vec<(Vec<PathStep<'static>>, &str)> = vec![
        (vec![], "/"),
        (vec![PathStep::new("body")], "/body"),
        (
            vec![PathStep::new("body"), PathStep::at("p", Position::new(2))],
            "/body/p[3]",
        ),
        (
            vec![
                PathStep::new("body"),
                PathStep::identity("p", "paraId", Ident::Text("1A2B3C4D")),
                PathStep::at("r", Position::new(0)),
            ],
            "/body/p[@paraId=1A2B3C4D]/r[1]",
        ),
        (
            vec![
                PathStep::at("slide", Position::new(0)),
                PathStep::identity("shape", "id", Ident::Number(5)),
            ],
            "/slide[1]/shape[@id=5]",
        ),
        (
            vec![
                PathStep::at("slide", Position::new(0)),
                PathStep::identity("shape", "name", Ident::Text("Title 1")),
            ],
            r#"/slide[1]/shape[@name="Title 1"]"#,
        ),
        (
            vec![PathStep::new("Sheet1"), PathStep::new("A1:C5")],
            "/Sheet1/A1:C5",
        ),
        (
            vec![PathStep::new("My Data"), PathStep::new("B2")],
            "/My Data/B2",
        ),
        (
            vec![PathStep::new("Q1/Q2"), PathStep::new("A1")],
            r#"/"Q1/Q2"/A1"#,
        ),
        (
            vec![PathStep::new("Data [old]"), PathStep::new("A1")],
            r#"/"Data [old]"/A1"#,
        ),
        (vec![PathStep::new("'quoted")], r#"/"'quoted""#),
        (vec![PathStep::new("say \"hi\"")], r#"/say "hi""#),
        (vec![PathStep::new("2024")], "/2024"),
        (
            vec![PathStep::identity(
                "comment",
                "commentId",
                Ident::Number(u64::MAX),
            )],
            "/comment[@commentId=18446744073709551615]",
        ),
        (
            vec![PathStep::identity("bookmark", "name", Ident::Text("=x]"))],
            r#"/bookmark[@name="=x]"]"#,
        ),
        (
            vec![PathStep::at("p", Position::new(usize::MAX - 1))],
            "/p[18446744073709551615]",
        ),
    ];
    for (steps, expected) in &cases {
        let path = NodePath::new(steps);
        let text = path.to_string();
        assert_eq!(&text, expected);
        let grammar = path.to_path();
        assert_eq!(grammar.to_string(), text, "display matches the grammar's");
        let parsed = Path::parse(&text).unwrap_or_else(|error| panic!("{text:?}: {error}"));
        assert_eq!(parsed, grammar, "{text:?} round-trips");
        assert_eq!(path.is_root(), steps.is_empty());
        assert_eq!(path.steps(), &steps[..]);
        for (step, parsed_step) in steps.iter().zip(parsed.steps()) {
            let back = super::raw::step_from_grammar(parsed_step).expect("an address step");
            assert_eq!(back.name(), step.name());
            assert_eq!(back.to_string(), step.to_string());
            assert_eq!(step.to_step().to_string(), step.to_string());
        }
    }
    assert_eq!(NodePath::ROOT.to_string(), "/");
    assert_eq!(PathStep::new("tc").select(), Select::Bare);
    assert_eq!(
        PathStep::at("tc", Position::new(1)).select(),
        Select::Position(Position::new(1))
    );
    assert_eq!(Ident::from(3_u32), Ident::Number(3));
    assert_eq!(Ident::from(3_u64), Ident::Number(3));
    let parsed = Path::parse("/a[1][2]/b[@x=1 and @y=2]/c[last()]").expect("parses");
    assert!(
        parsed
            .steps()
            .iter()
            .all(|step| PathStep::from_grammar(step).is_none())
    );
}

#[test]
fn node_paths_past_the_step_limit_render_but_do_not_parse() {
    let steps = vec![PathStep::at("item", Position::new(0)); limits::MAX_STEPS];
    let text = NodePath::new(&steps).to_string();
    assert_eq!(
        Path::parse(&text).expect("64 steps parse"),
        NodePath::new(&steps).to_path()
    );

    let steps = vec![PathStep::at("item", Position::new(0)); limits::MAX_STEPS + 1];
    let text = NodePath::new(&steps).to_string();
    assert_eq!(text, "/item[1]".repeat(limits::MAX_STEPS + 1));
    let error = Path::parse(&text).expect_err("65 steps exceed the grammar's limit");
    assert_eq!(error.code(), ErrorCode::MaxDepthExceeded);
}

#[test]
fn node_path_display_writes_into_a_fixed_buffer() {
    let steps = [
        PathStep::new("body"),
        PathStep::identity("p", "paraId", Ident::Number(1_234)),
    ];
    let mut out = String::with_capacity(64);
    write!(out, "{}", NodePath::new(&steps)).expect("writes");
    assert_eq!(out, "/body/p[@paraId=1234]");
}

// ---------------------------------------------------------------------------
// Spec tables
// ---------------------------------------------------------------------------

#[test]
fn the_fake_spec_table_is_consistent() {
    let spec = &fake::SPEC;
    assert_eq!(spec.root, Kind::Document);
    for (index, element) in spec.elements.iter().enumerate() {
        assert_eq!(element.kind as usize, index);
        assert!(
            !element.paths.is_empty(),
            "{} lists a path template",
            element.name
        );
        for path in element.paths {
            assert!(path.starts_with('/'), "{path}");
        }
        for identity in element.identities {
            assert!(
                element
                    .paths
                    .iter()
                    .any(|path| path.contains(&format!("@{identity}=")))
                    || *identity == "name",
                "{} identity {identity}",
                element.name
            );
        }
        for parent in element.parents {
            assert!(
                spec.elements[*parent as usize]
                    .children
                    .contains(&element.kind)
            );
        }
        assert_eq!(element.prop("no such key"), None);
    }
    let paragraph = &spec.elements[Kind::Paragraph as usize];
    assert_eq!(paragraph.prop("align").map(|(index, _)| index), Some(3));
    let chart = &spec.elements[Kind::Chart as usize];
    let gap = chart.gap.expect("the chart kind is a gap");
    let error = gap.error();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.message(), "the `chart` element is not supported yet");
    assert_eq!(
        error.suggestion(),
        Some("the fake leaf has no chart reader (follow-up: L-FAKE-R1)")
    );
    let axis = spec.gaps[0].error();
    assert_eq!(axis.code(), Code::UnsupportedPath);
    assert_eq!(
        axis.message(),
        "the `/header` axis does not exist in this format"
    );
}

#[test]
fn spec_checks_reject_bad_tables() {
    assert!(spec::aliases_valid(&["p", "paragraph"]));
    assert!(spec::aliases_valid(&[]));
    assert!(!spec::aliases_valid(&["paragraph", "p"]), "unsorted");
    assert!(!spec::aliases_valid(&["p", "p"]), "repeated");
    assert!(!spec::aliases_valid(&["P"]), "uppercase");
    assert!(!spec::aliases_valid(&[""]), "empty");
    assert!(!spec::aliases_valid(&["é"]), "non-ASCII");
    assert!(spec::aliases_overlap(&["a", "b"], &["c", "b"]));
    assert!(!spec::aliases_overlap(&["a", "b"], &["ab", "c"]));

    let bold = PropSpec::new("bold", ValueType::Bool);
    let size = PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE));
    assert!(spec::props_valid(&[bold, size]));
    assert!(!spec::props_valid(&[bold, size, bold]), "repeated key");
    assert!(
        !spec::props_valid(&[PropSpec::new("", ValueType::Text)]),
        "empty key"
    );
    assert_eq!(size.ty.schema_type(), "font-size");
    assert_eq!(bold.ty.schema_type(), "bool");

    let gap = |code, follow_up, reason| GapSpec {
        subject: Subject::Property {
            element: "run",
            key: "color",
        },
        code,
        follow_up,
        reason,
    };
    assert!(spec::gaps_valid(&[gap(
        Code::UnsupportedCapability,
        "L-DOCX-R4",
        "no reader"
    )]));
    assert!(!spec::gaps_valid(&[gap(
        Code::NotFound,
        "L-DOCX-R4",
        "no reader"
    )]));
    assert!(!spec::gaps_valid(&[gap(
        Code::UnsupportedPath,
        "",
        "no reader"
    )]));
    assert!(!spec::gaps_valid(&[gap(Code::UnsupportedPath, "M6", "")]));
    let property = gap(Code::UnsupportedCapability, "L-DOCX-R4", "no reader").error();
    assert_eq!(
        property.message(),
        "the `color` property of `run` is not supported yet"
    );
    let view = GapSpec {
        subject: Subject::View("html"),
        code: Code::UnsupportedCapability,
        follow_up: "M3",
        reason: "rendering is out of scope",
    };
    assert_eq!(
        view.error().message(),
        "the `html` view is not supported yet"
    );
    assert_eq!(
        format!("{}", Subject::Element("chart")),
        "the `chart` element"
    );
}

#[test]
fn authored_node_values_paths_and_sink_errors_have_content_free_debug() {
    let value = Value::Text(Cow::Borrowed("PRIVATE"));
    let preview = Preview::Text(Cow::Borrowed("PRIVATE"));
    let steps = [PathStep::new("PRIVATE")];
    for debug in [
        format!("{value:?}"),
        format!("{preview:?}"),
        format!("{:?}", NodePath::new(&steps)),
        format!("{:?}", steps[0]),
        format!("{:?}", Ident::Text("PRIVATE")),
    ] {
        assert!(!debug.contains("PRIVATE"));
    }
    let external = SerializeError::Serializer(io::Error::other("PRIVATE"));
    assert!(std::error::Error::source(&external).is_none());
    assert!(!external.to_string().contains("PRIVATE"));
    let SerializeError::Serializer(cause) = external else {
        panic!("typed cause lost")
    };
    assert_eq!(cause.to_string(), "PRIVATE");
    let error = SerializeError::Serializer("PRIVATE");
    assert!(!format!("{error:?}").contains("PRIVATE"));
    assert!(!error.to_string().contains("PRIVATE"));
}

#[test]
fn nonfinite_node_numbers_fail_with_the_typed_view_error() {
    let document = Document::new(
        Element::new(Kind::Item, "/item[1]").prop("weight", Value::Number(f64::INFINITY)),
    );
    let node = document.root();
    let mut output = Vec::new();
    let result = node
        .view()
        .serialize_into(&mut serde_json::Serializer::new(&mut output));
    assert!(
        matches!(result, Err(SerializeError::View(error)) if error.code() == Code::InvalidValue)
    );
}

#[test]
fn ordinary_grammar_conversion_refuses_native_identity() {
    let path = Path::parse("/p[@paraId=1234]").unwrap();
    assert!(PathStep::from_grammar(&path.steps()[0]).is_none());
    assert!(super::raw::step_from_grammar(&path.steps()[0]).is_some());
}
