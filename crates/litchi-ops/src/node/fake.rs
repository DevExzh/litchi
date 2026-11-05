//! A fake format for testing the node view (record 0010 §10.2, step R1).
//!
//! Compiled for this crate's tests and under the `test-support` feature only.
//! Its element kinds and property tables mirror the keys, order and JSON
//! types OfficeCLI's handlers write for Word, PowerPoint and Excel nodes, so
//! tests can compare the serializer's output with OfficeCLI's `DocumentNode`
//! JSON. A [`Document`] owns a tree of [`Element`]s built during a test's
//! setup; viewing and serializing it only borrow that tree.

use std::borrow::Cow;

use litchi_core::FileFormat;
use litchi_core::Position;

use super::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use super::view::{Children as ViewChildren, Format, Project, Source, ViewOptions};
use super::{Ident, NodePath, PathStep, Preview, Value};
use crate::error::{Code, Error, Result};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// The fake format's element kinds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// A Word document root (`/`, type `document`).
    Document,
    /// A Word body (`/body`).
    Body,
    /// A Word paragraph (`/body/p[N]`).
    Paragraph,
    /// A Word run (`/body/p[N]/r[M]`).
    Run,
    /// A presentation root (`/`, type `presentation`).
    Presentation,
    /// A slide (`/slide[N]`).
    Slide,
    /// A slide shape (`/slide[N]/shape[@id=K]`); its type varies by flavour
    /// (`textbox`, `title`, `shape`), set with [`Element::kind_name`].
    Shape,
    /// A workbook root (`/`, type `workbook`).
    Workbook,
    /// A worksheet (`/Sheet1`).
    Sheet,
    /// A cell (`/Sheet1/A1`).
    Cell,
    /// A chart: the format has it, the fake leaf does not (an element gap).
    Chart,
    /// A generic nestable item, for deep synthetic trees (`/item[1]/item[1]`).
    Item,
}

static ALIGN: Tokens = Tokens::new("alignment", &["center", "justify", "left", "right"], &[]);
static UNDERLINE: Tokens = Tokens::new("underline", &["double", "none", "single", "wave"], &[]);
static CELL_TYPE: Tokens = Tokens::new(
    "cell type",
    &[
        "Boolean",
        "Date",
        "Error",
        "InlineString",
        "Number",
        "SharedString",
        "String",
    ],
    &[],
);

/// Word paragraph spacing: bare numbers are unsigned twips, printed in points.
const SPACING: LengthSpec = LengthSpec::new(LengthContext::WordSpacing, LengthStyle::Points);

/// Slide sizes: printed points-first (`FormatEmuPtFirst`).
const SLIDE_SIZE: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::PointsFirst);

// `WordHandler.Navigation.cs` lines 118-477 (`GetRootNode`): document
// properties on the Word root.
const DOCUMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
];

// Word paragraph keys in `WordHandler.Navigation.cs` order (`ParagraphToNode`,
// lines 4011-4352): `style` and `styleId` both carry the style id;
// `numId`/`numLevel` are strings there, `outlineLvl` an int.
const PARAGRAPH_PROPS: &[PropSpec] = &[
    PropSpec::new("paraId", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
    PropSpec::new("styleId", ValueType::Text),
    PropSpec::new("align", ValueType::Token(&ALIGN)),
    PropSpec::new("spaceBefore", ValueType::Length(SPACING)),
    PropSpec::new("spaceAfter", ValueType::Length(SPACING)),
    PropSpec::new("keepNext", ValueType::Bool),
    PropSpec::new("numId", ValueType::Text),
    PropSpec::new("numLevel", ValueType::Text),
    PropSpec::new("outlineLvl", ValueType::Int),
];

// Word run keys in `WordHandler.Navigation.cs` order (lines 2427-2527).
const RUN_PROPS: &[PropSpec] = &[
    PropSpec::new("font.latin", ValueType::Text),
    PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE)),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("underline", ValueType::Token(&UNDERLINE)),
    PropSpec::new("strike", ValueType::Bool),
    PropSpec::new("superscript", ValueType::Bool),
];

// `PowerPointHandler.Query.cs` lines 28-60: the presentation root.
const PRESENTATION_PROPS: &[PropSpec] = &[
    PropSpec::new("slideWidth", ValueType::Length(SLIDE_SIZE)),
    PropSpec::new("slideHeight", ValueType::Length(SLIDE_SIZE)),
];

// `PowerPointHandler.Query.cs` lines 1297-1312: a slide.
const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("layout", ValueType::Text),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("hidden", ValueType::Bool),
];

// `PowerPointHandler.NodeBuilder.cs` lines 911-1026: a shape.
const SHAPE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("phIndex", ValueType::Int),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("id", ValueType::Int),
    PropSpec::new("isTitle", ValueType::Bool),
    PropSpec::new("x", ValueType::Length(LengthSpec::POSITION)),
    PropSpec::new("y", ValueType::Length(LengthSpec::POSITION)),
    PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("fill", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("rotation", ValueType::Number),
];

// `ExcelHandler.Query.cs` line 129: a sheet listed under the workbook.
const SHEET_PROPS: &[PropSpec] = &[PropSpec::new("name", ValueType::Text)];

// `ExcelHandler.Helpers.Node.cs` lines 368-563: a cell.
const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Token(&CELL_TYPE)),
    PropSpec::new("formula", ValueType::Text),
    PropSpec::new("cachedValue", ValueType::Text),
    PropSpec::new("evaluated", ValueType::Bool),
    PropSpec::new("empty", ValueType::Bool),
    PropSpec::new("font.size", ValueType::Length(LengthSpec::FONT_SIZE)),
    PropSpec::new("font.color", ValueType::Color(ColorContext::Spreadsheet)),
];

const ITEM_PROPS: &[PropSpec] = &[
    PropSpec::new("level", ValueType::Int),
    PropSpec::new("weight", ValueType::Number),
];

static CHART_GAP: GapSpec = GapSpec {
    subject: Subject::Element("chart"),
    code: Code::UnsupportedCapability,
    follow_up: "L-FAKE-R1",
    reason: "the fake leaf has no chart reader",
};

/// The fake format's spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[
        ElementSpec {
            kind: Kind::Document,
            name: "document",
            aliases: &["document"],
            parents: &[],
            children: &[Kind::Body],
            identities: &[],
            paths: &["/"],
            props: DOCUMENT_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Body,
            name: "body",
            aliases: &["body"],
            parents: &[Kind::Document],
            children: &[Kind::Paragraph],
            identities: &[],
            paths: &["/body"],
            props: &[],
            gap: None,
        },
        ElementSpec {
            kind: Kind::Paragraph,
            name: "paragraph",
            aliases: &["p", "paragraph"],
            parents: &[Kind::Body],
            children: &[Kind::Run],
            identities: &["paraId"],
            paths: &["/body/p[N]", "/body/p[@paraId=X]"],
            props: PARAGRAPH_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Run,
            name: "run",
            aliases: &["r", "run"],
            parents: &[Kind::Paragraph],
            children: &[],
            identities: &[],
            paths: &["/body/p[N]/r[M]"],
            props: RUN_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Presentation,
            name: "presentation",
            aliases: &["presentation"],
            parents: &[],
            children: &[Kind::Slide],
            identities: &[],
            paths: &["/"],
            props: PRESENTATION_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Slide,
            name: "slide",
            aliases: &["slide"],
            parents: &[Kind::Presentation],
            children: &[Kind::Shape, Kind::Chart],
            identities: &[],
            paths: &["/slide[N]"],
            props: SLIDE_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Shape,
            name: "shape",
            aliases: &["shape", "textbox", "title"],
            parents: &[Kind::Slide],
            children: &[],
            identities: &["id", "name"],
            paths: &["/slide[N]/shape[K]", "/slide[N]/shape[@id=K]"],
            props: SHAPE_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Workbook,
            name: "workbook",
            aliases: &["workbook"],
            parents: &[],
            children: &[Kind::Sheet],
            identities: &[],
            paths: &["/"],
            props: &[],
            gap: None,
        },
        ElementSpec {
            kind: Kind::Sheet,
            name: "sheet",
            aliases: &["sheet"],
            parents: &[Kind::Workbook],
            children: &[Kind::Cell],
            identities: &[],
            paths: &["/Sheet", "/sheet[N]"],
            props: SHEET_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Cell,
            name: "cell",
            aliases: &["cell"],
            parents: &[Kind::Sheet],
            children: &[],
            identities: &[],
            paths: &["/Sheet/A1"],
            props: CELL_PROPS,
            gap: None,
        },
        ElementSpec {
            kind: Kind::Chart,
            name: "chart",
            aliases: &["chart"],
            parents: &[Kind::Slide],
            children: &[],
            identities: &[],
            paths: &["/slide[N]/chart[K]"],
            props: &[],
            gap: Some(&CHART_GAP),
        },
        ElementSpec {
            kind: Kind::Item,
            name: "item",
            aliases: &["item"],
            parents: &[Kind::Item],
            children: &[Kind::Item],
            identities: &[],
            paths: &["/item[N]"],
            props: ITEM_PROPS,
            gap: None,
        },
    ],
    gaps: &[GapSpec {
        subject: Subject::Axis("/header"),
        code: Code::UnsupportedPath,
        follow_up: "M6",
        reason: "the fake format has no stories",
    }],
};

check_format_spec!(SPEC);

impl Kind {
    fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }
}

/// How an element reports `childCount`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChildCount {
    /// The number of children the element holds.
    Listed,
    /// A fixed count, whether or not the children are listed.
    Known(u32),
    /// Not known without decoding an untouched part (M1-Δ1): omitted.
    Unknown,
}

/// A member read the test asks to fail.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Failure {
    None,
    Text(Code),
    Format(Code),
    Child(usize, Code),
}

/// One element of a fake document, built during a test's setup.
#[derive(Debug, Clone)]
pub struct Element {
    kind: Kind,
    kind_name: Option<&'static str>,
    path: Vec<PathStep<'static>>,
    positional: Option<Vec<PathStep<'static>>>,
    text: Option<Cow<'static, str>>,
    preview: Option<Preview<'static>>,
    style: Option<Cow<'static, str>>,
    child_count: ChildCount,
    values: Vec<Option<Value<'static>>>,
    children: Vec<Element>,
    failure: Failure,
}

impl Element {
    /// An element of `kind` at `path`, written in the subset of the path
    /// grammar node addresses use: `/`, then `/`-separated names, each with
    /// at most one `[N]` or `[@key=value]` (no quoting).
    ///
    /// # Panics
    ///
    /// Panics when `path` is outside that subset.
    #[must_use]
    pub fn new(kind: Kind, path: &'static str) -> Self {
        Self::at(kind, steps(path))
    }

    /// An element of `kind` addressed by `path`.
    #[must_use]
    pub fn at(kind: Kind, path: Vec<PathStep<'static>>) -> Self {
        Self {
            kind,
            kind_name: None,
            path,
            positional: None,
            text: None,
            preview: None,
            style: None,
            child_count: ChildCount::Listed,
            values: vec![None; kind.spec().props.len()],
            children: Vec::new(),
            failure: Failure::None,
        }
    }

    /// The same element reporting `name` as its `type`.
    #[must_use]
    pub fn kind_name(mut self, name: &'static str) -> Self {
        self.kind_name = Some(name);
        self
    }

    /// The same element with a positional path distinct from its canonical
    /// one.
    #[must_use]
    pub fn positional(mut self, path: &'static str) -> Self {
        self.positional = Some(steps(path));
        self
    }

    /// The same element with `text`.
    #[must_use]
    pub fn text(mut self, text: impl Into<Cow<'static, str>>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// The same element with `preview`.
    #[must_use]
    pub fn preview(mut self, preview: Preview<'static>) -> Self {
        self.preview = Some(preview);
        self
    }

    /// The same element with the `style` member.
    #[must_use]
    pub fn style(mut self, style: impl Into<Cow<'static, str>>) -> Self {
        self.style = Some(style.into());
        self
    }

    /// The same element reporting `count` children, or omitting `childCount`
    /// for `None`, instead of counting the children it holds.
    #[must_use]
    pub fn child_count(mut self, count: Option<u32>) -> Self {
        self.child_count = count.map_or(ChildCount::Unknown, ChildCount::Known);
        self
    }

    /// The same element with the property `key` set to `value`.
    ///
    /// # Panics
    ///
    /// Panics when the kind's property table has no `key`.
    #[must_use]
    pub fn prop(mut self, key: &str, value: impl Into<Value<'static>>) -> Self {
        let Some((index, _)) = self.kind.spec().prop(key) else {
            panic!(
                "the fake `{}` kind has no `{key}` property",
                self.kind.spec().name
            );
        };
        self.values[index] = Some(value.into());
        self
    }

    /// The same element with `child` appended.
    #[must_use]
    pub fn child(mut self, child: Self) -> Self {
        self.children.push(child);
        self
    }

    /// The same element failing to read its text with `code`.
    #[must_use]
    pub fn fail_text(mut self, code: Code) -> Self {
        self.failure = Failure::Text(code);
        self
    }

    /// The same element failing to read its properties with `code`.
    #[must_use]
    pub fn fail_format(mut self, code: Code) -> Self {
        self.failure = Failure::Format(code);
        self
    }

    /// The same element failing to produce its child at `index` with `code`.
    #[must_use]
    pub fn fail_child(mut self, index: usize, code: Code) -> Self {
        self.failure = Failure::Child(index, code);
        self
    }

    /// The canonical path.
    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        NodePath::new(&self.path)
    }
}

/// A chain of `levels` nested [`Kind::Item`] elements below a root item:
/// `/item[1]`, `/item[1]/item[1]`, and so on, `levels + 1` elements in all.
#[must_use]
pub fn chain(levels: usize) -> Element {
    let mut path = vec![PathStep::at("item", Position::new(0))];
    let mut paths = Vec::with_capacity(levels + 1);
    for _ in 0..=levels {
        paths.push(path.clone());
        path.push(PathStep::at("item", Position::new(0)));
    }
    let mut element: Option<Element> = None;
    for (depth, path) in paths.into_iter().enumerate().rev() {
        let level = i64::try_from(depth).unwrap_or(i64::MAX);
        let mut item = Element::at(Kind::Item, path).prop("level", level);
        if let Some(child) = element.take() {
            item = item.child(child);
        }
        element = Some(item);
    }
    element.unwrap_or_else(|| Element::new(Kind::Item, "/item[1]"))
}

/// A fake document: a tree of elements, viewed through [`Document::root`].
#[derive(Debug, Clone)]
pub struct Document {
    format: FileFormat,
    root: Element,
}

impl Document {
    /// A document whose root is `root`, reporting [`FileFormat::Docx`].
    #[must_use]
    pub fn new(root: Element) -> Self {
        Self {
            format: FileFormat::Docx,
            root,
        }
    }

    /// The same document reporting `format`.
    #[must_use]
    pub fn with_format(mut self, format: FileFormat) -> Self {
        self.format = format;
        self
    }

    /// The root node.
    #[must_use]
    pub fn root(&self) -> super::Node<'_> {
        super::Node::Fake(Node {
            format: self.format,
            element: &self.root,
        })
    }
}

/// A node of a fake document.
#[derive(Debug, Clone, Copy)]
pub struct Node<'f> {
    format: FileFormat,
    element: &'f Element,
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.element
            .kind_name
            .unwrap_or(self.element.kind.spec().name)
    }

    fn file_format(&self) -> FileFormat {
        self.format
    }

    fn path(&self) -> NodePath<'_> {
        NodePath::new(&self.element.path)
    }

    fn positional_path(&self) -> NodePath<'_> {
        NodePath::new(
            self.element
                .positional
                .as_deref()
                .unwrap_or(&self.element.path),
        )
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        if let Failure::Text(code) = self.element.failure {
            return Err(Error::new(code, "the fake text read failed"));
        }
        Ok(self.element.text.as_deref().map(Cow::Borrowed))
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(self.element.preview.as_ref().map(Preview::borrowed))
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(self.element.style.as_deref().map(Cow::Borrowed))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        Ok(match self.element.child_count {
            ChildCount::Listed => {
                Some(u32::try_from(self.element.children.len()).unwrap_or(u32::MAX))
            },
            ChildCount::Known(count) => Some(count),
            ChildCount::Unknown => None,
        })
    }

    fn format(&self) -> Result<Format<'_>> {
        if let Failure::Format(code) = self.element.failure {
            return Err(Error::new(code, "the fake property read failed"));
        }
        Ok(Format::new(
            self.element.kind.spec().props,
            Source::Fake(&self.element.values),
        ))
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        let failure = match self.element.failure {
            Failure::Child(index, code) => Some((index, code)),
            _ => None,
        };
        ViewChildren::fake(Children {
            format: self.format,
            children: self.element.children.iter().enumerate(),
            failure,
            _node: std::marker::PhantomData,
        })
    }
}

/// The children of a fake node.
pub(crate) struct Children<'n, 'f> {
    format: FileFormat,
    children: std::iter::Enumerate<std::slice::Iter<'f, Element>>,
    failure: Option<(usize, Code)>,
    _node: std::marker::PhantomData<&'n ()>,
}

impl<'f> Iterator for Children<'_, 'f> {
    type Item = Result<super::Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        let (index, element) = self.children.next()?;
        if let Some((failing, code)) = self.failure
            && failing == index
        {
            return Some(Err(Error::new(code, "the fake child read failed")));
        }
        Some(Ok(super::Node::Fake(Node {
            format: self.format,
            element,
        })))
    }
}

/// The steps of a path in the subset [`Element::new`] accepts.
fn steps(path: &'static str) -> Vec<PathStep<'static>> {
    assert!(path.starts_with('/'), "a fake path is absolute: {path:?}");
    let rest = &path[1..];
    if rest.is_empty() {
        return Vec::new();
    }
    rest.split('/')
        .map(|segment| {
            let Some((name, selector)) = segment.split_once('[') else {
                return PathStep::new(segment);
            };
            let selector = selector
                .strip_suffix(']')
                .unwrap_or_else(|| panic!("an unclosed selector in {path:?}"));
            if let Some(identity) = selector.strip_prefix('@') {
                let (key, value) = identity
                    .split_once('=')
                    .unwrap_or_else(|| panic!("an identity without `=` in {path:?}"));
                PathStep::identity(name, key, Ident::Text(value))
            } else {
                let index: usize = selector
                    .parse()
                    .unwrap_or_else(|_| panic!("a selector that is not an index in {path:?}"));
                let position = Position::from_one_based(index)
                    .unwrap_or_else(|| panic!("index 0 in {path:?}"));
                PathStep::at(name, position)
            }
        })
        .collect()
}
