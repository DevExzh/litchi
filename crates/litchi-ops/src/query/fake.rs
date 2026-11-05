//! Test-only fake formats for the query walker (record 0010 §10.2, step R3).
//!
//! Three small formats, one per family vocabulary: a Word-like document, a
//! presentation and a workbook. Each has a field-less kind enum that
//! implements [`ElementKind`] over the family vocabulary (record 0010 §4.2)
//! and an R1 spec table checked by `check_format_spec!` (record 0010 §7.2),
//! and its documents are trees of [`Item`]s built during a test's setup. The
//! walker reaches them only through [`Navigate`] and [`Inspect`], so the
//! tests exercise exactly the seam a real adapter implements (M1 step R4).
//!
//! The workbook keeps rows and cells as sibling children of a worksheet, in
//! row-major order, and lists its tables (ListObjects) beside the tree for
//! header-row binding. Nothing here allocates on the walk: nodes are
//! references, properties are stored aligned with the spec tables, and text
//! is borrowed.

use std::borrow::Cow;
use std::cell::Cell;
use std::fmt::{self, Write as _};
use std::marker::PhantomData;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRef, Position};

use super::inspect::{HeaderBinding, HeaderKey, Inspect};
use crate::error::{Code, Error, Result};
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use crate::node::{Format, PathStep, Source, Value};
use crate::path::navigate::{ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// A fake format's kind enum: its vocabulary names, spec table and roots.
pub(crate) trait FakeKind: ElementKind {
    /// The spec table, indexed by ordinal.
    const SPEC: &'static FormatSpec<Self>;
    /// The format the fake reports.
    const FORMAT: FileFormat;
    /// The containers `*` lists the children of.
    const TOP: &'static [Self];
    /// Vocabulary names the format lacks as a concept.
    const ABSENT: &'static [&'static str];
    /// The worksheet row kind, for rows without a record.
    const ROW: Option<Self>;
    /// Kinds that occur once under their parent, whose path step has no
    /// index (`/body`, as OfficeCLI writes it).
    const SINGLETONS: &'static [Self] = &[];
}

static ALIGN: Tokens = Tokens::new("alignment", &["left", "center", "right", "justify"], &[]);
static UNDERLINE: Tokens = Tokens::new("underline", &["none", "single", "double", "wave"], &[]);
static HEADER_TYPE: Tokens = Tokens::new("header type", &["default", "first", "even"], &[]);
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

/// Word paragraph spacing: any length, printed in points.
const SPACING: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::Points);

static CHART_GAP: GapSpec = GapSpec {
    subject: Subject::Element("chart"),
    code: Code::UnsupportedCapability,
    follow_up: "L-FAKE-R1",
    reason: "the fake leaf has no chart reader",
};

static PIVOT_GAP: GapSpec = GapSpec {
    subject: Subject::Element("pivottable"),
    code: Code::UnsupportedCapability,
    follow_up: "L-FAKE-R2",
    reason: "the fake leaf has no pivot-table reader",
};

/// Declares a field-less kind enum with its vocabulary mapping.
macro_rules! kinds {
    (
        $(#[$meta:meta])*
        $name:ident: $vocabulary:expr, root $root:ident {
            $($kind:ident => $canonical:literal,)+
        }
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        pub(crate) enum $name {
            $(
                #[doc = concat!("The `", $canonical, "` element.")]
                $kind,
            )+
        }

        impl ElementKind for $name {
            const VOCABULARY: &'static Vocabulary = $vocabulary;
            const ROOT: Self = Self::$root;

            fn from_name(name: &str) -> Option<Self> {
                match name {
                    $($canonical => Some(Self::$kind),)+
                    _ => None,
                }
            }

            fn name(self) -> &'static str {
                match self {
                    $(Self::$kind => $canonical,)+
                }
            }

            fn ordinal(self) -> u8 {
                self as u8
            }
        }
    };
}

kinds! {
    /// The kinds of the fake Word-like format.
    WordKind: &vocabulary::WORD, root Document {
        Document => "document",
        Body => "body",
        Paragraph => "p",
        Run => "r",
        Hyperlink => "hyperlink",
        Table => "tbl",
        Row => "tr",
        Cell => "tc",
        Sdt => "sdt",
        Header => "header",
        Footer => "footer",
        Picture => "picture",
        Chart => "chart",
    }
}

kinds! {
    /// The kinds of the fake presentation format.
    SlideKind: &vocabulary::PRESENTATION, root Presentation {
        Presentation => "presentation",
        Slide => "slide",
        Shape => "shape",
        Picture => "picture",
        Table => "table",
        Row => "tr",
        Cell => "tc",
        Paragraph => "paragraph",
        Run => "run",
        Group => "group",
        Chart => "chart",
        Notes => "notes",
    }
}

kinds! {
    /// The kinds of the fake workbook format.
    SheetKind: &vocabulary::WORKBOOK, root Workbook {
        Workbook => "workbook",
        Sheet => "sheet",
        Row => "row",
        Cell => "cell",
        Column => "col",
        Table => "table",
        NamedRange => "namedrange",
        Chart => "chart",
        PivotTable => "pivottable",
    }
}

/// An element spec with the fields the fake does not vary filled in.
const fn element<K: Copy>(
    kind: K,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [K],
    children: &'static [K],
    props: &'static [PropSpec],
) -> ElementSpec<K> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities: &[],
        paths: &["/"],
        props,
        gap: None,
    }
}

const fn with_gap<K>(mut spec: ElementSpec<K>, gap: &'static GapSpec) -> ElementSpec<K> {
    spec.gap = Some(gap);
    spec
}

const fn with_identities<K>(
    mut spec: ElementSpec<K>,
    identities: &'static [&'static str],
) -> ElementSpec<K> {
    spec.identities = identities;
    spec
}

pub(crate) static WORD_SPEC: FormatSpec<WordKind> = {
    use WordKind::{
        Body, Cell, Chart, Document, Footer, Header, Hyperlink, Paragraph, Picture, Row, Run, Sdt,
        Table,
    };
    const STORIES: &[WordKind] = &[Body, Cell, Sdt, Header, Footer];
    FormatSpec {
        root: Document,
        elements: &[
            element(
                Document,
                "document",
                &["document"],
                &[],
                &[Body, Header, Footer],
                &[
                    PropSpec::new("title", ValueType::Text),
                    PropSpec::new("author", ValueType::Text),
                ],
            ),
            element(
                Body,
                "body",
                &["body"],
                &[Document],
                &[Paragraph, Table, Sdt],
                &[],
            ),
            with_identities(
                element(
                    Paragraph,
                    "paragraph",
                    &["p", "paragraph"],
                    STORIES,
                    &[Run, Hyperlink, Picture, Chart],
                    &[
                        PropSpec::new("paraId", ValueType::Text),
                        PropSpec::new("style", ValueType::Text),
                        PropSpec::new("align", ValueType::Token(&ALIGN)),
                        PropSpec::new("spaceBefore", ValueType::Length(SPACING)),
                        PropSpec::new("keepNext", ValueType::Bool),
                        PropSpec::new("numId", ValueType::Int),
                        PropSpec::new("outlineLvl", ValueType::Int),
                    ],
                ),
                &["paraId"],
            ),
            element(
                Run,
                "run",
                &["r", "run"],
                &[Paragraph, Hyperlink],
                &[],
                &[
                    PropSpec::new("font.latin", ValueType::Text),
                    PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE)),
                    PropSpec::new("bold", ValueType::Bool),
                    PropSpec::new("italic", ValueType::Bool),
                    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
                    PropSpec::new("underline", ValueType::Token(&UNDERLINE)),
                ],
            ),
            element(
                Hyperlink,
                "hyperlink",
                &["hyperlink"],
                &[Paragraph],
                &[Run],
                &[PropSpec::new("url", ValueType::Text)],
            ),
            element(
                Table,
                "table",
                &["table", "tbl"],
                STORIES,
                &[Row],
                &[PropSpec::new("style", ValueType::Text)],
            ),
            element(Row, "row", &["row", "tr"], &[Table], &[Cell], &[]),
            element(
                Cell,
                "cell",
                &["cell", "tc"],
                &[Row],
                &[Paragraph, Table],
                &[PropSpec::new("gridSpan", ValueType::Int)],
            ),
            element(
                Sdt,
                "sdt",
                &["contentcontrol", "sdt"],
                &[Body],
                &[Paragraph, Table],
                &[
                    PropSpec::new("tag", ValueType::Text),
                    PropSpec::new("sdtId", ValueType::Int),
                ],
            ),
            element(
                Header,
                "header",
                &["header"],
                &[Document],
                &[Paragraph, Table],
                &[PropSpec::new("type", ValueType::Token(&HEADER_TYPE))],
            ),
            element(
                Footer,
                "footer",
                &["footer"],
                &[Document],
                &[Paragraph, Table],
                &[PropSpec::new("type", ValueType::Token(&HEADER_TYPE))],
            ),
            element(
                Picture,
                "picture",
                &["image", "img", "picture"],
                &[Paragraph],
                &[],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("alt", ValueType::Text),
                    PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
                    PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
                ],
            ),
            with_gap(
                element(Chart, "chart", &["chart"], &[Paragraph], &[], &[]),
                &CHART_GAP,
            ),
        ],
        gaps: &[GapSpec {
            subject: Subject::Property {
                element: "run",
                key: "highlight",
            },
            code: Code::UnsupportedCapability,
            follow_up: "L-DOCX-R4",
            reason: "the fake leaf has no highlight reader",
        }],
    }
};

check_format_spec!(WORD_SPEC);

impl FakeKind for WordKind {
    const SPEC: &'static FormatSpec<Self> = &WORD_SPEC;
    const FORMAT: FileFormat = FileFormat::Docx;
    const TOP: &'static [Self] = &[Self::Body];
    const ABSENT: &'static [&'static str] = &["customXml"];
    const ROW: Option<Self> = None;
    const SINGLETONS: &'static [Self] = &[Self::Body];
}

pub(crate) static SLIDE_SPEC: FormatSpec<SlideKind> = {
    use SlideKind::{
        Cell, Chart, Group, Notes, Paragraph, Picture, Presentation, Row, Run, Shape, Slide, Table,
    };
    FormatSpec {
        root: Presentation,
        elements: &[
            element(
                Presentation,
                "presentation",
                &["presentation"],
                &[],
                &[Slide],
                &[],
            ),
            element(
                Slide,
                "slide",
                &["slide"],
                &[Presentation],
                &[Shape, Picture, Table, Group, Chart, Notes],
                &[
                    PropSpec::new("layout", ValueType::Text),
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("hidden", ValueType::Bool),
                ],
            ),
            with_identities(
                element(
                    Shape,
                    "shape",
                    &["shape", "textbox", "title"],
                    &[Slide, Group],
                    &[Paragraph],
                    &[
                        PropSpec::new("name", ValueType::Text),
                        PropSpec::new("id", ValueType::Int),
                        PropSpec::new("isTitle", ValueType::Bool),
                        PropSpec::new("alt", ValueType::Text),
                        PropSpec::new("x", ValueType::Length(LengthSpec::POSITION)),
                        PropSpec::new("y", ValueType::Length(LengthSpec::POSITION)),
                        PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
                        PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
                        PropSpec::new("fill", ValueType::Color(ColorContext::Drawing)),
                    ],
                ),
                &["id", "name"],
            ),
            element(
                Picture,
                "picture",
                &["pic", "picture"],
                &[Slide, Group],
                &[],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("id", ValueType::Int),
                    PropSpec::new("alt", ValueType::Text),
                    PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
                    PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
                ],
            ),
            element(
                Table,
                "table",
                &["table"],
                &[Slide],
                &[Row],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("id", ValueType::Int),
                ],
            ),
            element(Row, "tr", &["row", "tr"], &[Table], &[Cell], &[]),
            element(Cell, "tc", &["cell", "tc"], &[Row], &[Paragraph], &[]),
            element(
                Paragraph,
                "paragraph",
                &["p", "paragraph"],
                &[Shape, Cell, Notes],
                &[Run],
                &[PropSpec::new("align", ValueType::Token(&ALIGN))],
            ),
            element(
                Run,
                "run",
                &["r", "run"],
                &[Paragraph],
                &[],
                &[
                    PropSpec::new("bold", ValueType::Bool),
                    PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE)),
                    PropSpec::new("font", ValueType::Text),
                ],
            ),
            element(
                Group,
                "group",
                &["group"],
                &[Slide, Group],
                &[Shape, Picture, Group],
                &[PropSpec::new("name", ValueType::Text)],
            ),
            with_gap(
                element(Chart, "chart", &["chart"], &[Slide], &[], &[]),
                &CHART_GAP,
            ),
            element(Notes, "notes", &["notes"], &[Slide], &[Paragraph], &[]),
        ],
        gaps: &[GapSpec {
            subject: Subject::Property {
                element: "group",
                key: "alt",
            },
            code: Code::UnsupportedCapability,
            follow_up: "L-PPTX-R6",
            reason: "the fake leaf does not keep a group's description",
        }],
    }
};

check_format_spec!(SLIDE_SPEC);

impl FakeKind for SlideKind {
    const SPEC: &'static FormatSpec<Self> = &SLIDE_SPEC;
    const FORMAT: FileFormat = FileFormat::Pptx;
    const TOP: &'static [Self] = &[Self::Slide];
    const ABSENT: &'static [&'static str] = &[];
    const ROW: Option<Self> = None;
    const SINGLETONS: &'static [Self] = &[Self::Notes];
}

pub(crate) static SHEET_SPEC: FormatSpec<SheetKind> = {
    use SheetKind::{Cell, Chart, Column, NamedRange, PivotTable, Row, Sheet, Table, Workbook};
    FormatSpec {
        root: Workbook,
        elements: &[
            element(
                Workbook,
                "workbook",
                &["workbook"],
                &[],
                &[Sheet, NamedRange],
                &[],
            ),
            element(
                Sheet,
                "sheet",
                &["sheet"],
                &[Workbook],
                &[Column, Row, Cell, Table, Chart, PivotTable],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("hidden", ValueType::Bool),
                ],
            ),
            element(
                Row,
                "row",
                &["row"],
                &[Sheet],
                &[],
                &[
                    PropSpec::new("height", ValueType::Number),
                    PropSpec::new("hidden", ValueType::Bool),
                    PropSpec::new("outlineLevel", ValueType::Int),
                    PropSpec::new("collapsed", ValueType::Bool),
                    PropSpec::new("customHeight", ValueType::Bool),
                ],
            ),
            element(
                Cell,
                "cell",
                &["cell"],
                &[Sheet],
                &[],
                &[
                    PropSpec::new("type", ValueType::Token(&CELL_TYPE)),
                    PropSpec::new("formula", ValueType::Text),
                    PropSpec::new("cachedValue", ValueType::Text),
                    PropSpec::new("font.bold", ValueType::Bool),
                    PropSpec::new("font.italic", ValueType::Bool),
                    PropSpec::new("font.name", ValueType::Text),
                    PropSpec::new("font.size", ValueType::Length(LengthSpec::FONT_SIZE)),
                    PropSpec::new("font.color", ValueType::Color(ColorContext::Spreadsheet)),
                ],
            ),
            element(
                Column,
                "column",
                &["col", "column"],
                &[Sheet],
                &[],
                &[
                    PropSpec::new("width", ValueType::Number),
                    PropSpec::new("hidden", ValueType::Bool),
                ],
            ),
            element(
                Table,
                "table",
                &["listobject", "table"],
                &[Sheet],
                &[],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("ref", ValueType::Text),
                ],
            ),
            element(
                NamedRange,
                "namedrange",
                &["definedname", "namedrange"],
                &[Workbook],
                &[],
                &[
                    PropSpec::new("name", ValueType::Text),
                    PropSpec::new("ref", ValueType::Text),
                ],
            ),
            with_gap(
                element(Chart, "chart", &["chart"], &[Sheet], &[], &[]),
                &CHART_GAP,
            ),
            with_gap(
                element(
                    PivotTable,
                    "pivottable",
                    &["pivot", "pivottable"],
                    &[Sheet],
                    &[],
                    &[],
                ),
                &PIVOT_GAP,
            ),
        ],
        gaps: &[],
    }
};

check_format_spec!(SHEET_SPEC);

impl FakeKind for SheetKind {
    const SPEC: &'static FormatSpec<Self> = &SHEET_SPEC;
    const FORMAT: FileFormat = FileFormat::Xlsx;
    const TOP: &'static [Self] = &[Self::Sheet];
    const ABSENT: &'static [&'static str] = &[];
    const ROW: Option<Self> = Some(Self::Row);
}

fn spec<K: FakeKind>(kind: K) -> &'static ElementSpec<K> {
    &K::SPEC.elements[usize::from(kind.ordinal())]
}

/// One element of a fake document, built during a test's setup.
#[derive(Debug, Clone)]
pub(crate) struct Item<K> {
    kind: K,
    flavour: Option<&'static str>,
    text: Option<String>,
    style: Option<String>,
    values: Vec<Option<Value<'static>>>,
    children: Vec<Item<K>>,
    number: Option<usize>,
    cell: Option<CellRef>,
    name: Option<String>,
    raw: Option<Value<'static>>,
    fail_child: Option<usize>,
    fail_text: bool,
    /// A worksheet's rows by number and cells by `(row, column)`, zero-based
    /// and sorted, with each child's index: built by [`Doc::new`], so row
    /// and cell lookups cost O(log n), as a real sheet store's do.
    rows: Vec<(usize, usize)>,
    cells: Vec<((u32, u32), usize)>,
}

impl<K: FakeKind> Item<K> {
    /// An element of `kind`.
    pub(crate) fn new(kind: K) -> Self {
        Self {
            kind,
            flavour: None,
            text: None,
            style: None,
            values: vec![None; spec(kind).props.len()],
            children: Vec::new(),
            number: None,
            cell: None,
            name: None,
            raw: None,
            fail_child: None,
            fail_text: false,
            rows: Vec::new(),
            cells: Vec::new(),
        }
    }

    /// Indexes a worksheet's rows and cells.
    fn index(&mut self) {
        self.rows.clear();
        self.cells.clear();
        for (index, child) in self.children.iter().enumerate() {
            if Some(child.kind) == K::ROW {
                if let Some(number) = child.number {
                    self.rows.push((number, index));
                }
            }
            if let Some(cell) = child.cell {
                self.cells.push(((cell.row(), cell.column()), index));
            }
        }
        self.rows.sort_unstable();
        self.cells.sort_unstable();
    }

    /// The row child numbered `number`.
    fn row_numbered(&self, number: usize) -> Option<&Item<K>> {
        let found = self
            .rows
            .binary_search_by_key(&number, |&(held, _)| held)
            .ok()?;
        self.children.get(self.rows[found].1)
    }

    /// The cell child at zero-based `(row, column)`.
    fn cell_at(&self, row: u32, column: u32) -> Option<&Item<K>> {
        let found = self
            .cells
            .binary_search_by_key(&(row, column), |&(held, _)| held)
            .ok()?;
        self.children.get(self.cells[found].1)
    }

    /// The element's kind.
    pub(crate) fn kind(&self) -> K {
        self.kind
    }

    /// The same element with text.
    #[must_use]
    pub(crate) fn text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// The element's text.
    pub(crate) fn text_str(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// The same element with a `style` member.
    #[must_use]
    pub(crate) fn style(mut self, style: impl Into<String>) -> Self {
        self.style = Some(style.into());
        self
    }

    /// The same element reporting `flavour` as its OfficeCLI type.
    #[must_use]
    pub(crate) fn flavour(mut self, flavour: &'static str) -> Self {
        self.flavour = Some(flavour);
        self
    }

    /// The same element with property `key` set.
    ///
    /// # Panics
    ///
    /// When the kind declares no `key`.
    #[must_use]
    pub(crate) fn prop(mut self, key: &str, value: impl Into<Value<'static>>) -> Self {
        let Some(index) = spec(self.kind)
            .props
            .iter()
            .position(|prop| prop.name == key)
        else {
            panic!(
                "the fake `{}` kind has no `{key}` property",
                spec(self.kind).name
            );
        };
        self.values[index] = Some(value.into());
        self
    }

    /// The same element with `child` appended.
    #[must_use]
    pub(crate) fn child(mut self, child: Self) -> Self {
        self.children.push(child);
        self
    }

    /// The same element with `children` appended.
    #[must_use]
    pub(crate) fn children(mut self, children: impl IntoIterator<Item = Self>) -> Self {
        self.children.extend(children);
        self
    }

    /// The same element carrying `number` (a row number, a level).
    #[must_use]
    pub(crate) fn number(mut self, number: usize) -> Self {
        self.number = Some(number);
        self
    }

    /// The same cell at `a1`.
    ///
    /// # Panics
    ///
    /// When `a1` is not an A1 reference.
    #[must_use]
    pub(crate) fn cell(mut self, a1: &str) -> Self {
        self.cell = Some(CellRef::parse(a1).unwrap_or_else(|| panic!("not a cell: {a1}")));
        self
    }

    /// The same element named `name` (a worksheet).
    #[must_use]
    pub(crate) fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The same cell with a raw stored value that differs from its text.
    #[must_use]
    pub(crate) fn raw(mut self, value: impl Into<Value<'static>>) -> Self {
        self.raw = Some(value.into());
        self
    }

    /// The same element failing to produce its child at `index`.
    #[must_use]
    pub(crate) fn fail_child(mut self, index: usize) -> Self {
        self.fail_child = Some(index);
        self
    }

    /// The same element failing to read its text.
    #[must_use]
    pub(crate) fn fail_text(mut self) -> Self {
        self.fail_text = true;
        self
    }

    /// The text form of the property `key`, for identity lookups.
    fn ident(&self, key: &str) -> Option<String> {
        let index = spec(self.kind)
            .props
            .iter()
            .position(|prop| prop.name.eq_ignore_ascii_case(key))?;
        self.values[index].as_ref().map(ToString::to_string)
    }
}

/// A ListObject of the fake workbook.
#[derive(Debug, Clone)]
pub(crate) struct Table {
    /// The worksheet's position.
    pub(crate) sheet: usize,
    /// The header names.
    pub(crate) columns: Vec<String>,
    /// The 1-based column of the first header.
    pub(crate) first_column: u32,
    /// The 1-based header row.
    pub(crate) header_row: u32,
    /// The 1-based last row, totals row included.
    pub(crate) last_row: u32,
    /// Whether the last row is a totals row.
    pub(crate) totals: bool,
}

/// A fake document.
#[derive(Debug, Clone)]
pub(crate) struct Doc<K> {
    /// The root element.
    pub(crate) root: Item<K>,
    /// The workbook's tables.
    pub(crate) tables: Vec<Table>,
}

impl<K: FakeKind> Doc<K> {
    /// A document with `root` and no tables.
    pub(crate) fn new(mut root: Item<K>) -> Self {
        for child in &mut root.children {
            if child.kind.name() == "sheet" {
                child.index();
            }
        }
        Self {
            root,
            tables: Vec::new(),
        }
    }

    /// The same document with `table`.
    #[must_use]
    pub(crate) fn table(mut self, table: Table) -> Self {
        self.tables.push(table);
        self
    }
}

/// A node of a fake document.
#[derive(Debug)]
pub(crate) enum Node<'d, K> {
    /// An element of the tree; `sheet` is its worksheet when it is a direct
    /// child of one.
    Item {
        item: &'d Item<K>,
        sheet: Option<&'d Item<K>>,
    },
    /// A worksheet row without a record.
    Row { sheet: &'d Item<K>, number: u32 },
}

impl<K> Clone for Node<'_, K> {
    fn clone(&self) -> Self {
        *self
    }
}

impl<K> Copy for Node<'_, K> {}

impl<'d, K> Node<'d, K> {
    /// The element, unless the node is a row without a record.
    pub(crate) fn item(&self) -> Option<&'d Item<K>> {
        match *self {
            Self::Item { item, .. } => Some(item),
            Self::Row { .. } => None,
        }
    }
}

/// The fake adapter over one document; counts the nodes it produces.
#[derive(Debug)]
pub(crate) struct Nav<'d, K> {
    doc: &'d Doc<K>,
    produced: Cell<u64>,
}

impl<'d, K: FakeKind> Nav<'d, K> {
    pub(crate) fn new(doc: &'d Doc<K>) -> Self {
        Self {
            doc,
            produced: Cell::new(0),
        }
    }

    /// The nodes every children iterator has produced so far.
    pub(crate) fn produced(&self) -> u64 {
        self.produced.get()
    }

    /// The worksheets in a header-binding scope: every sheet, or the scope
    /// itself.
    fn sheets_in(&self, scope: &Node<'d, K>) -> Vec<usize> {
        let sheets = self
            .doc
            .root
            .children
            .iter()
            .filter(|child| child.kind.name() == "sheet");
        match scope.item() {
            Some(item) if std::ptr::eq(item, &self.doc.root) => (0..sheets.count()).collect(),
            Some(item) => sheets
                .enumerate()
                .filter(|(_, sheet)| std::ptr::eq(*sheet, item))
                .map(|(index, _)| index)
                .collect(),
            None => Vec::new(),
        }
    }
}

/// The children of a fake node whose kinds are in a set.
pub(crate) struct Children<'a, 'd, K> {
    nav: &'a Nav<'d, K>,
    items: std::slice::Iter<'d, Item<K>>,
    kinds: KindSet<K>,
    sheet: Option<&'d Item<K>>,
    fail_at: Option<usize>,
    index: usize,
}

impl<'d, K: FakeKind> Iterator for Children<'_, 'd, K> {
    type Item = Result<Node<'d, K>>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let item = self.items.next()?;
            let index = self.index;
            self.index += 1;
            if self.fail_at == Some(index) {
                self.items = [].iter();
                return Some(Err(Error::new(
                    Code::CorruptFile,
                    "the fake child read failed",
                )));
            }
            if !self.kinds.contains(item.kind) {
                continue;
            }
            self.nav.produced.set(self.nav.produced.get() + 1);
            return Some(Ok(Node::Item {
                item,
                sheet: self.sheet,
            }));
        }
    }
}

/// The fake adapter's children iterators count what they produce through the
/// adapter, so the traits are implemented on a reference to it: a walk then
/// owns the reference, and its iterators borrow the adapter.
impl<'a, 'd, K: FakeKind> Navigate<'d> for &'a Nav<'d, K> {
    type Kind = K;
    type Node = Node<'d, K>;
    type Gap = GapSpec;
    type Children = Children<'a, 'd, K>;

    const FORMAT: FileFormat = K::FORMAT;
    const ABSENT: &'static [&'static str] = K::ABSENT;

    fn root(&self) -> Result<Node<'d, K>> {
        Ok(Node::Item {
            item: &self.doc.root,
            sheet: None,
        })
    }

    fn child(
        &self,
        parent: &Node<'d, K>,
        step: &Resolved<'_, K>,
    ) -> Result<Lookup<Node<'d, K>, GapSpec>> {
        let Some(parent) = parent.item() else {
            return Ok(Lookup::NotApplicable);
        };
        let kind = step.kind;
        if let Some(gap) = spec(kind).gap {
            return Ok(Lookup::Gap(gap));
        }
        let sheet = (parent.kind.name() == "sheet").then_some(parent);
        let mut of_kind = parent.children.iter().filter(|child| child.kind == kind);
        let found = match step.select {
            Select::None => of_kind.next(),
            Select::Ordinal(position) => of_kind.nth(position.get()),
            Select::Last => of_kind.last(),
            Select::Number(number) if Some(kind) == K::ROW && sheet.is_some() => {
                let hit = parent.row_numbered(number);
                if hit.is_none() {
                    if let (Some(sheet), Ok(number)) = (sheet, u32::try_from(number)) {
                        return Ok(Lookup::Found(Node::Row { sheet, number }));
                    }
                }
                hit
            },
            Select::Number(number) => of_kind.find(|child| child.number == Some(number)),
            Select::Name { name, case } => of_kind.find(|child| {
                child
                    .name
                    .as_deref()
                    .is_some_and(|held| case.matches(held, name))
            }),
            Select::Identity(identities) => of_kind.find(|child| {
                identities.iter().all(|(identity, value)| {
                    child
                        .ident(identity.key)
                        .is_some_and(|held| identity.case.matches(&held, value))
                })
            }),
            Select::Cell(at) => parent
                .cell_at(at.row(), at.column())
                .filter(|child| child.kind == kind),
            Select::Token(_) | Select::Range(_) => return Ok(Lookup::NotApplicable),
        };
        Ok(found.map_or(Lookup::Missing, |item| {
            Lookup::Found(Node::Item { item, sheet })
        }))
    }

    fn children(&self, parent: &Node<'d, K>, kinds: KindSet<K>) -> Result<Children<'a, 'd, K>> {
        let (items, sheet, fail_at) = match parent.item() {
            Some(item) => (
                item.children.iter(),
                (item.kind.name() == "sheet").then_some(item),
                item.fail_child,
            ),
            None => ([].iter(), None, None),
        };
        Ok(Children {
            nav: *self,
            items,
            kinds,
            sheet,
            fail_at,
            index: 0,
        })
    }
}

impl<'d, K: FakeKind> Inspect<'d> for &Nav<'d, K> {
    const SPEC: &'static FormatSpec<K> = K::SPEC;
    const TOP: &'static [K] = K::TOP;

    fn kind(&self, node: &Node<'d, K>) -> K {
        match node {
            Node::Item { item, .. } => item.kind,
            Node::Row { .. } => K::ROW.unwrap_or(K::ROOT),
        }
    }

    fn type_name(&self, node: &Node<'d, K>) -> &'static str {
        match node {
            Node::Item { item, .. } => item.flavour.unwrap_or(spec(item.kind).name),
            Node::Row { .. } => "row",
        }
    }

    fn text<'n>(&'n self, node: &'n Node<'d, K>) -> Result<Option<Cow<'n, str>>> {
        match node {
            Node::Item { item, .. } if item.fail_text => {
                Err(Error::new(Code::CorruptFile, "the fake text read failed"))
            },
            Node::Item { item, .. } => Ok(item.text.as_deref().map(Cow::Borrowed)),
            Node::Row { .. } => Ok(None),
        }
    }

    fn style<'n>(&'n self, node: &'n Node<'d, K>) -> Result<Option<Cow<'n, str>>> {
        Ok(node
            .item()
            .and_then(|item| item.style.as_deref())
            .map(Cow::Borrowed))
    }

    fn format<'n>(&'n self, node: &'n Node<'d, K>) -> Result<Format<'n>> {
        let kind = self.kind(node);
        Ok(match node.item() {
            Some(item) => Format::new(spec(kind).props, Source::Fake(&item.values)),
            None => Format::new(spec(kind).props, Source::Empty(PhantomData)),
        })
    }

    fn number(&self, node: &Node<'d, K>) -> Option<usize> {
        match node {
            Node::Item { item, .. } => item.number,
            Node::Row { number, .. } => usize::try_from(*number).ok(),
        }
    }

    fn cell(&self, node: &Node<'d, K>) -> Option<CellRef> {
        node.item().and_then(|item| item.cell)
    }

    fn write_step(
        &self,
        node: &Node<'d, K>,
        kind: K,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        if let Some(number) = self.number(node) {
            if Some(kind) == K::ROW || kind.name() == "col" {
                return write!(out, "{}[{number}]", kind.name());
            }
        }
        let Some(item) = node.item() else {
            return fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out);
        };
        if let Some(name) = item.name.as_deref().filter(|_| kind.name() == "sheet") {
            return fmt::Display::fmt(&PathStep::new(name), out);
        }
        if let Some(cell) = item.cell {
            return write_cell(cell, out);
        }
        if K::SINGLETONS.contains(&kind) {
            return fmt::Display::fmt(&PathStep::new(kind.name()), out);
        }
        fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out)
    }

    fn bind_header(&self, scope: &Node<'d, K>, keys: &[HeaderKey<'_>]) -> Result<HeaderBinding> {
        let sheets = self.sheets_in(scope);
        let mut hits = Vec::new();
        for table in self
            .doc
            .tables
            .iter()
            .filter(|table| sheets.contains(&table.sheet))
        {
            let last_column = table.first_column + table.columns.len() as u32 - 1;
            let columns: Option<Vec<u32>> = keys
                .iter()
                .map(|key| {
                    let by_name = table
                        .columns
                        .iter()
                        .position(|column| column.eq_ignore_ascii_case(key.name))
                        .map(|index| table.first_column + index as u32);
                    by_name.or_else(|| {
                        crate::path::column_number(key.name)
                            .filter(|column| (table.first_column..=last_column).contains(column))
                    })
                })
                .collect();
            if let Some(columns) = columns {
                hits.push((table, columns));
            }
        }
        match hits.as_slice() {
            [] => Err(Error::new(
                Code::NotFound,
                "no table contains all requested query columns",
            )
            .with_suggestion("select a table with the requested columns")),
            [(table, columns)] => Ok(HeaderBinding {
                sheet: Position::new(table.sheet),
                rows: (
                    table.header_row + 1,
                    table.last_row - u32::from(table.totals),
                ),
                columns: columns.clone(),
            }),
            several => Err(Error::new(
                Code::InvalidSelector,
                format!(
                    "query columns are ambiguous across {} tables",
                    several.len()
                ),
            )
            .with_suggestion("scope the query to one worksheet, e.g. Sheet1!row[...]")),
        }
    }

    fn names_column(&self, scope: &Node<'d, K>, key: &str) -> Result<bool> {
        let sheets = self.sheets_in(scope);
        Ok(self
            .doc
            .tables
            .iter()
            .filter(|table| sheets.contains(&table.sheet))
            .any(|table| {
                table
                    .columns
                    .iter()
                    .any(|column| column.eq_ignore_ascii_case(key))
            }))
    }

    fn row_cell<'n>(
        &'n self,
        row: &'n Node<'d, K>,
        column: u32,
        raw: bool,
    ) -> Result<Option<Value<'n>>> {
        let (sheet, number) = match *row {
            Node::Item {
                item,
                sheet: Some(sheet),
            } => match item.number {
                Some(number) => (sheet, number),
                None => return Ok(None),
            },
            Node::Row { sheet, number } => (sheet, number as usize),
            Node::Item { sheet: None, .. } => return Ok(None),
        };
        let found = u32::try_from(number)
            .ok()
            .and_then(|number| sheet.cell_at(number.checked_sub(1)?, column.checked_sub(1)?));
        Ok(found.and_then(|cell| match (&cell.raw, raw) {
            (Some(value), true) => Some(value.borrowed()),
            _ => cell.text.as_deref().map(Value::from),
        }))
    }
}

/// Writes an A1 reference without allocating.
fn write_cell(cell: CellRef, out: &mut fmt::Formatter<'_>) -> fmt::Result {
    let mut letters = [0_u8; 3];
    let mut start = letters.len();
    let mut column = cell.column() + 1;
    while column > 0 && start > 0 {
        start -= 1;
        // `(column - 1) % 26` is below 26.
        letters[start] = b'A' + ((column - 1) % 26) as u8;
        column = (column - 1) / 26;
    }
    for &letter in &letters[start..] {
        out.write_char(char::from(letter))?;
    }
    write!(out, "{}", cell.row() + 1)
}
