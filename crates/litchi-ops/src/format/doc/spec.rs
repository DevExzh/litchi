//! The source-backed DOC vocabulary exposed by the unified adapter: the main
//! story's paragraphs, runs and top-level tables, the header and footer
//! stories, notes and comments.

use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// A DOC node kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The document root (`/`).
    Document,
    /// The main story (`/body`).
    Body,
    /// An ordinary paragraph (`/body/p[N]`, or a paragraph of a cell, a
    /// header, a footer, a note or a comment).
    Paragraph,
    /// A character run of a paragraph (`p[N]/r[M]`).
    Run,
    /// A top-level main-story table (`/body/tbl[N]`).
    Table,
    /// A table row (`tr[N]`).
    Row,
    /// A table cell (`tc[N]`).
    Cell,
    /// A non-empty header story (`/header[N]`).
    Header,
    /// A non-empty footer story (`/footer[N]`).
    Footer,
    /// A footnote (`/footnote[N]`, the Nth in reference order).
    Footnote,
    /// An endnote (`/endnote[N]`, the Nth in reference order).
    Endnote,
    /// The comment collection (`/comments`).
    Comments,
    /// A comment (`/comments/comment[N]`).
    Comment,
}

impl Kind {
    /// The OfficeCLI type name.
    #[must_use]
    pub const fn type_name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "paragraph",
            Self::Run => "run",
            Self::Table => "table",
            Self::Row => "row",
            Self::Cell => "cell",
            Self::Header => "header",
            Self::Footer => "footer",
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comments => "comments",
            Self::Comment => "comment",
        }
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::DOC;
    const ROOT: Self = Self::Document;

    fn from_name(name: &str) -> Option<Self> {
        match name {
            "document" => Some(Self::Document),
            "body" => Some(Self::Body),
            "p" | "paragraph" => Some(Self::Paragraph),
            "r" | "run" => Some(Self::Run),
            "tbl" | "table" => Some(Self::Table),
            "tr" | "row" => Some(Self::Row),
            "tc" | "cell" => Some(Self::Cell),
            "header" => Some(Self::Header),
            "footer" => Some(Self::Footer),
            "footnote" => Some(Self::Footnote),
            "endnote" => Some(Self::Endnote),
            "comments" => Some(Self::Comments),
            "comment" => Some(Self::Comment),
            _ => None,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "p",
            Self::Run => "r",
            Self::Table => "tbl",
            Self::Row => "tr",
            Self::Cell => "tc",
            Self::Header => "header",
            Self::Footer => "footer",
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comments => "comments",
            Self::Comment => "comment",
        }
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

/// Paragraph alignment as the DOC paragraph properties store it.
pub(crate) static ALIGNMENT: Tokens = Tokens::new(
    "alignment",
    &[
        "left",
        "center",
        "right",
        "justify",
        "distribute",
        "mediumKashida",
        "indented",
        "highKashida",
        "lowKashida",
        "thaiDistribute",
    ],
    &[],
);

/// Every DOC underline style the native reader distinguishes, under the
/// OOXML `ST_Underline` spelling. Writes accept `none` and `single`.
pub(crate) static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "none",
        "single",
        "double",
        "dotted",
        "dash",
        "wave",
        "thick",
        "words",
        "dotDash",
        "dotDotDash",
    ],
    &[],
);

/// Header and footer story kinds.
pub(crate) static STORY_TYPE: Tokens = Tokens::new("type", &["default", "even", "first"], &[]);

/// Nonnegative paragraph spacing, stored as twips and printed in points.
pub(crate) const SPACING: LengthSpec =
    LengthSpec::new(LengthContext::WordSpacing, LengthStyle::Points);

/// Signed paragraph indents, printed in points.
pub(crate) const SIGNED_POINTS: LengthSpec =
    LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

const READ_SET: Access = Access {
    get: true,
    set: true,
    add: false,
};

const SET_ONLY: Access = Access {
    get: false,
    set: true,
    add: false,
};

const READ_SET_ADD: Access = Access {
    get: true,
    set: true,
    add: true,
};

const fn prop(name: &'static str, ty: ValueType, access: Access) -> PropSpec {
    PropSpec {
        name,
        ty,
        access,
        raw: false,
        admission: crate::registry::Admission::NONE,
    }
}

/// Paragraph properties, in [`super::node::NodeProps`] index order. The run
/// formatting keys are write-only here: a paragraph-level assignment formats
/// every character of the paragraph.
pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET_ADD),
    prop("align", ValueType::Token(&ALIGNMENT), Access::READ),
    prop("spaceBefore", ValueType::Length(SPACING), Access::READ),
    prop("spaceAfter", ValueType::Length(SPACING), Access::READ),
    prop("indent", ValueType::Length(SIGNED_POINTS), Access::READ),
    prop(
        "rightIndent",
        ValueType::Length(SIGNED_POINTS),
        Access::READ,
    ),
    prop(
        "firstLineIndent",
        ValueType::Length(SIGNED_POINTS),
        Access::READ,
    ),
    prop("keepNext", ValueType::Bool, Access::READ),
    prop("keepLines", ValueType::Bool, Access::READ),
    prop("pageBreakBefore", ValueType::Bool, Access::READ),
    prop("bold", ValueType::Bool, SET_ONLY),
    prop("italic", ValueType::Bool, SET_ONLY),
    prop("underline", ValueType::Token(&UNDERLINE), SET_ONLY),
    prop("size", ValueType::Length(LengthSpec::FONT_SIZE), SET_ONLY),
    prop("color", ValueType::Color(ColorContext::Word), SET_ONLY),
    prop("font", ValueType::Text, SET_ONLY),
];

/// Run properties, in [`super::node::NodeProps`] index order. Values are the
/// reader's resolved run properties (direct formatting over the style chain).
pub(crate) const RUN_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET),
    prop("font", ValueType::Text, READ_SET),
    prop("size", ValueType::Length(LengthSpec::FONT_SIZE), READ_SET),
    prop("bold", ValueType::Bool, READ_SET),
    prop("italic", ValueType::Bool, READ_SET),
    prop("color", ValueType::Color(ColorContext::Word), READ_SET),
    prop("underline", ValueType::Token(&UNDERLINE), READ_SET),
    prop("strike", ValueType::Bool, Access::READ),
    prop("caps", ValueType::Bool, Access::READ),
    prop("smallcaps", ValueType::Bool, Access::READ),
    prop("superscript", ValueType::Bool, Access::READ),
    prop("subscript", ValueType::Bool, Access::READ),
];

/// Table properties, in [`super::node::NodeProps`] index order.
pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    prop("rows", ValueType::Int, Access::READ),
    prop("cols", ValueType::Int, Access::READ),
];

/// Row properties, in [`super::node::NodeProps`] index order.
pub(crate) const ROW_PROPS: &[PropSpec] = &[prop("header", ValueType::Bool, Access::READ)];

/// Cell properties, in [`super::node::NodeProps`] index order.
pub(crate) const CELL_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, Access::READ),
    prop("fill", ValueType::Color(ColorContext::Word), Access::READ),
];

/// Header and footer properties, in [`super::node::NodeProps`] index order.
pub(crate) const STORY_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, Access::READ),
    prop("type", ValueType::Token(&STORY_TYPE), Access::READ),
];

/// Footnote and endnote properties, in [`super::node::NodeProps`] index order.
pub(crate) const NOTE_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, Access::READ),
    prop("number", ValueType::Int, Access::READ),
];

/// Comment properties, in [`super::node::NodeProps`] index order.
pub(crate) const COMMENT_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, Access::READ),
    prop("author", ValueType::Text, Access::READ),
    prop("initials", ValueType::Text, Access::READ),
];

const DOCUMENT: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Document,
    name: "document",
    aliases: &["document"],
    parents: &[],
    children: &[
        Kind::Body,
        Kind::Header,
        Kind::Footer,
        Kind::Footnote,
        Kind::Endnote,
        Kind::Comments,
    ],
    identities: &[],
    paths: &["/"],
    props: &[],
    gap: None,
};

const BODY: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Body,
    name: "body",
    aliases: &["body"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph, Kind::Table],
    identities: &[],
    paths: &["/body"],
    props: &[],
    gap: None,
};

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[
        Kind::Body,
        Kind::Cell,
        Kind::Header,
        Kind::Footer,
        Kind::Footnote,
        Kind::Endnote,
        Kind::Comment,
    ],
    children: &[Kind::Run],
    identities: &[],
    paths: &[
        "/body/p[N]",
        "/body/tbl[N]/tr[N]/tc[N]/p[N]",
        "/header[N]/p[N]",
        "/footer[N]/p[N]",
        "/footnote[N]/p[N]",
        "/endnote[N]/p[N]",
        "/comments/comment[N]/p[N]",
    ],
    props: PARAGRAPH_PROPS,
    gap: None,
};

const RUN: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Run,
    name: "run",
    aliases: &["r", "run"],
    parents: &[Kind::Paragraph],
    children: &[],
    identities: &[],
    paths: &["/body/p[N]/r[N]"],
    props: RUN_PROPS,
    gap: None,
};

const TABLE: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Table,
    name: "table",
    aliases: &["table", "tbl"],
    parents: &[Kind::Body],
    children: &[Kind::Row],
    identities: &[],
    paths: &["/body/tbl[N]"],
    props: TABLE_PROPS,
    gap: None,
};

const ROW: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Row,
    name: "row",
    aliases: &["row", "tr"],
    parents: &[Kind::Table],
    children: &[Kind::Cell],
    identities: &[],
    paths: &["/body/tbl[N]/tr[N]"],
    props: ROW_PROPS,
    gap: None,
};

const CELL: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Cell,
    name: "cell",
    aliases: &["cell", "tc"],
    parents: &[Kind::Row],
    children: &[Kind::Paragraph],
    identities: &[],
    paths: &["/body/tbl[N]/tr[N]/tc[N]"],
    props: CELL_PROPS,
    gap: None,
};

const HEADER: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Header,
    name: "header",
    aliases: &["header"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph],
    identities: &["type"],
    paths: &["/header[N]"],
    props: STORY_PROPS,
    gap: None,
};

const FOOTER: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Footer,
    name: "footer",
    aliases: &["footer"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph],
    identities: &["type"],
    paths: &["/footer[N]"],
    props: STORY_PROPS,
    gap: None,
};

const FOOTNOTE: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Footnote,
    name: "footnote",
    aliases: &["footnote"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph],
    identities: &["footnoteId"],
    paths: &["/footnote[N]"],
    props: NOTE_PROPS,
    gap: None,
};

const ENDNOTE: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Endnote,
    name: "endnote",
    aliases: &["endnote"],
    parents: &[Kind::Document],
    children: &[Kind::Paragraph],
    identities: &["endnoteId"],
    paths: &["/endnote[N]"],
    props: NOTE_PROPS,
    gap: None,
};

const COMMENTS: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Comments,
    name: "comments",
    aliases: &["comments"],
    parents: &[Kind::Document],
    children: &[Kind::Comment],
    identities: &[],
    paths: &["/comments"],
    props: &[],
    gap: None,
};

const COMMENT: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Comment,
    name: "comment",
    aliases: &["comment"],
    parents: &[Kind::Comments],
    children: &[Kind::Paragraph],
    identities: &["commentId"],
    paths: &["/comments/comment[N]"],
    props: COMMENT_PROPS,
    gap: None,
};

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[
        DOCUMENT, BODY, PARAGRAPH, RUN, TABLE, ROW, CELL, HEADER, FOOTER, FOOTNOTE, ENDNOTE,
        COMMENTS, COMMENT,
    ],
    // Concepts the native `litchi-doc` reader models but the unified adapter
    // does not yet project, under their OfficeCLI DOCX type names.
    gaps: &[
        element_gap("hyperlink"),
        element_gap("equation"),
        element_gap("section"),
        element_gap("style"),
        element_gap("numbering"),
        element_gap("bookmark"),
        element_gap("field"),
        element_gap("formfield"),
        element_gap("picture"),
        element_gap("ole"),
        element_gap("revision"),
        element_gap("textbox"),
        element_gap("shape"),
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);

const fn element_gap(element: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up: "record 0258 follow-up (DOC read+write loop beyond the P0 surface)",
        reason: "this native DOC concept is not yet projected into generic paths and queries",
    }
}

pub(crate) const fn alignment_name(
    justification: litchi_doc::parts::pap::Justification,
) -> &'static str {
    use litchi_doc::parts::pap::Justification as J;
    match justification {
        J::Left => "left",
        J::Center => "center",
        J::Right => "right",
        J::Justified => "justify",
        J::Distributed => "distribute",
        J::MediumKashida => "mediumKashida",
        J::Indented => "indented",
        J::HighKashida => "highKashida",
        J::LowKashida => "lowKashida",
        J::ThaiDistributed => "thaiDistribute",
    }
}

pub(crate) const fn underline_name(style: litchi_doc::parts::chp::UnderlineStyle) -> &'static str {
    use litchi_doc::parts::chp::UnderlineStyle as U;
    match style {
        U::None => "none",
        U::Single => "single",
        U::Double => "double",
        U::Dotted => "dotted",
        U::Dashed => "dash",
        U::Wavy => "wave",
        U::Thick => "thick",
        U::WordsOnly => "words",
        U::DashDot => "dotDash",
        U::DashDotDot => "dotDotDash",
    }
}
