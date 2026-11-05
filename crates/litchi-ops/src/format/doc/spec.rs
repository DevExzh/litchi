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
    /// The stylesheet (`/styles`).
    Styles,
    /// A style (`/styles/<id or name>`).
    Style,
    /// The list tables (`/numbering`).
    Numbering,
    /// A list definition, an `LSTF` (`/numbering/abstractNum[N]`).
    AbstractNum,
    /// A list instance, an `LFO` (`/numbering/num[N]`).
    Num,
    /// A list level, an `LVL` (`/numbering/abstractNum[N]/level[L]`).
    Level,
    /// A main-story `HYPERLINK` field (`/body/p[N]/hyperlink[M]`).
    Hyperlink,
    /// A main-story bookmark (`/bookmark[N]`, `/bookmark[@name=X]`).
    Bookmark,
    /// A main-story field (`/field[N]`).
    Field,
    /// A main-story `TOC` field (`/toc[N]`).
    Toc,
    /// A section (`/section[N]`).
    Section,
    /// A paragraph tab stop (`/body/p[N]/tab[M]`).
    Tab,
    /// A main-story text box (`/body/textbox[N]`).
    TextBox,
    /// A main-story floating shape anchor (`/body/shape[N]`).
    Shape,
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
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Numbering => "numbering",
            Self::AbstractNum => "abstractNum",
            Self::Num => "num",
            Self::Level => "level",
            Self::Hyperlink => "hyperlink",
            Self::Bookmark => "bookmark",
            Self::Field => "field",
            Self::Toc => "toc",
            Self::Section => "section",
            Self::Tab => "tab",
            Self::TextBox => "textbox",
            Self::Shape => "shape",
        }
    }

    /// Whether the kind is one of the main-story structure kinds the generic
    /// views list. The parts and marks (styles, numbering, hyperlinks,
    /// bookmarks, fields, sections, tab stops, text boxes and shapes) are
    /// reached by path and query only, so every view stays as it was.
    pub(crate) const fn is_structural(self) -> bool {
        (self as u8) <= (Self::Comment as u8)
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
            "styles" => Some(Self::Styles),
            "style" => Some(Self::Style),
            "numbering" => Some(Self::Numbering),
            "abstractNum" | "abstractnum" => Some(Self::AbstractNum),
            "num" => Some(Self::Num),
            "level" | "lvl" => Some(Self::Level),
            "hyperlink" => Some(Self::Hyperlink),
            "bookmark" => Some(Self::Bookmark),
            "field" => Some(Self::Field),
            "toc" | "tableofcontents" => Some(Self::Toc),
            "section" => Some(Self::Section),
            "tab" => Some(Self::Tab),
            "textbox" => Some(Self::TextBox),
            "shape" => Some(Self::Shape),
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
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Numbering => "numbering",
            Self::AbstractNum => "abstractNum",
            Self::Num => "num",
            Self::Level => "level",
            Self::Hyperlink => "hyperlink",
            Self::Bookmark => "bookmark",
            Self::Field => "field",
            Self::Toc => "toc",
            Self::Section => "section",
            Self::Tab => "tab",
            Self::TextBox => "textbox",
            Self::Shape => "shape",
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

const READ_ADD: Access = Access {
    get: true,
    set: false,
    add: true,
};

const ADD_ONLY: Access = Access {
    get: false,
    set: false,
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
    // Record 0324: the paragraph style (`istd`) and its list binding.
    prop("style", ValueType::Text, READ_SET),
    prop("styleId", ValueType::Text, READ_SET),
    prop("styleName", ValueType::Text, READ_SET),
    prop("numFmt", ValueType::Text, Access::READ),
    prop("listStyle", ValueType::Token(&LIST_STYLE), Access::READ),
    prop("start", ValueType::Int, Access::READ),
    prop("numLevel", ValueType::Int, Access::READ),
];

/// A paragraph's list kind, as DOCX reports it.
pub(crate) static LIST_STYLE: Tokens = Tokens::new("list style", &["bullet", "ordered"], &[]);

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
    // Record 0324: the run's character style (`sprmCIstd`).
    prop("rStyle", ValueType::Text, READ_SET),
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

/// A style's `type` (`StdfBase.stk`), under the `ST_StyleType` spelling.
pub(crate) static STYLE_TYPE: Tokens = Tokens::new(
    "style type",
    &["paragraph", "character", "table", "numbering"],
    &[("char", 1), ("para", 0)],
);

/// Highlight colors (`sprmCHighlight`), under the `ST_HighlightColor`
/// spelling.
pub(crate) static HIGHLIGHT: Tokens = Tokens::new(
    "highlight",
    &[
        "none",
        "black",
        "blue",
        "cyan",
        "green",
        "magenta",
        "red",
        "yellow",
        "white",
        "darkBlue",
        "darkCyan",
        "darkGreen",
        "darkMagenta",
        "darkRed",
        "darkYellow",
        "darkGray",
        "lightGray",
    ],
    &[],
);

/// A list definition's kind (`LSTF.fSimpleList`, `LSTF.fHybrid`).
pub(crate) static MULTI_LEVEL_TYPE: Tokens = Tokens::new(
    "multi-level type",
    &["hybridMultilevel", "multilevel", "singleLevel"],
    &[("hybrid", 0), ("multi", 1), ("single", 2)],
);

/// A level's number alignment (`LVLF.jc`).
pub(crate) static LEVEL_JUSTIFICATION: Tokens =
    Tokens::new("level justification", &["left", "center", "right"], &[]);

/// The character after a level's number (`LVLF.ixchFollow`).
pub(crate) static LEVEL_SUFFIX: Tokens =
    Tokens::new("level suffix", &["tab", "space", "nothing"], &[("none", 2)]);

/// A tab stop's alignment (`TBD.jc`), under the `ST_TabJc` spelling.
pub(crate) static TAB_VAL: Tokens = Tokens::new(
    "tab alignment",
    &["left", "center", "right", "decimal", "bar", "num"],
    &[],
);

/// A tab stop's leader (`TBD.tlc`), under the `ST_TabTlc` spelling.
pub(crate) static TAB_LEADER: Tokens = Tokens::new(
    "tab leader",
    &["none", "dot", "hyphen", "underscore", "heavy", "middleDot"],
    &[],
);

/// A section's break kind (`sprmSBkc`), under the `ST_SectionMark` spelling.
pub(crate) static SECTION_TYPE: Tokens = Tokens::new(
    "section type",
    &[
        "continuous",
        "nextColumn",
        "nextPage",
        "evenPage",
        "oddPage",
    ],
    &[],
);

/// Page orientation (`sprmSBOrientation`).
pub(crate) static ORIENTATION: Tokens = Tokens::new("orientation", &["portrait", "landscape"], &[]);

/// Line numbering restart (`sprmSLnc`), OfficeCLI's `lineNumbers` values.
pub(crate) static LINE_NUMBERS: Tokens = Tokens::new(
    "line numbering",
    &["restartPage", "restartSection", "continuous"],
    &[],
);

/// A section's vertical alignment (`sprmSVjc`).
pub(crate) static SECTION_VALIGN: Tokens = Tokens::new(
    "vertical alignment",
    &["top", "center", "both", "bottom"],
    &[],
);

/// Text wrapping around a floating shape (`SPA.wr`).
pub(crate) static SHAPE_WRAP: Tokens = Tokens::new(
    "shape wrap",
    &["topAndBottom", "none", "square", "tight", "through"],
    &[],
);

/// A floating shape's horizontal origin (`SPA.bx`).
pub(crate) static SHAPE_H_RELATIVE: Tokens =
    Tokens::new("horizontal origin", &["margin", "page", "column"], &[]);

/// A floating shape's vertical origin (`SPA.by`).
pub(crate) static SHAPE_V_RELATIVE: Tokens =
    Tokens::new("vertical origin", &["margin", "page", "paragraph"], &[]);

/// Page and paragraph extents, printed as OfficeCLI prints extents.
pub(crate) const EXTENT: LengthSpec = LengthSpec::EXTENT;

const fn raw(name: &'static str, ty: ValueType) -> PropSpec {
    PropSpec {
        name,
        ty,
        access: Access::READ,
        raw: true,
        admission: crate::registry::Admission::NONE,
    }
}

const fn read(name: &'static str, ty: ValueType) -> PropSpec {
    prop(name, ty, Access::READ)
}

/// `/styles`: the number of styles.
pub(crate) const STYLES_PROPS: &[PropSpec] = &[read("count", ValueType::Int)];

/// A style, in the DOCX style order (record 0319): identity and UI flags,
/// then the style's own character properties, then its own paragraph
/// properties. Values are the style's direct `UPX` properties, never the
/// inherited chain.
pub(crate) const STYLE_PROPS: &[PropSpec] = &[
    read("id", ValueType::Text),
    read("name", ValueType::Text),
    read("type", ValueType::Token(&STYLE_TYPE)),
    read("default", ValueType::Bool),
    read("basedOn", ValueType::Text),
    read("next", ValueType::Text),
    read("linked", ValueType::Text),
    read("autoRedefine", ValueType::Bool),
    read("hidden", ValueType::Bool),
    read("qFormat", ValueType::Bool),
    read("uiPriority", ValueType::Int),
    read("semiHidden", ValueType::Bool),
    read("unhideWhenUsed", ValueType::Bool),
    read("locked", ValueType::Bool),
    read("customStyle", ValueType::Bool),
    read("font", ValueType::Text),
    read("font.ea", ValueType::Text),
    read("font.cs", ValueType::Text),
    read("size", ValueType::Length(LengthSpec::FONT_SIZE)),
    read("size.cs", ValueType::Length(LengthSpec::FONT_SIZE)),
    read("bold", ValueType::Bool),
    read("bold.cs", ValueType::Bool),
    read("italic", ValueType::Bool),
    read("italic.cs", ValueType::Bool),
    read("color", ValueType::Color(ColorContext::Word)),
    read("underline", ValueType::Token(&UNDERLINE)),
    read("strike", ValueType::Bool),
    read("dstrike", ValueType::Bool),
    read("caps", ValueType::Bool),
    read("smallCaps", ValueType::Bool),
    read("vanish", ValueType::Bool),
    read("highlight", ValueType::Token(&HIGHLIGHT)),
    read("superscript", ValueType::Bool),
    read("subscript", ValueType::Bool),
    read("charSpacing", ValueType::Length(SIGNED_POINTS)),
    read("align", ValueType::Token(&ALIGNMENT)),
    read("spaceBefore", ValueType::Length(SPACING)),
    read("spaceAfter", ValueType::Length(SPACING)),
    read("firstLineIndent", ValueType::Length(SIGNED_POINTS)),
    read("leftIndent", ValueType::Length(SIGNED_POINTS)),
    read("rightIndent", ValueType::Length(SIGNED_POINTS)),
    read("outlineLvl", ValueType::Int),
    read("keepNext", ValueType::Bool),
    read("keepLines", ValueType::Bool),
    read("pageBreakBefore", ValueType::Bool),
    read("widowControl", ValueType::Bool),
    read("contextualSpacing", ValueType::Bool),
];

/// `/numbering`: the numbers of list definitions and instances.
pub(crate) const NUMBERING_PROPS: &[PropSpec] = &[
    read("abstractNumCount", ValueType::Int),
    read("numCount", ValueType::Int),
];

/// A list definition (`LSTF`). Its `id` is the native index, a raw
/// identity (record 0319 §3).
pub(crate) const ABSTRACT_NUM_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    read("type", ValueType::Token(&MULTI_LEVEL_TYPE)),
    read("name", ValueType::Text),
];

/// A list instance (`LFO`): the definition it uses and its start overrides.
pub(crate) const NUM_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    read("abstractNumId", ValueType::Int),
    read("startOverride.0", ValueType::Int),
    read("startOverride.1", ValueType::Int),
    read("startOverride.2", ValueType::Int),
    read("startOverride.3", ValueType::Int),
    read("startOverride.4", ValueType::Int),
    read("startOverride.5", ValueType::Int),
    read("startOverride.6", ValueType::Int),
    read("startOverride.7", ValueType::Int),
    read("startOverride.8", ValueType::Int),
];

/// A list level (`LVL`), in the DOCX level order.
pub(crate) const LEVEL_PROPS: &[PropSpec] = &[
    raw("ilvl", ValueType::Int),
    read("start", ValueType::Int),
    read("format", ValueType::Text),
    read("lvlText", ValueType::Text),
    read("justification", ValueType::Token(&LEVEL_JUSTIFICATION)),
    read("suff", ValueType::Token(&LEVEL_SUFFIX)),
    read("lvlRestart", ValueType::Int),
    read("isLgl", ValueType::Bool),
    read("indent", ValueType::Int),
    read("hanging", ValueType::Int),
    read("font", ValueType::Text),
    read("size", ValueType::Length(LengthSpec::FONT_SIZE)),
    read("color", ValueType::Color(ColorContext::Word)),
    read("bold", ValueType::Bool),
    read("italic", ValueType::Bool),
];

/// A `HYPERLINK` field: its instruction switches, then the formatting of
/// its first result run, as DOCX reads a hyperlink's first run.
pub(crate) const HYPERLINK_PROPS: &[PropSpec] = &[
    prop("url", ValueType::Text, READ_SET_ADD),
    prop("anchor", ValueType::Text, READ_ADD),
    prop("tooltip", ValueType::Text, READ_ADD),
    read("tgtFrame", ValueType::Text),
    read("font", ValueType::Text),
    read("size", ValueType::Length(LengthSpec::FONT_SIZE)),
    read("bold", ValueType::Bool),
    read("italic", ValueType::Bool),
    read("color", ValueType::Color(ColorContext::Word)),
    read("underline", ValueType::Token(&UNDERLINE)),
    read("strike", ValueType::Bool),
    read("rStyle", ValueType::Text),
    prop("text", ValueType::Text, READ_SET_ADD),
];

/// A bookmark: its name; `text` seeds a new bookmark's covered text.
pub(crate) const BOOKMARK_PROPS: &[PropSpec] = &[
    prop("name", ValueType::Text, READ_SET_ADD),
    prop("text", ValueType::Text, ADD_ONLY),
];

/// A field: its type, instruction, lock and whether it has a result.
pub(crate) const FIELD_PROPS: &[PropSpec] = &[
    read("fieldType", ValueType::Text),
    read("instruction", ValueType::Text),
    read("fldLock", ValueType::Bool),
    read("evaluated", ValueType::Bool),
];

/// A table of contents, from its `TOC` field's switches.
pub(crate) const TOC_PROPS: &[PropSpec] = &[
    read("levels", ValueType::Text),
    read("hyperlinks", ValueType::Bool),
    read("pageNumbers", ValueType::Bool),
];

/// A section's page setup, under OfficeCLI's section keys.
pub(crate) const SECTION_PROPS: &[PropSpec] = &[
    prop("type", ValueType::Token(&SECTION_TYPE), READ_SET),
    prop("pageWidth", ValueType::Length(EXTENT), READ_SET),
    prop("pageHeight", ValueType::Length(EXTENT), READ_SET),
    prop("orientation", ValueType::Token(&ORIENTATION), READ_SET),
    prop("marginTop", ValueType::Length(SIGNED_POINTS), READ_SET),
    prop("marginBottom", ValueType::Length(SIGNED_POINTS), READ_SET),
    prop("marginLeft", ValueType::Length(EXTENT), READ_SET),
    prop("marginRight", ValueType::Length(EXTENT), READ_SET),
    prop("marginHeader", ValueType::Length(EXTENT), READ_SET),
    prop("marginFooter", ValueType::Length(EXTENT), READ_SET),
    prop("marginGutter", ValueType::Length(EXTENT), READ_SET),
    prop("columns", ValueType::Int, READ_SET),
    prop("columnSpace", ValueType::Length(EXTENT), READ_SET),
    read("columns.separator", ValueType::Bool),
    prop("titlePage", ValueType::Bool, READ_SET),
    read("pageNumFmt", ValueType::Text),
    read("pageStart", ValueType::Int),
    read("lineNumbers", ValueType::Token(&LINE_NUMBERS)),
    read("lineNumberCountBy", ValueType::Int),
    read("lineNumberDistance", ValueType::Int),
    read("direction", ValueType::Token(&DIRECTION)),
    read("rtlGutter", ValueType::Bool),
    prop("vAlign", ValueType::Token(&SECTION_VALIGN), READ_SET),
];

/// Paragraph or section direction.
pub(crate) static DIRECTION: Tokens = Tokens::new("direction", &["ltr", "rtl"], &[]);

/// A paragraph tab stop: position in twips, alignment and leader.
pub(crate) const TAB_PROPS: &[PropSpec] = &[
    prop("pos", ValueType::Int, READ_SET_ADD),
    prop("val", ValueType::Token(&TAB_VAL), READ_SET_ADD),
    prop("leader", ValueType::Token(&TAB_LEADER), READ_SET_ADD),
];

/// A main-story text box: its anchor's extent; its story text is the text.
pub(crate) const TEXT_BOX_PROPS: &[PropSpec] = &[
    read("width", ValueType::Length(EXTENT)),
    read("height", ValueType::Length(EXTENT)),
];

/// A floating shape's anchor (`FSPA`): extent, wrapping and origins.
pub(crate) const SHAPE_PROPS: &[PropSpec] = &[
    read("width", ValueType::Length(EXTENT)),
    read("height", ValueType::Length(EXTENT)),
    read("anchor.x", ValueType::Length(SIGNED_POINTS)),
    read("anchor.y", ValueType::Length(SIGNED_POINTS)),
    read("wrap", ValueType::Token(&SHAPE_WRAP)),
    read("hRelative", ValueType::Token(&SHAPE_H_RELATIVE)),
    read("vRelative", ValueType::Token(&SHAPE_V_RELATIVE)),
    read("behindDoc", ValueType::Bool),
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
        Kind::Styles,
        Kind::Numbering,
        Kind::Bookmark,
        Kind::Field,
        Kind::Toc,
        Kind::Section,
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
    children: &[Kind::Paragraph, Kind::Table, Kind::TextBox, Kind::Shape],
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
    children: &[Kind::Run, Kind::Hyperlink, Kind::Tab],
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

const fn element(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    children: &'static [Kind],
    identities: &'static [&'static str],
    paths: &'static [&'static str],
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities,
        paths,
        props,
        gap: None,
    }
}

const STYLES: ElementSpec<Kind> = element(
    Kind::Styles,
    "styles",
    &["styles"],
    &[Kind::Document],
    &[Kind::Style],
    &[],
    &["/styles"],
    STYLES_PROPS,
);

const STYLE: ElementSpec<Kind> = element(
    Kind::Style,
    "style",
    &["style"],
    &[Kind::Styles],
    &[],
    &[],
    &["/styles/<id or name>"],
    STYLE_PROPS,
);

const NUMBERING: ElementSpec<Kind> = element(
    Kind::Numbering,
    "numbering",
    &["numbering"],
    &[Kind::Document],
    &[Kind::AbstractNum, Kind::Num],
    &[],
    &["/numbering"],
    NUMBERING_PROPS,
);

const ABSTRACT_NUM: ElementSpec<Kind> = element(
    Kind::AbstractNum,
    "abstractNum",
    &["abstractnum"],
    &[Kind::Numbering],
    &[Kind::Level],
    &["id"],
    &["/numbering/abstractNum[N]"],
    ABSTRACT_NUM_PROPS,
);

const NUM: ElementSpec<Kind> = element(
    Kind::Num,
    "num",
    &["num"],
    &[Kind::Numbering],
    &[],
    &["id"],
    &["/numbering/num[N]"],
    NUM_PROPS,
);

const LEVEL: ElementSpec<Kind> = element(
    Kind::Level,
    "level",
    &["level", "lvl"],
    &[Kind::AbstractNum],
    &[],
    &["ilvl"],
    &["/numbering/abstractNum[N]/level[L]"],
    LEVEL_PROPS,
);

const HYPERLINK: ElementSpec<Kind> = element(
    Kind::Hyperlink,
    "hyperlink",
    &["hyperlink"],
    &[Kind::Paragraph],
    &[],
    &[],
    &["/body/p[N]/hyperlink[M]"],
    HYPERLINK_PROPS,
);

const BOOKMARK: ElementSpec<Kind> = element(
    Kind::Bookmark,
    "bookmark",
    &["bookmark"],
    &[Kind::Document],
    &[],
    &["name"],
    &["/bookmark[N]", "/bookmark[@name=X]"],
    BOOKMARK_PROPS,
);

const FIELD: ElementSpec<Kind> = element(
    Kind::Field,
    "field",
    &["field"],
    &[Kind::Document],
    &[],
    &[],
    &["/field[N]"],
    FIELD_PROPS,
);

const TOC: ElementSpec<Kind> = element(
    Kind::Toc,
    "toc",
    &["tableofcontents", "toc"],
    &[Kind::Document],
    &[],
    &[],
    &["/toc[N]"],
    TOC_PROPS,
);

const SECTION: ElementSpec<Kind> = element(
    Kind::Section,
    "section",
    &["section"],
    &[Kind::Document],
    &[],
    &[],
    &["/section[N]"],
    SECTION_PROPS,
);

const TAB: ElementSpec<Kind> = element(
    Kind::Tab,
    "tab",
    &["tab"],
    &[Kind::Paragraph],
    &[],
    &[],
    &["/body/p[N]/tab[M]"],
    TAB_PROPS,
);

const TEXT_BOX: ElementSpec<Kind> = element(
    Kind::TextBox,
    "textbox",
    &["textbox"],
    &[Kind::Body],
    &[],
    &[],
    &["/body/textbox[N]"],
    TEXT_BOX_PROPS,
);

const SHAPE: ElementSpec<Kind> = element(
    Kind::Shape,
    "shape",
    &["shape"],
    &[Kind::Body],
    &[],
    &[],
    &["/body/shape[N]"],
    SHAPE_PROPS,
);

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[
        DOCUMENT,
        BODY,
        PARAGRAPH,
        RUN,
        TABLE,
        ROW,
        CELL,
        HEADER,
        FOOTER,
        FOOTNOTE,
        ENDNOTE,
        COMMENTS,
        COMMENT,
        STYLES,
        STYLE,
        NUMBERING,
        ABSTRACT_NUM,
        NUM,
        LEVEL,
        HYPERLINK,
        BOOKMARK,
        FIELD,
        TOC,
        SECTION,
        TAB,
        TEXT_BOX,
        SHAPE,
    ],
    // Concepts the native `litchi-doc` reader models but the unified adapter
    // does not yet project, under their OfficeCLI DOCX type names.
    gaps: &[
        element_gap(
            "equation",
            "record 0324 follow-up 4: DOC EQ fields and MTEF equation objects as `equation` nodes",
        ),
        element_gap(
            "formfield",
            "record 0324 follow-up 3: DOC FORMTEXT, FORMCHECKBOX and FORMDROPDOWN nodes and writes",
        ),
        element_gap(
            "picture",
            "record 0324 follow-up 5: DOC inline and floating pictures as `picture` nodes",
        ),
        element_gap(
            "ole",
            "record 0324 follow-up 5: DOC embedded objects as `ole` nodes",
        ),
        element_gap(
            "revision",
            "record 0324 follow-up 2: DOC revision nodes and accept/reject (P2)",
        ),
        element_gap(
            "watermark",
            "record 0324 follow-up 6: the DOC header WordArt watermark",
        ),
        GapSpec {
            subject: Subject::Element("sdt"),
            code: crate::Code::UnsupportedPath,
            follow_up: "none: content controls are DOCX-only",
            reason: "Word 97-2003 binary documents have no structured document tags",
        },
        // OfficeCLI elements the DOC adapter has no kind for (record 0339).
        word_family_gap(
            "chart",
            "record 0324 (deferred-P1): DOC charts as chart nodes",
            "embedded charts are kept as objects; they are not projected as chart nodes",
        ),
        word_family_gap(
            "series",
            "record 0324 (deferred-P1): DOC charts as chart nodes",
            "chart series follow the chart gap",
        ),
        word_family_gap(
            "axis",
            "record 0324 (deferred-P1): DOC charts as chart nodes",
            "chart axes follow the chart gap",
        ),
        word_family_gap(
            "bookmarkEnd",
            "record 0324 (deferred-P1)",
            "a bookmark is one node over its start and end; a standalone end marker is not addressed on its own",
        ),
        word_family_gap(
            "permStart",
            "record 0324 (deferred-P1)",
            "DOC editing-permission ranges are not projected as nodes",
        ),
        word_family_gap(
            "ptab",
            "record 0324 (deferred-P1)",
            "positional tabs are not projected as nodes",
        ),
        word_family_gap(
            "col",
            "record 0324 (deferred-P1)",
            "table columns are not nodes; cell widths are read on each cell",
        ),
        word_family_gap(
            "fieldChar",
            "none: out of scope (record 0332)",
            "a complex field is one field node; its field-character and instruction parts are not separate elements",
        ),
        word_family_gap(
            "instrText",
            "none: out of scope (record 0332)",
            "a complex field is one field node; its field-character and instruction parts are not separate elements",
        ),
        GapSpec {
            subject: Subject::Element("diagram"),
            code: crate::Code::UnsupportedPath,
            follow_up: "none: out of scope (record 0332, design §10)",
            reason: "OfficeCLI's diagram renders Mermaid text into shapes or a picture; rendering is outside the unified API",
        },
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);

/// Elements the unified DOC writer sets (record 0324 adds hyperlinks,
/// bookmarks, tab stops and sections).
pub(crate) const SETTABLE: &[&str] = &["p", "r", "hyperlink", "bookmark", "tab", "section"];
/// Elements the unified DOC writer creates.
pub(crate) const ADDABLE: &[&str] = &["p", "hyperlink", "bookmark", "tab"];
/// Elements the unified DOC writer removes.
pub(crate) const REMOVABLE: &[&str] = &["p", "r", "hyperlink", "bookmark", "tab"];

/// A gap for an OfficeCLI Word element the table has no kind for (record
/// 0339): naming it is `unsupported_capability`.
const fn word_family_gap(
    element: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

const fn element_gap(element: &'static str, follow_up: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up,
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
