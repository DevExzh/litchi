//! The source-backed RTF vocabulary exposed by the unified adapter: the main
//! story's paragraphs, runs, hyperlinks, pictures and tables, the
//! header/footer, note and comment stories, the stylesheet and the list
//! tables.

use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// A RTF node kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The document root (`/`).
    Document,
    /// The main story (`/body`).
    Body,
    /// An ordinary main-story paragraph (`/body/p[N]`) or a header/footer
    /// paragraph (`/header[N]/p[M]`).
    Paragraph,
    /// A contiguous character-format run (`/body/p[N]/r[N]`).
    Run,
    /// A top-level body table (`/body/tbl[N]`).
    Table,
    /// A table row (`/body/tbl[N]/tr[M]`).
    Row,
    /// A table cell (`/body/tbl[N]/tr[M]/tc[K]`).
    Cell,
    /// A header story (`/header[N]`).
    Header,
    /// A footer story (`/footer[N]`).
    Footer,
    /// A footnote (`/footnote[N]`, the Nth footnote in source order).
    Footnote,
    /// An endnote (`/endnote[N]`, the Nth endnote in source order).
    Endnote,
    /// The comment collection (`/comments`).
    Comments,
    /// A comment (`/comments/comment[N]`).
    Comment,
    /// A main-story `HYPERLINK` field (`/body/p[N]/hyperlink[M]`).
    Hyperlink,
    /// A main-story picture (`/body/p[N]/picture[M]`).
    Picture,
    /// The stylesheet (`/styles`).
    Styles,
    /// A stylesheet entry (`/styles/<name>`).
    Style,
    /// The list tables (`/numbering`).
    Numbering,
    /// A list definition (`/numbering/abstractNum[N]`).
    AbstractNum,
    /// A list override (`/numbering/num[N]`).
    Num,
    /// A list level (`/numbering/abstractNum[N]/level[L]`).
    Level,
    /// A section (`/section[N]`, record 0328).
    Section,
    /// A main-story bookmark (`/bookmark[N]`, `/bookmark[@name=X]`).
    Bookmark,
    /// A main-story field (`/field[N]`).
    Field,
    /// A main-story `TOC` field (`/toc[N]`).
    Toc,
    /// A main-story paragraph's tab stop (`/body/p[N]/tab[M]`).
    Tab,
    /// A main-story text box (`/body/textbox[N]`).
    TextBox,
    /// A main-story drawing shape (`/body/shape[N]`).
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
            Self::Hyperlink => "hyperlink",
            Self::Picture => "picture",
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Numbering => "numbering",
            Self::AbstractNum => "abstractNum",
            Self::Num => "num",
            Self::Level => "level",
            Self::Section => "section",
            Self::Bookmark => "bookmark",
            Self::Field => "field",
            Self::Toc => "toc",
            Self::Tab => "tab",
            Self::TextBox => "textbox",
            Self::Shape => "shape",
        }
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::RTF;
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
            "hyperlink" => Some(Self::Hyperlink),
            "picture" => Some(Self::Picture),
            "styles" => Some(Self::Styles),
            "style" => Some(Self::Style),
            "numbering" => Some(Self::Numbering),
            "abstractNum" => Some(Self::AbstractNum),
            "num" => Some(Self::Num),
            "level" => Some(Self::Level),
            "section" => Some(Self::Section),
            "bookmark" => Some(Self::Bookmark),
            "field" => Some(Self::Field),
            "toc" => Some(Self::Toc),
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
            Self::Hyperlink => "hyperlink",
            Self::Picture => "picture",
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Numbering => "numbering",
            Self::AbstractNum => "abstractNum",
            Self::Num => "num",
            Self::Level => "level",
            Self::Section => "section",
            Self::Bookmark => "bookmark",
            Self::Field => "field",
            Self::Toc => "toc",
            Self::Tab => "tab",
            Self::TextBox => "textbox",
            Self::Shape => "shape",
        }
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

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
        Kind::Section,
        Kind::Bookmark,
        Kind::Field,
        Kind::Toc,
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

pub(crate) static ALIGNMENT: Tokens =
    Tokens::new("alignment", &["left", "right", "center", "justify"], &[]);

/// The complete local underline vocabulary retained by the native RTF
/// formatter.  Every native variant has its own token so a read never
/// collapses a dotted, dashed, or wave underline into the generic `single`
/// value used by some other document families.
pub(crate) static RUN_UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "none",
        "single",
        "double",
        "dotted",
        "dashed",
        "dashDot",
        "dashDotDot",
        "words",
        "thick",
        "wave",
        "hairline",
        "thickDotted",
        "thickDashed",
        "thickDashDot",
        "thickDashDotDot",
        "thickLongDash",
        "longDash",
        "heavyWave",
        "doubleWave",
    ],
    &[],
);

/// Nonnegative paragraph spacing, read as Word twips and printed in points.
pub(crate) const SPACING: LengthSpec =
    LengthSpec::new(LengthContext::WordSpacing, LengthStyle::Points);

/// Signed paragraph offsets, read with the ordinary signed length coercion
/// and printed in points.
pub(crate) const SIGNED_POINTS: LengthSpec =
    LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    PropSpec {
        name: "text",
        ty: ValueType::Text,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "align",
        ty: ValueType::Token(&ALIGNMENT),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "spaceBefore",
        ty: ValueType::Length(SPACING),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "spaceAfter",
        ty: ValueType::Length(SPACING),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "indent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "rightIndent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "firstLineIndent",
        ty: ValueType::Length(SIGNED_POINTS),
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "keepNext",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "keepLines",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "pageBreakBefore",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: false,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    prop("style", ValueType::Text, Access::READ),
    prop("numId", ValueType::Int, Access::READ),
    prop("numLevel", ValueType::Int, Access::READ),
];

/// Local character properties for a native RTF run.  The scalar types follow
/// the common word-processing run registry, while the native RTF projection
/// keeps its own exact underline token vocabulary.  The ten format facets are
/// writable through the native omission-preserving run patch; font and
/// foreground colour remain read-only until their resource-preserving editors
/// exist.
pub(crate) const RUN_PROPS: &[PropSpec] = &[
    PropSpec {
        name: "text",
        ty: ValueType::Text,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "font",
        ty: ValueType::Text,
        access: Access::READ,
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "size",
        ty: ValueType::Length(LengthSpec::FONT_SIZE),
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "bold",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "italic",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "color",
        ty: ValueType::Color(ColorContext::Word),
        access: Access::READ,
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "underline",
        ty: ValueType::Token(&RUN_UNDERLINE),
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "strike",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "dstrike",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "vanish",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "caps",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "smallcaps",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    PropSpec {
        name: "outline",
        ty: ValueType::Bool,
        access: Access {
            get: true,
            set: true,
            add: true,
        },
        raw: false,
        admission: crate::registry::Admission::NONE,
    },
    prop("rStyle", ValueType::Text, Access::READ),
];

pub(crate) const fn alignment_name(alignment: litchi_rtf::Alignment) -> &'static str {
    match alignment {
        litchi_rtf::Alignment::Left => "left",
        litchi_rtf::Alignment::Right => "right",
        litchi_rtf::Alignment::Center => "center",
        litchi_rtf::Alignment::Justify => "justify",
    }
}

pub(crate) const fn underline_name(underline: litchi_rtf::text::Underline) -> &'static str {
    match underline {
        litchi_rtf::text::Underline::None => "none",
        litchi_rtf::text::Underline::Single => "single",
        litchi_rtf::text::Underline::Double => "double",
        litchi_rtf::text::Underline::Dotted => "dotted",
        litchi_rtf::text::Underline::Dashed => "dashed",
        litchi_rtf::text::Underline::DashDot => "dashDot",
        litchi_rtf::text::Underline::DashDotDot => "dashDotDot",
        litchi_rtf::text::Underline::Words => "words",
        litchi_rtf::text::Underline::Thick => "thick",
        litchi_rtf::text::Underline::Wave => "wave",
        litchi_rtf::text::Underline::Hairline => "hairline",
        litchi_rtf::text::Underline::ThickDotted => "thickDotted",
        litchi_rtf::text::Underline::ThickDashed => "thickDashed",
        litchi_rtf::text::Underline::ThickDashDot => "thickDashDot",
        litchi_rtf::text::Underline::ThickDashDotDot => "thickDashDotDot",
        litchi_rtf::text::Underline::ThickLongDash => "thickLongDash",
        litchi_rtf::text::Underline::LongDash => "longDash",
        litchi_rtf::text::Underline::HeavyWave => "heavyWave",
        litchi_rtf::text::Underline::DoubleWave => "doubleWave",
    }
}

const fn prop(name: &'static str, ty: ValueType, access: Access) -> PropSpec {
    PropSpec {
        name,
        ty,
        access,
        raw: false,
        admission: crate::registry::Admission::NONE,
    }
}

const READ_SET: Access = Access {
    get: true,
    set: true,
    add: false,
};

/// Header/footer kinds: `default` for every page, `first`, `even` (left) and
/// `odd` (right) pages.
pub(crate) static STORY_TYPE: Tokens =
    Tokens::new("type", &["default", "first", "even", "odd"], &[]);

/// Native table-cell vertical alignment.
pub(crate) static VALIGN: Tokens = Tokens::new("valign", &["top", "center", "bottom"], &[]);

/// Native cell merge roles.
pub(crate) static MERGE: Tokens = Tokens::new("merge", &["restart", "continue"], &[]);

/// Native picture formats.
pub(crate) static PICTURE_TYPE: Tokens = Tokens::new(
    "type",
    &[
        "png", "jpeg", "emf", "wmf", "pict", "bmp", "dib", "gif", "tiff", "unknown",
    ],
    &[],
);

/// Stylesheet entry kinds.
pub(crate) static STYLE_TYPE: Tokens =
    Tokens::new("type", &["paragraph", "character", "section", "table"], &[]);

/// List-level number formats, under the OOXML `ST_NumberFormat` spelling.
pub(crate) static NUMBER_FORMAT: Tokens = Tokens::new(
    "numFmt",
    &[
        "decimal",
        "upperRoman",
        "lowerRoman",
        "upperLetter",
        "lowerLetter",
        "ordinal",
        "cardinalText",
        "ordinalText",
        "bullet",
        "none",
        "other",
    ],
    &[],
);

/// Table properties, in node property-table order.
pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    prop("rows", ValueType::Int, Access::READ),
    prop("cols", ValueType::Int, Access::READ),
];

/// Row properties, in node property-table order.
pub(crate) const ROW_PROPS: &[PropSpec] = &[
    prop("header", ValueType::Bool, Access::READ),
    prop("cantSplit", ValueType::Bool, Access::READ),
    prop("height", ValueType::Length(SPACING), Access::READ),
];

/// Cell properties, in node property-table order. Only the text is writable:
/// the native destination editor replaces the cell's text story exactly.
pub(crate) const CELL_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET),
    prop("width", ValueType::Length(SIGNED_POINTS), Access::READ),
    prop("valign", ValueType::Token(&VALIGN), Access::READ),
    prop("vmerge", ValueType::Token(&MERGE), Access::READ),
    prop("hmerge", ValueType::Token(&MERGE), Access::READ),
];

/// Header and footer story properties.
pub(crate) const STORY_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET),
    prop("type", ValueType::Token(&STORY_TYPE), Access::READ),
];

/// Footnote and endnote properties.
pub(crate) const NOTE_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET),
    prop("number", ValueType::Int, Access::READ),
];

/// Comment properties.
pub(crate) const COMMENT_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, READ_SET),
    prop("author", ValueType::Text, Access::READ),
    prop("initials", ValueType::Text, Access::READ),
    prop("date", ValueType::Text, Access::READ),
];

/// Main-story `HYPERLINK` field properties; the cached result is the text.
pub(crate) const HYPERLINK_PROPS: &[PropSpec] = &[
    prop("text", ValueType::Text, Access::READ),
    prop("url", ValueType::Text, Access::READ),
    prop("anchor", ValueType::Text, Access::READ),
    prop("tooltip", ValueType::Text, Access::READ),
    prop("tgtFrame", ValueType::Text, Access::READ),
];

/// Picture properties: the displayed extent and the native media format.
pub(crate) const PICTURE_PROPS: &[PropSpec] = &[
    prop("width", ValueType::Length(LengthSpec::EXTENT), Access::READ),
    prop(
        "height",
        ValueType::Length(LengthSpec::EXTENT),
        Access::READ,
    ),
    prop("type", ValueType::Token(&PICTURE_TYPE), Access::READ),
];

/// Stylesheet entry properties; `basedOn` and `next` name other entries.
pub(crate) const STYLE_PROPS: &[PropSpec] = &[
    prop("name", ValueType::Text, Access::READ),
    prop("type", ValueType::Token(&STYLE_TYPE), Access::READ),
    prop("id", ValueType::Int, Access::READ),
    prop("basedOn", ValueType::Text, Access::READ),
    prop("next", ValueType::Text, Access::READ),
];

/// List-definition (`\list`) properties.
pub(crate) const ABSTRACT_NUM_PROPS: &[PropSpec] = &[
    prop("id", ValueType::Int, Access::READ),
    prop("name", ValueType::Text, Access::READ),
    prop("levels", ValueType::Int, Access::READ),
];

/// List-override (`\listoverride`) properties.
pub(crate) const NUM_PROPS: &[PropSpec] = &[
    prop("id", ValueType::Int, Access::READ),
    prop("abstractNumId", ValueType::Int, Access::READ),
];

/// List-level properties.
pub(crate) const LEVEL_PROPS: &[PropSpec] = &[
    prop("ilvl", ValueType::Int, Access::READ),
    prop("numFmt", ValueType::Token(&NUMBER_FORMAT), Access::READ),
    prop("lvlText", ValueType::Text, Access::READ),
    prop("start", ValueType::Int, Access::READ),
];

const READ_SET_ADD: Access = Access {
    get: true,
    set: true,
    add: true,
};

/// A section's break kind (`\sbkN`), under the `ST_SectionMark` spelling.
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

/// Page orientation (`\lndscpsxn`).
pub(crate) static ORIENTATION: Tokens = Tokens::new("orientation", &["portrait", "landscape"], &[]);

/// A section's vertical alignment (`\vertal*`).
pub(crate) static SECTION_VALIGN: Tokens = Tokens::new(
    "vertical alignment",
    &["top", "center", "both", "bottom"],
    &[],
);

/// Section direction.
pub(crate) static DIRECTION: Tokens = Tokens::new("direction", &["ltr", "rtl"], &[]);

/// Page extents and margins, printed as OfficeCLI prints extents.
pub(crate) const EXTENT: LengthSpec = LengthSpec::EXTENT;

/// A section's page setup (record 0328), under the DOCX section keys DOC
/// uses (record 0324).
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
    prop("columns.separator", ValueType::Bool, Access::READ),
    prop("titlePage", ValueType::Bool, READ_SET),
    prop("pageNumFmt", ValueType::Text, Access::READ),
    prop("pageStart", ValueType::Int, Access::READ),
    prop("direction", ValueType::Token(&DIRECTION), Access::READ),
    prop("vAlign", ValueType::Token(&SECTION_VALIGN), READ_SET),
];

/// A bookmark: its name.
pub(crate) const BOOKMARK_PROPS: &[PropSpec] = &[prop("name", ValueType::Text, READ_SET)];

/// A field: its type, instruction, lock and whether it has a result.
pub(crate) const FIELD_PROPS: &[PropSpec] = &[
    prop("fieldType", ValueType::Text, Access::READ),
    prop("instruction", ValueType::Text, READ_SET),
    prop("fldLock", ValueType::Bool, Access::READ),
    prop("evaluated", ValueType::Bool, Access::READ),
];

/// A table of contents, from its `TOC` field's switches; each is written
/// back as the field's instruction.
pub(crate) const TOC_PROPS: &[PropSpec] = &[
    prop("levels", ValueType::Text, READ_SET),
    prop("hyperlinks", ValueType::Bool, READ_SET),
    prop("pageNumbers", ValueType::Bool, READ_SET),
    prop("instruction", ValueType::Text, READ_SET),
];

/// A tab stop's alignment, under the `ST_TabJc` spelling.
pub(crate) static TAB_VAL: Tokens = Tokens::new(
    "tab alignment",
    &["left", "center", "right", "decimal", "bar"],
    &[],
);

/// A tab stop's leader, under the `ST_TabTlc` spelling (`\tleq` reads as
/// `equal`).
pub(crate) static TAB_LEADER: Tokens = Tokens::new(
    "tab leader",
    &[
        "none",
        "dot",
        "hyphen",
        "underscore",
        "heavy",
        "middleDot",
        "equal",
    ],
    &[],
);

/// A paragraph tab stop: position in twips, alignment and leader.
pub(crate) const TAB_PROPS: &[PropSpec] = &[
    prop("pos", ValueType::Int, READ_SET_ADD),
    prop("val", ValueType::Token(&TAB_VAL), READ_SET_ADD),
    prop("leader", ValueType::Token(&TAB_LEADER), READ_SET_ADD),
];

/// A text box: its extent; its story text is the text.
pub(crate) const TEXT_BOX_PROPS: &[PropSpec] = &[
    prop("width", ValueType::Length(EXTENT), Access::READ),
    prop("height", ValueType::Length(EXTENT), Access::READ),
];

/// Text wrapping around a shape, under the OfficeCLI shape `wrap` spelling.
pub(crate) static SHAPE_WRAP: Tokens = Tokens::new(
    "shape wrap",
    &[
        "none",
        "square",
        "tight",
        "through",
        "topAndBottom",
        "behind",
        "front",
    ],
    &[],
);

/// A drawing shape: extent, position, wrapping and identity.
pub(crate) const SHAPE_PROPS: &[PropSpec] = &[
    prop("width", ValueType::Length(EXTENT), Access::READ),
    prop("height", ValueType::Length(EXTENT), Access::READ),
    prop("anchor.x", ValueType::Length(SIGNED_POINTS), Access::READ),
    prop("anchor.y", ValueType::Length(SIGNED_POINTS), Access::READ),
    prop("wrap", ValueType::Token(&SHAPE_WRAP), Access::READ),
    prop("behindDoc", ValueType::Bool, Access::READ),
    prop("geometry", ValueType::Text, Access::READ),
    prop("name", ValueType::Text, Access::READ),
];

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[Kind::Body, Kind::Header, Kind::Footer],
    children: &[Kind::Run, Kind::Hyperlink, Kind::Picture, Kind::Tab],
    identities: &[],
    paths: &["/body/p[N]", "/header[N]/p[M]", "/footer[N]/p[M]"],
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

const fn element(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    children: &'static [Kind],
    paths: &'static [&'static str],
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities: &[],
        paths,
        props,
        gap: None,
    }
}

const fn identities(
    mut spec: ElementSpec<Kind>,
    identities: &'static [&'static str],
) -> ElementSpec<Kind> {
    spec.identities = identities;
    spec
}

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = {
    use Kind as K;
    FormatSpec {
        root: K::Document,
        elements: &[
            DOCUMENT,
            BODY,
            PARAGRAPH,
            RUN,
            element(
                K::Table,
                "table",
                &["table", "tbl"],
                &[K::Body],
                &[K::Row],
                &["/body/tbl[N]"],
                TABLE_PROPS,
            ),
            element(
                K::Row,
                "row",
                &["row", "tr"],
                &[K::Table],
                &[K::Cell],
                &["/body/tbl[N]/tr[M]"],
                ROW_PROPS,
            ),
            element(
                K::Cell,
                "cell",
                &["cell", "tc"],
                &[K::Row],
                &[],
                &["/body/tbl[N]/tr[M]/tc[K]"],
                CELL_PROPS,
            ),
            identities(
                element(
                    K::Header,
                    "header",
                    &["header"],
                    &[K::Document],
                    // Story paragraphs are addressed by path and listed by
                    // views; unscoped `p` queries stay in the main story.
                    &[],
                    &["/header[N]", "/header[@type=T]"],
                    STORY_PROPS,
                ),
                &["type"],
            ),
            identities(
                element(
                    K::Footer,
                    "footer",
                    &["footer"],
                    &[K::Document],
                    // Story paragraphs are addressed by path and listed by
                    // views; unscoped `p` queries stay in the main story.
                    &[],
                    &["/footer[N]", "/footer[@type=T]"],
                    STORY_PROPS,
                ),
                &["type"],
            ),
            identities(
                element(
                    K::Footnote,
                    "footnote",
                    &["footnote"],
                    &[K::Document],
                    &[],
                    &["/footnote[N]"],
                    NOTE_PROPS,
                ),
                &["footnoteId"],
            ),
            identities(
                element(
                    K::Endnote,
                    "endnote",
                    &["endnote"],
                    &[K::Document],
                    &[],
                    &["/endnote[N]"],
                    NOTE_PROPS,
                ),
                &["endnoteId"],
            ),
            element(
                K::Comments,
                "comments",
                &["comments"],
                &[K::Document],
                &[K::Comment],
                &["/comments"],
                &[],
            ),
            identities(
                element(
                    K::Comment,
                    "comment",
                    &["comment"],
                    &[K::Comments],
                    &[],
                    &["/comments/comment[N]"],
                    COMMENT_PROPS,
                ),
                &["commentId"],
            ),
            element(
                K::Hyperlink,
                "hyperlink",
                &["hyperlink"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/hyperlink[M]"],
                HYPERLINK_PROPS,
            ),
            element(
                K::Picture,
                "picture",
                &["image", "img", "picture"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/picture[M]"],
                PICTURE_PROPS,
            ),
            element(
                K::Styles,
                "styles",
                &["styles"],
                &[K::Document],
                &[K::Style],
                &["/styles"],
                &[],
            ),
            element(
                K::Style,
                "style",
                &["style"],
                &[K::Styles],
                &[],
                &["/styles/<name>"],
                STYLE_PROPS,
            ),
            element(
                K::Numbering,
                "numbering",
                &["numbering"],
                &[K::Document],
                &[K::AbstractNum, K::Num],
                &["/numbering"],
                &[],
            ),
            identities(
                element(
                    K::AbstractNum,
                    "abstractNum",
                    &["abstractnum"],
                    &[K::Numbering],
                    &[K::Level],
                    &["/numbering/abstractNum[N]"],
                    ABSTRACT_NUM_PROPS,
                ),
                &["id"],
            ),
            identities(
                element(
                    K::Num,
                    "num",
                    &["num"],
                    &[K::Numbering],
                    &[],
                    &["/numbering/num[N]"],
                    NUM_PROPS,
                ),
                &["id"],
            ),
            identities(
                element(
                    K::Level,
                    "level",
                    &["level", "lvl"],
                    &[K::AbstractNum],
                    &[],
                    &["/numbering/abstractNum[N]/level[L]"],
                    LEVEL_PROPS,
                ),
                &["ilvl"],
            ),
            element(
                K::Section,
                "section",
                &["section"],
                &[K::Document],
                &[],
                &["/section[N]"],
                SECTION_PROPS,
            ),
            identities(
                element(
                    K::Bookmark,
                    "bookmark",
                    &["bookmark"],
                    &[K::Document],
                    &[],
                    &["/bookmark[N]", "/bookmark[@name=X]"],
                    BOOKMARK_PROPS,
                ),
                &["name"],
            ),
            element(
                K::Field,
                "field",
                &["field"],
                &[K::Document],
                &[],
                &["/field[N]"],
                FIELD_PROPS,
            ),
            element(
                K::Toc,
                "toc",
                &["tableofcontents", "toc"],
                &[K::Document],
                &[],
                &["/toc[N]"],
                TOC_PROPS,
            ),
            element(
                K::Tab,
                "tab",
                &["tab"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/tab[M]"],
                TAB_PROPS,
            ),
            element(
                K::TextBox,
                "textbox",
                &["textbox"],
                &[K::Body],
                &[],
                &["/body/textbox[N]"],
                TEXT_BOX_PROPS,
            ),
            element(
                K::Shape,
                "shape",
                &["shape"],
                &[K::Body],
                &[],
                &["/body/shape[N]"],
                SHAPE_PROPS,
            ),
        ],
        gaps: &[
            GapSpec {
                subject: Subject::Element("formfield"),
                code: crate::Code::UnsupportedCapability,
                follow_up: "record 0328 follow-up: RTF form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN) as formfield nodes and their writes",
                reason: "RTF form fields are read by the native model but not yet projected into generic paths and queries",
            },
            read_gap("ole"),
            GapSpec {
                subject: Subject::Element("equation"),
                code: crate::Code::UnsupportedCapability,
                follow_up: "record 0328 follow-up: RTF EQ fields and math zones as equation nodes",
                reason: "RTF equation fields and math zones are read by the native model but not yet projected",
            },
            GapSpec {
                subject: Subject::Element("revision"),
                code: crate::Code::UnsupportedCapability,
                follow_up: "record 0328 follow-up: RTF revision nodes and accept/reject (P2)",
                reason: "RTF revision marks are read by the native model but not yet projected",
            },
            GapSpec {
                subject: Subject::Element("sdt"),
                code: crate::Code::UnsupportedPath,
                follow_up: "none: content controls are DOCX-only",
                reason: "Rich Text Format has no structured document tags",
            },
            GapSpec {
                subject: Subject::Element("watermark"),
                code: crate::Code::UnsupportedCapability,
                follow_up: "record 0328 follow-up: the RTF header WordArt watermark",
                reason: "a watermark is a header shape the story reads do not project",
            },
            // OfficeCLI elements the RTF adapter has no kind for (record 0339).
            word_family_gap(
                "chart",
                "record 0328 (deferred-P1): RTF charts as chart nodes",
                "embedded charts are kept as objects; they are not projected as chart nodes",
            ),
            word_family_gap(
                "series",
                "record 0328 (deferred-P1): RTF charts as chart nodes",
                "chart series follow the chart gap",
            ),
            word_family_gap(
                "axis",
                "record 0328 (deferred-P1): RTF charts as chart nodes",
                "chart axes follow the chart gap",
            ),
            word_family_gap(
                "bookmarkEnd",
                "record 0328 (deferred-P1)",
                "a bookmark is one node over its start and end; a standalone end marker is not addressed on its own",
            ),
            word_family_gap(
                "permStart",
                "record 0328 (deferred-P1)",
                "RTF editing-permission ranges are not projected as nodes",
            ),
            word_family_gap(
                "ptab",
                "record 0328 (deferred-P1)",
                "positional tabs are not projected as nodes",
            ),
            word_family_gap(
                "col",
                "record 0328 (deferred-P1)",
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
    }
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);

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

const fn read_gap(element: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up: "unified RTF retained-node projection",
        reason: "this native RTF concept is not yet integrated into generic paths and queries",
    }
}

/// Elements the unified RTF writer sets: main-story paragraphs and runs, and
/// the text of retained destination stories.
pub(crate) const SETTABLE: &[&str] = &[
    "p", "r", "tc", "header", "footer", "footnote", "endnote", "comment", "section", "bookmark",
    "field", "toc", "tab",
];
/// Elements the unified RTF writer creates: paragraphs (under the body)
/// and runs. The body itself is a parent, never added (record 0289).
pub(crate) const ADDABLE: &[&str] = &["p", "r", "tab"];
/// Elements the unified RTF writer removes.
pub(crate) const REMOVABLE: &[&str] = &["p", "r", "picture", "bookmark", "tab"];

const fn operation_gap(
    verb: &'static str,
    element: &'static str,
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element,
            context,
        },
        code: crate::Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

const RICH_RUN: &str = "M6 RTF rich run preservation";
const RICH_PARAGRAPH: &str = "M6 RTF rich paragraph preservation";
const TABLE_EDITOR: &str = "record 0275 follow-up: native RTF table structure editor";
const STORY_EDITOR: &str = "record 0275 follow-up: native RTF story lifecycle editor";
const FIELD_EDITOR: &str = "record 0275 follow-up: native RTF field instruction editor";
const PICTURE_EDITOR: &str = "record 0275 follow-up: native RTF picture insertion";
const STYLE_EDITOR: &str = "record 0275 follow-up: native RTF style application";
const LIST_EDITOR: &str = "record 0275 follow-up: native RTF list editor";

/// RTF write refusals that the format-level verb support would otherwise
/// hide: every element the reader projects but the native editor cannot yet
/// change, and every source closure the native editors refuse.
pub(crate) static WRITE_GAPS: [GapSpec; 33] = [
    operation_gap(
        "set",
        "run",
        "native-source-closure",
        RICH_RUN,
        "changed run text/formatting requires an ordinary uncompressed body without dependent destinations, review content, inline line breaks or paragraph style/list references; run text containing paragraph/line-break controls, native mixed-operation composition and durable replay remain unsupported; opaque body syntax is preserved when outside changed source spans with projectable UTF-8 anchors",
    ),
    operation_gap(
        "set",
        "paragraph",
        "native-source-closure",
        RICH_PARAGRAPH,
        "paragraph text, alignment and layout updates require the native editor to prove the source formatting and destination closure",
    ),
    operation_gap(
        "remove",
        "paragraph",
        "native-source-closure",
        RICH_PARAGRAPH,
        "paragraph removal requires the native editor to prove an ordinary literal body and retain every unselected source boundary",
    ),
    operation_gap(
        "add",
        "paragraph",
        "native-source-closure",
        RICH_PARAGRAPH,
        "paragraph insertion requires the native editor to prove the ordinary literal body and preserve every unselected source boundary",
    ),
    operation_gap(
        "swap",
        "body",
        "native-source-closure",
        RICH_PARAGRAPH,
        "paragraph swapping requires the native editor to prove an ordinary plain body and preserve paragraph order and boundaries",
    ),
    operation_gap(
        "remove",
        "run",
        "native-source-closure",
        RICH_RUN,
        "run removal preserves containing paragraphs through native empty text replacement; dependent destinations, compressed transport, inline line breaks, paragraph style/list references and opaque syntax overlapping changed source spans remain unsupported; opaque boundary anchors are projected through text edits",
    ),
    operation_gap(
        "move",
        "run",
        "native-run-edit",
        "M6 RTF atomic native run edits",
        "run reordering requires an atomic native character-run editor that preserves paragraph boundaries and formatting controls",
    ),
    operation_gap(
        "swap",
        "run",
        "native-run-edit",
        "M6 RTF atomic native run edits",
        "run swapping requires an atomic native character-run editor that preserves paragraph boundaries and formatting controls",
    ),
    operation_gap(
        "set",
        "paragraph",
        "style-application",
        STYLE_EDITOR,
        "paragraph `style`, `numId` and `numLevel` are read-only: applying a style or list requires emitting the referenced definition's properties with the handle",
    ),
    operation_gap(
        "set",
        "run",
        "style-application",
        STYLE_EDITOR,
        "run `rStyle` is read-only: applying a character style requires emitting its formatting with the `\\cs` handle",
    ),
    operation_gap(
        "get",
        "paragraph",
        "table-cell-paragraphs",
        TABLE_EDITOR,
        "table cells expose their flat text story; per-paragraph cell nodes (`tc/p`) are not projected",
    ),
    operation_gap(
        "get",
        "run",
        "header-footer-runs",
        RICH_RUN,
        "header/footer paragraphs expose text only; their character runs are not projected",
    ),
    operation_gap(
        "set",
        "cell",
        "native-source-closure",
        TABLE_EDITOR,
        "cell text is spliced into the retained source; cells whose text range contains paragraph resets, destinations or unbalanced groups, nested-table cells, compressed or non-ASCII transports are refused, and cell properties other than `text` are read-only",
    ),
    operation_gap(
        "add",
        "table",
        "native-table-structure",
        TABLE_EDITOR,
        "the native RTF editor cannot insert tables",
    ),
    operation_gap(
        "remove",
        "table",
        "native-table-structure",
        TABLE_EDITOR,
        "the native RTF editor cannot remove tables",
    ),
    operation_gap(
        "set",
        "table",
        "native-table-structure",
        TABLE_EDITOR,
        "table and row properties are read-only",
    ),
    operation_gap(
        "add",
        "row",
        "native-table-structure",
        TABLE_EDITOR,
        "the native RTF editor cannot insert or remove rows or cells",
    ),
    operation_gap(
        "set",
        "header",
        "native-source-closure",
        STORY_EDITOR,
        "header/footer text is spliced per paragraph; a multi-paragraph story is set through `/header[N]/p[M]`, and stories with drawings, fields or page-number events are refused",
    ),
    operation_gap(
        "add",
        "header",
        "native-story-lifecycle",
        STORY_EDITOR,
        "the native RTF editor cannot create or remove header/footer stories",
    ),
    operation_gap(
        "set",
        "footnote",
        "native-source-closure",
        STORY_EDITOR,
        "note text is spliced into the retained source; notes with drawings, fields or page breaks are refused",
    ),
    operation_gap(
        "add",
        "footnote",
        "native-story-lifecycle",
        STORY_EDITOR,
        "the native RTF editor cannot insert or remove footnotes or endnotes",
    ),
    operation_gap(
        "add",
        "comment",
        "native-story-lifecycle",
        STORY_EDITOR,
        "the native RTF editor cannot insert or remove comments",
    ),
    operation_gap(
        "set",
        "hyperlink",
        "native-field-editor",
        FIELD_EDITOR,
        "hyperlink targets and results are read-only field metadata",
    ),
    operation_gap(
        "add",
        "hyperlink",
        "native-field-editor",
        FIELD_EDITOR,
        "the native RTF editor cannot insert or remove hyperlink fields",
    ),
    operation_gap(
        "add",
        "picture",
        "native-picture-insertion",
        PICTURE_EDITOR,
        "the native RTF editor cannot insert `\\pict` groups; existing payloads can be replaced natively at equal size only",
    ),
    operation_gap(
        "set",
        "picture",
        "native-picture-properties",
        PICTURE_EDITOR,
        "picture extent and format are read-only",
    ),
    operation_gap(
        "remove",
        "picture",
        "native-source-closure",
        PICTURE_EDITOR,
        "picture removal requires every picture to be a standalone hexadecimal PNG/JPEG body group (no fields, shapes, objects, wrappers or binary payloads) and at most 64 pictures per transaction",
    ),
    operation_gap(
        "set",
        "style",
        "native-stylesheet-editor",
        STYLE_EDITOR,
        "stylesheet entries are read-only",
    ),
    operation_gap(
        "add",
        "style",
        "native-stylesheet-editor",
        STYLE_EDITOR,
        "the native RTF editor cannot add or remove stylesheet entries",
    ),
    operation_gap(
        "set",
        "abstractNum",
        "native-list-editor",
        LIST_EDITOR,
        "list definitions, overrides and levels are read-only",
    ),
    operation_gap(
        "add",
        "abstractNum",
        "native-list-editor",
        LIST_EDITOR,
        "the native RTF editor cannot add or remove lists",
    ),
    operation_gap(
        "get",
        "body",
        "table-order",
        TABLE_EDITOR,
        "the native model keeps tables apart from the paragraph flow, so `/body` lists paragraphs before tables rather than in source order",
    ),
    operation_gap(
        "set",
        "cell",
        "nested-table-cells",
        TABLE_EDITOR,
        "nested-table cells are neither projected nor editable",
    ),
];
