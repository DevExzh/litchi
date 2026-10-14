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

const PARAGRAPH: ElementSpec<Kind> = ElementSpec {
    kind: Kind::Paragraph,
    name: "paragraph",
    aliases: &["p", "paragraph"],
    parents: &[Kind::Body, Kind::Header, Kind::Footer],
    children: &[Kind::Run, Kind::Hyperlink, Kind::Picture],
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
        ],
        gaps: &[
            read_gap("shape"),
            read_gap("textbox"),
            read_gap("field"),
            read_gap("formfield"),
            read_gap("bookmark"),
            read_gap("ole"),
            read_gap("equation"),
            read_gap("revision"),
            read_gap("section"),
        ],
    }
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);

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
    "p", "r", "tc", "header", "footer", "footnote", "endnote", "comment",
];
/// Elements the unified RTF writer creates: paragraphs (under the body)
/// and runs. The body itself is a parent, never added (record 0289).
pub(crate) const ADDABLE: &[&str] = &["p", "r"];
/// Elements the unified RTF writer removes.
pub(crate) const REMOVABLE: &[&str] = &["p", "r", "picture"];

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
