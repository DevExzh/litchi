//! The ODT element vocabulary, property tables and capability gaps: the one
//! static table path resolution, query compilation, the node view and the
//! capability report read.

use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// An ODT node kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The document root (`/`).
    Document,
    /// The text body (`/body`, `office:text`).
    Body,
    /// A paragraph (`p`, `text:p`).
    Paragraph,
    /// A heading (`h`, `text:h`).
    Heading,
    /// A run: a `text:span`, or character data directly in a paragraph,
    /// heading or hyperlink (`r`).
    Run,
    /// A hyperlink (`text:a`).
    Hyperlink,
    /// A list (`text:list`).
    List,
    /// A list item or list header (`listItem`).
    ListItem,
    /// A table (`tbl`, `table:table`).
    Table,
    /// A table row (`tr`).
    Row,
    /// A table cell (`tc`), covered cells included.
    Cell,
    /// A section (`text:section`).
    Section,
    /// A frame (`draw:frame`).
    Frame,
    /// An image (`picture`, `draw:image`).
    Picture,
    /// A text box (`draw:text-box`).
    TextBox,
    /// A footnote (`text:note`).
    Footnote,
    /// An endnote (`text:note`).
    Endnote,
    /// A comment (`office:annotation`).
    Comment,
    /// A bookmark (`text:bookmark`, `text:bookmark-start`).
    Bookmark,
    /// The common styles (`/styles`).
    Styles,
    /// One common style (`/styles/<name>`).
    Style,
    /// The document properties (`/metadata`).
    Metadata,
}

impl Kind {
    /// The OfficeCLI type name.
    #[must_use]
    pub const fn type_name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "paragraph",
            Self::Heading => "heading",
            Self::Run => "run",
            Self::Hyperlink => "hyperlink",
            Self::List => "list",
            Self::ListItem => "listItem",
            Self::Table => "table",
            Self::Row => "row",
            Self::Cell => "cell",
            Self::Section => "section",
            Self::Frame => "frame",
            Self::Picture => "picture",
            Self::TextBox => "textbox",
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comment => "comment",
            Self::Bookmark => "bookmark",
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Metadata => "metadata",
        }
    }

    /// The adapter kind of a leaf body kind.
    pub(crate) const fn of(kind: litchi_odt::body::Kind) -> Option<Self> {
        use litchi_odt::body::Kind as L;
        Some(match kind {
            L::Body => Self::Body,
            L::Paragraph => Self::Paragraph,
            L::Heading => Self::Heading,
            L::Span | L::Text => Self::Run,
            L::Hyperlink => Self::Hyperlink,
            L::List => Self::List,
            L::ListItem => Self::ListItem,
            L::Table => Self::Table,
            L::Row => Self::Row,
            L::Cell | L::CoveredCell => Self::Cell,
            L::Section => Self::Section,
            L::Frame => Self::Frame,
            L::Image => Self::Picture,
            L::TextBox => Self::TextBox,
            L::Footnote => Self::Footnote,
            L::Endnote => Self::Endnote,
            L::Annotation => Self::Comment,
            L::Bookmark => Self::Bookmark,
            _ => return None,
        })
    }

    /// The number of kinds, for per-kind counters.
    pub(crate) const COUNT: usize = 22;
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::ODT;
    const ROOT: Self = Self::Document;

    fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "document" => Self::Document,
            "body" => Self::Body,
            "p" => Self::Paragraph,
            "h" => Self::Heading,
            "r" => Self::Run,
            "hyperlink" => Self::Hyperlink,
            "list" => Self::List,
            "listItem" => Self::ListItem,
            "tbl" => Self::Table,
            "tr" => Self::Row,
            "tc" => Self::Cell,
            "section" => Self::Section,
            "frame" => Self::Frame,
            "picture" => Self::Picture,
            "textbox" => Self::TextBox,
            "footnote" => Self::Footnote,
            "endnote" => Self::Endnote,
            "comment" => Self::Comment,
            "bookmark" => Self::Bookmark,
            "styles" => Self::Styles,
            "style" => Self::Style,
            "metadata" => Self::Metadata,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Body => "body",
            Self::Paragraph => "p",
            Self::Heading => "h",
            Self::Run => "r",
            Self::Hyperlink => "hyperlink",
            Self::List => "list",
            Self::ListItem => "listItem",
            Self::Table => "tbl",
            Self::Row => "tr",
            Self::Cell => "tc",
            Self::Section => "section",
            Self::Frame => "frame",
            Self::Picture => "picture",
            Self::TextBox => "textbox",
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comment => "comment",
            Self::Bookmark => "bookmark",
            Self::Styles => "styles",
            Self::Style => "style",
            Self::Metadata => "metadata",
        }
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

/// Paragraph alignment: `start`/`end` read as `left`/`right`.
pub(crate) static ALIGN: Tokens =
    Tokens::new("alignment", &["left", "center", "right", "justify"], &[]);

/// Frame anchoring (`text:anchor-type`).
pub(crate) static ANCHOR: Tokens = Tokens::new(
    "anchor",
    &["paragraph", "char", "as-char", "page", "frame"],
    &[],
);

/// Note classes.
pub(crate) static NOTE_CLASS: Tokens = Tokens::new("note class", &["footnote", "endnote"], &[]);

/// Paragraph spacing, printed in points.
pub(crate) const SPACING: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::Points);
/// Signed paragraph offsets, printed in points.
pub(crate) const SIGNED: LengthSpec = LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

const fn prop(name: &'static str, ty: ValueType, get: bool, set: bool, add: bool) -> PropSpec {
    PropSpec {
        name,
        ty,
        access: Access { get, set, add },
        raw: false,
        admission: crate::registry::Admission::NONE,
    }
}

const fn read(name: &'static str, ty: ValueType) -> PropSpec {
    prop(name, ty, true, false, false)
}

const fn write(name: &'static str, ty: ValueType) -> PropSpec {
    prop(name, ty, true, true, false)
}

const fn create(name: &'static str, ty: ValueType) -> PropSpec {
    prop(name, ty, true, true, true)
}

const COLOR: ValueType = ValueType::Color(ColorContext::Word);
const SIZE: ValueType = ValueType::Length(LengthSpec::FONT_SIZE);

/// Character properties shared by runs and (through their style) paragraphs.
macro_rules! character_props {
    () => {
        [
            write("bold", ValueType::Bool),
            write("italic", ValueType::Bool),
            write("underline", ValueType::Bool),
            write("strike", ValueType::Bool),
            write("size", SIZE),
            write("color", COLOR),
            write("font", ValueType::Text),
            write("highlight", COLOR),
        ]
    };
}

/// The indexes of [`PARAGRAPH_PROPS`]; the heading table extends it.
pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &{
    let character = character_props!();
    [
        create("text", ValueType::Text),
        create("style", ValueType::Text),
        create("align", ValueType::Token(&ALIGN)),
        write("spaceBefore", ValueType::Length(SPACING)),
        write("spaceAfter", ValueType::Length(SPACING)),
        write("indent", ValueType::Length(SIGNED)),
        write("rightIndent", ValueType::Length(SIGNED)),
        write("firstLineIndent", ValueType::Length(SIGNED)),
        write("lineSpacing", ValueType::Text),
        write("keepNext", ValueType::Bool),
        write("pageBreakBefore", ValueType::Bool),
        write("shading", COLOR),
        character[0],
        character[1],
        character[2],
        character[3],
        character[4],
        character[5],
        character[6],
        character[7],
    ]
};

/// The paragraph table plus the outline level.
pub(crate) const HEADING_PROPS: &[PropSpec] = &{
    let p = PARAGRAPH_PROPS;
    [
        p[0],
        p[1],
        p[2],
        p[3],
        p[4],
        p[5],
        p[6],
        p[7],
        p[8],
        p[9],
        p[10],
        p[11],
        p[12],
        p[13],
        p[14],
        p[15],
        p[16],
        p[17],
        p[18],
        p[19],
        create("level", ValueType::Int),
    ]
};

pub(crate) const RUN_PROPS: &[PropSpec] = &{
    let character = character_props!();
    [
        create("text", ValueType::Text),
        read("style", ValueType::Text),
        character[0],
        character[1],
        character[2],
        character[3],
        character[4],
        character[5],
        character[6],
        character[7],
        write("superscript", ValueType::Bool),
        write("subscript", ValueType::Bool),
        write("caps", ValueType::Bool),
        write("smallcaps", ValueType::Bool),
    ]
};

pub(crate) const HYPERLINK_PROPS: &[PropSpec] = &[
    create("text", ValueType::Text),
    create("url", ValueType::Text),
    write("tooltip", ValueType::Text),
    write("tgtFrame", ValueType::Text),
];

pub(crate) const LIST_PROPS: &[PropSpec] = &[read("style", ValueType::Text)];

pub(crate) const LIST_ITEM_PROPS: &[PropSpec] = &[create("text", ValueType::Text)];

pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    write("name", ValueType::Text),
    prop("rows", ValueType::Int, true, false, true),
    prop("cols", ValueType::Int, true, false, true),
    read("style", ValueType::Text),
];

pub(crate) const ROW_PROPS: &[PropSpec] = &[read("cells", ValueType::Int)];

pub(crate) const CELL_PROPS: &[PropSpec] = &[
    write("text", ValueType::Text),
    read("colspan", ValueType::Int),
    read("rowspan", ValueType::Int),
    read("covered", ValueType::Bool),
    read("valueType", ValueType::Text),
    read("shading", COLOR),
    read("valign", ValueType::Text),
];

pub(crate) const SECTION_PROPS: &[PropSpec] = &[
    create("name", ValueType::Text),
    write("protected", ValueType::Bool),
    write("hidden", ValueType::Bool),
];

pub(crate) const FRAME_PROPS: &[PropSpec] = &[
    write("name", ValueType::Text),
    write("width", ValueType::Length(LengthSpec::EXTENT)),
    write("height", ValueType::Length(LengthSpec::EXTENT)),
    write("x", ValueType::Length(LengthSpec::POSITION)),
    write("y", ValueType::Length(LengthSpec::POSITION)),
    write("anchor", ValueType::Token(&ANCHOR)),
];

pub(crate) const PICTURE_PROPS: &[PropSpec] = &[read("src", ValueType::Text)];

pub(crate) const TEXT_BOX_PROPS: &[PropSpec] = &[write("text", ValueType::Text)];

pub(crate) const NOTE_PROPS: &[PropSpec] = &[
    create("text", ValueType::Text),
    read("citation", ValueType::Text),
    read("class", ValueType::Token(&NOTE_CLASS)),
];

pub(crate) const COMMENT_PROPS: &[PropSpec] = &[
    create("text", ValueType::Text),
    prop("author", ValueType::Text, true, false, true),
    read("date", ValueType::Text),
];

pub(crate) const BOOKMARK_PROPS: &[PropSpec] = &[create("name", ValueType::Text)];

pub(crate) const STYLE_PROPS: &[PropSpec] = &[
    read("name", ValueType::Text),
    read("displayName", ValueType::Text),
    read("family", ValueType::Text),
    read("parent", ValueType::Text),
    read("align", ValueType::Token(&ALIGN)),
    read("bold", ValueType::Bool),
    read("italic", ValueType::Bool),
    read("size", SIZE),
    read("color", COLOR),
    read("font", ValueType::Text),
];

pub(crate) const METADATA_PROPS: &[PropSpec] = &[
    write("title", ValueType::Text),
    write("subject", ValueType::Text),
    write("description", ValueType::Text),
    write("author", ValueType::Text),
    write("lastModifiedBy", ValueType::Text),
    write("language", ValueType::Text),
    write("keywords", ValueType::Text),
    read("created", ValueType::Text),
    read("modified", ValueType::Text),
    read("application", ValueType::Text),
    read("revisionNumber", ValueType::Text),
];

use Kind as K;

/// Kinds that hold block content.
const BLOCK_PARENTS: &[Kind] = &[
    K::Body,
    K::Section,
    K::Cell,
    K::TextBox,
    K::Footnote,
    K::Endnote,
    K::Comment,
    K::ListItem,
];
/// Kinds that hold inline content.
const INLINE_PARENTS: &[Kind] = &[K::Paragraph, K::Heading, K::Hyperlink];
const FRAME_PARENTS: &[Kind] = &[
    K::Body,
    K::Section,
    K::Cell,
    K::TextBox,
    K::Footnote,
    K::Endnote,
    K::Comment,
    K::ListItem,
    K::Paragraph,
    K::Heading,
    K::Hyperlink,
];
const BLOCKS: &[Kind] = &[
    K::Paragraph,
    K::Heading,
    K::List,
    K::Table,
    K::Section,
    K::Frame,
];
const INLINES: &[Kind] = &[
    K::Run,
    K::Hyperlink,
    K::Bookmark,
    K::Footnote,
    K::Endnote,
    K::Comment,
    K::Frame,
];

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

/// The one static table shared by path resolution, query compilation and the
/// node-view projection.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Document,
    elements: &[
        element(
            K::Document,
            "document",
            &["document"],
            &[],
            &[K::Body, K::Styles, K::Metadata],
            &["/"],
            &[],
        ),
        element(
            K::Body,
            "body",
            &["body"],
            &[K::Document],
            BLOCKS,
            &["/body"],
            &[],
        ),
        element(
            K::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            BLOCK_PARENTS,
            INLINES,
            &["/body/p[N]"],
            PARAGRAPH_PROPS,
        ),
        element(
            K::Heading,
            "heading",
            &["h", "heading"],
            BLOCK_PARENTS,
            INLINES,
            &["/body/h[N]"],
            HEADING_PROPS,
        ),
        element(
            K::Run,
            "run",
            &["r", "run", "span"],
            INLINE_PARENTS,
            &[],
            &["/body/p[N]/r[N]"],
            RUN_PROPS,
        ),
        element(
            K::Hyperlink,
            "hyperlink",
            &["hyperlink"],
            INLINE_PARENTS,
            &[
                K::Run,
                K::Hyperlink,
                K::Bookmark,
                K::Footnote,
                K::Endnote,
                K::Comment,
                K::Frame,
            ],
            &["/body/p[N]/hyperlink[N]"],
            HYPERLINK_PROPS,
        ),
        element(
            K::List,
            "list",
            &["list"],
            &[
                K::Body,
                K::Section,
                K::Cell,
                K::TextBox,
                K::Footnote,
                K::Endnote,
                K::Comment,
                K::ListItem,
            ],
            &[K::ListItem],
            &["/body/list[N]"],
            LIST_PROPS,
        ),
        element(
            K::ListItem,
            "listItem",
            &["li", "listitem"],
            &[K::List],
            &[K::Paragraph, K::Heading, K::List],
            &["/body/list[N]/listItem[N]"],
            LIST_ITEM_PROPS,
        ),
        element(
            K::Table,
            "table",
            &["table", "tbl"],
            &[
                K::Body,
                K::Section,
                K::Cell,
                K::TextBox,
                K::Footnote,
                K::Endnote,
                K::Comment,
                K::ListItem,
            ],
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
            &["/body/tbl[N]/tr[N]"],
            ROW_PROPS,
        ),
        element(
            K::Cell,
            "cell",
            &["cell", "tc"],
            &[K::Row],
            BLOCKS,
            &["/body/tbl[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            K::Section,
            "section",
            &["section"],
            &[
                K::Body,
                K::Section,
                K::Cell,
                K::TextBox,
                K::Footnote,
                K::Endnote,
                K::Comment,
                K::ListItem,
            ],
            BLOCKS,
            &["/body/section[N]"],
            SECTION_PROPS,
        ),
        element(
            K::Frame,
            "frame",
            &["frame"],
            FRAME_PARENTS,
            &[K::Picture, K::TextBox],
            &["/body/p[N]/frame[N]"],
            FRAME_PROPS,
        ),
        element(
            K::Picture,
            "picture",
            &["image", "img", "picture"],
            &[K::Frame],
            &[],
            &["/body/p[N]/frame[N]/picture[N]"],
            PICTURE_PROPS,
        ),
        element(
            K::TextBox,
            "textbox",
            &["textbox"],
            &[K::Frame],
            BLOCKS,
            &["/body/p[N]/frame[N]/textbox"],
            TEXT_BOX_PROPS,
        ),
        element(
            K::Footnote,
            "footnote",
            &["footnote"],
            INLINE_PARENTS,
            BLOCKS,
            &["/body/p[N]/footnote[N]"],
            NOTE_PROPS,
        ),
        element(
            K::Endnote,
            "endnote",
            &["endnote"],
            INLINE_PARENTS,
            BLOCKS,
            &["/body/p[N]/endnote[N]"],
            NOTE_PROPS,
        ),
        element(
            K::Comment,
            "comment",
            &["annotation", "comment"],
            INLINE_PARENTS,
            BLOCKS,
            &["/body/p[N]/comment[N]"],
            COMMENT_PROPS,
        ),
        ElementSpec {
            identities: &["name"],
            ..element(
                K::Bookmark,
                "bookmark",
                &["bookmark"],
                INLINE_PARENTS,
                &[],
                &[
                    "/body/p[N]/bookmark[N]",
                    "/body/p[N]/bookmark[@name=<name>]",
                ],
                BOOKMARK_PROPS,
            )
        },
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
            K::Metadata,
            "metadata",
            &["metadata"],
            &[K::Document],
            &[],
            &["/metadata"],
            METADATA_PROPS,
        ),
    ],
    gaps: &[
        read_gap("header"),
        read_gap("footer"),
        read_gap("numbering"),
        read_gap("field"),
        read_gap("shape"),
        read_gap("chart"),
        read_gap("ole"),
        read_gap("toc"),
        read_gap("revision"),
        read_gap("col"),
        read_gap("equation"),
        read_gap("break"),
        read_gap("tab"),
        read_gap("group"),
        read_gap("formfield"),
        read_gap("settings"),
    ],
};

check_format_spec!(SPEC);

const _: () = {
    assert!(SPEC.elements.len() == Kind::COUNT);
    assert!(
        SPEC.elements.len() <= 128,
        "a query kind set holds 128 kinds"
    );
};

const fn read_gap(element: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: crate::Code::UnsupportedCapability,
        follow_up: "M6 ODT body index (record 0276)",
        reason: "the ODT body index does not model this element yet",
    }
}

/// Canonical names of the Word family that ODT has no concept of.
pub(crate) const ABSENT: &[&str] = &[
    "abstractNum",
    "axis",
    "bookmarkEnd",
    "comments",
    "customXml",
    "endnotes",
    "fieldChar",
    "footnotes",
    "instrText",
    "level",
    "num",
    "permStart",
    "ptab",
    "sdt",
    "series",
    "theme",
    "watermark",
];

/// Canonical element names whose properties `set` writes.
pub(crate) const SETTABLE: &[&str] = &[
    "bookmark",
    "comment",
    "endnote",
    "footnote",
    "frame",
    "h",
    "hyperlink",
    "listItem",
    "metadata",
    "p",
    "r",
    "section",
    "tbl",
    "tc",
    "textbox",
];

/// Canonical element names `add` creates.
pub(crate) const ADDABLE: &[&str] = &[
    "bookmark",
    "comment",
    "endnote",
    "footnote",
    "h",
    "hyperlink",
    "list",
    "listItem",
    "p",
    "r",
    "section",
    "tbl",
    "tr",
];

/// Canonical element names `remove` deletes (rows without vertical merges).
pub(crate) const REMOVABLE: &[&str] = &[
    "bookmark",
    "comment",
    "endnote",
    "footnote",
    "frame",
    "h",
    "hyperlink",
    "list",
    "listItem",
    "p",
    "picture",
    "r",
    "section",
    "tbl",
    "textbox",
    "tr",
];

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

/// The write operations the ODT adapter refuses, with their follow-ups
/// (record 0276 §gaps).
pub(crate) static WRITE_GAPS: [GapSpec; 12] = [
    operation_gap(
        "remove",
        "cell",
        "table-grid",
        "M6 ODT table grid edits",
        "cells keep the table grid: remove or move the whole row, or set the cell's text",
    ),
    operation_gap(
        "add",
        "cell",
        "table-grid",
        "M6 ODT table grid edits",
        "adding a cell would break the table grid; add a row, which spans the first row's columns",
    ),
    operation_gap(
        "move",
        "cell",
        "table-grid",
        "M6 ODT table grid edits",
        "cells move only with their row",
    ),
    operation_gap(
        "remove",
        "row",
        "vertical-merge",
        "M6 ODT table grid edits",
        "a row holding vertically merged or covered cells cannot leave the table",
    ),
    operation_gap(
        "add",
        "picture",
        "image-resources",
        "M6 ODT drawing and media edits",
        "adding an image needs a package media resource and a frame; frames and images are readable, settable (frame geometry) and removable",
    ),
    operation_gap(
        "add",
        "frame",
        "drawing-resources",
        "M6 ODT drawing and media edits",
        "frames are created with their content (image or text box), which the body edits do not author yet",
    ),
    operation_gap(
        "set",
        "style",
        "common-styles",
        "M6 ODT style definitions",
        "common styles in styles.xml are read-only; paragraph and run formatting derives automatic styles instead",
    ),
    operation_gap(
        "set",
        "run",
        "clear-formatting",
        "M6 ODT formatting removal",
        "a null value cannot remove a character or paragraph property; set the property's off value instead",
    ),
    operation_gap(
        "set",
        "cell",
        "typed-values",
        "M6 ODT typed cell values",
        "cell text writes a string cell; numeric, date and formula cell values are not authored",
    ),
    operation_gap(
        "set",
        "document",
        "signed-or-encrypted-package",
        "ADR 0034 ODF security vocabulary (record 0268 G5)",
        "the leaf preserves signed and encrypted packages only as exact no-ops; changed writes are refused",
    ),
    operation_gap(
        "get",
        "document",
        "flat-odf-input",
        "M6 flat ODF owner",
        "single-XML .fodt input is detected but refused; the leaf flat snapshot has no body index or package owner the adapter can publish through",
    ),
    operation_gap(
        "create",
        "document",
        "template",
        "use Flavor::Document",
        "ODT template creation (.ott) is not supported",
    ),
];
