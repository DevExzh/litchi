//! The OpenDocument presentation vocabulary of the unified adapter.
//!
//! ODP shares the presentation family's paths with PPTX and PPT. The table
//! lists every presentation concept ODF has; the ones the adapter cannot
//! answer yet carry a typed gap, and the concepts ODF lacks altogether
//! (SmartArt, slide zoom, threaded comments and replies, 3D models and
//! themes) are `ABSENT` in the navigator, so naming them is
//! `unsupported_path`. Property tables list only values the scene index or
//! the leaf projection answers.

use crate::Code;
use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthSpec, ValueType};

/// A node kind in an OpenDocument presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
#[non_exhaustive]
pub enum Kind {
    /// The presentation root (`/`).
    Presentation,
    /// A `draw:page` (`/slide[N]`).
    Slide,
    /// A text frame or drawing shape (`/slide[N]/shape[N]`).
    Shape,
    /// A presentation placeholder (`presentation:class`), addressed among
    /// shapes.
    Placeholder,
    /// An image frame.
    Picture,
    /// A table frame.
    Table,
    /// A table row.
    Row,
    /// A table cell.
    Cell,
    /// A chart object frame.
    Chart,
    /// A `draw:g` group.
    Group,
    /// A line or connector.
    Connector,
    /// An embedded object frame.
    Ole,
    /// A plugin (audio or video) frame.
    Media,
    /// A video plugin frame (record 0328), addressed among media.
    Video,
    /// An audio plugin frame (record 0328), addressed among media.
    Audio,
    /// A paragraph of a text story.
    Paragraph,
    /// A run of a paragraph.
    Run,
    /// A `text:line-break` of a paragraph (record 0328).
    LineBreak,
    /// A shape's click action or a link run's target (record 0328).
    Hyperlink,
    /// A slide's speaker notes.
    Notes,
    /// A slide annotation.
    Comment,
    /// A formula object frame (record 0328), addressed among objects.
    Equation,
    /// A master page.
    SlideMaster,
    /// A presentation page layout.
    SlideLayout,
    /// The notes master (gap).
    NotesMaster,
    /// The notes page as a slide (gap).
    NoteSlide,
    /// An animation (gap).
    Animation,
    /// A chart series (record 0328).
    Series,
    /// A chart axis (record 0328).
    Axis,
    /// A declared table column (record 0328).
    Col,
}

impl Kind {
    /// Every kind, in the ordinal order of [`SPEC`].
    pub(crate) const ALL: &'static [Self] = &[
        Self::Presentation,
        Self::Slide,
        Self::Shape,
        Self::Placeholder,
        Self::Picture,
        Self::Table,
        Self::Row,
        Self::Cell,
        Self::Chart,
        Self::Group,
        Self::Connector,
        Self::Ole,
        Self::Media,
        Self::Video,
        Self::Audio,
        Self::Paragraph,
        Self::Run,
        Self::LineBreak,
        Self::Hyperlink,
        Self::Notes,
        Self::Comment,
        Self::Equation,
        Self::SlideMaster,
        Self::SlideLayout,
        Self::NotesMaster,
        Self::NoteSlide,
        Self::Animation,
        Self::Series,
        Self::Axis,
        Self::Col,
    ];

    pub(crate) const COUNT: usize = Self::ALL.len();

    pub(crate) const fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }

    /// The canonical path segment.
    #[must_use]
    pub const fn name(self) -> &'static str {
        self.spec().name
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::PRESENTATION;
    const ROOT: Self = Self::Presentation;

    fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "presentation" => Self::Presentation,
            "slide" => Self::Slide,
            "shape" | "textbox" | "title" => Self::Shape,
            "ph" | "placeholder" => Self::Placeholder,
            "image" | "img" | "pic" | "picture" => Self::Picture,
            "table" => Self::Table,
            "row" | "tr" => Self::Row,
            "cell" | "tc" => Self::Cell,
            "chart" => Self::Chart,
            "group" => Self::Group,
            "connection" | "connector" => Self::Connector,
            "embed" | "object" | "ole" | "oleobject" => Self::Ole,
            "media" => Self::Media,
            "video" => Self::Video,
            "audio" => Self::Audio,
            "p" | "paragraph" => Self::Paragraph,
            "r" | "run" => Self::Run,
            "br" | "linebreak" => Self::LineBreak,
            "hlink" | "hyperlink" => Self::Hyperlink,
            "notes" => Self::Notes,
            "comment" => Self::Comment,
            "equation" | "formula" | "math" => Self::Equation,
            "slideMaster" | "slidemaster" => Self::SlideMaster,
            "layout" | "slideLayout" | "slidelayout" => Self::SlideLayout,
            "notesMaster" | "notesmaster" => Self::NotesMaster,
            "noteSlide" | "noteslide" => Self::NoteSlide,
            "animate" | "animation" => Self::Animation,
            "series" => Self::Series,
            "axis" => Self::Axis,
            "col" | "column" => Self::Col,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        self.spec().name
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

const fn raw(name: &'static str, ty: ValueType) -> PropSpec {
    PropSpec {
        name,
        ty,
        access: Access::READ,
        raw: true,
        admission: crate::registry::Admission::NONE,
    }
}

const fn length(name: &'static str, ty: LengthSpec) -> PropSpec {
    PropSpec::new(name, ValueType::Length(ty))
}

/// The presentation root: slide count and the `meta.xml` projection.
pub(crate) const PRESENTATION_PROPS: &[PropSpec] = &[
    PropSpec::new("slideCount", ValueType::Int),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
    PropSpec::new("subject", ValueType::Text),
    PropSpec::new("keywords", ValueType::List),
    PropSpec::new("description", ValueType::Text),
];

/// A `draw:page`.
pub(crate) const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("slideNumber", ValueType::Int),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("layout", ValueType::Text),
    PropSpec::new("master", ValueType::Text),
];

/// A text frame or drawing shape, and a placeholder.
pub(crate) const SHAPE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    raw("id", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("shapeType", ValueType::Text),
    PropSpec::new("placeholderType", ValueType::Text),
    PropSpec::new("isTitle", ValueType::Bool),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("zorder", ValueType::Int),
    PropSpec::new("style", ValueType::Text),
    // The formatting every run of the shape's text shares (record 0316).
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    length("size", LengthSpec::FONT_SIZE),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("font", ValueType::Text),
];

/// An image, chart, embedded-object or media frame.
pub(crate) const FRAME_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    raw("id", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("source", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("zorder", ValueType::Int),
];

/// A table frame.
pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    raw("id", ValueType::Text),
    PropSpec::new("rows", ValueType::Int),
    PropSpec::new("cols", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("zorder", ValueType::Int),
    PropSpec::new("style", ValueType::Text),
];

/// A group or a line/connector.
pub(crate) const GROUP_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    raw("id", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("zorder", ValueType::Int),
];

/// A table cell.
pub(crate) const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("colSpan", ValueType::Int),
    PropSpec::new("rowSpan", ValueType::Int),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    length("size", LengthSpec::FONT_SIZE),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("font", ValueType::Text),
];

/// The character properties of paragraphs and runs (record 0308), and of
/// shapes, table cells and notes (record 0316): read through the node's
/// style chain, common styles included, and written as derived automatic
/// styles.
pub(crate) const CHARACTER_KEYS: &[&str] =
    &["bold", "italic", "underline", "size", "color", "font"];

/// A paragraph: its text is the node text.
pub(crate) const PARAGRAPH_PROPS: &[PropSpec] = &[
    PropSpec::new("style", ValueType::Text),
    PropSpec::new("level", ValueType::Int),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    length("size", LengthSpec::FONT_SIZE),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("font", ValueType::Text),
];

/// A run: its text is the node text.
pub(crate) const RUN_PROPS: &[PropSpec] = &[
    PropSpec::new("style", ValueType::Text),
    PropSpec::new("link", ValueType::Text),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    length("size", LengthSpec::FONT_SIZE),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("font", ValueType::Text),
];

/// Speaker notes.
pub(crate) const NOTES_PROPS: &[PropSpec] = &[
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    length("size", LengthSpec::FONT_SIZE),
    PropSpec::new("color", ValueType::Color(ColorContext::Word)),
    PropSpec::new("font", ValueType::Text),
];

/// An annotation.
pub(crate) const COMMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
    PropSpec::new("initials", ValueType::Text),
    PropSpec::new("date", ValueType::Text),
];

/// A master page.
pub(crate) const MASTER_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("displayName", ValueType::Text),
    PropSpec::new("pageLayout", ValueType::Text),
];

/// A presentation page layout.
pub(crate) const LAYOUT_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("displayName", ValueType::Text),
];

/// A readable property `set` and `add` write.
const fn writable(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.set = true;
    spec.access.add = true;
    spec
}

/// A hyperlink (record 0328): a shape's click action or a link run's
/// target, as PPTX names them.
pub(crate) const HYPERLINK_PROPS: &[PropSpec] = &[
    writable("link", ValueType::Text),
    writable("tooltip", ValueType::Text),
];

/// A video or audio plugin frame (record 0328): the frame's keys, then the
/// media.
pub(crate) const MEDIA_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    raw("id", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("source", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("zorder", ValueType::Int),
    PropSpec::new("mediaType", ValueType::Text),
    PropSpec::new("mimeType", ValueType::Text),
    PropSpec::new("loop", ValueType::Bool),
    PropSpec::new("mute", ValueType::Bool),
];

/// A formula object (record 0328): its StarMath (or MathML token) text and
/// its frame.
pub(crate) const EQUATION_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Text),
    PropSpec::new("formula", ValueType::Text),
    PropSpec::new("name", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

/// A declared table column (record 0328).
pub(crate) const COLUMN_PROPS: &[PropSpec] = &[length("width", LengthSpec::EXTENT)];

/// A chart series (record 0328), with the ODC series keys.
pub(crate) const SERIES_PROPS: &[PropSpec] = &[
    PropSpec::new("values", ValueType::Text),
    PropSpec::new("nameRef", ValueType::Text),
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("axis", ValueType::Text),
    PropSpec::new("pointCount", ValueType::Int),
];

/// A chart axis (record 0328): its PPTX role, then the ODC axis keys.
pub(crate) const AXIS_PROPS: &[PropSpec] = &[
    PropSpec::new("role", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("dimension", ValueType::Text),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
];

const TREE_PARENTS: &[Kind] = &[Kind::Slide, Kind::Group];

const TREE: &[Kind] = &[
    Kind::Shape,
    Kind::Placeholder,
    Kind::Picture,
    Kind::Table,
    Kind::Chart,
    Kind::Group,
    Kind::Connector,
    Kind::Ole,
    Kind::Media,
    Kind::Video,
    Kind::Audio,
    Kind::Equation,
];

const SHAPE_IDS: &[&str] = &["id", "name"];

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

const fn gap(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    paths: &'static [&'static str],
    gap: &'static GapSpec,
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children: &[],
        identities: &[],
        paths,
        props: &[],
        gap: Some(gap),
    }
}

const fn element_gap(
    element: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

static NOTES_MASTER_GAP: GapSpec = element_gap(
    "notesMaster",
    "M6 ODP notes master",
    "the master page's presentation:notes is not projected as a node",
);
static NOTE_SLIDE_GAP: GapSpec = element_gap(
    "noteSlide",
    "M6 ODP notes pages",
    "speaker notes are exposed as /slide[N]/notes; the notes page's drawing objects are not nodes",
);
static ANIMATION_GAP: GapSpec = element_gap(
    "animation",
    "M6 ODP animation nodes",
    "the leaf keeps slide animation trees inert; they are not projected as nodes",
);

/// A PresentationML concept ODF has no counterpart of (record 0328).
const fn pptx_only(element: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: Code::UnsupportedPath,
        follow_up: "none: a PresentationML-only concept",
        reason: "OpenDocument presentations have no themes, modern comments, 3D models or zoom frames",
    }
}

static SHAPE_FILL_GAP: GapSpec = GapSpec {
    subject: Subject::Property {
        element: "shape",
        key: "fill",
    },
    code: Code::UnsupportedCapability,
    follow_up: "M6 ODP graphic styles",
    reason: "fill and line come from automatic graphic styles, which the adapter does not resolve",
};
static STATS_FONT_GAP: GapSpec = GapSpec {
    subject: Subject::View("stats.fonts"),
    code: Code::UnsupportedCapability,
    follow_up: "M6 ODP text styles",
    reason: "font usage comes from text styles, which the stats view does not resolve",
};

/// The OpenDocument presentation vocabulary.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Presentation,
    elements: &[
        element(
            Kind::Presentation,
            "presentation",
            &["presentation"],
            &[],
            &[Kind::Slide, Kind::SlideMaster, Kind::SlideLayout],
            &[],
            &["/"],
            PRESENTATION_PROPS,
        ),
        element(
            Kind::Slide,
            "slide",
            &["slide"],
            &[Kind::Presentation],
            &[
                Kind::Shape,
                Kind::Placeholder,
                Kind::Picture,
                Kind::Table,
                Kind::Chart,
                Kind::Group,
                Kind::Connector,
                Kind::Ole,
                Kind::Media,
                Kind::Video,
                Kind::Audio,
                Kind::Notes,
                Kind::Comment,
                Kind::Equation,
                Kind::NoteSlide,
                Kind::Animation,
            ],
            &[],
            &["/slide[N]"],
            SLIDE_PROPS,
        ),
        element(
            Kind::Shape,
            "shape",
            &["shape", "textbox", "title"],
            TREE_PARENTS,
            &[Kind::Paragraph, Kind::Hyperlink],
            SHAPE_IDS,
            &["/slide[N]/shape[N]"],
            SHAPE_PROPS,
        ),
        element(
            Kind::Placeholder,
            "placeholder",
            &["ph", "placeholder"],
            TREE_PARENTS,
            &[Kind::Paragraph],
            SHAPE_IDS,
            &["/slide[N]/placeholder[N]", "/slide[N]/placeholder[TYPE]"],
            SHAPE_PROPS,
        ),
        element(
            Kind::Picture,
            "picture",
            &["image", "img", "pic", "picture"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/picture[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table"],
            TREE_PARENTS,
            &[Kind::Row, Kind::Col],
            SHAPE_IDS,
            &["/slide[N]/table[N]"],
            TABLE_PROPS,
        ),
        element(
            Kind::Row,
            "tr",
            &["row", "tr"],
            &[Kind::Table],
            &[Kind::Cell],
            &[],
            &["/slide[N]/table[N]/tr[N]"],
            &[],
        ),
        element(
            Kind::Cell,
            "tc",
            &["cell", "tc"],
            &[Kind::Row],
            &[Kind::Paragraph],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            Kind::Chart,
            "chart",
            &["chart"],
            TREE_PARENTS,
            &[Kind::Series, Kind::Axis],
            SHAPE_IDS,
            &["/slide[N]/chart[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Group,
            "group",
            &["group"],
            TREE_PARENTS,
            TREE,
            SHAPE_IDS,
            &["/slide[N]/group[N]"],
            GROUP_PROPS,
        ),
        element(
            Kind::Connector,
            "connector",
            &["connection", "connector"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/connector[N]"],
            GROUP_PROPS,
        ),
        element(
            Kind::Ole,
            "ole",
            &["embed", "object", "ole", "oleobject"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/ole[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Media,
            "media",
            &["media"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/media[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Video,
            "video",
            &["video"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/video[N]"],
            MEDIA_PROPS,
        ),
        element(
            Kind::Audio,
            "audio",
            &["audio"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/audio[N]"],
            MEDIA_PROPS,
        ),
        element(
            Kind::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Kind::Shape, Kind::Placeholder, Kind::Cell, Kind::Notes],
            &[Kind::Run, Kind::LineBreak],
            &[],
            &["/slide[N]/shape[N]/paragraph[N]"],
            PARAGRAPH_PROPS,
        ),
        element(
            Kind::Run,
            "run",
            &["r", "run"],
            &[Kind::Paragraph],
            &[Kind::Hyperlink],
            &[],
            &["/slide[N]/shape[N]/paragraph[N]/run[N]"],
            RUN_PROPS,
        ),
        element(
            Kind::LineBreak,
            "linebreak",
            &["br", "linebreak"],
            &[Kind::Paragraph],
            &[],
            &[],
            &["/slide[N]/shape[N]/paragraph[N]/linebreak[N]"],
            &[],
        ),
        element(
            Kind::Hyperlink,
            "hyperlink",
            &["hlink", "hyperlink"],
            &[Kind::Shape, Kind::Run],
            &[],
            &[],
            &[
                "/slide[N]/shape[N]/hyperlink",
                "/slide[N]/shape[N]/paragraph[N]/run[N]/hyperlink",
            ],
            HYPERLINK_PROPS,
        ),
        element(
            Kind::Notes,
            "notes",
            &["notes"],
            &[Kind::Slide],
            &[Kind::Paragraph],
            &[],
            &["/slide[N]/notes"],
            NOTES_PROPS,
        ),
        element(
            Kind::Comment,
            "comment",
            &["comment"],
            &[Kind::Slide],
            &[],
            &[],
            &["/slide[N]/comment[N]"],
            COMMENT_PROPS,
        ),
        element(
            Kind::Equation,
            "equation",
            &["equation", "formula", "math"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/equation[N]"],
            EQUATION_PROPS,
        ),
        element(
            Kind::SlideMaster,
            "slideMaster",
            &["slidemaster"],
            &[Kind::Presentation],
            &[],
            &[],
            &["/slidemaster[N]"],
            MASTER_PROPS,
        ),
        element(
            Kind::SlideLayout,
            "slideLayout",
            &["layout", "slidelayout"],
            &[Kind::Presentation],
            &[],
            &[],
            &["/slidelayout[N]"],
            LAYOUT_PROPS,
        ),
        gap(
            Kind::NotesMaster,
            "notesMaster",
            &["notesmaster"],
            &[],
            &["/notesMaster"],
            &NOTES_MASTER_GAP,
        ),
        gap(
            Kind::NoteSlide,
            "noteSlide",
            &["noteslide"],
            &[Kind::Slide],
            &["/slide[N]/noteSlide"],
            &NOTE_SLIDE_GAP,
        ),
        gap(
            Kind::Animation,
            "animation",
            &["animate", "animation"],
            &[Kind::Slide],
            &["/slide[N]/animation[N]"],
            &ANIMATION_GAP,
        ),
        element(
            Kind::Series,
            "series",
            &["series"],
            &[Kind::Chart],
            &[],
            &[],
            &["/slide[N]/chart[N]/series[N]"],
            SERIES_PROPS,
        ),
        element(
            Kind::Axis,
            "axis",
            &["axis"],
            &[Kind::Chart],
            &[],
            &["role"],
            &[
                "/slide[N]/chart[N]/axis[N]",
                "/slide[N]/chart[N]/axis[@role=ROLE]",
            ],
            AXIS_PROPS,
        ),
        element(
            Kind::Col,
            "col",
            &["col", "column"],
            &[Kind::Table],
            &[],
            &[],
            &["/slide[N]/table[N]/col[N]"],
            COLUMN_PROPS,
        ),
    ],
    gaps: &[
        SHAPE_FILL_GAP,
        STATS_FONT_GAP,
        pptx_only("theme"),
        pptx_only("modernComment"),
        pptx_only("model3d"),
        pptx_only("zoom"),
        // OfficeCLI's Mermaid `diagram` (record 0339).
        GapSpec {
            subject: Subject::Element("diagram"),
            code: Code::UnsupportedPath,
            follow_up: "none: out of scope (record 0332, design §10)",
            reason: "OfficeCLI's diagram renders Mermaid text into shapes or a picture; rendering is outside the unified API",
        },
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(SPEC.elements.len() <= 128);
