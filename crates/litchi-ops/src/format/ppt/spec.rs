//! The read-side vocabulary of the legacy PowerPoint adapter.
//!
//! The binary reader exposes one presentation catalog, lazily decoded slides,
//! and the native shape tree of each slide.  The table contains the complete
//! shared presentation vocabulary so a known but unimplemented axis is
//! reported as a typed capability gap instead of looking like an unknown
//! path.  Property tables contain only values the native projection answers.

use crate::Code;
use crate::node::spec::{
    Access, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthSpec, Tokens, ValueType};

/// A node kind in a legacy PowerPoint presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
#[non_exhaustive]
pub(crate) enum Kind {
    Presentation,
    Slide,
    Shape,
    Placeholder,
    Picture,
    Table,
    Row,
    Cell,
    Chart,
    Group,
    Connector,
    SmartArt,
    Ole,
    Paragraph,
    Run,
    LineBreak,
    Notes,
    Comment,
    Video,
    Audio,
    Media,
    Equation,
    Zoom,
    Model3d,
    SlideMaster,
    SlideLayout,
    Theme,
    ModernComment,
    Reply,
    Hyperlink,
    Animation,
    Series,
    Axis,
    Col,
    NoteSlide,
    NotesMaster,
}

impl Kind {
    /// Every kind, in the ordinal order used by [`SPEC`].
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
        Self::SmartArt,
        Self::Ole,
        Self::Paragraph,
        Self::Run,
        Self::LineBreak,
        Self::Notes,
        Self::Comment,
        Self::Video,
        Self::Audio,
        Self::Media,
        Self::Equation,
        Self::Zoom,
        Self::Model3d,
        Self::SlideMaster,
        Self::SlideLayout,
        Self::Theme,
        Self::ModernComment,
        Self::Reply,
        Self::Hyperlink,
        Self::Animation,
        Self::Series,
        Self::Axis,
        Self::Col,
        Self::NoteSlide,
        Self::NotesMaster,
    ];

    pub(crate) const COUNT: usize = Self::ALL.len();

    pub(crate) const fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }

    pub(crate) const fn name(self) -> &'static str {
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
            "smartart" => Self::SmartArt,
            "embed" | "object" | "ole" | "oleobject" => Self::Ole,
            "p" | "paragraph" => Self::Paragraph,
            "r" | "run" => Self::Run,
            "br" | "linebreak" => Self::LineBreak,
            "notes" => Self::Notes,
            "comment" => Self::Comment,
            "video" => Self::Video,
            "audio" => Self::Audio,
            "media" => Self::Media,
            "equation" | "formula" | "math" => Self::Equation,
            "slidezoom" | "zoom" => Self::Zoom,
            "3dmodel" | "model3d" => Self::Model3d,
            "layout" | "slidelayout" => Self::SlideLayout,
            "slidemaster" => Self::SlideMaster,
            "theme" => Self::Theme,
            "modern-comment" | "moderncomment" | "thread" | "threadedcomment" => {
                Self::ModernComment
            },
            "reply" => Self::Reply,
            "hlink" | "hyperlink" => Self::Hyperlink,
            "animate" | "animation" => Self::Animation,
            "series" => Self::Series,
            "axis" => Self::Axis,
            "col" | "column" => Self::Col,
            "noteslide" => Self::NoteSlide,
            "notesmaster" => Self::NotesMaster,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        self.name()
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

const PRESENTATION_PROPS: &[PropSpec] = &[PropSpec::new("slideCount", ValueType::Int)];

const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("slideNumber", ValueType::Int),
    raw("slideId", ValueType::Int),
    raw("persistId", ValueType::Int),
    PropSpec::new("hidden", ValueType::Bool),
];

const SHAPE_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    PropSpec::new("shapeType", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("hidden", ValueType::Bool),
];

const PLACEHOLDER_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    PropSpec::new("placeholderType", ValueType::Text),
    raw("placeholderIndex", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("hidden", ValueType::Bool),
];

const PICTURE_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    PropSpec::new("frameKind", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("hidden", ValueType::Bool),
];

const TABLE_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    PropSpec::new("rows", ValueType::Int),
    PropSpec::new("columns", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

const GROUP_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

const CONNECTOR_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    raw("sourceIndex", ValueType::Int),
    PropSpec::new("shapeType", ValueType::Text),
    PropSpec::new("length", ValueType::Number),
];

const TREE: &[Kind] = &[
    Kind::Shape,
    Kind::Placeholder,
    Kind::Picture,
    Kind::Table,
    Kind::Chart,
    Kind::Group,
    Kind::Connector,
    Kind::SmartArt,
    Kind::Ole,
    Kind::Video,
    Kind::Audio,
    Kind::Media,
];

const TREE_PARENTS: &[Kind] = &[
    Kind::Slide,
    Kind::Group,
    Kind::SlideMaster,
    Kind::SlideLayout,
];

const SHAPE_IDS: &[&str] = &["id", "name"];

/// Speaker notes: the text of the notes page's body placeholder.
const NOTES_PROPS: &[PropSpec] = &[PropSpec::new("text", ValueType::Text)];

/// A slide comment (`Comment10Container`, record 0323), in the PPTX comment
/// vocabulary: the anchor is in master units, read as a length.
pub(crate) const COMMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
    PropSpec::new("initials", ValueType::Text),
    PropSpec::new("index", ValueType::Int),
    PropSpec::new("date", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
];

/// Run underline (record 0323): the binary format's `CFStyle` underline is
/// single or none, printed as the PPTX `underline` tokens.
pub(crate) static UNDERLINE: Tokens = Tokens::new("underline", &["single", "none"], &[("sng", 0)]);

/// A text run (record 0323): one `StyleTextPropAtom` character run's piece
/// of a paragraph, in the PPTX run vocabulary. `font.latin` is the
/// `FontCollection` typeface; `color` reads an sRGB color only.
pub(crate) const RUN_PROPS: &[PropSpec] = &[
    PropSpec::new("font.latin", ValueType::Text),
    PropSpec::new("size", ValueType::Length(LengthSpec::FONT_SIZE)),
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Token(&UNDERLINE)),
    PropSpec::new("color", ValueType::Color(ColorContext::Drawing)),
];

/// An OLE object frame (record 0323): a picture frame showing an embedded or
/// linked object, with the object's `ProgID`.
const OLE_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("progId", ValueType::Text),
    PropSpec::new("linked", ValueType::Bool),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

/// A media frame (record 0323): a picture frame playing a movie (`video`) or
/// a sound (`audio`); `media` names both.
const MEDIA_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

/// An embedded chart (record 0323): a Microsoft Graph or Excel chart object
/// shown by an OLE picture frame, read through its MS-OGRAPH chart. The
/// frame keys are the picture's; `chartType` uses the PPTX chart names.
const CHART_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    PropSpec::new("name", ValueType::Text),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
];

/// The presentation theme (record 0323): the `PowerPoint` 2007 theme part a
/// `RoundTripThemeAtom` carries, read with the PPTX theme vocabulary.
pub(crate) const THEME_PROPS: &[PropSpec] = &[
    PropSpec::new("dk1", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("lt1", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("dk2", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("lt2", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent1", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent2", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent3", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent4", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent5", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("accent6", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("hyperlink", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("followedhyperlink", ValueType::Color(ColorContext::Drawing)),
    PropSpec::new("headingFont", ValueType::Text),
    PropSpec::new("bodyFont", ValueType::Text),
    PropSpec::new("headingFont.ea", ValueType::Text),
    PropSpec::new("bodyFont.ea", ValueType::Text),
    PropSpec::new("headingFont.cs", ValueType::Text),
    PropSpec::new("bodyFont.cs", ValueType::Text),
    PropSpec::new("name", ValueType::Text),
];

/// A click hyperlink (record 0323), in the PPTX vocabulary: `link` is a URL,
/// `slide[N]`, or `firstslide`, `lastslide`, `nextslide`, `previousslide`,
/// `endshow`; `tooltip` is the screen tip.
pub(crate) const HYPERLINK_PROPS: &[PropSpec] = &[
    PropSpec::new("link", ValueType::Text),
    PropSpec::new("tooltip", ValueType::Text),
];

/// A table row (record 0323): a band of the table group's cell shapes; its
/// height is the distance to the next row's top.
const ROW_PROPS: &[PropSpec] = &[length("height", LengthSpec::EXTENT)];

/// A table cell (record 0323): one cell shape; its text is the node's text
/// member and `text` a write-only registry key, as for paragraphs.
const CELL_PROPS: &[PropSpec] = &[];

/// A table column (record 0323): its width.
const COLUMN_PROPS: &[PropSpec] = &[length("width", LengthSpec::EXTENT)];

/// A paragraph of a shape's text body: the text between paragraph marks. Its
/// text is the node's text member, as for PPTX paragraphs; `text` is a
/// write-only registry key.
const PARAGRAPH_PROPS: &[PropSpec] = &[];

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
    children: &'static [Kind],
    paths: &'static [&'static str],
    gap: &'static GapSpec,
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
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

const fn property_gap(
    element: &'static str,
    key: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

static SMARTART_GAP: GapSpec = element_gap(
    "smartart",
    "0339-F3 (deferred-P1): PPT SmartArt and diagram groups",
    "a legacy PPT file keeps SmartArt and PowerPoint 97 diagrams as drawn OfficeArt groups, which read as group and shape nodes; no semantic diagram model is projected",
);
static EQUATION_GAP: GapSpec = element_gap(
    "equation",
    "0323-F6 PPT Equation Editor objects",
    "legacy PPT equations are Equation Editor OLE objects (Equation.3) whose MTEF payload the unified layer does not read as an equation element",
);
static ZOOM_GAP: GapSpec = element_gap(
    "zoom",
    "wontfix: PPTX-only (record 0323)",
    "slide and section zoom are PowerPoint 2016 OOXML objects; the legacy binary format has no zoom record",
);
static MODEL3D_GAP: GapSpec = element_gap(
    "model3d",
    "wontfix: PPTX-only (record 0323)",
    "3D models are PowerPoint 2019 OOXML objects; the legacy binary format has no 3D-model record",
);
static MASTER_GAP: GapSpec = element_gap(
    "slideMaster",
    "0339-F4 (deferred-P1): PPT main and title masters as nodes",
    "legacy PPT master records are not exposed as navigable nodes",
);
static LAYOUT_GAP: GapSpec = element_gap(
    "slideLayout",
    "0339-F4 (deferred-P1): PPT layouts as nodes",
    "a legacy PPT slide names a layout type in its slide atom (PowerPoint 2007 and later add one master per layout); there are no layout records to expose as navigable nodes yet",
);
static MODERN_COMMENT_GAP: GapSpec = element_gap(
    "modernComment",
    "wontfix: PPTX-only (record 0323)",
    "modern threaded comments are a PowerPoint 2019 OOXML part; legacy PPT comments are Comment10Container records, read and written as comment elements",
);
static REPLY_GAP: GapSpec = element_gap(
    "reply",
    "wontfix: PPTX-only (record 0323)",
    "threaded-comment replies are a PowerPoint 2019 OOXML part; legacy PPT comments have no replies",
);
static ANIMATION_GAP: GapSpec = element_gap(
    "animation",
    "0323-F5 PPT animation nodes",
    "the leaf animation codec reads the timing tree and the per-shape AnimationInfo atoms; they are not projected as per-shape nodes",
);
static SERIES_GAP: GapSpec = element_gap(
    "series",
    "0323-F7 PPT chart series and axes",
    "the MS-OGRAPH chart model reads series, but the unified layer projects a chart node only (type, title, frame)",
);
static AXIS_GAP: GapSpec = element_gap(
    "axis",
    "0323-F7 PPT chart series and axes",
    "the MS-OGRAPH chart model reads axes, but the unified layer projects a chart node only (type, title, frame)",
);
static NOTE_SLIDE_GAP: GapSpec = element_gap(
    "noteSlide",
    "M1-extended",
    "legacy PPT notes slides are not exposed as navigable nodes",
);
static NOTES_MASTER_GAP: GapSpec = element_gap(
    "notesMaster",
    "M1-extended",
    "legacy PPT notes masters are not exposed as navigable nodes",
);

static PICTURE_MEDIA_GAP: GapSpec = property_gap(
    "picture",
    "mediaType",
    "M1-extended",
    "legacy PPT frame classification is exposed as frameKind",
);
static SHAPE_FILL_GAP: GapSpec = property_gap(
    "shape",
    "fill",
    "M1-extended",
    "legacy PPT exposes only common geometry in the unified node view",
);
static STATS_CHART_GAP: GapSpec = GapSpec {
    subject: Subject::View("stats.charts"),
    code: Code::UnsupportedCapability,
    follow_up: "M1-extended",
    reason: "legacy PPT has no chart-part model from which the stats view can count charts",
};

/// The complete shared presentation vocabulary with native PPT read support
/// for the presentation, slide and shape-tree projections.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Presentation,
    elements: &[
        element(
            Kind::Presentation,
            "presentation",
            &["presentation"],
            &[],
            &[
                Kind::Slide,
                Kind::SlideMaster,
                Kind::Theme,
                Kind::NotesMaster,
            ],
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
                Kind::SmartArt,
                Kind::Ole,
                Kind::Video,
                Kind::Audio,
                Kind::Media,
                Kind::Notes,
                Kind::Comment,
                Kind::Equation,
                Kind::Zoom,
                Kind::Model3d,
                Kind::ModernComment,
                Kind::NoteSlide,
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
            &[
                Kind::Paragraph,
                Kind::Equation,
                Kind::Hyperlink,
                Kind::Animation,
            ],
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
            &["type", "idx", "id", "name"],
            &["/slide[N]/placeholder[N]", "/slide[N]/placeholder[TYPE]"],
            PLACEHOLDER_PROPS,
        ),
        element(
            Kind::Picture,
            "picture",
            &["image", "img", "pic", "picture"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/picture[N]"],
            PICTURE_PROPS,
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
            ROW_PROPS,
        ),
        element(
            Kind::Cell,
            "tc",
            &["cell", "tc"],
            &[Kind::Row],
            &[],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            Kind::Chart,
            "chart",
            &["chart"],
            TREE_PARENTS,
            &[Kind::Series, Kind::Axis, Kind::Animation],
            &[],
            &["/slide[N]/chart[N]"],
            CHART_PROPS,
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
            CONNECTOR_PROPS,
        ),
        gap(
            Kind::SmartArt,
            "smartart",
            &["smartart"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/smartart[N]"],
            &SMARTART_GAP,
        ),
        element(
            Kind::Ole,
            "ole",
            &["embed", "object", "ole", "oleobject"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/ole[N]"],
            OLE_PROPS,
        ),
        element(
            Kind::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Kind::Shape, Kind::Placeholder],
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
            Kind::Notes,
            "notes",
            &["notes"],
            &[Kind::Slide],
            &[],
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
            Kind::Media,
            "media",
            &["media"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/media[N]"],
            MEDIA_PROPS,
        ),
        gap(
            Kind::Equation,
            "equation",
            &["equation", "formula", "math"],
            &[Kind::Slide, Kind::Shape],
            &[],
            &["/slide[N]/shape[N]/equation[N]"],
            &EQUATION_GAP,
        ),
        gap(
            Kind::Zoom,
            "zoom",
            &["slidezoom", "zoom"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/zoom[N]"],
            &ZOOM_GAP,
        ),
        gap(
            Kind::Model3d,
            "model3d",
            &["3dmodel", "model3d"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/model3d[N]"],
            &MODEL3D_GAP,
        ),
        gap(
            Kind::SlideMaster,
            "slideMaster",
            &["slidemaster"],
            &[Kind::Presentation],
            TREE,
            &["/slidemaster[N]"],
            &MASTER_GAP,
        ),
        gap(
            Kind::SlideLayout,
            "slideLayout",
            &["layout", "slidelayout"],
            &[Kind::SlideMaster],
            TREE,
            &["/slidemaster[N]/slidelayout[N]", "/slidelayout[N]"],
            &LAYOUT_GAP,
        ),
        element(
            Kind::Theme,
            "theme",
            &["theme"],
            &[Kind::Presentation],
            &[],
            &[],
            &["/theme"],
            THEME_PROPS,
        ),
        gap(
            Kind::ModernComment,
            "modernComment",
            &[
                "modern-comment",
                "moderncomment",
                "thread",
                "threadedcomment",
            ],
            &[Kind::Slide],
            &[Kind::Reply],
            &["/slide[N]/moderncomment[N]"],
            &MODERN_COMMENT_GAP,
        ),
        gap(
            Kind::Reply,
            "reply",
            &["reply"],
            &[Kind::ModernComment],
            &[],
            &["/slide[N]/moderncomment[N]/reply[N]"],
            &REPLY_GAP,
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
        gap(
            Kind::Animation,
            "animation",
            &["animate", "animation"],
            &[Kind::Shape, Kind::Chart],
            &[],
            &["/slide[N]/shape[N]/animation[N]"],
            &ANIMATION_GAP,
        ),
        gap(
            Kind::Series,
            "series",
            &["series"],
            &[Kind::Chart],
            &[],
            &["/slide[N]/chart[N]/series[N]"],
            &SERIES_GAP,
        ),
        gap(
            Kind::Axis,
            "axis",
            &["axis"],
            &[Kind::Chart],
            &[],
            &["/slide[N]/chart[N]/axis[@role=ROLE]"],
            &AXIS_GAP,
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
        gap(
            Kind::NoteSlide,
            "noteSlide",
            &["noteslide"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/noteSlide"],
            &NOTE_SLIDE_GAP,
        ),
        gap(
            Kind::NotesMaster,
            "notesMaster",
            &["notesmaster"],
            &[Kind::Presentation],
            &[],
            &["/notesMaster"],
            &NOTES_MASTER_GAP,
        ),
    ],
    gaps: &[
        PICTURE_MEDIA_GAP,
        SHAPE_FILL_GAP,
        STATS_CHART_GAP,
        // OfficeCLI's Mermaid `diagram` (record 0339).
        GapSpec {
            subject: Subject::Element("diagram"),
            code: Code::UnsupportedPath,
            follow_up: "none: out of scope (record 0332, design §10)",
            reason: "OfficeCLI's diagram renders Mermaid text into shapes or a picture; rendering is outside the unified API, and SmartArt is the smartart element",
        },
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(SPEC.elements.len() <= 128);

/// Setter keys shared by the PPT shape and placeholder registries.
#[derive(Clone, Copy)]
pub(super) enum ShapeSetter {
    Text,
    X,
    Y,
    Width,
    Height,
}

/// Setter key for the fixed-width slide visibility owner.
#[derive(Clone, Copy)]
pub(super) enum SlideSetter {
    Hidden,
}

const fn shape_write_entries<const N: usize>(
    readable: &[PropSpec],
) -> [crate::registry::Entry<ShapeSetter>; N] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};

    let mut text_spec = PropSpec::new("text", ValueType::Text);
    text_spec.access.get = false;
    let text = Entry {
        spec: text_spec,
        access: WriteAccess::Set(ShapeSetter::Text),
        constraint: Constraint::None,
        clear: true,
    };
    let mut entries = [text; N];
    let mut index = 0;
    while index < readable.len() {
        entries[index] = Entry::read(readable[index]);
        let setter = if same_name(readable[index].name, "x") {
            Some(ShapeSetter::X)
        } else if same_name(readable[index].name, "y") {
            Some(ShapeSetter::Y)
        } else if same_name(readable[index].name, "width") {
            Some(ShapeSetter::Width)
        } else if same_name(readable[index].name, "height") {
            Some(ShapeSetter::Height)
        } else {
            None
        };
        if let Some(setter) = setter {
            entries[index].access = WriteAccess::Set(setter);
            if matches!(setter, ShapeSetter::Width | ShapeSetter::Height) {
                entries[index].constraint = Constraint::Length {
                    min: crate::value::Length::from_emu(0),
                    max: crate::value::Length::from_emu(i64::MAX),
                };
            }
        }
        index += 1;
    }
    entries[N - 1] = text;
    entries
}

const fn same_name(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if !left[index].eq_ignore_ascii_case(&right[index]) {
            return false;
        }
        index += 1;
    }
    true
}

const SHAPE_WRITE_COUNT: usize = SHAPE_PROPS.len() + 1;
const PLACEHOLDER_WRITE_COUNT: usize = PLACEHOLDER_PROPS.len() + 1;

pub(super) const SHAPE_WRITE_TABLE: crate::registry::Table<ShapeSetter, SHAPE_WRITE_COUNT, 0> =
    crate::registry::Table::new(shape_write_entries::<SHAPE_WRITE_COUNT>(SHAPE_PROPS), []);
pub(crate) const SHAPE_CAPABILITY_PROPS: &[PropSpec] = &SHAPE_WRITE_TABLE.readable();

pub(super) const PLACEHOLDER_WRITE_TABLE: crate::registry::Table<
    ShapeSetter,
    PLACEHOLDER_WRITE_COUNT,
    0,
> = crate::registry::Table::new(
    shape_write_entries::<PLACEHOLDER_WRITE_COUNT>(PLACEHOLDER_PROPS),
    [],
);
pub(crate) const PLACEHOLDER_CAPABILITY_PROPS: &[PropSpec] = &PLACEHOLDER_WRITE_TABLE.readable();

const SLIDE_WRITE_COUNT: usize = SLIDE_PROPS.len();
const fn slide_write_entries() -> [crate::registry::Entry<SlideSetter>; SLIDE_WRITE_COUNT] {
    use crate::registry::{Access as WriteAccess, Entry};

    let mut entries = [Entry::read(SLIDE_PROPS[0]); SLIDE_WRITE_COUNT];
    let mut index = 0;
    while index < SLIDE_PROPS.len() {
        entries[index] = Entry::read(SLIDE_PROPS[index]);
        if same_name(SLIDE_PROPS[index].name, "hidden") {
            entries[index].access = WriteAccess::Set(SlideSetter::Hidden);
        }
        index += 1;
    }
    entries
}

pub(super) const SLIDE_WRITE_TABLE: crate::registry::Table<SlideSetter, SLIDE_WRITE_COUNT, 0> =
    crate::registry::Table::new(slide_write_entries(), []);
pub(crate) const SLIDE_CAPABILITY_PROPS: &[PropSpec] = &SLIDE_WRITE_TABLE.readable();

/// Setter key of the text-only notes and paragraph registries.
#[derive(Clone, Copy)]
pub(super) enum TextSetter {
    Text,
}

/// The one settable `text` entry of a notes or paragraph registry; `get`
/// follows whether the element's `format` lists the text.
const fn text_write_entries(readable: bool) -> [crate::registry::Entry<TextSetter>; 1] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};

    let mut spec = PropSpec::new("text", ValueType::Text);
    spec.access.get = readable;
    [Entry {
        spec,
        access: WriteAccess::Set(TextSetter::Text),
        constraint: Constraint::None,
        clear: true,
    }]
}

const _: () = assert!(NOTES_PROPS.len() == 1 && same_name(NOTES_PROPS[0].name, "text"));

pub(super) const NOTES_WRITE_TABLE: crate::registry::Table<TextSetter, 1, 0> =
    crate::registry::Table::new(text_write_entries(true), []);
pub(crate) const NOTES_CAPABILITY_PROPS: &[PropSpec] = &NOTES_WRITE_TABLE.readable();

pub(super) const PARAGRAPH_WRITE_TABLE: crate::registry::Table<TextSetter, 1, 0> =
    crate::registry::Table::new(text_write_entries(false), []);
pub(crate) const PARAGRAPH_CAPABILITY_PROPS: &[PropSpec] = &PARAGRAPH_WRITE_TABLE.readable();

/// The cell registry (record 0323): the write-only `text` key.
pub(super) const CELL_WRITE_TABLE: crate::registry::Table<TextSetter, 1, 0> =
    crate::registry::Table::new(text_write_entries(false), []);
pub(crate) const CELL_CAPABILITY_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

const _: () = assert!(CELL_PROPS.is_empty());

/// Comment assignments (record 0323), as the PPTX comment registry: `text`,
/// `x` and `y` are settable; `author` (required), `initials` and `date` are
/// given when the comment is added; `index` is assigned.
#[derive(Clone, Copy)]
pub(super) enum CommentSetter {
    Text,
    Author,
    Initials,
    Date,
    X,
    Y,
}

const fn comment_write_entries() -> [crate::registry::Entry<CommentSetter>; COMMENT_PROPS.len()] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};

    let mut entries = [Entry::read(COMMENT_PROPS[0]); COMMENT_PROPS.len()];
    let mut index = 0;
    while index < COMMENT_PROPS.len() {
        let spec = COMMENT_PROPS[index];
        let access = if same_name(spec.name, "text") {
            Some(WriteAccess::SetAndAdd(CommentSetter::Text))
        } else if same_name(spec.name, "author") {
            Some(WriteAccess::Add(CommentSetter::Author))
        } else if same_name(spec.name, "initials") {
            Some(WriteAccess::Add(CommentSetter::Initials))
        } else if same_name(spec.name, "date") {
            Some(WriteAccess::Add(CommentSetter::Date))
        } else if same_name(spec.name, "x") {
            Some(WriteAccess::SetAndAdd(CommentSetter::X))
        } else if same_name(spec.name, "y") {
            Some(WriteAccess::SetAndAdd(CommentSetter::Y))
        } else {
            None
        };
        entries[index] = match access {
            Some(access) => Entry {
                spec,
                access,
                constraint: if same_name(spec.name, "author") {
                    Constraint::NotEmpty
                } else {
                    Constraint::None
                },
                clear: false,
            },
            None => Entry::read(spec),
        };
        index += 1;
    }
    entries
}

pub(super) const COMMENT_WRITE_TABLE: crate::registry::Table<
    CommentSetter,
    { COMMENT_PROPS.len() },
    0,
> = crate::registry::Table::new(comment_write_entries(), []);
pub(crate) const COMMENT_CAPABILITY_PROPS: &[PropSpec] = &COMMENT_WRITE_TABLE.readable();

/// Hyperlink assignments (record 0323): `link` (alias `url`) and `tooltip`
/// on `add` and `set`, as the PPTX registry.
#[derive(Clone, Copy)]
pub(super) enum HyperlinkSetter {
    Link,
    Tooltip,
}

const fn hyperlink_write_entries() -> [crate::registry::Entry<HyperlinkSetter>; 2] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};
    [
        Entry {
            spec: HYPERLINK_PROPS[0],
            access: WriteAccess::SetAndAdd(HyperlinkSetter::Link),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: HYPERLINK_PROPS[1],
            access: WriteAccess::SetAndAdd(HyperlinkSetter::Tooltip),
            constraint: Constraint::None,
            clear: true,
        },
    ]
}

pub(super) const HYPERLINK_WRITE_TABLE: crate::registry::Table<HyperlinkSetter, 2, 1> =
    crate::registry::Table::new(
        hyperlink_write_entries(),
        [crate::registry::Alias {
            name: "url",
            entry: 0,
        }],
    );
pub(crate) const HYPERLINK_CAPABILITY_PROPS: &[PropSpec] = &HYPERLINK_WRITE_TABLE.readable();

/// Run assignments (record 0323): the write-only `text` key and the
/// character formatting the binary `TextCFException` holds.
#[derive(Clone, Copy)]
pub(super) enum RunSetter {
    Text,
    Size,
    Bold,
    Italic,
    Underline,
    Color,
}

const RUN_WRITE_COUNT: usize = RUN_PROPS.len() + 1;
const fn run_write_entries() -> [crate::registry::Entry<RunSetter>; RUN_WRITE_COUNT] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};

    let mut entries = [Entry::read(RUN_PROPS[0]); RUN_WRITE_COUNT];
    let mut index = 0;
    while index < RUN_PROPS.len() {
        let spec = RUN_PROPS[index];
        let setter = if same_name(spec.name, "size") {
            Some(RunSetter::Size)
        } else if same_name(spec.name, "bold") {
            Some(RunSetter::Bold)
        } else if same_name(spec.name, "italic") {
            Some(RunSetter::Italic)
        } else if same_name(spec.name, "underline") {
            Some(RunSetter::Underline)
        } else if same_name(spec.name, "color") {
            Some(RunSetter::Color)
        } else {
            None
        };
        entries[index] = match setter {
            Some(setter) => Entry {
                spec,
                access: WriteAccess::Set(setter),
                constraint: Constraint::None,
                clear: false,
            },
            None => Entry::read(spec),
        };
        index += 1;
    }
    let mut text = PropSpec::new("text", ValueType::Text);
    text.access.get = false;
    entries[RUN_PROPS.len()] = Entry {
        spec: text,
        access: WriteAccess::Set(RunSetter::Text),
        constraint: Constraint::None,
        clear: true,
    };
    entries
}

pub(super) const RUN_WRITE_TABLE: crate::registry::Table<RunSetter, RUN_WRITE_COUNT, 0> =
    crate::registry::Table::new(run_write_entries(), []);
pub(crate) const RUN_CAPABILITY_PROPS: &[PropSpec] = &RUN_WRITE_TABLE.readable();
