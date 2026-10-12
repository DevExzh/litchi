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
use crate::value::{LengthSpec, ValueType};

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
];

const SHAPE_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
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
    PropSpec::new("rows", ValueType::Int),
    PropSpec::new("columns", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

const GROUP_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
    length("x", LengthSpec::POSITION),
    length("y", LengthSpec::POSITION),
    length("width", LengthSpec::EXTENT),
    length("height", LengthSpec::EXTENT),
];

const CONNECTOR_PROPS: &[PropSpec] = &[
    raw("id", ValueType::Int),
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

const SHAPE_IDS: &[&str] = &["id"];

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

static ROW_GAP: GapSpec = element_gap(
    "tr",
    "M1-extended",
    "the legacy PPT table projection exposes dimensions and cell text but no row nodes",
);
static CELL_GAP: GapSpec = element_gap(
    "tc",
    "M1-extended",
    "the legacy PPT table projection exposes cell text through the table node only",
);
static CHART_GAP: GapSpec = element_gap(
    "chart",
    "M1-extended",
    "legacy PPT chart frames have no chart-part projection",
);
static SMARTART_GAP: GapSpec = element_gap(
    "smartart",
    "M1-extended",
    "legacy PPT diagram frames have no semantic diagram projection",
);
static OLE_GAP: GapSpec = element_gap(
    "ole",
    "M1-extended",
    "legacy PPT object frames are retained as picture metadata only",
);
static PARAGRAPH_GAP: GapSpec = element_gap(
    "paragraph",
    "M1-extended",
    "legacy PPT exposes aggregate shape text, not paragraph nodes",
);
static RUN_GAP: GapSpec = element_gap(
    "run",
    "M1-extended",
    "legacy PPT exposes aggregate shape text, not text-run nodes",
);
static LINEBREAK_GAP: GapSpec = element_gap(
    "linebreak",
    "M1-extended",
    "legacy PPT exposes normalized shape text, not line-break nodes",
);
static NOTES_GAP: GapSpec = element_gap(
    "notes",
    "M1-extended",
    "speaker notes are outside the initial legacy PPT navigation projection",
);
static COMMENT_GAP: GapSpec = element_gap(
    "comment",
    "M1-extended",
    "legacy PPT comments are parsed by the presentation leaf but not navigable nodes",
);
static VIDEO_GAP: GapSpec = element_gap(
    "video",
    "M1-extended",
    "media frames are retained as native picture metadata only",
);
static AUDIO_GAP: GapSpec = element_gap(
    "audio",
    "M1-extended",
    "media frames are retained as native picture metadata only",
);
static MEDIA_GAP: GapSpec = element_gap(
    "media",
    "M1-extended",
    "media frames are retained as native picture metadata only",
);
static EQUATION_GAP: GapSpec = element_gap(
    "equation",
    "M1-extended",
    "legacy PPT has no equation semantic model",
);
static ZOOM_GAP: GapSpec = element_gap(
    "zoom",
    "M1-extended",
    "legacy PPT has no zoom semantic model",
);
static MODEL3D_GAP: GapSpec = element_gap(
    "model3d",
    "M1-extended",
    "legacy PPT has no 3D-model semantic model",
);
static MASTER_GAP: GapSpec = element_gap(
    "slideMaster",
    "M1-extended",
    "legacy PPT master records are not exposed as navigable nodes",
);
static LAYOUT_GAP: GapSpec = element_gap(
    "slideLayout",
    "M1-extended",
    "legacy PPT layout records are not exposed as navigable nodes",
);
static THEME_GAP: GapSpec = element_gap(
    "theme",
    "M1-extended",
    "legacy PPT has no theme semantic projection",
);
static MODERN_COMMENT_GAP: GapSpec = element_gap(
    "modernComment",
    "M1-extended",
    "legacy PPT has no threaded-comment semantic projection",
);
static REPLY_GAP: GapSpec = element_gap(
    "reply",
    "M1-extended",
    "legacy PPT has no threaded-comment semantic projection",
);
static HYPERLINK_GAP: GapSpec = element_gap(
    "hyperlink",
    "M1-extended",
    "legacy PPT hyperlink targets require relationship resolution",
);
static ANIMATION_GAP: GapSpec = element_gap(
    "animation",
    "M1-extended",
    "legacy PPT timing records are not exposed as nodes",
);
static SERIES_GAP: GapSpec = element_gap(
    "series",
    "M1-extended",
    "legacy PPT chart series are not exposed as nodes",
);
static AXIS_GAP: GapSpec = element_gap(
    "axis",
    "M1-extended",
    "legacy PPT chart axes are not exposed as nodes",
);
static COL_GAP: GapSpec = element_gap(
    "col",
    "M1-extended",
    "legacy PPT table columns have no separate node projection",
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

static PICTURE_ALT_GAP: GapSpec = property_gap(
    "picture",
    "alt",
    "M1-extended",
    "the native PPT picture projection does not expose alternate text",
);
static PICTURE_MEDIA_GAP: GapSpec = property_gap(
    "picture",
    "mediaType",
    "M1-extended",
    "legacy PPT frame classification is exposed as frameKind",
);
static CHART_TYPE_GAP: GapSpec = property_gap(
    "chart",
    "chartType",
    "M1-extended",
    "legacy PPT chart frames have no chart-part projection",
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
static STATS_PICTURE_ALT_GAP: GapSpec = GapSpec {
    subject: Subject::View("stats.picturesWithoutAlt"),
    code: Code::UnsupportedCapability,
    follow_up: "M1-extended",
    reason: "legacy PPT picture frames have no alternate-text field for a picturesWithoutAlt count",
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
            &["type", "idx", "id"],
            &["/slide[N]/placeholder[N]", "/slide[N]/placeholder[TYPE]"],
            PLACEHOLDER_PROPS,
        ),
        element(
            Kind::Picture,
            "picture",
            &["image", "img", "pic", "picture"],
            TREE_PARENTS,
            &[],
            &["id", "name"],
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
        gap(
            Kind::Row,
            "tr",
            &["row", "tr"],
            &[Kind::Table],
            &[Kind::Cell],
            &["/slide[N]/table[N]/tr[N]"],
            &ROW_GAP,
        ),
        gap(
            Kind::Cell,
            "tc",
            &["cell", "tc"],
            &[Kind::Row],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            &CELL_GAP,
        ),
        gap(
            Kind::Chart,
            "chart",
            &["chart"],
            TREE_PARENTS,
            &[Kind::Series, Kind::Axis, Kind::Animation],
            &["/slide[N]/chart[N]"],
            &CHART_GAP,
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
        gap(
            Kind::Ole,
            "ole",
            &["embed", "object", "ole", "oleobject"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/ole[N]"],
            &OLE_GAP,
        ),
        gap(
            Kind::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Kind::Shape, Kind::Placeholder],
            &[Kind::Run, Kind::LineBreak],
            &["/slide[N]/shape[N]/paragraph[N]"],
            &PARAGRAPH_GAP,
        ),
        gap(
            Kind::Run,
            "run",
            &["r", "run"],
            &[Kind::Paragraph],
            &[Kind::Hyperlink],
            &["/slide[N]/shape[N]/run[N]"],
            &RUN_GAP,
        ),
        gap(
            Kind::LineBreak,
            "linebreak",
            &["br", "linebreak"],
            &[Kind::Paragraph],
            &[],
            &["/slide[N]/shape[N]/linebreak[N]"],
            &LINEBREAK_GAP,
        ),
        gap(
            Kind::Notes,
            "notes",
            &["notes"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/notes"],
            &NOTES_GAP,
        ),
        gap(
            Kind::Comment,
            "comment",
            &["comment"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/comment[N]"],
            &COMMENT_GAP,
        ),
        gap(
            Kind::Video,
            "video",
            &["video"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/video[N]"],
            &VIDEO_GAP,
        ),
        gap(
            Kind::Audio,
            "audio",
            &["audio"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/audio[N]"],
            &AUDIO_GAP,
        ),
        gap(
            Kind::Media,
            "media",
            &["media"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/media[N]"],
            &MEDIA_GAP,
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
        gap(
            Kind::Theme,
            "theme",
            &["theme"],
            &[Kind::Presentation],
            &[],
            &["/theme"],
            &THEME_GAP,
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
        gap(
            Kind::Hyperlink,
            "hyperlink",
            &["hlink", "hyperlink"],
            &[Kind::Shape, Kind::Run],
            &[],
            &["/slide[N]/shape[N]/hyperlink"],
            &HYPERLINK_GAP,
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
        gap(
            Kind::Col,
            "col",
            &["col", "column"],
            &[Kind::Table],
            &[],
            &["/slide[N]/table[N]/col[N]"],
            &COL_GAP,
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
        PICTURE_ALT_GAP,
        PICTURE_MEDIA_GAP,
        CHART_TYPE_GAP,
        SHAPE_FILL_GAP,
        STATS_CHART_GAP,
        STATS_PICTURE_ALT_GAP,
    ],
};

check_format_spec!(SPEC);

const _: () = assert!(SPEC.elements.len() <= 128);
