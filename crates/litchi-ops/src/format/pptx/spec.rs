//! The PPTX element kinds and their spec table (record 0010 §4.6, §7.2).
//!
//! One [`Kind`] per element of the presentation vocabulary the adapter
//! knows; [`SPEC`] holds, at each kind's ordinal, its OfficeCLI type,
//! aliases, structure, identity keys, path templates, readable properties
//! and capability gap. The resolver, the query compiler, the node view and
//! `capabilities()` all read this one table.
//!
//! Property keys and their order are OfficeCLI's (`PowerPointHandler.Query.cs`
//! for the root and slides, `PowerPointHandler.NodeBuilder.cs` for shapes,
//! paragraphs and runs, `PowerPointHandler.Comments.cs` for comments).
//! Only keys the leaf answers are declared; the others are listed in
//! [`SPEC`]'s gaps.

use crate::Code;
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// The element kinds of a PPTX presentation; each one's discriminant is its
/// index in [`SPEC`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub(crate) enum Kind {
    /// The presentation root (`/`).
    Presentation,
    /// A slide (`/slide[N]`).
    Slide,
    /// A `p:sp` shape: a title, placeholder, text box or geometry shape.
    Shape,
    /// A `p:sp` with a `p:ph`, addressed by placeholder rules; resolves to
    /// the shape.
    Placeholder,
    /// A `p:pic`.
    Picture,
    /// A graphic frame holding a table.
    Table,
    /// A table row.
    Row,
    /// A table cell.
    Cell,
    /// A graphic frame holding a chart.
    Chart,
    /// A `p:grpSp`.
    Group,
    /// A `p:cxnSp`.
    Connector,
    /// A graphic frame holding a SmartArt diagram (record 0010 M1-Δ15).
    SmartArt,
    /// A graphic frame holding an OLE object.
    Ole,
    /// A paragraph of a shape's text body.
    Paragraph,
    /// A text run (`a:r`).
    Run,
    /// A line break (`a:br`).
    LineBreak,
    /// A slide's speaker notes.
    Notes,
    /// A legacy slide comment.
    Comment,
    /// A video picture (gap).
    Video,
    /// An audio picture (gap).
    Audio,
    /// A media element (gap).
    Media,
    /// An equation (gap).
    Equation,
    /// A slide or section zoom (gap).
    Zoom,
    /// A 3D model (gap).
    Model3d,
    /// A slide master catalog entry and shape scene.
    SlideMaster,
    /// A slide layout catalog entry and shape scene.
    SlideLayout,
    /// The theme (gap).
    Theme,
    /// A modern threaded comment.
    ModernComment,
    /// A reply to a modern comment.
    Reply,
    /// A hyperlink (gap).
    Hyperlink,
    /// An animation (gap).
    Animation,
    /// A chart series (gap).
    Series,
    /// A chart axis (gap).
    Axis,
    /// A table grid column (gap).
    Col,
    /// A notes slide part (gap).
    NoteSlide,
    /// The notes master (gap).
    NotesMaster,
}

impl Kind {
    /// Every kind, in ordinal order.
    const ALL: [Self; 36] = [
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

    /// The kind's spec entry.
    pub(crate) fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }

    /// Whether the kind is an element of a shape tree: what a slide or a
    /// group lists, each kind with its own counter (record 0010 D5).
    pub(crate) const fn is_tree_kind(self) -> bool {
        matches!(
            self,
            Self::Shape
                | Self::Picture
                | Self::Table
                | Self::Chart
                | Self::Group
                | Self::Connector
                | Self::SmartArt
                | Self::Ole
        )
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::PRESENTATION;
    const ROOT: Self = Self::Presentation;

    fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "presentation" => Self::Presentation,
            "slide" => Self::Slide,
            "shape" => Self::Shape,
            "placeholder" => Self::Placeholder,
            "picture" => Self::Picture,
            "table" => Self::Table,
            "tr" => Self::Row,
            "tc" => Self::Cell,
            "chart" => Self::Chart,
            "group" => Self::Group,
            "connector" => Self::Connector,
            "smartart" => Self::SmartArt,
            "ole" => Self::Ole,
            "paragraph" => Self::Paragraph,
            "run" => Self::Run,
            "linebreak" => Self::LineBreak,
            "notes" => Self::Notes,
            "comment" => Self::Comment,
            "video" => Self::Video,
            "audio" => Self::Audio,
            "media" => Self::Media,
            "equation" => Self::Equation,
            "zoom" => Self::Zoom,
            "model3d" => Self::Model3d,
            "slideMaster" => Self::SlideMaster,
            "slideLayout" => Self::SlideLayout,
            "theme" => Self::Theme,
            "modernComment" => Self::ModernComment,
            "reply" => Self::Reply,
            "hyperlink" => Self::Hyperlink,
            "animation" => Self::Animation,
            "series" => Self::Series,
            "axis" => Self::Axis,
            "col" => Self::Col,
            "noteSlide" => Self::NoteSlide,
            "notesMaster" => Self::NotesMaster,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::Presentation => "presentation",
            Self::Slide => "slide",
            Self::Shape => "shape",
            Self::Placeholder => "placeholder",
            Self::Picture => "picture",
            Self::Table => "table",
            Self::Row => "tr",
            Self::Cell => "tc",
            Self::Chart => "chart",
            Self::Group => "group",
            Self::Connector => "connector",
            Self::SmartArt => "smartart",
            Self::Ole => "ole",
            Self::Paragraph => "paragraph",
            Self::Run => "run",
            Self::LineBreak => "linebreak",
            Self::Notes => "notes",
            Self::Comment => "comment",
            Self::Video => "video",
            Self::Audio => "audio",
            Self::Media => "media",
            Self::Equation => "equation",
            Self::Zoom => "zoom",
            Self::Model3d => "model3d",
            Self::SlideMaster => "slideMaster",
            Self::SlideLayout => "slideLayout",
            Self::Theme => "theme",
            Self::ModernComment => "modernComment",
            Self::Reply => "reply",
            Self::Hyperlink => "hyperlink",
            Self::Animation => "animation",
            Self::Series => "series",
            Self::Axis => "axis",
            Self::Col => "col",
            Self::NoteSlide => "noteSlide",
            Self::NotesMaster => "notesMaster",
        }
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

// Every kind's name maps back to the kind, and every vocabulary element is a
// kind: the adapter models or lists a gap for the whole presentation
// vocabulary.
const _: () = {
    let mut index = 0;
    while index < Kind::ALL.len() {
        assert!(Kind::ALL[index] as usize == index);
        index += 1;
    }
    assert!(vocabulary::PRESENTATION.elements.len() == Kind::ALL.len());
};

// ---------------------------------------------------------------------------
// Value vocabularies
// ---------------------------------------------------------------------------

/// Placeholder types, as OfficeCLI's `FormatPlaceholderType` prints them,
/// with the OOXML tokens and OfficeCLI's `ParsePlaceholderType` spellings as
/// aliases.
pub(crate) static PLACEHOLDER_TYPE: Tokens = Tokens::new(
    "placeholder type",
    &[
        "title", "ctrTitle", "body", "subtitle", "date", "footer", "slidenum", "header", "obj",
        "chart", "table", "clipart", "diagram", "media", "picture", "sldImg",
    ],
    &[
        ("centeredtitle", 1),
        ("centertitle", 1),
        ("content", 2),
        ("ctitle", 1),
        ("datetime", 4),
        ("dgm", 12),
        ("dt", 4),
        ("ftr", 5),
        ("hdr", 7),
        ("object", 8),
        ("pic", 14),
        ("sldnum", 6),
        ("slideimage", 15),
        ("slidenumber", 6),
        ("sub", 3),
        ("subtitlepres", 3),
        ("tbl", 10),
    ],
);

/// Run underline styles: OfficeCLI prints `single` and `double` for `sng`
/// and `dbl`, and every other `ST_TextUnderlineType` token as written.
pub(crate) static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "single",
        "double",
        "words",
        "heavy",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dotDashHeavy",
        "dotDotDash",
        "dotDotDashHeavy",
        "wavy",
        "wavyHeavy",
        "wavyDbl",
        "none",
    ],
    &[("dbl", 1), ("sng", 0)],
);

/// Strikethrough (`ST_TextStrikeType`), as OfficeCLI prints it.
pub(crate) static STRIKE: Tokens = Tokens::new(
    "strike",
    &["single", "double", "none"],
    &[("dblstrike", 1), ("nostrike", 2), ("sngstrike", 0)],
);

/// Capitalization (`ST_TextCapsType`).
pub(crate) static CAPS: Tokens = Tokens::new("cap", &["all", "small", "none"], &[]);

/// Paragraph alignment, as OfficeCLI's `MapTextAlignToFriendly` prints it.
pub(crate) static ALIGN: Tokens = Tokens::new(
    "alignment",
    &[
        "left", "center", "right", "justify", "justLow", "dist", "thaiDist",
    ],
    &[
        ("ctr", 1),
        ("distributed", 5),
        ("just", 3),
        ("l", 0),
        ("r", 2),
    ],
);

/// Reading direction.
pub(crate) static DIRECTION: Tokens = Tokens::new("direction", &["ltr", "rtl"], &[]);

/// Slide-layout type tokens accepted by PresentationML.
pub(crate) static LAYOUT_TYPE: Tokens = Tokens::new(
    "slide layout type",
    &[
        "title",
        "tx",
        "twoColTx",
        "tbl",
        "txAndChart",
        "chartAndTx",
        "dgm",
        "chart",
        "txAndClipArt",
        "clipArtAndTx",
        "titleOnly",
        "blank",
        "txAndObj",
        "objAndTx",
        "objOnly",
        "obj",
        "txAndMedia",
        "mediaAndTx",
        "objOverTx",
        "txOverObj",
        "txAndTwoObj",
        "twoObjAndTx",
        "twoObjOverTx",
        "fourObj",
        "vertTx",
        "clipArtAndVertTx",
        "vertTitleAndTx",
        "vertTitleAndTxOverChart",
        "twoObj",
        "objAndTwoObj",
        "twoObjAndObj",
        "cust",
        "secHead",
        "twoTxTwoObj",
        "objTx",
        "picTx",
    ],
    &[],
);

/// Effective status values of a modern threaded comment or reply.
pub(crate) static MODERN_STATUS: Tokens = Tokens::new(
    "modern comment status",
    &["active", "resolved", "closed"],
    &[],
);

/// Slide size presets, as OfficeCLI's root node names them.
pub(crate) static SLIDE_SIZE: Tokens = Tokens::new(
    "slide size",
    &[
        "widescreen",
        "standard",
        "16:10",
        "a4",
        "a3",
        "letter",
        "b4",
        "b5",
        "35mm",
        "overhead",
        "banner",
        "ledger",
        "custom",
    ],
    &[],
);

// ---------------------------------------------------------------------------
// Length types
// ---------------------------------------------------------------------------

/// Slide sizes: printed points first (`FormatEmuPaired`).
const SLIDE_LENGTH: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::PointsFirst);
/// Table and SmartArt positions: `FormatEmuLossy`.
const LOSSY_POSITION: LengthSpec = LengthSpec::new(LengthContext::Emu, LengthStyle::Lossy);
/// Table and SmartArt extents: `FormatEmuLossy`.
const LOSSY_EXTENT: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::Lossy);
/// Paragraph spacing in points (`FormatPptSpacing`).
const SPACING: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::Points);
/// Paragraph indents and margins in points (`FormatPptIndentPoints`).
const INDENT: LengthSpec = LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

const POSITION: ValueType = ValueType::Length(LengthSpec::POSITION);
const EXTENT: ValueType = ValueType::Length(LengthSpec::EXTENT);
const FONT_SIZE: ValueType = ValueType::Length(LengthSpec::FONT_SIZE);
const DRAWING_COLOR: ValueType = ValueType::Color(ColorContext::Drawing);

/// Declares a property table and a key enum whose discriminants are the
/// table's indices, so a node's values are looked up by index without
/// comparing key text.
macro_rules! props {
    ($(#[$meta:meta])* $table:ident, $keys:ident { $($key:ident = $name:literal : $ty:expr $(; $policy:ident)?,)+ }) => {
        $(#[$meta])*
        pub(crate) const $table: &[PropSpec] = &[$(pptx_prop!($name, $ty $(; $policy)?),)+];

        #[doc = concat!("The keys of [`", stringify!($table), "`], by index.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub(crate) enum $keys {
            $(
                #[doc = concat!("`", $name, "`.")]
                $key,
            )+
        }

        impl $keys {
            const ALL: &'static [Self] = &[$(Self::$key,)+];

            /// The key at `index` of the table.
            pub(crate) fn at(index: usize) -> Option<Self> {
                Self::ALL.get(index).copied()
            }
        }

        const _: () = assert!($keys::ALL.len() == $table.len());
    };
}

/// Declares a property table whose values are dispatched by the node layer
/// through the table's string key.  Catalog and related-part nodes currently
/// have no keyed value reader, so generating an unused enum and `at` helper
/// for them would fail the warning-denied build.
macro_rules! props_values {
    ($(#[$meta:meta])* $table:ident { $($key:ident = $name:literal : $ty:expr $(; $policy:ident)?,)+ }) => {
        $(#[$meta])*
        pub(crate) const $table: &[PropSpec] = &[$(pptx_prop!($name, $ty $(; $policy)?),)+];
    };
}

/// Build a property entry while keeping native metadata available to the
/// explicit raw query surface.  The workbook adapter has the same policy
/// marker; keep this local so the PPTX-only build does not depend on that
/// feature's `PropSpec` extension method.
macro_rules! pptx_prop {
    ($name:literal, $ty:expr) => {
        PropSpec::new($name, $ty)
    };
    ($name:literal, $ty:expr; raw) => {
        PropSpec {
            name: $name,
            ty: $ty,
            access: crate::node::spec::Access::READ,
            raw: true,
        }
    };
}

props! {
    /// The presentation root (`PowerPointHandler.Query.cs` `Get("/")`).
    PRESENTATION_PROPS, RootKey {
        SlideWidth = "slideWidth": ValueType::Length(SLIDE_LENGTH),
        SlideHeight = "slideHeight": ValueType::Length(SLIDE_LENGTH),
        SlideSize = "slideSize": ValueType::Token(&SLIDE_SIZE),
        Title = "title": ValueType::Text,
        Author = "author": ValueType::Text,
        Subject = "subject": ValueType::Text,
        Keywords = "keywords": ValueType::Text,
        Description = "description": ValueType::Text,
        LastModifiedBy = "lastModifiedBy": ValueType::Text,
        RevisionNumber = "revisionNumber": ValueType::Text,
        Category = "category": ValueType::Text,
        Created = "created": ValueType::Text,
        Modified = "modified": ValueType::Text,
        ExtendedTemplate = "extended.template": ValueType::Text,
        ExtendedManager = "extended.manager": ValueType::Text,
        ExtendedCompany = "extended.company": ValueType::Text,
        ExtendedApplication = "extended.application": ValueType::Text,
        ExtendedPages = "extended.pages": ValueType::Int,
        ExtendedWords = "extended.words": ValueType::Int,
        ExtendedCharacters = "extended.characters": ValueType::Int,
        ExtendedTotalTime = "extended.totalTime": ValueType::Int,
    }
}

props! {
    /// A slide (`PowerPointHandler.Query.cs` slide branch).
    SLIDE_PROPS, SlideKey {
        Layout = "layout": ValueType::Text,
        LayoutType = "layoutType": ValueType::Text,
        Name = "name": ValueType::Text,
        Hidden = "hidden": ValueType::Bool,
    }
}

props! {
    /// A `p:sp` (`ShapeToNode`): semantic name, placeholder kind and geometry,
    /// then the first run's character properties and the first paragraph's.
    ///
    /// The producer `cNvPr@id` and placeholder `ph@idx` are native metadata.
    /// They remain usable only through the raw path/property surface and are
    /// intentionally absent from this ordinary property registry.
    SHAPE_PROPS, ShapeKey {
        Name = "name": ValueType::Text,
        PhType = "phType": ValueType::Token(&PLACEHOLDER_TYPE),
        PhIndex = "phIndex": ValueType::Int; raw,
        Alt = "alt": ValueType::Text,
        Id = "id": ValueType::Int; raw,
        IsTitle = "isTitle": ValueType::Bool,
        X = "x": POSITION,
        Y = "y": POSITION,
        Width = "width": EXTENT,
        Height = "height": EXTENT,
        Font = "font": ValueType::Text,
        FontLatin = "font.latin": ValueType::Text,
        FontEa = "font.ea": ValueType::Text,
        FontCs = "font.cs": ValueType::Text,
        Size = "size": FONT_SIZE,
        Bold = "bold": ValueType::Bool,
        Italic = "italic": ValueType::Bool,
        Cap = "cap": ValueType::Token(&CAPS),
        Underline = "underline": ValueType::Token(&UNDERLINE),
        Strike = "strike": ValueType::Token(&STRIKE),
        Highlight = "highlight": DRAWING_COLOR,
        Spacing = "spacing": ValueType::Number,
        Color = "color": DRAWING_COLOR,
        FlipH = "flipH": ValueType::Bool,
        FlipV = "flipV": ValueType::Bool,
        Zorder = "zorder": ValueType::Int,
        Rotation = "rotation": ValueType::Number,
        Align = "align": ValueType::Token(&ALIGN),
        LineSpacing = "lineSpacing": ValueType::Text,
        SpaceBefore = "spaceBefore": ValueType::Length(SPACING),
        SpaceAfter = "spaceAfter": ValueType::Length(SPACING),
        Indent = "indent": ValueType::Length(INDENT),
        MarginLeft = "marginLeft": ValueType::Length(INDENT),
        MarginRight = "marginRight": ValueType::Length(INDENT),
        Direction = "direction": ValueType::Token(&DIRECTION),
    }
}

props! {
    /// A picture, a connector, a chart, an OLE frame or a group: the
    /// properties every shape-tree element has (`PictureToNode`,
    /// `ConnectorToNode`, `ChartToNode`, `BuildGroupNode`).
    /// Producer IDs are native metadata and therefore are kept out of the
    /// ordinary property table; raw callers may address them explicitly.
    FRAME_PROPS, FrameKey {
        Name = "name": ValueType::Text,
        Id = "id": ValueType::Int; raw,
        Alt = "alt": ValueType::Text,
        X = "x": POSITION,
        Y = "y": POSITION,
        Width = "width": EXTENT,
        Height = "height": EXTENT,
        Rotation = "rotation": ValueType::Number,
        FlipH = "flipH": ValueType::Bool,
        FlipV = "flipV": ValueType::Bool,
        Zorder = "zorder": ValueType::Int,
    }
}

props! {
    /// A table (`TableToNode`), whose geometry OfficeCLI prints lossily.
    TABLE_PROPS, TableKey {
        Name = "name": ValueType::Text,
        Id = "id": ValueType::Int; raw,
        Rows = "rows": ValueType::Int,
        Cols = "cols": ValueType::Int,
        Style = "style": ValueType::Text,
        FirstRow = "firstRow": ValueType::Bool,
        LastRow = "lastRow": ValueType::Bool,
        FirstCol = "firstCol": ValueType::Bool,
        LastCol = "lastCol": ValueType::Bool,
        BandedRows = "bandedRows": ValueType::Bool,
        BandedCols = "bandedCols": ValueType::Bool,
        X = "x": ValueType::Length(LOSSY_POSITION),
        Y = "y": ValueType::Length(LOSSY_POSITION),
        Width = "width": ValueType::Length(LOSSY_EXTENT),
        Height = "height": ValueType::Length(LOSSY_EXTENT),
        Zorder = "zorder": ValueType::Int,
    }
}

props! {
    /// A SmartArt frame (`SmartArtToNode`).
    SMARTART_PROPS, SmartArtKey {
        Name = "name": ValueType::Text,
        Id = "id": ValueType::Int; raw,
        X = "x": ValueType::Length(LOSSY_POSITION),
        Y = "y": ValueType::Length(LOSSY_POSITION),
        Width = "width": ValueType::Length(LOSSY_EXTENT),
        Height = "height": ValueType::Length(LOSSY_EXTENT),
        Zorder = "zorder": ValueType::Int,
    }
}

props! {
    /// A paragraph (`ShapeToNode` paragraph children).
    PARAGRAPH_PROPS, ParagraphKey {
        Align = "align": ValueType::Token(&ALIGN),
        Level = "level": ValueType::Int,
        Indent = "indent": ValueType::Length(INDENT),
        MarginLeft = "marginLeft": ValueType::Length(INDENT),
        MarginRight = "marginRight": ValueType::Length(INDENT),
        Direction = "direction": ValueType::Token(&DIRECTION),
        LineSpacing = "lineSpacing": ValueType::Text,
        SpaceBefore = "spaceBefore": ValueType::Length(SPACING),
        SpaceAfter = "spaceAfter": ValueType::Length(SPACING),
        Size = "size": FONT_SIZE,
    }
}

props! {
    /// A run (`RunToNode`).
    RUN_PROPS, RunKey {
        FontLatin = "font.latin": ValueType::Text,
        FontEa = "font.ea": ValueType::Text,
        FontCs = "font.cs": ValueType::Text,
        Size = "size": FONT_SIZE,
        Bold = "bold": ValueType::Bool,
        Italic = "italic": ValueType::Bool,
        Underline = "underline": ValueType::Token(&UNDERLINE),
        Strike = "strike": ValueType::Token(&STRIKE),
        Spacing = "spacing": ValueType::Number,
        Baseline = "baseline": ValueType::Text,
        Color = "color": DRAWING_COLOR,
        Highlight = "highlight": DRAWING_COLOR,
        Cap = "cap": ValueType::Token(&CAPS),
        Kern = "kern": ValueType::Int,
        Lang = "lang": ValueType::Text,
    }
}

props! {
    /// Speaker notes (`PowerPointHandler.Query.cs` notes branch).
    NOTES_PROPS, NotesKey {
        Text = "text": ValueType::Text,
    }
}

props! {
    /// A legacy comment (`CommentToNode`).
    COMMENT_PROPS, CommentKey {
        Text = "text": ValueType::Text,
        Author = "author": ValueType::Text,
        Initials = "initials": ValueType::Text,
        Index = "index": ValueType::Int,
        Date = "date": ValueType::Text,
        X = "x": POSITION,
        Y = "y": POSITION,
    }
}

props_values! {
    /// A slide master catalog entry and its detached shape scene.
    MASTER_PROPS {
        Name = "name": ValueType::Text,
        Preserved = "preserved": ValueType::Bool,
    }
}

props_values! {
    /// A slide layout catalog entry and its detached shape scene.
    LAYOUT_PROPS {
        Name = "name": ValueType::Text,
        Type = "type": ValueType::Token(&LAYOUT_TYPE),
    }
}

props_values! {
    /// A modern threaded comment, with replies as child nodes.
    MODERN_COMMENT_PROPS {
        Text = "text": ValueType::Text,
        Author = "author": ValueType::Text,
        Created = "created": ValueType::Text,
        StartDate = "startDate": ValueType::Text,
        DueDate = "dueDate": ValueType::Text,
        Title = "title": ValueType::Text,
        Status = "status": ValueType::Token(&MODERN_STATUS),
        Progress = "progress": ValueType::Number,
        X = "x": POSITION,
        Y = "y": POSITION,
    }
}

props_values! {
    /// A reply in a modern threaded comment.
    REPLY_PROPS {
        Text = "text": ValueType::Text,
        Author = "author": ValueType::Text,
        Created = "created": ValueType::Text,
        Status = "status": ValueType::Token(&MODERN_STATUS),
    }
}

// ---------------------------------------------------------------------------
// Gaps
// ---------------------------------------------------------------------------

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

const MEDIA: &str =
    "the shape scene does not classify media pictures yet, so every p:pic is a picture";
const ALTERNATE: &str =
    "zoom and 3D-model frames sit in markup-compatibility content the scene reads as its fallback";
const MASTERS: &str = "the presentation theme is not readable on the owner yet";
const CHART_PARTS: &str = "chart frames are not mapped to their chart parts yet";
const NOTES_PARTS: &str =
    "notes slides are read through /slide[N]/notes; their parts have no adapter node";

static VIDEO_GAP: GapSpec = element_gap("video", "L-PPTX-R11", MEDIA);
static AUDIO_GAP: GapSpec = element_gap("audio", "L-PPTX-R11", MEDIA);
static MEDIA_GAP: GapSpec = element_gap("media", "L-PPTX-R11", MEDIA);
static EQUATION_GAP: GapSpec = element_gap(
    "equation",
    "L-PPTX-R11",
    "the text-body reader does not expose Office math yet",
);
static ZOOM_GAP: GapSpec = element_gap("zoom", "M1-extended", ALTERNATE);
static MODEL3D_GAP: GapSpec = element_gap("model3d", "M1-extended", ALTERNATE);
static THEME_GAP: GapSpec = element_gap("theme", "L-PPTX-R4", MASTERS);
static HYPERLINK_GAP: GapSpec = element_gap(
    "hyperlink",
    "M1-extended",
    "hyperlink targets need the slide's relationships, which nodes do not keep",
);
static ANIMATION_GAP: GapSpec = element_gap(
    "animation",
    "M1-extended",
    "the timing tree is not read per shape yet",
);
static SERIES_GAP: GapSpec = element_gap("series", "L-PPTX-R8", CHART_PARTS);
static AXIS_GAP: GapSpec = element_gap("axis", "L-PPTX-R8", CHART_PARTS);
static COLUMN_GAP: GapSpec = element_gap(
    "col",
    "L-PPTX-R8",
    "the table reader does not read the grid columns yet",
);
static NOTE_SLIDE_GAP: GapSpec = element_gap("noteSlide", "M1-extended", NOTES_PARTS);
static NOTES_MASTER_GAP: GapSpec = element_gap("notesMaster", "M1-extended", NOTES_PARTS);

/// Paragraphs inside a table cell.
pub(crate) static CELL_TEXT_GAP: GapSpec = GapSpec {
    subject: Subject::Axis("/slide[N]/table[N]/tr[N]/tc[N]/paragraph"),
    code: Code::UnsupportedCapability,
    follow_up: "L-PPTX-R8",
    reason: "the table reader gives a cell's text, not its text body",
};

/// Paragraphs of the speaker notes.
pub(crate) static NOTES_TEXT_GAP: GapSpec = GapSpec {
    subject: Subject::Axis("/slide[N]/notes/paragraph"),
    code: Code::UnsupportedCapability,
    follow_up: "M1-extended",
    reason: "the notes node reads the notes text, not the notes body's paragraphs",
};

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

const SHAPE_FILL: &str = "the shape record keeps no spPr fill, line or geometry span";
const BODY_PROPS: &str =
    "the text-body reader applies body-property defaults, so local values are unknown";
const LINKS: &str = "hyperlink targets need the slide's relationships, which nodes do not keep";

// ---------------------------------------------------------------------------
// The table
// ---------------------------------------------------------------------------

use Kind::{
    Animation, Audio, Axis, Cell, Chart, Col, Comment, Connector, Equation, Group, Hyperlink,
    LineBreak, Media, Model3d, ModernComment, NoteSlide, Notes, NotesMaster, Ole, Paragraph,
    Picture, Placeholder, Presentation, Reply, Row, Run, Series, Shape, Slide, SlideLayout,
    SlideMaster, SmartArt, Table, Theme, Video, Zoom,
};

/// What a slide or a group may hold.
const TREE: &[Kind] = &[
    Shape,
    Placeholder,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    SmartArt,
    Ole,
    Video,
    Audio,
    Media,
];

/// Shape-tree children of a master, which also include its layout axis.
const MASTER_CHILDREN: &[Kind] = &[
    SlideLayout,
    Shape,
    Placeholder,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    SmartArt,
    Ole,
    Video,
    Audio,
    Media,
];

/// Shape-tree children of a layout.
const LAYOUT_CHILDREN: &[Kind] = &[
    Shape,
    Placeholder,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    SmartArt,
    Ole,
    Video,
    Audio,
    Media,
];

/// The containers of a shape-tree element.
const TREE_PARENTS: &[Kind] = &[Slide, Group, SlideMaster, SlideLayout];

/// The identities of a shape-tree element. `name` is the ordinary semantic
/// selector; `id` remains registered so an explicitly raw path can resolve
/// producer metadata, but it is filtered from ordinary capability output.
const SHAPE_IDS: &[&str] = &["id", "name"];

/// One element with no gap.
#[allow(
    clippy::too_many_arguments,
    reason = "a table row of the spec; named fields would triple the table's length"
)]
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

/// One element the adapter lists as a gap.
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

/// The PPTX spec table (record 0010 §7.2).
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Presentation,
    elements: &[
        element(
            Presentation,
            "presentation",
            &["presentation"],
            &[],
            &[Slide, SlideMaster, Theme, NotesMaster],
            &[],
            &["/"],
            PRESENTATION_PROPS,
        ),
        element(
            Slide,
            "slide",
            &["slide"],
            &[Presentation],
            &[
                Shape,
                Placeholder,
                Picture,
                Table,
                Chart,
                Group,
                Connector,
                SmartArt,
                Ole,
                Video,
                Audio,
                Media,
                Notes,
                Comment,
                Equation,
                Zoom,
                Model3d,
                ModernComment,
                NoteSlide,
            ],
            &[],
            &["/slide[N]"],
            SLIDE_PROPS,
        ),
        element(
            Shape,
            "shape",
            &["shape", "textbox", "title"],
            &[Slide, Group, SlideMaster, SlideLayout],
            &[Paragraph, Equation, Hyperlink, Animation],
            SHAPE_IDS,
            &[
                "/slide[N]/shape[@id=ID]",
                "/slide[N]/shape[N]",
                "/slide[N]/shape[@name=NAME]",
                "/slide[N]/group[@id=ID]/shape[@id=ID]",
            ],
            SHAPE_PROPS,
        ),
        element(
            Placeholder,
            "placeholder",
            &["ph", "placeholder"],
            TREE_PARENTS,
            &[Paragraph],
            &["type", "idx", "id", "name"],
            &[
                "/slide[N]/placeholder[N]",
                "/slide[N]/placeholder[TYPE]",
                "/slide[N]/placeholder[@type=TYPE][@idx=K]",
            ],
            SHAPE_PROPS,
        ),
        element(
            Picture,
            "picture",
            &["image", "img", "pic", "picture"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/picture[@id=ID]", "/slide[N]/picture[N]"],
            FRAME_PROPS,
        ),
        element(
            Table,
            "table",
            &["table"],
            TREE_PARENTS,
            &[Row, Col],
            SHAPE_IDS,
            &["/slide[N]/table[@id=ID]", "/slide[N]/table[N]"],
            TABLE_PROPS,
        ),
        element(
            Row,
            "tr",
            &["row", "tr"],
            &[Table],
            &[Cell],
            &[],
            &["/slide[N]/table[N]/tr[N]"],
            &[],
        ),
        element(
            Cell,
            "tc",
            &["cell", "tc"],
            &[Row],
            &[],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            &[],
        ),
        element(
            Chart,
            "chart",
            &["chart"],
            TREE_PARENTS,
            &[Series, Axis, Animation],
            SHAPE_IDS,
            &["/slide[N]/chart[@id=ID]", "/slide[N]/chart[N]"],
            FRAME_PROPS,
        ),
        element(
            Group,
            "group",
            &["group"],
            TREE_PARENTS,
            TREE,
            SHAPE_IDS,
            &["/slide[N]/group[@id=ID]", "/slide[N]/group[N]"],
            FRAME_PROPS,
        ),
        element(
            Connector,
            "connector",
            &["connection", "connector"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/connector[@id=ID]", "/slide[N]/connector[N]"],
            FRAME_PROPS,
        ),
        element(
            SmartArt,
            "smartart",
            &["smartart"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/smartart[@id=ID]", "/slide[N]/smartart[N]"],
            SMARTART_PROPS,
        ),
        element(
            Ole,
            "ole",
            &["embed", "object", "ole", "oleobject"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/ole[@id=ID]", "/slide[N]/ole[N]"],
            FRAME_PROPS,
        ),
        element(
            Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Shape, Placeholder],
            &[Run, LineBreak],
            &[],
            &["/slide[N]/shape[@id=ID]/paragraph[N]"],
            PARAGRAPH_PROPS,
        ),
        element(
            Run,
            "run",
            &["r", "run"],
            &[Paragraph],
            &[Hyperlink],
            &[],
            &[
                "/slide[N]/shape[@id=ID]/paragraph[N]/run[N]",
                "/slide[N]/shape[@id=ID]/run[N]",
            ],
            RUN_PROPS,
        ),
        element(
            LineBreak,
            "linebreak",
            &["br", "linebreak"],
            &[Paragraph],
            &[],
            &[],
            &["/slide[N]/shape[@id=ID]/paragraph[N]/linebreak[N]"],
            &[],
        ),
        element(
            Notes,
            "notes",
            &["notes"],
            &[Slide],
            &[],
            &[],
            &["/slide[N]/notes"],
            NOTES_PROPS,
        ),
        element(
            Comment,
            "comment",
            &["comment"],
            &[Slide],
            &[],
            &[],
            &["/slide[N]/comment[N]"],
            COMMENT_PROPS,
        ),
        gap(
            Video,
            "video",
            &["video"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/video[N]"],
            &VIDEO_GAP,
        ),
        gap(
            Audio,
            "audio",
            &["audio"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/audio[N]"],
            &AUDIO_GAP,
        ),
        gap(
            Media,
            "media",
            &["media"],
            TREE_PARENTS,
            &[],
            &["/slide[N]/media[N]"],
            &MEDIA_GAP,
        ),
        gap(
            Equation,
            "equation",
            &["equation", "formula", "math"],
            &[Slide, Shape],
            &[],
            &["/slide[N]/shape[N]/equation[N]"],
            &EQUATION_GAP,
        ),
        gap(
            Zoom,
            "zoom",
            &["slidezoom", "zoom"],
            &[Slide],
            &[],
            &["/slide[N]/zoom[N]"],
            &ZOOM_GAP,
        ),
        gap(
            Model3d,
            "model3d",
            &["3dmodel", "model3d"],
            &[Slide],
            &[],
            &["/slide[N]/model3d[N]"],
            &MODEL3D_GAP,
        ),
        element(
            SlideMaster,
            "slideMaster",
            &["slidemaster"],
            &[Presentation],
            MASTER_CHILDREN,
            &[],
            &["/slidemaster[N]"],
            MASTER_PROPS,
        ),
        element(
            SlideLayout,
            "slideLayout",
            &["layout", "slidelayout"],
            &[SlideMaster],
            LAYOUT_CHILDREN,
            &[],
            &["/slidemaster[N]/slidelayout[N]", "/slidelayout[N]"],
            LAYOUT_PROPS,
        ),
        gap(
            Theme,
            "theme",
            &["theme"],
            &[Presentation],
            &[],
            &["/theme"],
            &THEME_GAP,
        ),
        element(
            ModernComment,
            "modernComment",
            &[
                "modern-comment",
                "moderncomment",
                "thread",
                "threadedcomment",
            ],
            &[Slide],
            &[Reply],
            &[],
            &["/slide[N]/moderncomment[N]"],
            MODERN_COMMENT_PROPS,
        ),
        element(
            Reply,
            "reply",
            &["reply"],
            &[ModernComment],
            &[],
            &[],
            &["/slide[N]/moderncomment[N]/reply[N]"],
            REPLY_PROPS,
        ),
        gap(
            Hyperlink,
            "hyperlink",
            &["hlink", "hyperlink"],
            &[Shape, Run],
            &[],
            &["/slide[N]/shape[N]/hyperlink"],
            &HYPERLINK_GAP,
        ),
        gap(
            Animation,
            "animation",
            &["animate", "animation"],
            &[Shape, Chart],
            &[],
            &["/slide[N]/shape[N]/animation[N]"],
            &ANIMATION_GAP,
        ),
        gap(
            Series,
            "series",
            &["series"],
            &[Chart],
            &[],
            &["/slide[N]/chart[N]/series[N]"],
            &SERIES_GAP,
        ),
        gap(
            Axis,
            "axis",
            &["axis"],
            &[Chart],
            &[],
            &["/slide[N]/chart[N]/axis[@role=ROLE]"],
            &AXIS_GAP,
        ),
        gap(
            Col,
            "col",
            &["col", "column"],
            &[Table],
            &[],
            &["/slide[N]/table[N]/col[N]"],
            &COLUMN_GAP,
        ),
        gap(
            NoteSlide,
            "noteSlide",
            &["noteslide"],
            &[Slide],
            &[],
            &["/slide[N]/noteSlide"],
            &NOTE_SLIDE_GAP,
        ),
        gap(
            NotesMaster,
            "notesMaster",
            &["notesmaster"],
            &[Presentation],
            &[],
            &["/notesMaster"],
            &NOTES_MASTER_GAP,
        ),
    ],
    gaps: &[
        property_gap(
            "presentation",
            "defaultFont",
            "L-PPTX-R4",
            "the theme is not readable on the source-backed owner yet",
        ),
        property_gap(
            "slide",
            "background",
            "M1-extended",
            "slide backgrounds are not read per slide yet",
        ),
        property_gap(
            "slide",
            "transition",
            "M1-extended",
            "slide transitions are not read per slide yet",
        ),
        property_gap(
            "slide",
            "notes",
            "M1-Δ1",
            "the slide's notes live in another part; read /slide[N]/notes",
        ),
        property_gap("shape", "fill", "L-PPTX-R6b", SHAPE_FILL),
        property_gap("shape", "opacity", "L-PPTX-R6b", SHAPE_FILL),
        property_gap("shape", "geometry", "L-PPTX-R6b", SHAPE_FILL),
        property_gap("shape", "line", "L-PPTX-R6b", SHAPE_FILL),
        property_gap("shape", "lineWidth", "L-PPTX-R6b", SHAPE_FILL),
        property_gap("shape", "valign", "L-PPTX-R7b", BODY_PROPS),
        property_gap("shape", "margin", "L-PPTX-R7b", BODY_PROPS),
        property_gap("shape", "wrap", "L-PPTX-R7b", BODY_PROPS),
        property_gap("shape", "autoFit", "L-PPTX-R7b", BODY_PROPS),
        property_gap("shape", "textDirection", "L-PPTX-R7b", BODY_PROPS),
        property_gap("shape", "link", "M1-extended", LINKS),
        property_gap("shape", "tooltip", "M1-extended", LINKS),
        property_gap(
            "shape",
            "list",
            "M1-extended",
            "bullets are read but not yet named in OfficeCLI's list vocabulary",
        ),
        property_gap(
            "paragraph",
            "list",
            "M1-extended",
            "bullets are read but not yet named in OfficeCLI's list vocabulary",
        ),
        property_gap("run", "link", "M1-extended", LINKS),
        property_gap("run", "tooltip", "M1-extended", LINKS),
        property_gap(
            "table",
            "colWidths",
            "L-PPTX-R8",
            "the table reader does not read the grid columns yet",
        ),
        property_gap(
            "tr",
            "height",
            "L-PPTX-R8",
            "the table reader does not read row properties yet",
        ),
        property_gap(
            "tc",
            "colspan",
            "L-PPTX-R8",
            "the table reader does not read cell properties yet",
        ),
        property_gap(
            "tc",
            "rowspan",
            "L-PPTX-R8",
            "the table reader does not read cell properties yet",
        ),
        property_gap(
            "tc",
            "fill",
            "L-PPTX-R8",
            "the table reader does not read cell properties yet",
        ),
        property_gap(
            "chart",
            "chartType",
            "L-PPTX-R8",
            "chart frames are not mapped to their chart parts yet",
        ),
        property_gap(
            "picture",
            "mediaType",
            "L-PPTX-R11",
            "the shape scene does not classify media pictures yet",
        ),
    ],
};

check_format_spec!(SPEC);

#[derive(Clone, Copy)]
pub(super) enum ShapeSetter {
    Text,
    X,
    Y,
    Width,
    Height,
}
const SHAPE_WRITE_COUNT: usize = SHAPE_PROPS.len() + 1;
const fn shape_write_entries() -> [crate::registry::Entry<ShapeSetter>; SHAPE_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut text_spec = PropSpec::new("text", ValueType::Text);
    text_spec.access.get = false;
    let text = Entry {
        spec: text_spec,
        access: Access::SetAndAdd(ShapeSetter::Text),
        constraint: Constraint::None,
        clear: true,
    };
    let mut entries = [text; SHAPE_WRITE_COUNT];
    let mut index = 0;
    while index < SHAPE_PROPS.len() {
        entries[index] = Entry::read(SHAPE_PROPS[index]);
        let setter = if index == ShapeKey::X as usize {
            Some(ShapeSetter::X)
        } else if index == ShapeKey::Y as usize {
            Some(ShapeSetter::Y)
        } else if index == ShapeKey::Width as usize {
            Some(ShapeSetter::Width)
        } else if index == ShapeKey::Height as usize {
            Some(ShapeSetter::Height)
        } else {
            None
        };
        if let Some(setter) = setter {
            entries[index].access = Access::SetAndAdd(setter);
            if index == ShapeKey::Width as usize || index == ShapeKey::Height as usize {
                entries[index].constraint = Constraint::Length {
                    min: crate::value::Length::from_emu(1),
                    max: crate::value::Length::from_emu(i64::MAX),
                };
            }
        }
        index += 1;
    }
    entries
}
pub(super) const SHAPE_WRITE_TABLE: crate::registry::Table<ShapeSetter, SHAPE_WRITE_COUNT, 0> =
    crate::registry::Table::new(shape_write_entries(), []);
pub(crate) const SHAPE_CAPABILITY_PROPS: &[PropSpec] = &SHAPE_WRITE_TABLE.readable();

const fn placeholder_capability_props() -> [PropSpec; SHAPE_WRITE_COUNT] {
    let mut props = SHAPE_WRITE_TABLE.readable();
    let mut index = 0;
    while index < props.len() {
        let mut access = props[index].access;
        access.add = false;
        props[index].access = access;
        index += 1;
    }
    props
}
const PLACEHOLDER_CAPABILITY_PROPS_ARRAY: [PropSpec; SHAPE_WRITE_COUNT] =
    placeholder_capability_props();
pub(crate) const PLACEHOLDER_CAPABILITY_PROPS: &[PropSpec] = &PLACEHOLDER_CAPABILITY_PROPS_ARRAY;

#[derive(Clone, Copy)]
pub(super) enum SlideProperty {
    Layout,
    Name,
    Hidden,
}
const fn slide_write_entries() -> [crate::registry::Entry<SlideProperty>; SLIDE_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(SLIDE_PROPS[0]); SLIDE_PROPS.len()];
    let mut index = 0;
    while index < SLIDE_PROPS.len() {
        entries[index] = Entry::read(SLIDE_PROPS[index]);
        index += 1;
    }
    entries[0] = Entry {
        spec: SLIDE_PROPS[0],
        access: Access::Add(SlideProperty::Layout),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries[2] = Entry {
        spec: SLIDE_PROPS[2],
        access: Access::SetAndAdd(SlideProperty::Name),
        constraint: Constraint::None,
        clear: true,
    };
    entries[3] = Entry {
        spec: SLIDE_PROPS[3],
        access: Access::SetAndAdd(SlideProperty::Hidden),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const SLIDE_WRITE_TABLE: crate::registry::Table<
    SlideProperty,
    { SLIDE_PROPS.len() },
    0,
> = crate::registry::Table::new(slide_write_entries(), []);
pub(crate) const SLIDE_CAPABILITY_PROPS: &[PropSpec] = &SLIDE_WRITE_TABLE.readable();

/// The presentation root with its settable document-property keys (record
/// 0253).
pub(crate) const PRESENTATION_CAPABILITY_PROPS: &[PropSpec] =
    &crate::format::metadata::root_capabilities::<{ PRESENTATION_PROPS.len() }>(PRESENTATION_PROPS);
