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
    /// A content node of a SmartArt graphic (record 0284).
    SmartArtNode,
}

impl Kind {
    /// Every kind, in ordinal order.
    pub(super) const ALL: [Self; 37] = [
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
        Self::SmartArtNode,
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
            "node" => Self::SmartArtNode,
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
            Self::SmartArtNode => "node",
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
    // OfficeCLI also takes `true` for single and `false` for none.
    &[("dbl", 1), ("false", 17), ("sng", 0), ("true", 0)],
);

/// A table cell's vertical anchor, as OfficeCLI prints it (`top`, `center`,
/// `bottom` for `t`, `ctr`, `b`; record 0333).
pub(crate) static VALIGN: Tokens = Tokens::new(
    "vertical alignment",
    &["top", "center", "bottom"],
    &[("b", 2), ("ctr", 1), ("middle", 1), ("t", 0)],
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
            admission: crate::registry::Admission::NONE,
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
        Keywords = "keywords": ValueType::List,
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
        Link = "link": ValueType::Text,
        Tooltip = "tooltip": ValueType::Text,
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
    /// A chart frame (`ChartToNode`): the frame properties, then what its
    /// chart part says (record 0270). The keys up to `zorder` are
    /// [`FRAME_PROPS`] in the same order.
    CHART_PROPS, ChartKey {
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
        ChartType = "chartType": ValueType::Text,
        Title = "title": ValueType::Text,
        Legend = "legend": ValueType::Text,
        Categories = "categories": ValueType::Text,
        CatTitle = "catTitle": ValueType::Text,
    }
}

props! {
    /// A chart series (`chart-series`): its name, values, categories and
    /// fill color.
    SERIES_PROPS, SeriesKey {
        Name = "name": ValueType::Text,
        Values = "values": ValueType::Text,
        Categories = "categories": ValueType::Text,
        Color = "color": ValueType::Text,
    }
}

props! {
    /// A chart axis (`chart-axis`): its role, title and the range of a
    /// value axis.
    AXIS_PROPS, AxisKey {
        Role = "role": ValueType::Text,
        Title = "title": ValueType::Text,
        Min = "min": ValueType::Number,
        Max = "max": ValueType::Number,
    }
}

const _: () = assert!(ChartKey::Zorder as usize + 1 == FRAME_PROPS.len());

props! {
    /// A picture or a group (`PictureToNode`, `BuildGroupNode`): the frame
    /// properties, then its click action (record 0325). The keys up to
    /// `zorder` are [`FRAME_PROPS`] in the same order.
    LINKED_FRAME_PROPS, LinkedFrameKey {
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
        Link = "link": ValueType::Text,
        Tooltip = "tooltip": ValueType::Text,
    }
}

const _: () = assert!(LinkedFrameKey::Zorder as usize + 1 == FRAME_PROPS.len());

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
    /// A SmartArt frame (`SmartArtToNode`), then what its data model says
    /// (record 0284): the layout token (`list`, `process`, ... for the
    /// layouts Litchi authors, else the last segment of the layout's unique
    /// identifier), the layout's unique identifier and the node count.
    SMARTART_PROPS, SmartArtKey {
        Name = "name": ValueType::Text,
        Id = "id": ValueType::Int; raw,
        X = "x": ValueType::Length(LOSSY_POSITION),
        Y = "y": ValueType::Length(LOSSY_POSITION),
        Width = "width": ValueType::Length(LOSSY_EXTENT),
        Height = "height": ValueType::Length(LOSSY_EXTENT),
        Zorder = "zorder": ValueType::Int,
        Layout = "layout": ValueType::Text,
        LayoutId = "layoutId": ValueType::Text,
        NodeCount = "nodeCount": ValueType::Int,
    }
}

props! {
    /// A SmartArt content node (record 0284): its one-based nesting level.
    NODE_PROPS, NodeKey {
        Level = "level": ValueType::Int,
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
        Link = "link": ValueType::Text,
        Tooltip = "tooltip": ValueType::Text,
    }
}

props! {
    /// A table row (`table-row.json`, record 0333): its height.
    ROW_PROPS, RowKey {
        Height = "height": EXTENT,
    }
}

props! {
    /// A table cell (`TableCellToNode`): its local solid fill, vertical
    /// anchor and insets (record 0333).
    CELL_PROPS, CellKey {
        Fill = "fill": DRAWING_COLOR,
        Valign = "valign": ValueType::Token(&VALIGN),
        PaddingLeft = "padding.left": EXTENT,
        PaddingRight = "padding.right": EXTENT,
        PaddingTop = "padding.top": EXTENT,
        PaddingBottom = "padding.bottom": EXTENT,
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

props! {
    /// A modern threaded comment, with replies as child nodes. `initials`,
    /// `resolved` (the thread's `resolved` status) and `parent` (a reply's
    /// thread; never set on a thread, add only) follow OfficeCLI's
    /// `ModernCommentToNode` (record 0346); the native GUID `id` is raw.
    MODERN_COMMENT_PROPS, ModernCommentKey {
        Text = "text": ValueType::Text,
        Author = "author": ValueType::Text,
        Initials = "initials": ValueType::Text,
        Created = "created": ValueType::Text,
        StartDate = "startDate": ValueType::Text,
        DueDate = "dueDate": ValueType::Text,
        Title = "title": ValueType::Text,
        Status = "status": ValueType::Token(&MODERN_STATUS),
        Resolved = "resolved": ValueType::Bool,
        Progress = "progress": ValueType::Number,
        X = "x": POSITION,
        Y = "y": POSITION,
        Parent = "parent": ValueType::Text,
        Id = "id": ValueType::Text; raw,
    }
}

props! {
    /// A reply in a modern threaded comment; `parent` is its thread's path
    /// (record 0346), and the native GUID `id` is raw.
    REPLY_PROPS, ReplyKey {
        Text = "text": ValueType::Text,
        Author = "author": ValueType::Text,
        Initials = "initials": ValueType::Text,
        Created = "created": ValueType::Text,
        Status = "status": ValueType::Token(&MODERN_STATUS),
        Parent = "parent": ValueType::Text,
        Id = "id": ValueType::Text; raw,
    }
}

props! {
    /// An Office Math equation of slide text (`equation.json`, record
    /// 0321): its OMML as a LaTeX-like formula.
    EQUATION_PROPS, EquationKey {
        Formula = "formula": ValueType::Text,
    }
}

props! {
    /// A media picture (`media.json`, record 0321): the frame keys of
    /// [`FRAME_PROPS`], in order, then the PowerPoint 2010 playback trim in
    /// milliseconds.
    MEDIA_PROPS, MediaKey {
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
        TrimStart = "trimStart": ValueType::Text,
        TrimEnd = "trimEnd": ValueType::Text,
    }
}

props! {
    /// A shape's or run's click action (`PowerPointHandler.Query.cs`
    /// hyperlink branch, record 0321).
    HYPERLINK_PROPS, HyperlinkKey {
        Link = "link": ValueType::Text,
        Tooltip = "tooltip": ValueType::Text,
    }
}

props! {
    /// A table grid column (`PowerPointHandler.Query.cs` `col` branch,
    /// record 0321).
    COLUMN_PROPS, ColumnKey {
        Width = "width": EXTENT,
    }
}

props! {
    /// The presentation theme (`PowerPointHandler.Theme.cs` `GetThemeNode`,
    /// record 0321): the color scheme's twelve slots, the major and minor
    /// fonts, and the color scheme's name.
    THEME_PROPS, ThemeKey {
        Dark1 = "dk1": DRAWING_COLOR,
        Light1 = "lt1": DRAWING_COLOR,
        Dark2 = "dk2": DRAWING_COLOR,
        Light2 = "lt2": DRAWING_COLOR,
        Accent1 = "accent1": DRAWING_COLOR,
        Accent2 = "accent2": DRAWING_COLOR,
        Accent3 = "accent3": DRAWING_COLOR,
        Accent4 = "accent4": DRAWING_COLOR,
        Accent5 = "accent5": DRAWING_COLOR,
        Accent6 = "accent6": DRAWING_COLOR,
        Hyperlink = "hyperlink": DRAWING_COLOR,
        FollowedHyperlink = "followedhyperlink": DRAWING_COLOR,
        HeadingFont = "headingFont": ValueType::Text,
        BodyFont = "bodyFont": ValueType::Text,
        HeadingFontEastAsian = "headingFont.ea": ValueType::Text,
        BodyFontEastAsian = "bodyFont.ea": ValueType::Text,
        HeadingFontComplexScript = "headingFont.cs": ValueType::Text,
        BodyFontComplexScript = "bodyFont.cs": ValueType::Text,
        Name = "name": ValueType::Text,
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

const ALTERNATE: &str =
    "zoom and 3D-model frames sit in markup-compatibility content the scene reads as its fallback";
const NOTES_PARTS: &str =
    "notes slides are read through /slide[N]/notes; their parts have no adapter node";

static ZOOM_GAP: GapSpec = element_gap("zoom", "M1-extended", ALTERNATE);
static MODEL3D_GAP: GapSpec = element_gap("model3d", "M1-extended", ALTERNATE);
static ANIMATION_GAP: GapSpec = element_gap(
    "animation",
    "M1-extended",
    "the timing tree is not read per shape yet",
);
static NOTE_SLIDE_GAP: GapSpec = element_gap("noteSlide", "M1-extended", NOTES_PARTS);
static NOTES_MASTER_GAP: GapSpec = element_gap("notesMaster", "M1-extended", NOTES_PARTS);

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

const CROP: &str = "the picture reader does not model the blip fill's a:srcRect yet";
const SHAPE_FILL: &str = "the shape record keeps no spPr fill, line or geometry span";
const BODY_PROPS: &str =
    "the text-body reader applies body-property defaults, so local values are unknown";

// ---------------------------------------------------------------------------
// The table
// ---------------------------------------------------------------------------

use Kind::{
    Animation, Audio, Axis, Cell, Chart, Col, Comment, Connector, Equation, Group, Hyperlink,
    LineBreak, Media, Model3d, ModernComment, NoteSlide, Notes, NotesMaster, Ole, Paragraph,
    Picture, Placeholder, Presentation, Reply, Row, Run, Series, Shape, Slide, SlideLayout,
    SlideMaster, SmartArt, SmartArtNode, Table, Theme, Video, Zoom,
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
            LINKED_FRAME_PROPS,
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
            ROW_PROPS,
        ),
        element(
            Cell,
            "tc",
            &["cell", "tc"],
            &[Row],
            &[Paragraph],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            Chart,
            "chart",
            &["chart"],
            TREE_PARENTS,
            &[Series, Axis, Animation],
            SHAPE_IDS,
            &["/slide[N]/chart[@id=ID]", "/slide[N]/chart[N]"],
            CHART_PROPS,
        ),
        element(
            Group,
            "group",
            &["group"],
            TREE_PARENTS,
            TREE,
            SHAPE_IDS,
            &["/slide[N]/group[@id=ID]", "/slide[N]/group[N]"],
            LINKED_FRAME_PROPS,
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
            &[SmartArtNode],
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
            &[Shape, Placeholder, Cell, Notes],
            &[Run, LineBreak],
            &[],
            &[
                "/slide[N]/shape[@id=ID]/paragraph[N]",
                "/slide[N]/table[N]/tr[N]/tc[N]/paragraph[N]",
                "/slide[N]/notes/paragraph[N]",
            ],
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
            &[Paragraph],
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
        element(
            Video,
            "video",
            &["video"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/video[N]"],
            MEDIA_PROPS,
        ),
        element(
            Audio,
            "audio",
            &["audio"],
            TREE_PARENTS,
            &[],
            SHAPE_IDS,
            &["/slide[N]/audio[N]"],
            MEDIA_PROPS,
        ),
        element(
            Media,
            "media",
            &["media"],
            TREE_PARENTS,
            &[],
            &[],
            &["/slide[N]/media[N]"],
            MEDIA_PROPS,
        ),
        element(
            Equation,
            "equation",
            &["equation", "formula", "math"],
            &[Slide, Shape],
            &[],
            &[],
            &["/slide[N]/shape[N]/equation[N]", "/slide[N]/equation[N]"],
            EQUATION_PROPS,
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
        element(
            Theme,
            "theme",
            &["theme"],
            &[Presentation],
            &[],
            &[],
            &["/theme"],
            THEME_PROPS,
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
        element(
            Hyperlink,
            "hyperlink",
            &["hlink", "hyperlink"],
            &[Shape, Run],
            &[],
            &[],
            &[
                "/slide[N]/shape[N]/hyperlink",
                "/slide[N]/shape[N]/paragraph[N]/run[N]/hyperlink",
            ],
            HYPERLINK_PROPS,
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
        element(
            Series,
            "series",
            &["series"],
            &[Chart],
            &[],
            &["id"],
            &[
                "/slide[N]/chart[N]/series[N]",
                "/slide[N]/chart[N]/series[@id=ID]",
            ],
            SERIES_PROPS,
        ),
        element(
            Axis,
            "axis",
            &["axis"],
            &[Chart],
            &[],
            &["role"],
            &[
                "/slide[N]/chart[N]/axis[@role=ROLE]",
                "/slide[N]/chart[N]/axis[N]",
            ],
            AXIS_PROPS,
        ),
        element(
            Col,
            "col",
            &["col", "column"],
            &[Table],
            &[],
            &[],
            &["/slide[N]/table[N]/col[N]"],
            COLUMN_PROPS,
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
        element(
            SmartArtNode,
            "node",
            &["node"],
            &[SmartArt, SmartArtNode],
            &[SmartArtNode],
            &[],
            &[
                "/slide[N]/smartart[N]/node[N]",
                "/slide[N]/smartart[N]/node[N]/node[N]",
            ],
            NODE_PROPS,
        ),
    ],
    gaps: &[
        // OfficeCLI's Mermaid `diagram` (record 0339).
        GapSpec {
            subject: Subject::Element("diagram"),
            code: Code::UnsupportedPath,
            follow_up: "none: out of scope (record 0332, design §10)",
            reason: "OfficeCLI's diagram renders Mermaid text into shapes or a picture; rendering is outside the unified API, and SmartArt is the smartart element",
        },
        property_gap(
            "presentation",
            "defaultFont",
            "L-PPTX-R4",
            "the root does not read the theme part; read bodyFont of /theme",
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
        property_gap(
            "table",
            "colWidths",
            "L-PPTX-R8",
            "the table node does not list its grid widths; read /slide[N]/table[N]/col[N]",
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
        property_gap("picture", "crop", "0333-F4", CROP),
        property_gap("picture", "cropLeft", "0333-F4", CROP),
        property_gap("picture", "cropTop", "0333-F4", CROP),
        property_gap("picture", "cropRight", "0333-F4", CROP),
        property_gap("picture", "cropBottom", "0333-F4", CROP),
        property_gap(
            "picture",
            "mediaType",
            "L-PPTX-R11",
            "a picture node does not report its media kind; read /slide[N]/video[N] or audio[N]",
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
    /// Rich formatting applied to every paragraph and run of the text body.
    Format(TextSetter),
    /// The click action (`link`, `tooltip`; record 0325), written through the
    /// hyperlink verbs rather than the text edit.
    Click(ClickKey),
}

/// The click-action keys of shapes, runs, pictures and groups (record 0325).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum ClickKey {
    Link,
    Tooltip,
}

/// Rich text assignments for shapes, paragraphs, runs and table cells
/// (`litchi_pptx::opened::TextProperties`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum TextSetter {
    Text,
    Bold,
    Italic,
    Size,
    Color,
    Font,
    FontLatin,
    Align,
    Fill,
    /// A shape's own fill (`p:spPr`), `none` or a colour.
    ShapeFill,
    /// A shape's preset geometry (`a:prstGeom`).
    Geometry,
    /// A run's click action (record 0325), written through the hyperlink
    /// verbs rather than the text edit.
    Click(ClickKey),
    /// `a:rPr@u` (record 0333).
    Underline,
    /// `a:rPr@strike` (record 0333).
    Strike,
    /// A table cell's vertical anchor (record 0333).
    Valign,
    /// One table cell inset (record 0333).
    Margin(litchi_pptx::opened::CellMargin),
    /// Every table cell inset at once (`padding`, record 0333).
    Padding,
}

/// A text write entry that a new paragraph or run also takes (record 0333).
const fn text_add_entry(spec: PropSpec, setter: TextSetter) -> crate::registry::Entry<TextSetter> {
    let mut entry = text_entry(spec, setter);
    entry.access = crate::registry::Access::SetAndAdd(setter);
    entry
}

/// A write entry for `spec` with a text setter; `get` keeps the spec's
/// read flag (write-only entries have none).
const fn text_entry(spec: PropSpec, setter: TextSetter) -> crate::registry::Entry<TextSetter> {
    use crate::registry::{Access, Constraint, Entry};
    Entry {
        spec,
        access: Access::Set(setter),
        constraint: if matches!(setter, TextSetter::Font | TextSetter::FontLatin) {
            Constraint::NotEmpty
        } else {
            Constraint::None
        },
        clear: !matches!(setter, TextSetter::Text),
    }
}

/// A `link` or `tooltip` write entry: settable, and given at creation where
/// the element is added; an empty `link` (or `none`) removes the action.
const fn click_entry<S: Copy>(spec: PropSpec, setter: S) -> crate::registry::Entry<S> {
    use crate::registry::{Access, Constraint, Entry};
    Entry {
        spec,
        access: Access::SetAndAdd(setter),
        constraint: Constraint::None,
        clear: true,
    }
}

/// A run's `link` or `tooltip`: settable only, as runs are not added.
const fn run_click_entry(spec: PropSpec, setter: TextSetter) -> crate::registry::Entry<TextSetter> {
    use crate::registry::{Access, Constraint, Entry};
    Entry {
        spec,
        access: Access::Set(setter),
        constraint: Constraint::None,
        clear: true,
    }
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

const RUN_WRITE_COUNT: usize = RUN_PROPS.len() + 2;
const fn run_write_entries() -> [crate::registry::Entry<TextSetter>; RUN_WRITE_COUNT] {
    use crate::registry::Entry;
    let mut entries = [Entry::read(RUN_PROPS[0]); RUN_WRITE_COUNT];
    let mut index = 0;
    while index < RUN_PROPS.len() {
        let spec = RUN_PROPS[index];
        entries[index] = if index == RunKey::FontLatin as usize {
            text_add_entry(spec, TextSetter::FontLatin)
        } else if index == RunKey::Size as usize {
            text_add_entry(spec, TextSetter::Size)
        } else if index == RunKey::Bold as usize {
            text_add_entry(spec, TextSetter::Bold)
        } else if index == RunKey::Italic as usize {
            text_add_entry(spec, TextSetter::Italic)
        } else if index == RunKey::Color as usize {
            text_add_entry(spec, TextSetter::Color)
        } else if index == RunKey::Underline as usize {
            text_add_entry(spec, TextSetter::Underline)
        } else if index == RunKey::Strike as usize {
            text_add_entry(spec, TextSetter::Strike)
        } else if index == RunKey::Link as usize {
            run_click_entry(spec, TextSetter::Click(ClickKey::Link))
        } else if index == RunKey::Tooltip as usize {
            run_click_entry(spec, TextSetter::Click(ClickKey::Tooltip))
        } else {
            Entry::read(spec)
        };
        index += 1;
    }
    entries[RUN_PROPS.len()] =
        text_add_entry(write_only("text", ValueType::Text), TextSetter::Text);
    entries[RUN_PROPS.len() + 1] =
        text_add_entry(write_only("font", ValueType::Text), TextSetter::Font);
    entries
}
/// OfficeCLI's run aliases of writable properties (`run.json`).
const RUN_WRITE_ALIASES: [crate::registry::Alias; 3] = {
    use crate::registry::Alias;
    let font = RUN_PROPS.len() + 1;
    [
        Alias {
            name: "fontname",
            entry: font,
        },
        Alias {
            name: "fontFamily",
            entry: font,
        },
        Alias {
            name: "font.name",
            entry: font,
        },
    ]
};
pub(super) const RUN_WRITE_TABLE: crate::registry::Table<TextSetter, RUN_WRITE_COUNT, 3> =
    crate::registry::Table::new(run_write_entries(), RUN_WRITE_ALIASES);
pub(crate) const RUN_CAPABILITY_PROPS: &[PropSpec] = &RUN_WRITE_TABLE.readable();

const PARAGRAPH_WRITE_COUNT: usize = PARAGRAPH_PROPS.len() + 7;
const fn paragraph_write_entries() -> [crate::registry::Entry<TextSetter>; PARAGRAPH_WRITE_COUNT] {
    use crate::registry::Entry;
    let mut entries = [Entry::read(PARAGRAPH_PROPS[0]); PARAGRAPH_WRITE_COUNT];
    let mut index = 0;
    while index < PARAGRAPH_PROPS.len() {
        let spec = PARAGRAPH_PROPS[index];
        entries[index] = if index == ParagraphKey::Align as usize {
            text_add_entry(spec, TextSetter::Align)
        } else if index == ParagraphKey::Size as usize {
            text_add_entry(spec, TextSetter::Size)
        } else {
            Entry::read(spec)
        };
        index += 1;
    }
    let base = PARAGRAPH_PROPS.len();
    entries[base] = text_add_entry(write_only("text", ValueType::Text), TextSetter::Text);
    entries[base + 1] = text_add_entry(write_only("bold", ValueType::Bool), TextSetter::Bold);
    entries[base + 2] = text_add_entry(write_only("italic", ValueType::Bool), TextSetter::Italic);
    entries[base + 3] = text_add_entry(write_only("color", DRAWING_COLOR), TextSetter::Color);
    entries[base + 4] = text_add_entry(write_only("font", ValueType::Text), TextSetter::Font);
    entries[base + 5] = text_add_entry(
        write_only("underline", ValueType::Token(&UNDERLINE)),
        TextSetter::Underline,
    );
    entries[base + 6] = text_add_entry(
        write_only("strike", ValueType::Token(&STRIKE)),
        TextSetter::Strike,
    );
    entries
}
/// OfficeCLI's paragraph aliases of writable properties (`paragraph.json`).
const PARAGRAPH_WRITE_ALIASES: [crate::registry::Alias; 2] = [
    crate::registry::Alias {
        name: "alignment",
        entry: ParagraphKey::Align as usize,
    },
    crate::registry::Alias {
        name: "halign",
        entry: ParagraphKey::Align as usize,
    },
];
pub(super) const PARAGRAPH_WRITE_TABLE: crate::registry::Table<
    TextSetter,
    PARAGRAPH_WRITE_COUNT,
    2,
> = crate::registry::Table::new(paragraph_write_entries(), PARAGRAPH_WRITE_ALIASES);
pub(crate) const PARAGRAPH_CAPABILITY_PROPS: &[PropSpec] = &PARAGRAPH_WRITE_TABLE.readable();

const CELL_WRITE_COUNT: usize = CELL_PROPS.len() + 10;
const fn cell_write_entries() -> [crate::registry::Entry<TextSetter>; CELL_WRITE_COUNT] {
    use litchi_pptx::opened::CellMargin;
    let mut entries = [text_entry(CELL_PROPS[0], TextSetter::Fill); CELL_WRITE_COUNT];
    entries[CellKey::Valign as usize] =
        text_entry(CELL_PROPS[CellKey::Valign as usize], TextSetter::Valign);
    entries[CellKey::PaddingLeft as usize] = text_entry(
        CELL_PROPS[CellKey::PaddingLeft as usize],
        TextSetter::Margin(CellMargin::Left),
    );
    entries[CellKey::PaddingRight as usize] = text_entry(
        CELL_PROPS[CellKey::PaddingRight as usize],
        TextSetter::Margin(CellMargin::Right),
    );
    entries[CellKey::PaddingTop as usize] = text_entry(
        CELL_PROPS[CellKey::PaddingTop as usize],
        TextSetter::Margin(CellMargin::Top),
    );
    entries[CellKey::PaddingBottom as usize] = text_entry(
        CELL_PROPS[CellKey::PaddingBottom as usize],
        TextSetter::Margin(CellMargin::Bottom),
    );
    let base = CELL_PROPS.len();
    entries[base] = text_entry(write_only("text", ValueType::Text), TextSetter::Text);
    entries[base + 1] = text_entry(write_only("bold", ValueType::Bool), TextSetter::Bold);
    entries[base + 2] = text_entry(write_only("italic", ValueType::Bool), TextSetter::Italic);
    entries[base + 3] = text_entry(write_only("size", FONT_SIZE), TextSetter::Size);
    entries[base + 4] = text_entry(write_only("color", DRAWING_COLOR), TextSetter::Color);
    entries[base + 5] = text_entry(write_only("font", ValueType::Text), TextSetter::Font);
    entries[base + 6] = text_entry(
        write_only("align", ValueType::Token(&ALIGN)),
        TextSetter::Align,
    );
    entries[base + 7] = text_entry(
        write_only("underline", ValueType::Token(&UNDERLINE)),
        TextSetter::Underline,
    );
    entries[base + 8] = text_entry(
        write_only("strike", ValueType::Token(&STRIKE)),
        TextSetter::Strike,
    );
    entries[base + 9] = text_entry(write_only("padding", EXTENT), TextSetter::Padding);
    entries
}
/// OfficeCLI's table-cell aliases of writable properties
/// (`table-cell.json`, record 0333).
const CELL_WRITE_ALIASES: [crate::registry::Alias; 17] = {
    use crate::registry::Alias;
    const BASE: usize = CELL_PROPS.len();
    const fn alias(name: &'static str, entry: usize) -> Alias {
        Alias { name, entry }
    }
    [
        alias("background", CellKey::Fill as usize),
        alias("shd", CellKey::Fill as usize),
        alias("shading", CellKey::Fill as usize),
        alias("margin.bottom", CellKey::PaddingBottom as usize),
        alias("font.bold", BASE + 1),
        alias("font.italic", BASE + 2),
        alias("font.size", BASE + 3),
        alias("fontsize", BASE + 3),
        alias("font.color", BASE + 4),
        alias("font.name", BASE + 5),
        alias("font.latin", BASE + 5),
        alias("alignment", BASE + 6),
        alias("halign", BASE + 6),
        alias("font.underline", BASE + 7),
        alias("strikethrough", BASE + 8),
        alias("font.strike", BASE + 8),
        alias("margin", BASE + 9),
    ]
};
pub(super) const CELL_WRITE_TABLE: crate::registry::Table<TextSetter, CELL_WRITE_COUNT, 17> =
    crate::registry::Table::new(cell_write_entries(), CELL_WRITE_ALIASES);
pub(crate) const CELL_CAPABILITY_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();
const SHAPE_WRITE_COUNT: usize = SHAPE_PROPS.len() + 3;
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
    // Write-only until the shape record keeps its `p:spPr` fill and geometry
    // (read gap L-PPTX-R6b); writable since record 0269.
    entries[SHAPE_PROPS.len() + 1] = Entry {
        spec: write_only("fill", ValueType::Text),
        access: Access::SetAndAdd(ShapeSetter::Format(TextSetter::ShapeFill)),
        constraint: Constraint::NotEmpty,
        clear: true,
    };
    entries[SHAPE_PROPS.len() + 2] = Entry {
        spec: write_only("geometry", ValueType::Text),
        access: Access::SetAndAdd(ShapeSetter::Format(TextSetter::Geometry)),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
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
        } else if index == ShapeKey::Font as usize {
            Some(ShapeSetter::Format(TextSetter::Font))
        } else if index == ShapeKey::FontLatin as usize {
            Some(ShapeSetter::Format(TextSetter::FontLatin))
        } else if index == ShapeKey::Size as usize {
            Some(ShapeSetter::Format(TextSetter::Size))
        } else if index == ShapeKey::Bold as usize {
            Some(ShapeSetter::Format(TextSetter::Bold))
        } else if index == ShapeKey::Italic as usize {
            Some(ShapeSetter::Format(TextSetter::Italic))
        } else if index == ShapeKey::Color as usize {
            Some(ShapeSetter::Format(TextSetter::Color))
        } else if index == ShapeKey::Align as usize {
            Some(ShapeSetter::Format(TextSetter::Align))
        } else if index == ShapeKey::Underline as usize {
            Some(ShapeSetter::Format(TextSetter::Underline))
        } else if index == ShapeKey::Strike as usize {
            Some(ShapeSetter::Format(TextSetter::Strike))
        } else if index == ShapeKey::Link as usize {
            Some(ShapeSetter::Click(ClickKey::Link))
        } else if index == ShapeKey::Tooltip as usize {
            Some(ShapeSetter::Click(ClickKey::Tooltip))
        } else {
            None
        };
        if let Some(setter) = setter {
            // Text boxes take their formatting at creation too (record 0269):
            // the native add and the body formatting stage in one edit.
            entries[index].access = Access::SetAndAdd(setter);
            if matches!(setter, ShapeSetter::Format(_)) {
                entries[index].clear = true;
                if index == ShapeKey::Font as usize || index == ShapeKey::FontLatin as usize {
                    entries[index].constraint = Constraint::NotEmpty;
                }
            }
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
/// OfficeCLI's shape aliases of writable properties (`shape.json`): the
/// corpus scripts size shapes with `w`/`h` and set `fontSize`.
const SHAPE_WRITE_ALIASES: [crate::registry::Alias; 8] = {
    use crate::registry::Alias;
    [
        Alias {
            name: "left",
            entry: ShapeKey::X as usize,
        },
        Alias {
            name: "top",
            entry: ShapeKey::Y as usize,
        },
        Alias {
            name: "w",
            entry: ShapeKey::Width as usize,
        },
        Alias {
            name: "h",
            entry: ShapeKey::Height as usize,
        },
        // Matched case-insensitively, so this also admits `fontsize`.
        Alias {
            name: "fontSize",
            entry: ShapeKey::Size as usize,
        },
        Alias {
            name: "font.size",
            entry: ShapeKey::Size as usize,
        },
        Alias {
            name: "preset",
            entry: SHAPE_PROPS.len() + 2,
        },
        Alias {
            name: "shape",
            entry: SHAPE_PROPS.len() + 2,
        },
    ]
};
pub(super) const SHAPE_WRITE_TABLE: crate::registry::Table<ShapeSetter, SHAPE_WRITE_COUNT, 8> =
    crate::registry::Table::new(shape_write_entries(), SHAPE_WRITE_ALIASES);
pub(crate) const SHAPE_CAPABILITY_PROPS: &[PropSpec] = &SHAPE_WRITE_TABLE.readable();
/// The schema export's per-property aliases of shapes, paragraphs and runs.
pub(crate) const SHAPE_PROP_ALIASES: &[crate::registry::PropAlias] =
    &SHAPE_WRITE_TABLE.prop_aliases();
pub(crate) const PARAGRAPH_PROP_ALIASES: &[crate::registry::PropAlias] =
    &PARAGRAPH_WRITE_TABLE.prop_aliases();
pub(crate) const RUN_PROP_ALIASES: &[crate::registry::PropAlias] = &RUN_WRITE_TABLE.prop_aliases();

/// The placeholder's capability rows: the shape's, with `add` only for the
/// keys a new placeholder takes, `phType` and `text` (record 0346).
const fn placeholder_capability_props() -> [PropSpec; SHAPE_WRITE_COUNT] {
    let mut props = SHAPE_WRITE_TABLE.readable();
    let mut index = 0;
    while index < props.len() {
        let mut access = props[index].access;
        access.add = const_eq(props[index].name, "phType") || const_eq(props[index].name, "text");
        props[index].access = access;
        index += 1;
    }
    props
}

const fn const_eq(left: &str, right: &str) -> bool {
    let (left, right) = (left.as_bytes(), right.as_bytes());
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

/// A new placeholder's keys (record 0346): its type, required, and text.
#[derive(Clone, Copy)]
pub(super) enum PlaceholderSetter {
    Kind,
    Text,
}
pub(super) const PLACEHOLDER_ADD_TABLE: crate::registry::Table<PlaceholderSetter, 2, 1> =
    crate::registry::Table::new(
        [
            crate::registry::Entry {
                spec: PropSpec::new("phType", ValueType::Token(&PLACEHOLDER_TYPE)),
                access: crate::registry::Access::Add(PlaceholderSetter::Kind),
                constraint: crate::registry::Constraint::None,
                clear: false,
            },
            crate::registry::Entry {
                spec: PropSpec::new("text", ValueType::Text),
                access: crate::registry::Access::Add(PlaceholderSetter::Text),
                constraint: crate::registry::Constraint::None,
                clear: false,
            },
        ],
        [crate::registry::Alias {
            name: "type",
            entry: 0,
        }],
    );
const PLACEHOLDER_CAPABILITY_PROPS_ARRAY: [PropSpec; SHAPE_WRITE_COUNT] =
    placeholder_capability_props();
pub(crate) const PLACEHOLDER_CAPABILITY_PROPS: &[PropSpec] = &PLACEHOLDER_CAPABILITY_PROPS_ARRAY;

#[derive(Clone, Copy)]
pub(super) enum SlideProperty {
    Layout,
    Name,
    Hidden,
    /// Speaker-notes text; set only, routed before slide metadata.
    Notes,
    /// The slide's direct transition, OfficeCLI's shorthand (record 0269).
    Transition,
    /// OfficeCLI's `title` shorthand on add: the title placeholder's text
    /// (record 0346).
    Title,
    /// OfficeCLI's `text` shorthand on add: the body placeholder's text
    /// (record 0346).
    Text,
}
const SLIDE_WRITE_COUNT: usize = SLIDE_PROPS.len() + 4;
const fn slide_write_entries() -> [crate::registry::Entry<SlideProperty>; SLIDE_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(SLIDE_PROPS[0]); SLIDE_WRITE_COUNT];
    let mut index = 0;
    while index < SLIDE_PROPS.len() {
        entries[index] = Entry::read(SLIDE_PROPS[index]);
        index += 1;
    }
    // Settable since record 0346: the slide's layout relationship.
    entries[0] = Entry {
        spec: SLIDE_PROPS[0],
        access: Access::SetAndAdd(SlideProperty::Layout),
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
    entries[SLIDE_PROPS.len()] = Entry {
        spec: write_only("notes", ValueType::Text),
        access: Access::Set(SlideProperty::Notes),
        constraint: Constraint::None,
        clear: true,
    };
    // Write-only until slides read their transition (read gap M1-extended).
    entries[SLIDE_PROPS.len() + 1] = Entry {
        spec: write_only("transition", ValueType::Text),
        // Set only: a source-backed slide insertion refuses markup
        // compatibility in the new slide, which Morph and the PowerPoint
        // 2010+ effects need; set the transition after adding the slide.
        access: Access::Set(SlideProperty::Transition),
        constraint: Constraint::None,
        clear: true,
    };
    // Add only, as OfficeCLI: afterwards the placeholders hold the text.
    entries[SLIDE_PROPS.len() + 2] = Entry {
        spec: write_only("title", ValueType::Text),
        access: Access::Add(SlideProperty::Title),
        constraint: Constraint::None,
        clear: false,
    };
    entries[SLIDE_PROPS.len() + 3] = Entry {
        spec: write_only("text", ValueType::Text),
        access: Access::Add(SlideProperty::Text),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const SLIDE_WRITE_TABLE: crate::registry::Table<SlideProperty, SLIDE_WRITE_COUNT, 0> =
    crate::registry::Table::new(slide_write_entries(), []);

/// Speaker-notes `text` (set and, since record 0346, add; null clears the
/// text).
const fn notes_write_entries() -> [crate::registry::Entry<()>; NOTES_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    [Entry {
        spec: NOTES_PROPS[0],
        access: Access::SetAndAdd(()),
        constraint: Constraint::None,
        clear: true,
    }]
}
const NOTES_WRITE_TABLE: crate::registry::Table<(), { NOTES_PROPS.len() }, 0> =
    crate::registry::Table::new(notes_write_entries(), []);
pub(crate) const NOTES_CAPABILITY_PROPS: &[PropSpec] = &NOTES_WRITE_TABLE.readable();

/// Legacy comment assignments.
#[derive(Clone, Copy)]
pub(super) enum CommentSetter {
    Text,
    Author,
    Initials,
    Date,
    X,
    Y,
    /// The reading direction of the text (`add` only, record 0346):
    /// `rtl` prefixes a right-to-left mark, as OfficeCLI does, since
    /// `p:cm` has no direction attribute.
    Direction,
}
const COMMENT_WRITE_COUNT: usize = COMMENT_PROPS.len() + 1;
const fn comment_write_entries() -> [crate::registry::Entry<CommentSetter>; COMMENT_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(COMMENT_PROPS[0]); COMMENT_WRITE_COUNT];
    let mut index = 0;
    while index < COMMENT_PROPS.len() {
        let spec = COMMENT_PROPS[index];
        // Author, initials and date are settable since record 0346.
        let access = if index == CommentKey::Text as usize {
            Some(Access::SetAndAdd(CommentSetter::Text))
        } else if index == CommentKey::Author as usize {
            Some(Access::SetAndAdd(CommentSetter::Author))
        } else if index == CommentKey::Initials as usize {
            Some(Access::SetAndAdd(CommentSetter::Initials))
        } else if index == CommentKey::Date as usize {
            Some(Access::SetAndAdd(CommentSetter::Date))
        } else if index == CommentKey::X as usize {
            Some(Access::SetAndAdd(CommentSetter::X))
        } else if index == CommentKey::Y as usize {
            Some(Access::SetAndAdd(CommentSetter::Y))
        } else {
            None
        };
        entries[index] = match access {
            Some(access) => Entry {
                spec,
                access,
                constraint: if index == CommentKey::Author as usize {
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
    entries[COMMENT_PROPS.len()] = Entry {
        spec: write_only("direction", ValueType::Token(&DIRECTION)),
        access: Access::Add(CommentSetter::Direction),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const COMMENT_WRITE_TABLE: crate::registry::Table<
    CommentSetter,
    COMMENT_WRITE_COUNT,
    1,
> = crate::registry::Table::new(
    comment_write_entries(),
    [crate::registry::Alias {
        name: "dir",
        entry: COMMENT_PROPS.len(),
    }],
);
pub(crate) const COMMENT_CAPABILITY_PROPS: &[PropSpec] = &COMMENT_WRITE_TABLE.readable();

/// Modern comment and reply assignments (record 0346).
#[derive(Clone, Copy)]
pub(super) enum ModernSetter {
    Text,
    Author,
    Initials,
    Created,
    /// The thread's resolution (threads only).
    Resolved,
    /// The thread a new comment replies to (`add` only).
    Parent,
}
const fn modern_comment_write_entries()
-> [crate::registry::Entry<ModernSetter>; MODERN_COMMENT_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(MODERN_COMMENT_PROPS[0]); MODERN_COMMENT_PROPS.len()];
    let mut index = 0;
    while index < MODERN_COMMENT_PROPS.len() {
        let spec = MODERN_COMMENT_PROPS[index];
        let access = match ModernCommentKey::ALL[index] {
            ModernCommentKey::Text => Some(Access::SetAndAdd(ModernSetter::Text)),
            ModernCommentKey::Author => Some(Access::SetAndAdd(ModernSetter::Author)),
            ModernCommentKey::Initials => Some(Access::SetAndAdd(ModernSetter::Initials)),
            ModernCommentKey::Created => Some(Access::SetAndAdd(ModernSetter::Created)),
            ModernCommentKey::Resolved => Some(Access::SetAndAdd(ModernSetter::Resolved)),
            ModernCommentKey::Parent => Some(Access::Add(ModernSetter::Parent)),
            _ => None,
        };
        entries[index] = match access {
            Some(access) => Entry {
                spec,
                access,
                constraint: if matches!(
                    ModernCommentKey::ALL[index],
                    ModernCommentKey::Author | ModernCommentKey::Parent
                ) {
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
pub(super) const MODERN_COMMENT_WRITE_TABLE: crate::registry::Table<
    ModernSetter,
    { MODERN_COMMENT_PROPS.len() },
    0,
> = crate::registry::Table::new(modern_comment_write_entries(), []);
pub(crate) const MODERN_COMMENT_CAPABILITY_PROPS: &[PropSpec] =
    &MODERN_COMMENT_WRITE_TABLE.readable();

const fn reply_write_entries() -> [crate::registry::Entry<ModernSetter>; REPLY_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(REPLY_PROPS[0]); REPLY_PROPS.len()];
    let mut index = 0;
    while index < REPLY_PROPS.len() {
        let spec = REPLY_PROPS[index];
        let setter = match ReplyKey::ALL[index] {
            ReplyKey::Text => Some(ModernSetter::Text),
            ReplyKey::Author => Some(ModernSetter::Author),
            ReplyKey::Initials => Some(ModernSetter::Initials),
            ReplyKey::Created => Some(ModernSetter::Created),
            _ => None,
        };
        entries[index] = match setter {
            Some(setter) => Entry {
                spec,
                access: Access::Set(setter),
                constraint: if matches!(ReplyKey::ALL[index], ReplyKey::Author) {
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
pub(super) const REPLY_WRITE_TABLE: crate::registry::Table<ModernSetter, { REPLY_PROPS.len() }, 0> =
    crate::registry::Table::new(reply_write_entries(), []);
pub(crate) const REPLY_CAPABILITY_PROPS: &[PropSpec] = &REPLY_WRITE_TABLE.readable();

/// Picture insertion assignments (`add` only).
#[derive(Clone, Copy)]
pub(super) enum PictureSetter {
    Src,
    Name,
    /// The alternative text (`cNvPr@descr`, record 0333).
    Alt,
    X,
    Y,
    Width,
    Height,
    /// The click action (record 0325).
    Click(ClickKey),
}
const PICTURE_WRITE_COUNT: usize = LINKED_FRAME_PROPS.len() + 1;
const fn picture_write_entries() -> [crate::registry::Entry<PictureSetter>; PICTURE_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(LINKED_FRAME_PROPS[0]); PICTURE_WRITE_COUNT];
    let mut index = 0;
    while index < LINKED_FRAME_PROPS.len() {
        let spec = LINKED_FRAME_PROPS[index];
        let setter = if index == LinkedFrameKey::Name as usize {
            Some(PictureSetter::Name)
        } else if index == LinkedFrameKey::Alt as usize {
            Some(PictureSetter::Alt)
        } else if index == LinkedFrameKey::X as usize {
            Some(PictureSetter::X)
        } else if index == LinkedFrameKey::Y as usize {
            Some(PictureSetter::Y)
        } else if index == LinkedFrameKey::Width as usize {
            Some(PictureSetter::Width)
        } else if index == LinkedFrameKey::Height as usize {
            Some(PictureSetter::Height)
        } else {
            None
        };
        entries[index] = if index == LinkedFrameKey::Link as usize {
            click_entry(spec, PictureSetter::Click(ClickKey::Link))
        } else if index == LinkedFrameKey::Tooltip as usize {
            click_entry(spec, PictureSetter::Click(ClickKey::Tooltip))
        } else {
            match setter {
                Some(setter) => Entry {
                    spec,
                    // Pictures are named at creation only, as in OfficeCLI;
                    // their frame, alternative text and image are also set
                    // (record 0333).
                    access: if matches!(setter, PictureSetter::Name) {
                        Access::Add(setter)
                    } else {
                        Access::SetAndAdd(setter)
                    },
                    constraint: if matches!(setter, PictureSetter::Width | PictureSetter::Height) {
                        Constraint::Length {
                            min: crate::value::Length::from_emu(1),
                            max: crate::value::Length::from_emu(i64::MAX),
                        }
                    } else {
                        Constraint::None
                    },
                    clear: matches!(setter, PictureSetter::Alt),
                },
                None => Entry::read(spec),
            }
        };
        index += 1;
    }
    entries[LINKED_FRAME_PROPS.len()] = Entry {
        spec: write_only("src", ValueType::Text),
        access: Access::SetAndAdd(PictureSetter::Src),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries
}
pub(super) const PICTURE_WRITE_TABLE: crate::registry::Table<
    PictureSetter,
    PICTURE_WRITE_COUNT,
    3,
> = crate::registry::Table::new(
    picture_write_entries(),
    [
        crate::registry::Alias {
            name: "path",
            entry: LINKED_FRAME_PROPS.len(),
        },
        crate::registry::Alias {
            name: "altText",
            entry: LinkedFrameKey::Alt as usize,
        },
        crate::registry::Alias {
            name: "description",
            entry: LinkedFrameKey::Alt as usize,
        },
    ],
);
pub(crate) const PICTURE_CAPABILITY_PROPS: &[PropSpec] = &PICTURE_WRITE_TABLE.readable();

/// Group assignments: its click action (record 0325).
const fn group_write_entries() -> [crate::registry::Entry<ClickKey>; LINKED_FRAME_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(LINKED_FRAME_PROPS[0]); LINKED_FRAME_PROPS.len()];
    let mut index = 0;
    while index < LINKED_FRAME_PROPS.len() {
        let spec = LINKED_FRAME_PROPS[index];
        entries[index] = if index == LinkedFrameKey::Link as usize {
            Entry {
                spec,
                access: Access::Set(ClickKey::Link),
                constraint: Constraint::None,
                clear: true,
            }
        } else if index == LinkedFrameKey::Tooltip as usize {
            Entry {
                spec,
                access: Access::Set(ClickKey::Tooltip),
                constraint: Constraint::None,
                clear: true,
            }
        } else {
            Entry::read(spec)
        };
        index += 1;
    }
    entries
}
pub(super) const GROUP_WRITE_TABLE: crate::registry::Table<
    ClickKey,
    { LINKED_FRAME_PROPS.len() },
    0,
> = crate::registry::Table::new(group_write_entries(), []);
pub(crate) const GROUP_CAPABILITY_PROPS: &[PropSpec] = &GROUP_WRITE_TABLE.readable();
pub(crate) const SLIDE_CAPABILITY_PROPS: &[PropSpec] = &SLIDE_WRITE_TABLE.readable();

/// Chart assignments (record 0270): `add` builds a chart, `set` edits the
/// title, legend, categories and category-axis title of an existing one.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum ChartSetter {
    Name,
    X,
    Y,
    Width,
    Height,
    ChartType,
    Title,
    Legend,
    Categories,
    CatTitle,
    Data,
    AxisTitle,
    /// A formatting key ([`crate::format::chart::FORMAT_PROPS`]).
    Format(crate::format::chart::FormatKey),
    /// `series{N}`, by its zero-based number.
    Numbered(u8),
    /// `series{N}.name`, `.values` or `.color`, by the zero-based number
    /// (record 0271).
    Dotted(u8, crate::format::chart::SeriesField),
}
const CHART_FORMAT_START: usize = CHART_PROPS.len() + 2;
const CHART_NUMBERED_START: usize = CHART_FORMAT_START + crate::format::chart::FORMAT_PROPS.len();
const CHART_DOTTED_START: usize =
    CHART_NUMBERED_START + crate::format::chart::NUMBERED_SERIES_PROPS.len();
const CHART_WRITE_COUNT: usize =
    CHART_DOTTED_START + crate::format::chart::DOTTED_SERIES_PROPS.len();
const fn chart_write_entries() -> [crate::registry::Entry<ChartSetter>; CHART_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(CHART_PROPS[0]); CHART_WRITE_COUNT];
    let mut index = 0;
    while index < CHART_PROPS.len() {
        let spec = CHART_PROPS[index];
        let access = if index == ChartKey::Name as usize {
            Some(Access::Add(ChartSetter::Name))
        } else if index == ChartKey::X as usize {
            Some(Access::SetAndAdd(ChartSetter::X))
        } else if index == ChartKey::Y as usize {
            Some(Access::SetAndAdd(ChartSetter::Y))
        } else if index == ChartKey::Width as usize {
            Some(Access::SetAndAdd(ChartSetter::Width))
        } else if index == ChartKey::Height as usize {
            Some(Access::SetAndAdd(ChartSetter::Height))
        } else if index == ChartKey::ChartType as usize {
            Some(Access::SetAndAdd(ChartSetter::ChartType))
        } else if index == ChartKey::Title as usize {
            Some(Access::SetAndAdd(ChartSetter::Title))
        } else if index == ChartKey::Legend as usize {
            Some(Access::SetAndAdd(ChartSetter::Legend))
        } else if index == ChartKey::Categories as usize {
            Some(Access::SetAndAdd(ChartSetter::Categories))
        } else if index == ChartKey::CatTitle as usize {
            Some(Access::SetAndAdd(ChartSetter::CatTitle))
        } else {
            None
        };
        entries[index] = match access {
            Some(access) => Entry {
                spec,
                access,
                constraint: if index == ChartKey::Width as usize
                    || index == ChartKey::Height as usize
                {
                    Constraint::Length {
                        min: crate::value::Length::from_emu(1),
                        max: crate::value::Length::from_emu(i64::MAX),
                    }
                } else {
                    Constraint::None
                },
                clear: index == ChartKey::Title as usize || index == ChartKey::CatTitle as usize,
            },
            None => Entry::read(spec),
        };
        index += 1;
    }
    entries[CHART_PROPS.len()] = Entry {
        spec: write_only("data", ValueType::Text),
        access: Access::SetAndAdd(ChartSetter::Data),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries[CHART_PROPS.len() + 1] = Entry {
        spec: write_only("axisTitle", ValueType::Text),
        access: Access::SetAndAdd(ChartSetter::AxisTitle),
        constraint: Constraint::None,
        clear: true,
    };
    let mut format = 0;
    while format < crate::format::chart::FORMAT_PROPS.len() {
        let (key, spec) = crate::format::chart::FORMAT_PROPS[format];
        entries[CHART_FORMAT_START + format] = Entry {
            spec,
            access: Access::SetAndAdd(ChartSetter::Format(key)),
            constraint: Constraint::None,
            clear: false,
        };
        format += 1;
    }
    let mut number = 0;
    while number < crate::format::chart::NUMBERED_SERIES_PROPS.len() {
        entries[CHART_NUMBERED_START + number] = Entry {
            spec: crate::format::chart::NUMBERED_SERIES_PROPS[number],
            access: Access::Add(ChartSetter::Numbered(number as u8)),
            constraint: Constraint::NotEmpty,
            clear: false,
        };
        number += 1;
    }
    let mut dotted = 0;
    while dotted < crate::format::chart::DOTTED_SERIES_PROPS.len() {
        let field = crate::format::chart::SERIES_FIELDS[dotted % 3];
        let setter = ChartSetter::Dotted((dotted / 3) as u8, field);
        entries[CHART_DOTTED_START + dotted] = Entry {
            spec: crate::format::chart::DOTTED_SERIES_PROPS[dotted],
            // A series' name, values and color are also settable on the
            // chart (OfficeCLI's chart-level `set`).
            access: Access::SetAndAdd(setter),
            constraint: Constraint::NotEmpty,
            clear: false,
        };
        dotted += 1;
    }
    entries
}
const _: () = assert!(
    CHART_WRITE_COUNT <= crate::format::chart::MAX_TABLE,
    "chart adds track assignments in an Assigned set"
);
/// The first formatting key of [`CHART_WRITE_TABLE`].
pub(super) const CHART_FORMAT_INDEX: usize = CHART_FORMAT_START;
pub(super) const CHART_WRITE_TABLE: crate::registry::Table<ChartSetter, CHART_WRITE_COUNT, 3> =
    crate::registry::Table::new(
        chart_write_entries(),
        [
            crate::registry::Alias {
                name: "type",
                entry: ChartKey::ChartType as usize,
            },
            crate::registry::Alias {
                name: "htitle",
                entry: ChartKey::CatTitle as usize,
            },
            crate::registry::Alias {
                name: "vtitle",
                entry: CHART_PROPS.len() + 1,
            },
        ],
    );
pub(crate) const CHART_CAPABILITY_PROPS: &[PropSpec] = &CHART_WRITE_TABLE.readable();

/// Chart-series assignments (record 0270).
#[derive(Clone, Copy)]
pub(super) enum SeriesSetter {
    Name,
    Values,
    Color,
}
const fn series_write_entries() -> [crate::registry::Entry<SeriesSetter>; SERIES_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(SERIES_PROPS[0]); SERIES_PROPS.len()];
    entries[SeriesKey::Name as usize] = Entry {
        spec: SERIES_PROPS[SeriesKey::Name as usize],
        access: Access::Set(SeriesSetter::Name),
        constraint: Constraint::None,
        clear: false,
    };
    entries[SeriesKey::Values as usize] = Entry {
        spec: SERIES_PROPS[SeriesKey::Values as usize],
        access: Access::Set(SeriesSetter::Values),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries[SeriesKey::Categories as usize] =
        Entry::read(SERIES_PROPS[SeriesKey::Categories as usize]);
    entries[SeriesKey::Color as usize] = Entry {
        spec: SERIES_PROPS[SeriesKey::Color as usize],
        access: Access::Set(SeriesSetter::Color),
        constraint: Constraint::NotEmpty,
        clear: true,
    };
    entries
}
pub(super) const SERIES_WRITE_TABLE: crate::registry::Table<
    SeriesSetter,
    { SERIES_PROPS.len() },
    0,
> = crate::registry::Table::new(series_write_entries(), []);
pub(crate) const SERIES_CAPABILITY_PROPS: &[PropSpec] = &SERIES_WRITE_TABLE.readable();

/// Chart-axis assignments (record 0270).
#[derive(Clone, Copy)]
pub(super) enum AxisSetter {
    Title,
    Min,
    Max,
}
const fn axis_write_entries() -> [crate::registry::Entry<AxisSetter>; AXIS_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(AXIS_PROPS[0]); AXIS_PROPS.len()];
    entries[AxisKey::Title as usize] = Entry {
        spec: AXIS_PROPS[AxisKey::Title as usize],
        access: Access::Set(AxisSetter::Title),
        constraint: Constraint::None,
        clear: true,
    };
    entries[AxisKey::Min as usize] = Entry {
        spec: AXIS_PROPS[AxisKey::Min as usize],
        access: Access::Set(AxisSetter::Min),
        constraint: Constraint::None,
        clear: true,
    };
    entries[AxisKey::Max as usize] = Entry {
        spec: AXIS_PROPS[AxisKey::Max as usize],
        access: Access::Set(AxisSetter::Max),
        constraint: Constraint::None,
        clear: true,
    };
    entries
}
pub(super) const AXIS_WRITE_TABLE: crate::registry::Table<AxisSetter, { AXIS_PROPS.len() }, 0> =
    crate::registry::Table::new(axis_write_entries(), []);
pub(crate) const AXIS_CAPABILITY_PROPS: &[PropSpec] = &AXIS_WRITE_TABLE.readable();

/// SmartArt assignments (record 0284): `add` authors a graphic from a text
/// outline with a layout; `set` rebuilds it with another layout or outline.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum SmartArtSetter {
    Name,
    X,
    Y,
    Width,
    Height,
    Layout,
    Text,
}
const SMARTART_WRITE_COUNT: usize = SMARTART_PROPS.len() + 1;
const fn smartart_write_entries() -> [crate::registry::Entry<SmartArtSetter>; SMARTART_WRITE_COUNT]
{
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(SMARTART_PROPS[0]); SMARTART_WRITE_COUNT];
    let mut index = 0;
    while index < SMARTART_PROPS.len() {
        let spec = SMARTART_PROPS[index];
        let access = if index == SmartArtKey::Name as usize {
            Some(Access::Add(SmartArtSetter::Name))
        } else if index == SmartArtKey::X as usize {
            Some(Access::Add(SmartArtSetter::X))
        } else if index == SmartArtKey::Y as usize {
            Some(Access::Add(SmartArtSetter::Y))
        } else if index == SmartArtKey::Width as usize {
            Some(Access::Add(SmartArtSetter::Width))
        } else if index == SmartArtKey::Height as usize {
            Some(Access::Add(SmartArtSetter::Height))
        } else if index == SmartArtKey::Layout as usize {
            Some(Access::SetAndAdd(SmartArtSetter::Layout))
        } else {
            None
        };
        entries[index] = match access {
            Some(access) => Entry {
                spec,
                access,
                constraint: if index == SmartArtKey::Width as usize
                    || index == SmartArtKey::Height as usize
                {
                    Constraint::Length {
                        min: crate::value::Length::from_emu(1),
                        max: crate::value::Length::from_emu(i64::MAX),
                    }
                } else if index == SmartArtKey::Layout as usize {
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
    entries[SMARTART_PROPS.len()] = Entry {
        spec: write_only("text", ValueType::Text),
        access: Access::SetAndAdd(SmartArtSetter::Text),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const SMARTART_WRITE_TABLE: crate::registry::Table<
    SmartArtSetter,
    SMARTART_WRITE_COUNT,
    0,
> = crate::registry::Table::new(smartart_write_entries(), []);
pub(crate) const SMARTART_CAPABILITY_PROPS: &[PropSpec] = &SMARTART_WRITE_TABLE.readable();

/// SmartArt node assignments (record 0284): the node's text.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum NodeSetter {
    Text,
}
const NODE_WRITE_COUNT: usize = NODE_PROPS.len() + 1;
const fn node_write_entries() -> [crate::registry::Entry<NodeSetter>; NODE_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(NODE_PROPS[0]); NODE_WRITE_COUNT];
    entries[NODE_PROPS.len()] = Entry {
        spec: write_only("text", ValueType::Text),
        access: Access::SetAndAdd(NodeSetter::Text),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const NODE_WRITE_TABLE: crate::registry::Table<NodeSetter, NODE_WRITE_COUNT, 0> =
    crate::registry::Table::new(node_write_entries(), []);
pub(crate) const NODE_CAPABILITY_PROPS: &[PropSpec] = &NODE_WRITE_TABLE.readable();

/// Theme assignments (record 0321): every theme property is settable; the
/// setter is the key it writes.
const fn theme_write_entries() -> [crate::registry::Entry<ThemeKey>; THEME_PROPS.len()] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(THEME_PROPS[0]); THEME_PROPS.len()];
    let mut index = 0;
    while index < THEME_PROPS.len() {
        let key = ThemeKey::ALL[index];
        entries[index] = Entry {
            spec: THEME_PROPS[index],
            access: Access::Set(key),
            constraint: if matches!(key, ThemeKey::Name) {
                Constraint::NotEmpty
            } else {
                Constraint::None
            },
            clear: false,
        };
        index += 1;
    }
    entries
}
/// OfficeCLI's theme aliases (`theme.json`).
const THEME_WRITE_ALIASES: [crate::registry::Alias; 14] = {
    use crate::registry::Alias;
    const fn alias(name: &'static str, key: ThemeKey) -> Alias {
        Alias {
            name,
            entry: key as usize,
        }
    }
    [
        alias("dark1", ThemeKey::Dark1),
        alias("light1", ThemeKey::Light1),
        alias("dark2", ThemeKey::Dark2),
        alias("light2", ThemeKey::Light2),
        alias("hlink", ThemeKey::Hyperlink),
        alias("folhlink", ThemeKey::FollowedHyperlink),
        alias("majorFont", ThemeKey::HeadingFont),
        alias("major", ThemeKey::HeadingFont),
        alias("minorFont", ThemeKey::BodyFont),
        alias("minor", ThemeKey::BodyFont),
        alias("majorFont.ea", ThemeKey::HeadingFontEastAsian),
        alias("majorFont.cs", ThemeKey::HeadingFontComplexScript),
        alias("minorFont.ea", ThemeKey::BodyFontEastAsian),
        alias("minorFont.cs", ThemeKey::BodyFontComplexScript),
    ]
};
pub(super) const THEME_WRITE_TABLE: crate::registry::Table<ThemeKey, { THEME_PROPS.len() }, 14> =
    crate::registry::Table::new(theme_write_entries(), THEME_WRITE_ALIASES);
pub(crate) const THEME_CAPABILITY_PROPS: &[PropSpec] = &THEME_WRITE_TABLE.readable();

/// A catalog entry's write table: `name` (index 0 of its read table) is
/// settable (record 0321).
const fn named_write_entries<const N: usize>(
    props: &[PropSpec],
) -> [crate::registry::Entry<()>; N] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(props[0]); N];
    let mut index = 0;
    while index < N {
        entries[index] = Entry::read(props[index]);
        index += 1;
    }
    entries[0] = Entry {
        spec: props[0],
        access: Access::Set(()),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries
}
pub(super) const MASTER_WRITE_TABLE: crate::registry::Table<(), { MASTER_PROPS.len() }, 0> =
    crate::registry::Table::new(named_write_entries(MASTER_PROPS), []);
pub(crate) const MASTER_CAPABILITY_PROPS: &[PropSpec] = &MASTER_WRITE_TABLE.readable();
pub(super) const LAYOUT_WRITE_TABLE: crate::registry::Table<(), { LAYOUT_PROPS.len() }, 0> =
    crate::registry::Table::new(named_write_entries(LAYOUT_PROPS), []);
pub(crate) const LAYOUT_CAPABILITY_PROPS: &[PropSpec] = &LAYOUT_WRITE_TABLE.readable();

/// Table insertion assignments (record 0333): OfficeCLI's `add --type
/// table` keys Litchi authors: `rows`, `cols`, `data`, `name` and the frame.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum TableSetter {
    Name,
    Rows,
    Cols,
    X,
    Y,
    Width,
    Height,
    Data,
    /// An `a:tblPr` style switch, by its position in [`TABLE_SWITCHES`]
    /// (record 0342: set only).
    Switch(u8),
}

/// The table style switches a table-level `set` writes (record 0342), in
/// [`TableSetter::Switch`] order.
pub(super) const TABLE_SWITCHES: [TableKey; 6] = [
    TableKey::FirstRow,
    TableKey::LastRow,
    TableKey::FirstCol,
    TableKey::LastCol,
    TableKey::BandedRows,
    TableKey::BandedCols,
];
const TABLE_WRITE_COUNT: usize = TABLE_PROPS.len() + 1;
const fn table_write_entries() -> [crate::registry::Entry<TableSetter>; TABLE_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(TABLE_PROPS[0]); TABLE_WRITE_COUNT];
    let mut index = 0;
    while index < TABLE_PROPS.len() {
        let spec = TABLE_PROPS[index];
        let setter = if index == TableKey::Name as usize {
            Some(TableSetter::Name)
        } else if index == TableKey::Rows as usize {
            Some(TableSetter::Rows)
        } else if index == TableKey::Cols as usize {
            Some(TableSetter::Cols)
        } else if index == TableKey::X as usize {
            Some(TableSetter::X)
        } else if index == TableKey::Y as usize {
            Some(TableSetter::Y)
        } else if index == TableKey::Width as usize {
            Some(TableSetter::Width)
        } else if index == TableKey::Height as usize {
            Some(TableSetter::Height)
        } else {
            let mut switch = 0;
            let mut found = None;
            while switch < TABLE_SWITCHES.len() {
                if TABLE_SWITCHES[switch] as usize == index {
                    found = Some(TableSetter::Switch(switch as u8));
                }
                switch += 1;
            }
            found
        };
        entries[index] = match setter {
            Some(setter) => Entry {
                spec,
                // Record 0342: `name` is set too, and the style switches
                // are set only.
                access: match setter {
                    TableSetter::Name => Access::SetAndAdd(setter),
                    TableSetter::Switch(_) => Access::Set(setter),
                    _ => Access::Add(setter),
                },
                constraint: match setter {
                    TableSetter::Rows | TableSetter::Cols => Constraint::Int { min: 1, max: 5_000 },
                    TableSetter::Width | TableSetter::Height => Constraint::Length {
                        min: crate::value::Length::from_emu(1),
                        max: crate::value::Length::from_emu(i64::MAX),
                    },
                    TableSetter::Name => Constraint::NotEmpty,
                    _ => Constraint::None,
                },
                clear: false,
            },
            None => Entry::read(spec),
        };
        index += 1;
    }
    entries[TABLE_PROPS.len()] = Entry {
        spec: write_only("data", ValueType::Text),
        access: Access::Add(TableSetter::Data),
        constraint: Constraint::None,
        clear: false,
    };
    entries
}
pub(super) const TABLE_WRITE_TABLE: crate::registry::Table<TableSetter, TABLE_WRITE_COUNT, 2> =
    crate::registry::Table::new(
        table_write_entries(),
        [
            crate::registry::Alias {
                name: "w",
                entry: TableKey::Width as usize,
            },
            crate::registry::Alias {
                name: "h",
                entry: TableKey::Height as usize,
            },
        ],
    );
pub(crate) const TABLE_CAPABILITY_PROPS: &[PropSpec] = &TABLE_WRITE_TABLE.readable();

/// Table row assignments (record 0333): `height` on `add` and `set`; on
/// `add`, OfficeCLI's `cols` check and `c1`…`c20` seed texts.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RowSetter {
    Height,
    Cols,
    /// `c{N}`, by its zero-based column.
    Cell(u8),
}
/// The seed-text keys of a new row.
const ROW_CELL_KEYS: [&str; 20] = [
    "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", "c11", "c12", "c13", "c14", "c15",
    "c16", "c17", "c18", "c19", "c20",
];
const ROW_WRITE_COUNT: usize = ROW_PROPS.len() + 1 + ROW_CELL_KEYS.len();
const fn row_write_entries() -> [crate::registry::Entry<RowSetter>; ROW_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(ROW_PROPS[0]); ROW_WRITE_COUNT];
    entries[RowKey::Height as usize] = Entry {
        spec: ROW_PROPS[RowKey::Height as usize],
        access: Access::SetAndAdd(RowSetter::Height),
        constraint: Constraint::Length {
            min: crate::value::Length::from_emu(0),
            max: crate::value::Length::from_emu(27_273_042_316_900),
        },
        clear: false,
    };
    entries[ROW_PROPS.len()] = Entry {
        spec: write_only("cols", ValueType::Int),
        access: Access::Add(RowSetter::Cols),
        constraint: Constraint::None,
        clear: false,
    };
    let mut index = 0;
    while index < ROW_CELL_KEYS.len() {
        entries[ROW_PROPS.len() + 1 + index] = Entry {
            spec: write_only(ROW_CELL_KEYS[index], ValueType::Text),
            access: Access::Add(RowSetter::Cell(index as u8)),
            constraint: Constraint::None,
            clear: false,
        };
        index += 1;
    }
    entries
}
pub(super) const ROW_WRITE_TABLE: crate::registry::Table<RowSetter, ROW_WRITE_COUNT, 0> =
    crate::registry::Table::new(row_write_entries(), []);
pub(crate) const ROW_CAPABILITY_PROPS: &[PropSpec] = &ROW_WRITE_TABLE.readable();

/// Table column assignments (record 0321): `width` on `add` and `set`, seed
/// `text` on `add` (OfficeCLI's `table-column.json`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum ColumnSetter {
    Width,
    Text,
}
const fn column_write_entries() -> [crate::registry::Entry<ColumnSetter>; 2] {
    use crate::registry::{Access, Constraint, Entry};
    [
        Entry {
            spec: COLUMN_PROPS[0],
            access: Access::SetAndAdd(ColumnSetter::Width),
            constraint: Constraint::Length {
                min: crate::value::Length::from_emu(0),
                max: crate::value::Length::from_emu(27_273_042_316_900),
            },
            clear: false,
        },
        Entry {
            spec: write_only("text", ValueType::Text),
            access: Access::Add(ColumnSetter::Text),
            constraint: Constraint::None,
            clear: false,
        },
    ]
}
pub(super) const COLUMN_WRITE_TABLE: crate::registry::Table<ColumnSetter, 2, 0> =
    crate::registry::Table::new(column_write_entries(), []);
pub(crate) const COLUMN_CAPABILITY_PROPS: &[PropSpec] = &COLUMN_WRITE_TABLE.readable();

/// Hyperlink assignments (record 0321): `link` (alias `url`) and `tooltip`
/// on `add` and `set`.
const fn hyperlink_write_entries() -> [crate::registry::Entry<HyperlinkKey>; 2] {
    use crate::registry::{Access, Constraint, Entry};
    [
        Entry {
            spec: HYPERLINK_PROPS[0],
            access: Access::SetAndAdd(HyperlinkKey::Link),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: HYPERLINK_PROPS[1],
            access: Access::SetAndAdd(HyperlinkKey::Tooltip),
            constraint: Constraint::None,
            clear: true,
        },
    ]
}
pub(super) const HYPERLINK_WRITE_TABLE: crate::registry::Table<HyperlinkKey, 2, 1> =
    crate::registry::Table::new(
        hyperlink_write_entries(),
        [crate::registry::Alias {
            name: "url",
            entry: 0,
        }],
    );
pub(crate) const HYPERLINK_CAPABILITY_PROPS: &[PropSpec] = &HYPERLINK_WRITE_TABLE.readable();

/// Media picture assignments (record 0321): `add` embeds a file (`src`)
/// with its poster image (`poster`), a name and a frame.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum MediaSetter {
    Src,
    Poster,
    Name,
    X,
    Y,
    Width,
    Height,
}
const MEDIA_WRITE_COUNT: usize = MEDIA_PROPS.len() + 2;
const fn media_write_entries() -> [crate::registry::Entry<MediaSetter>; MEDIA_WRITE_COUNT] {
    use crate::registry::{Access, Constraint, Entry};
    let mut entries = [Entry::read(MEDIA_PROPS[0]); MEDIA_WRITE_COUNT];
    let mut index = 0;
    while index < MEDIA_PROPS.len() {
        let spec = MEDIA_PROPS[index];
        let setter = if index == MediaKey::Name as usize {
            Some(MediaSetter::Name)
        } else if index == MediaKey::X as usize {
            Some(MediaSetter::X)
        } else if index == MediaKey::Y as usize {
            Some(MediaSetter::Y)
        } else if index == MediaKey::Width as usize {
            Some(MediaSetter::Width)
        } else if index == MediaKey::Height as usize {
            Some(MediaSetter::Height)
        } else {
            None
        };
        entries[index] = match setter {
            Some(setter) => Entry {
                spec,
                access: Access::Add(setter),
                constraint: if matches!(setter, MediaSetter::Width | MediaSetter::Height) {
                    Constraint::Length {
                        min: crate::value::Length::from_emu(1),
                        max: crate::value::Length::from_emu(i64::MAX),
                    }
                } else {
                    Constraint::None
                },
                clear: false,
            },
            None => Entry::read(spec),
        };
        index += 1;
    }
    entries[MEDIA_PROPS.len()] = Entry {
        spec: write_only("src", ValueType::Text),
        access: Access::Add(MediaSetter::Src),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries[MEDIA_PROPS.len() + 1] = Entry {
        spec: write_only("poster", ValueType::Text),
        access: Access::Add(MediaSetter::Poster),
        constraint: Constraint::NotEmpty,
        clear: false,
    };
    entries
}
pub(super) const MEDIA_WRITE_TABLE: crate::registry::Table<MediaSetter, MEDIA_WRITE_COUNT, 1> =
    crate::registry::Table::new(
        media_write_entries(),
        [crate::registry::Alias {
            name: "path",
            entry: MEDIA_PROPS.len(),
        }],
    );
pub(crate) const MEDIA_CAPABILITY_PROPS: &[PropSpec] = &MEDIA_WRITE_TABLE.readable();
