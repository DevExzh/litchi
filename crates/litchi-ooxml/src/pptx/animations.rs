//! Animation support for PowerPoint presentations.
//!
//! This module provides read/write support for slide animations and timing.

use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::ops::Range;

const MAX_TIMING_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_PRESERVED_TIMING_BYTES: usize = 8 * 1024 * 1024;
const MAX_TIMING_DEPTH: usize = 128;
const MAX_TIMING_NODES: usize = 250_000;
const MAX_TIMING_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TIMING_ATTRIBUTES: usize = 64;
const MAX_ANIMATIONS: usize = 10_000;
const MAX_ANIMATION_BUILDS: usize = 10_000;
const MAX_PARAGRAPH_TEMPLATES: usize = 9;
const MAX_TEMPLATE_TIME_NODE_BYTES: usize = 1024 * 1024;
const MAX_TIME_FILTER_BYTES: usize = 64 * 1024;
const MAX_TIME_FILTER_POINTS: usize = 4_096;
const MAX_NORMALIZED_TIME_DECIMALS: usize = 18;
pub(crate) const MAX_TIMING_MILLISECONDS: u32 = 2_147_483_625;
const DRAWINGML_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWINGML_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const CHART_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const DIAGRAM_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";

/// Animation effect type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationEffect {
    /// Appear effect
    Appear,
    /// Fade effect
    Fade,
    /// Fly in effect
    FlyIn,
    /// Float in effect
    FloatIn,
    /// Split effect
    Split,
    /// Wipe effect
    Wipe,
    /// Zoom effect
    Zoom,
    /// Bounce effect
    Bounce,
    /// Spin effect
    Spin,
    /// Grow/Shrink effect
    GrowShrink,
    /// Custom/Unknown effect
    Custom(String),
}

impl AnimationEffect {
    /// Parse from preset ID.
    pub fn from_preset_id(id: u32) -> Self {
        match id {
            1 => AnimationEffect::Appear,
            2 => AnimationEffect::FlyIn,
            6 => AnimationEffect::GrowShrink,
            8 => AnimationEffect::Spin,
            10 => AnimationEffect::Fade,
            16 => AnimationEffect::Split,
            22 => AnimationEffect::Wipe,
            23 => AnimationEffect::Zoom,
            24 => AnimationEffect::Bounce,
            42 => AnimationEffect::FloatIn,
            _ => AnimationEffect::Custom(format!("preset_{}", id)),
        }
    }

    fn from_preset_parts(class: &str, id: u32) -> Self {
        match (class, id) {
            ("entr", 1) => Self::Appear,
            ("entr", 2) => Self::FlyIn,
            ("entr", 10) => Self::Fade,
            ("entr", 16) => Self::Split,
            ("entr", 22) => Self::Wipe,
            ("entr", 23) => Self::Zoom,
            ("entr", 24) => Self::Bounce,
            ("entr", 42) => Self::FloatIn,
            ("emph", 6) => Self::GrowShrink,
            ("emph", 8) => Self::Spin,
            _ => Self::Custom(format!("{class}:{id}")),
        }
    }

    /// Parse from preset class string (for backwards compatibility).
    pub fn from_preset(preset: &str) -> Self {
        match preset.to_lowercase().as_str() {
            "entr" | "appear" => AnimationEffect::Appear,
            "fade" => AnimationEffect::Fade,
            "fly" | "flyin" => AnimationEffect::FlyIn,
            "float" | "floatin" => AnimationEffect::FloatIn,
            "split" => AnimationEffect::Split,
            "wipe" => AnimationEffect::Wipe,
            "zoom" => AnimationEffect::Zoom,
            "bounce" => AnimationEffect::Bounce,
            "spin" => AnimationEffect::Spin,
            "grow" | "growshrink" => AnimationEffect::GrowShrink,
            other => AnimationEffect::Custom(other.to_string()),
        }
    }

    /// Get the preset ID for this effect.
    /// These are defined in ECMA-376 Part 1.
    pub fn preset_id(&self) -> u32 {
        match self {
            AnimationEffect::Appear => 1,
            AnimationEffect::FlyIn => 2,
            AnimationEffect::FloatIn => 42,
            AnimationEffect::Split => 16,
            AnimationEffect::Fade => 10,
            AnimationEffect::Wipe => 22,
            AnimationEffect::Zoom => 23,
            AnimationEffect::Bounce => 24,
            AnimationEffect::Spin => 8, // Spin is emphasis, but using ID 8
            AnimationEffect::GrowShrink => 6, // GrowShrink is emphasis
            AnimationEffect::Custom(value) => value
                .split_once(':')
                .and_then(|(_, id)| id.parse().ok())
                .unwrap_or(1),
        }
    }

    /// Get the preset class for this effect.
    /// Valid values: "entr" (entrance), "exit", "emph" (emphasis), "path", "verb", "mediacall"
    pub fn preset_class(&self) -> &str {
        match self {
            // Entrance effects
            AnimationEffect::Appear => "entr",
            AnimationEffect::FlyIn => "entr",
            AnimationEffect::FloatIn => "entr",
            AnimationEffect::Split => "entr",
            AnimationEffect::Fade => "entr",
            AnimationEffect::Wipe => "entr",
            AnimationEffect::Zoom => "entr",
            AnimationEffect::Bounce => "entr",
            // Emphasis effects
            AnimationEffect::Spin => "emph",
            AnimationEffect::GrowShrink => "emph",
            // Default to entrance
            AnimationEffect::Custom(value) => value
                .split_once(':')
                .map(|(class, _)| class)
                .filter(|class| {
                    matches!(
                        *class,
                        "entr" | "exit" | "emph" | "path" | "verb" | "mediacall"
                    )
                })
                .unwrap_or("entr"),
        }
    }

    /// Get the preset class string (deprecated, use preset_class instead).
    #[deprecated(note = "Use preset_class() and preset_id() instead")]
    pub fn to_preset(&self) -> &str {
        self.preset_class()
    }
}

/// Animation trigger type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationTrigger {
    /// Start on click
    #[default]
    OnClick,
    /// Start with previous animation
    WithPrevious,
    /// Start after previous animation
    AfterPrevious,
}

/// Unsigned identifier linking a timing node to an entry in `p:bldLst`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AnimationGroupId(u32);

impl AnimationGroupId {
    /// Construct an OOXML timing group identifier.
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the encoded unsigned group identifier.
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl From<u32> for AnimationGroupId {
    fn from(value: u32) -> Self {
        Self::new(value)
    }
}

/// Paragraph build mode from `ST_TLParaBuildType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationParagraphBuildType {
    AllAtOnce,
    Paragraph,
    Custom,
    /// Schema default: build the text shape as a whole.
    #[default]
    Whole,
}

impl AnimationParagraphBuildType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "p" => Ok(Self::Paragraph),
            "cust" => Ok(Self::Custom),
            "whole" => Ok(Self::Whole),
            _ => Err(invalid("invalid paragraph build type")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Paragraph => "p",
            Self::Custom => "cust",
            Self::Whole => "whole",
        }
    }
}

/// Diagram build mode from `ST_TLDiagramBuildType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationDiagramBuildType {
    /// Schema default: animate the diagram as one object.
    #[default]
    Whole,
    DepthByNode,
    DepthByBranch,
    BreadthByNode,
    BreadthByLevel,
    Clockwise,
    ClockwiseIn,
    ClockwiseOut,
    CounterClockwise,
    CounterClockwiseIn,
    CounterClockwiseOut,
    InByRing,
    OutByRing,
    Up,
    Down,
    AllAtOnce,
    Custom,
}

impl AnimationDiagramBuildType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "whole" => Ok(Self::Whole),
            "depthByNode" => Ok(Self::DepthByNode),
            "depthByBranch" => Ok(Self::DepthByBranch),
            "breadthByNode" => Ok(Self::BreadthByNode),
            "breadthByLvl" => Ok(Self::BreadthByLevel),
            "cw" => Ok(Self::Clockwise),
            "cwIn" => Ok(Self::ClockwiseIn),
            "cwOut" => Ok(Self::ClockwiseOut),
            "ccw" => Ok(Self::CounterClockwise),
            "ccwIn" => Ok(Self::CounterClockwiseIn),
            "ccwOut" => Ok(Self::CounterClockwiseOut),
            "inByRing" => Ok(Self::InByRing),
            "outByRing" => Ok(Self::OutByRing),
            "up" => Ok(Self::Up),
            "down" => Ok(Self::Down),
            "allAtOnce" => Ok(Self::AllAtOnce),
            "cust" => Ok(Self::Custom),
            _ => Err(invalid("invalid diagram build type")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Whole => "whole",
            Self::DepthByNode => "depthByNode",
            Self::DepthByBranch => "depthByBranch",
            Self::BreadthByNode => "breadthByNode",
            Self::BreadthByLevel => "breadthByLvl",
            Self::Clockwise => "cw",
            Self::ClockwiseIn => "cwIn",
            Self::ClockwiseOut => "cwOut",
            Self::CounterClockwise => "ccw",
            Self::CounterClockwiseIn => "ccwIn",
            Self::CounterClockwiseOut => "ccwOut",
            Self::InByRing => "inByRing",
            Self::OutByRing => "outByRing",
            Self::Up => "up",
            Self::Down => "down",
            Self::AllAtOnce => "allAtOnce",
            Self::Custom => "cust",
        }
    }
}

/// Build information for a PowerPoint OLE diagram.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationDiagramBuild {
    /// OLE graphic-frame shape receiving the build.
    pub shape_id: u32,
    /// Timing group referenced by the build.
    pub group_id: AnimationGroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    /// Diagram build mode. Defaults to `Whole`.
    pub build_type: AnimationDiagramBuildType,
}

impl AnimationDiagramBuild {
    /// Construct a diagram build using schema defaults.
    pub const fn new(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: AnimationDiagramBuildType::Whole,
        }
    }

    /// Set whether this build appears expanded in animation UIs.
    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }

    /// Set the diagram build mode.
    pub fn with_build_type(mut self, build_type: AnimationDiagramBuildType) -> Self {
        self.build_type = build_type;
        self
    }
}

/// DrawingML diagram build mode used inside `p:bldGraphic/p:bldSub`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationGraphicDiagramBuildType {
    /// Schema default: animate all diagram content at once.
    #[default]
    AllAtOnce,
    One,
    LevelOne,
    LevelAtOnce,
}

impl AnimationGraphicDiagramBuildType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "one" => Ok(Self::One),
            "lvlOne" => Ok(Self::LevelOne),
            "lvlAtOnce" => Ok(Self::LevelAtOnce),
            _ => Err(invalid("invalid graphical-object diagram build type")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::One => "one",
            Self::LevelOne => "lvlOne",
            Self::LevelAtOnce => "lvlAtOnce",
        }
    }
}

/// DrawingML chart build mode used inside `p:bldGraphic/p:bldSub`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationGraphicChartBuildType {
    /// Schema default: animate all chart content at once.
    #[default]
    AllAtOnce,
    Series,
    Category,
    SeriesElement,
    CategoryElement,
}

impl AnimationGraphicChartBuildType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid graphical-object chart build type")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Series => "series",
            Self::Category => "category",
            Self::SeriesElement => "seriesEl",
            Self::CategoryElement => "categoryEl",
        }
    }
}

/// Required `p:bldGraphic` content choice.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationGraphicBuildMode {
    /// Animate the chart or diagram as one graphical object.
    AsOne,
    /// Animate SmartArt/diagram sub-elements.
    Diagram {
        build_type: AnimationGraphicDiagramBuildType,
        reverse: bool,
    },
    /// Animate chart sub-elements.
    Chart {
        build_type: AnimationGraphicChartBuildType,
        animate_background: bool,
    },
}

/// Build information for a chart or SmartArt graphical frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationGraphicBuild {
    pub shape_id: u32,
    pub group_id: AnimationGroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    pub mode: AnimationGraphicBuildMode,
}

impl AnimationGraphicBuild {
    pub const fn new(
        shape_id: u32,
        group_id: AnimationGroupId,
        mode: AnimationGraphicBuildMode,
    ) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            mode,
        }
    }

    pub const fn as_one(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self::new(shape_id, group_id, AnimationGraphicBuildMode::AsOne)
    }

    pub const fn diagram(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self::new(
            shape_id,
            group_id,
            AnimationGraphicBuildMode::Diagram {
                build_type: AnimationGraphicDiagramBuildType::AllAtOnce,
                reverse: false,
            },
        )
    }

    pub const fn chart(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self::new(
            shape_id,
            group_id,
            AnimationGraphicBuildMode::Chart {
                build_type: AnimationGraphicChartBuildType::AllAtOnce,
                animate_background: true,
            },
        )
    }

    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }
}

/// Embedded OLE chart build mode from `ST_TLOleChartBuildType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationOleChartBuildType {
    #[default]
    AllAtOnce,
    Series,
    Category,
    SeriesElement,
    CategoryElement,
}

impl AnimationOleChartBuildType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid OLE chart build type")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Series => "series",
            Self::Category => "category",
            Self::SeriesElement => "seriesEl",
            Self::CategoryElement => "categoryEl",
        }
    }
}

/// Build information for a legacy embedded OLE chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationOleChartBuild {
    pub shape_id: u32,
    pub group_id: AnimationGroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    /// Chart build mode. Defaults to `AllAtOnce`.
    pub build_type: AnimationOleChartBuildType,
    /// Whether the chart background participates. Defaults to `true`.
    pub animate_background: bool,
}

impl AnimationOleChartBuild {
    pub const fn new(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: AnimationOleChartBuildType::AllAtOnce,
            animate_background: true,
        }
    }

    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }

    pub fn with_build_type(mut self, build_type: AnimationOleChartBuildType) -> Self {
        self.build_type = build_type;
        self
    }

    pub fn with_animate_background(mut self, animate: bool) -> Self {
        self.animate_background = animate;
        self
    }
}

/// Validated root parallel time node embedded by a paragraph-build template.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationTemplateTimeNode {
    xml: Box<str>,
}

impl AnimationTemplateTimeNode {
    /// Validate and store one bounded `p:par` template time node.
    pub fn parse(xml: &str) -> Result<Self> {
        validate_template_time_node(xml)?;
        Ok(Self {
            xml: xml.to_string().into_boxed_str(),
        })
    }

    /// Exact validated XML for the root `p:par` node.
    pub fn as_xml(&self) -> &str {
        &self.xml
    }
}

/// Template effects for one paragraph level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationParagraphTemplate {
    /// PowerPoint paragraph level in the inclusive range `0..=9`.
    pub level: u8,
    /// Required single root time node.
    pub time_node: AnimationTemplateTimeNode,
}

impl AnimationParagraphTemplate {
    /// Construct a paragraph template with a PowerPoint-supported level.
    pub fn new(level: u8, time_node: AnimationTemplateTimeNode) -> Result<Self> {
        if level > 9 {
            return Err(invalid("paragraph template level exceeds PowerPoint limit"));
        }
        Ok(Self { level, time_node })
    }
}

/// A paragraph build associated with a text shape and timing group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationParagraphBuild {
    /// Shape whose paragraphs participate in the build.
    pub shape_id: u32,
    /// Timing group referenced by the build.
    pub group_id: AnimationGroupId,
    /// Whether the build is expanded in the animation UI. Defaults to `false`.
    pub ui_expand: bool,
    /// Paragraph build mode. Defaults to `Whole`.
    pub build_type: AnimationParagraphBuildType,
    /// Paragraph build level. Defaults to `1`; non-default values require `Paragraph` mode.
    pub build_level: u32,
    /// Whether to animate the text shape background. Defaults to `false`.
    pub animate_background: bool,
    /// Whether to update background animation automatically. Defaults to `true`.
    pub auto_update_animate_background: bool,
    /// Whether paragraph order is reversed. Defaults to `false` and requires `Paragraph` mode.
    pub reverse: bool,
    /// Automatic advance time. The schema default is `Indefinite`.
    pub auto_advance: Duration,
    /// Optional template effects, one per unique paragraph level.
    pub templates: Vec<AnimationParagraphTemplate>,
}

impl AnimationParagraphBuild {
    /// Construct a paragraph build reference.
    pub fn new(shape_id: u32, group_id: AnimationGroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: AnimationParagraphBuildType::Whole,
            build_level: 1,
            animate_background: false,
            auto_update_animate_background: true,
            reverse: false,
            auto_advance: Duration::Indefinite,
            templates: Vec::new(),
        }
    }

    /// Set whether the build appears expanded in animation UIs.
    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }

    /// Set the paragraph build mode.
    pub fn with_build_type(mut self, build_type: AnimationParagraphBuildType) -> Self {
        self.build_type = build_type;
        self
    }

    /// Set the paragraph build level.
    pub fn with_build_level(mut self, level: u32) -> Self {
        self.build_level = level;
        self
    }

    /// Set whether the text shape background participates in the animation.
    pub fn with_animate_background(mut self, animate: bool) -> Self {
        self.animate_background = animate;
        self
    }

    /// Set automatic background-animation updates.
    pub fn with_auto_update_animate_background(mut self, update: bool) -> Self {
        self.auto_update_animate_background = update;
        self
    }

    /// Set reverse paragraph ordering.
    pub fn with_reverse(mut self, reverse: bool) -> Self {
        self.reverse = reverse;
        self
    }

    /// Set the automatic advance time.
    pub fn with_auto_advance(mut self, auto_advance: impl Into<Duration>) -> Self {
        self.auto_advance = auto_advance.into();
        self
    }

    /// Add template effects for a paragraph level.
    pub fn with_template(mut self, template: AnimationParagraphTemplate) -> Self {
        self.templates.push(template);
        self
    }

    /// Effective PowerPoint auto-advance delay; `indefinite` is interpreted as zero.
    pub const fn powerpoint_auto_advance_milliseconds(&self) -> u32 {
        match self.auto_advance {
            Duration::Finite(value) => value,
            Duration::Indefinite => 0,
        }
    }
}

/// Event filtering supported by PowerPoint for a triggered sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationEventFilter {
    /// Prevent the trigger event from bubbling beyond the interactive sequence.
    CancelBubble,
}

impl AnimationEventFilter {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "cancelBubble" => Ok(Self::CancelBubble),
            _ => Err(invalid("invalid PowerPoint animation event filter")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::CancelBubble => "cancelBubble",
        }
    }
}

/// Structural sequence containing an animation effect.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum AnimationSequenceContext {
    /// The slide's ordinary click sequence.
    #[default]
    Main,
    /// A sequence activated by clicking a shape on the slide.
    Interactive {
        /// Shape whose click activates or advances the sequence.
        trigger_shape_id: u32,
        /// Optional PowerPoint event-bubbling filter on the `interactiveSeq` cTn.
        event_filter: Option<AnimationEventFilter>,
    },
}

/// Animation direction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationDirection {
    Up,
    Down,
    Left,
    Right,
    UpLeft,
    UpRight,
    DownLeft,
    DownRight,
    /// Toward the center, used by Zoom.
    In,
    /// Away from the center, used by Zoom.
    Out,
    /// Horizontal closing split.
    HorizontalIn,
    /// Horizontal opening split.
    HorizontalOut,
    /// Vertical closing split.
    VerticalIn,
    /// Vertical opening split.
    VerticalOut,
    /// A subtle zoom toward the center.
    InSlightly,
    /// A subtle zoom away from the center.
    OutSlightly,
    /// Zoom toward the center beginning at the screen center.
    InFromScreenCenter,
    /// Zoom away from the center ending at the screen center.
    OutFromScreenCenter,
}

/// Behavior of animated properties after a time node becomes inactive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationFill {
    Remove,
    Freeze,
    Hold,
    Transition,
}

impl AnimationFill {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "remove" => Ok(Self::Remove),
            "freeze" => Ok(Self::Freeze),
            "hold" => Ok(Self::Hold),
            "transition" => Ok(Self::Transition),
            _ => Err(invalid("invalid animation fill behavior")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Remove => "remove",
            Self::Freeze => "freeze",
            Self::Hold => "hold",
            Self::Transition => "transition",
        }
    }
}

/// Policy controlling whether a completed time node can restart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationRestart {
    Always,
    WhenNotActive,
    Never,
}

impl AnimationRestart {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "always" => Ok(Self::Always),
            "whenNotActive" => Ok(Self::WhenNotActive),
            "never" => Ok(Self::Never),
            _ => Err(invalid("invalid animation restart behavior")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Always => "always",
            Self::WhenNotActive => "whenNotActive",
            Self::Never => "never",
        }
    }
}

/// Repeat count for a time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationRepeat {
    /// Count in OOXML thousandths, where `1000` means one iteration.
    Finite(u32),
    Indefinite,
}

/// Nonzero playback speed in thousandths of a percent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationSpeed(i32);

impl AnimationSpeed {
    /// Construct a speed value. PowerPoint rejects zero speed.
    pub fn new(thousandths_percent: i32) -> Result<Self> {
        if thousandths_percent == 0 {
            Err(invalid("animation speed must be nonzero"))
        } else {
            Ok(Self(thousandths_percent))
        }
    }

    /// Return the encoded OOXML percentage value.
    pub const fn thousandths_percent(self) -> i32 {
        self.0
    }
}

/// Positive fixed percentage used for acceleration and deceleration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationProgress(u32);

impl AnimationProgress {
    /// Construct a value from thousandths of a percent (`100000` is 100%).
    pub fn new(thousandths_percent: u32) -> Result<Self> {
        if thousandths_percent > 100_000 {
            Err(invalid(
                "animation progression percentage exceeds 100 percent",
            ))
        } else {
            Ok(Self(thousandths_percent))
        }
    }

    /// Return the encoded OOXML percentage value.
    pub const fn thousandths_percent(self) -> u32 {
        self.0
    }
}

/// Synchronization policy between a time node and its containing group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationSyncBehavior {
    CanSlip,
    Locked,
    /// PowerPoint's assumed synchronization behavior.
    None,
}

/// Exact normalized time in the inclusive range `0..=1`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NormalizedTime {
    numerator: u64,
    scale: u64,
}

impl NormalizedTime {
    /// Construct a normalized time from millionths.
    pub fn from_millionths(value: u32) -> Result<Self> {
        if value > 1_000_000 {
            return Err(invalid("normalized time exceeds 1.0"));
        }
        Ok(Self::normalized(u64::from(value), 1_000_000))
    }

    /// Exact numerator of the normalized decimal value.
    pub const fn numerator(self) -> u64 {
        self.numerator
    }

    /// Exact power-of-ten scale of the normalized decimal value.
    pub const fn scale(self) -> u64 {
        self.scale
    }

    fn parse(value: &str) -> Result<Self> {
        if value.is_empty() {
            return Err(invalid("normalized time is empty"));
        }
        let (whole, fraction) = match value.split_once('.') {
            Some((whole, fraction)) => {
                if fraction.is_empty() || fraction.contains('.') {
                    return Err(invalid("invalid normalized time decimal"));
                }
                (whole, Some(fraction))
            },
            None => (value, None),
        };
        if !matches!(whole, "0" | "1") {
            return Err(invalid("normalized time must be between 0 and 1"));
        }
        let Some(fraction) = fraction else {
            return Ok(if whole == "1" {
                Self {
                    numerator: 1,
                    scale: 1,
                }
            } else {
                Self {
                    numerator: 0,
                    scale: 1,
                }
            });
        };
        if fraction.len() > MAX_NORMALIZED_TIME_DECIMALS
            || !fraction.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(invalid("invalid or over-precise normalized time"));
        }
        if whole == "1" && fraction.bytes().any(|byte| byte != b'0') {
            return Err(invalid("normalized time exceeds 1.0"));
        }
        if whole == "1" {
            return Ok(Self {
                numerator: 1,
                scale: 1,
            });
        }
        let numerator = fraction
            .parse::<u64>()
            .map_err(|_| invalid("normalized time decimal overflows"))?;
        let scale = 10u64
            .checked_pow(
                u32::try_from(fraction.len())
                    .map_err(|_| invalid("normalized time precision overflows"))?,
            )
            .ok_or_else(|| invalid("normalized time scale overflows"))?;
        Ok(Self::normalized(numerator, scale))
    }

    fn normalized(mut numerator: u64, mut scale: u64) -> Self {
        while scale > 1 && numerator % 10 == 0 {
            numerator /= 10;
            scale /= 10;
        }
        Self { numerator, scale }
    }

    fn strictly_before(self, other: Self) -> bool {
        u128::from(self.numerator) * u128::from(other.scale)
            < u128::from(other.numerator) * u128::from(self.scale)
    }

    fn write_value(self) -> String {
        if self.numerator == 0 {
            return "0".to_string();
        }
        if self.numerator == self.scale {
            return "1".to_string();
        }
        let decimals = self.scale.ilog10() as usize;
        format!("0.{:0decimals$}", self.numerator)
    }
}

/// A source-time to warped-time mapping point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AnimationTimePoint {
    pub local_time: NormalizedTime,
    pub warped_time: NormalizedTime,
}

impl AnimationTimePoint {
    pub const fn new(local_time: NormalizedTime, warped_time: NormalizedTime) -> Self {
        Self {
            local_time,
            warped_time,
        }
    }
}

/// Bounded piecewise time-warp filter for a common time node.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationTimeFilter {
    points: Box<[AnimationTimePoint]>,
}

impl AnimationTimeFilter {
    /// Construct a filter whose local times are strictly increasing.
    pub fn new(points: Vec<AnimationTimePoint>) -> Result<Self> {
        if points.is_empty() {
            return Err(invalid(
                "animation time filter must contain at least one point",
            ));
        }
        if points.len() > MAX_TIME_FILTER_POINTS {
            return Err(invalid(
                "animation time filter point count exceeds safety limit",
            ));
        }
        if points
            .windows(2)
            .any(|pair| !pair[0].local_time.strictly_before(pair[1].local_time))
        {
            return Err(invalid(
                "animation time filter local times must be strictly increasing",
            ));
        }
        Ok(Self {
            points: points.into_boxed_slice(),
        })
    }

    /// Mapping points in source-time order.
    pub fn points(&self) -> &[AnimationTimePoint] {
        &self.points
    }

    fn parse(value: &str) -> Result<Self> {
        if value.len() > MAX_TIME_FILTER_BYTES {
            return Err(invalid("animation time filter exceeds safety limit"));
        }
        let mut points = Vec::new();
        for pair in value.split(';') {
            if points.len() >= MAX_TIME_FILTER_POINTS {
                return Err(invalid(
                    "animation time filter point count exceeds safety limit",
                ));
            }
            let pair = pair.trim();
            let (local, warped) = pair
                .split_once(',')
                .ok_or_else(|| invalid("animation time filter point is missing a comma"))?;
            if warped.contains(',') {
                return Err(invalid("animation time filter point has too many values"));
            }
            points.push(AnimationTimePoint::new(
                NormalizedTime::parse(local.trim())?,
                NormalizedTime::parse(warped.trim())?,
            ));
        }
        Self::new(points)
    }

    fn write_value(&self) -> String {
        let mut output = String::new();
        for (index, point) in self.points.iter().enumerate() {
            if index != 0 {
                output.push(';');
            }
            output.push_str(&point.local_time.write_value());
            output.push(',');
            output.push_str(&point.warped_time.write_value());
        }
        output
    }
}

impl AnimationSyncBehavior {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "canSlip" => Ok(Self::CanSlip),
            "locked" => Ok(Self::Locked),
            "none" => Ok(Self::None),
            _ => Err(invalid("invalid animation synchronization behavior")),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::CanSlip => "canSlip",
            Self::Locked => "locked",
            Self::None => "none",
        }
    }
}

impl AnimationRepeat {
    fn write_value(self) -> String {
        match self {
            Self::Finite(value) => value.to_string(),
            Self::Indefinite => "indefinite".to_string(),
        }
    }
}

/// Duration of a simple animation time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Duration {
    /// A finite duration in milliseconds.
    Finite(u32),
    /// An animation that has no finite duration.
    Indefinite,
}

impl Duration {
    /// Construct a finite duration in milliseconds.
    pub const fn milliseconds(value: u32) -> Self {
        Self::Finite(value)
    }

    /// Return the finite millisecond value, or `None` for an indefinite duration.
    pub const fn as_milliseconds(self) -> Option<u32> {
        match self {
            Self::Finite(value) => Some(value),
            Self::Indefinite => None,
        }
    }

    fn write_value(self) -> String {
        match self {
            Self::Finite(value) => value.to_string(),
            Self::Indefinite => "indefinite".to_string(),
        }
    }
}

impl From<u32> for Duration {
    fn from(value: u32) -> Self {
        Self::Finite(value)
    }
}

impl PartialEq<u32> for Duration {
    fn eq(&self, other: &u32) -> bool {
        matches!(self, Self::Finite(value) if value == other)
    }
}

/// An animation applied to a shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Animation {
    /// Target shape ID
    pub shape_id: u32,
    /// Animation effect
    pub effect: AnimationEffect,
    /// Trigger type
    pub trigger: AnimationTrigger,
    /// Duration in milliseconds
    pub duration: Duration,
    /// Delay before starting (ms)
    pub delay: u32,
    /// Direction (for directional effects)
    pub direction: Option<AnimationDirection>,
    /// Property state retained after the animation becomes inactive.
    pub fill: Option<AnimationFill>,
    /// Policy for restarting this time node.
    pub restart: Option<AnimationRestart>,
    /// Whether the animation runs backward after reaching its end.
    pub auto_reverse: bool,
    /// Optional repeat count.
    pub repeat: Option<AnimationRepeat>,
    /// Optional nonzero playback speed.
    pub speed: Option<AnimationSpeed>,
    /// Optional acceleration fraction.
    pub acceleration: Option<AnimationProgress>,
    /// Optional deceleration fraction.
    pub deceleration: Option<AnimationProgress>,
    /// Whether this time node is visible in the animation user interface.
    pub display: Option<bool>,
    /// Optional total duration for repeated playback.
    pub repeat_duration: Option<Duration>,
    /// Optional synchronization policy with the containing time group.
    pub sync_behavior: Option<AnimationSyncBehavior>,
    /// Whether this node is an after-effect.
    pub after_effect: Option<bool>,
    /// Optional normalized-time warp filter.
    pub time_filter: Option<AnimationTimeFilter>,
    /// Main-sequence or shape-triggered interactive-sequence context.
    pub sequence_context: AnimationSequenceContext,
    /// Optional build-list group containing this effect time node.
    pub group_id: Option<AnimationGroupId>,
    /// Sequence order (1-based)
    pub order: u32,
}

impl Animation {
    /// Create a new animation.
    pub fn new(shape_id: u32, effect: AnimationEffect) -> Self {
        Self {
            shape_id,
            effect,
            trigger: AnimationTrigger::OnClick,
            duration: Duration::Finite(500),
            delay: 0,
            direction: None,
            fill: Some(AnimationFill::Hold),
            restart: None,
            auto_reverse: false,
            repeat: None,
            speed: None,
            acceleration: None,
            deceleration: None,
            display: None,
            repeat_duration: None,
            sync_behavior: None,
            after_effect: None,
            time_filter: None,
            sequence_context: AnimationSequenceContext::Main,
            group_id: None,
            order: 1,
        }
    }

    /// Set the trigger type.
    pub fn with_trigger(mut self, trigger: AnimationTrigger) -> Self {
        self.trigger = trigger;
        self
    }

    /// Set the duration.
    pub fn with_duration(mut self, duration: impl Into<Duration>) -> Self {
        self.duration = duration.into();
        self
    }

    /// Set a finite duration in milliseconds.
    pub fn with_duration_ms(mut self, duration: u32) -> Self {
        self.duration = Duration::Finite(duration);
        self
    }

    /// Set the delay.
    pub fn with_delay(mut self, delay: u32) -> Self {
        self.delay = delay;
        self
    }

    /// Set the direction.
    pub fn with_direction(mut self, direction: AnimationDirection) -> Self {
        self.direction = Some(direction);
        self
    }

    /// Set the fill behavior for the animation time node.
    pub fn with_fill(mut self, fill: AnimationFill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set the restart behavior for the animation time node.
    pub fn with_restart(mut self, restart: AnimationRestart) -> Self {
        self.restart = Some(restart);
        self
    }

    /// Enable or disable automatic reversal.
    pub fn with_auto_reverse(mut self, auto_reverse: bool) -> Self {
        self.auto_reverse = auto_reverse;
        self
    }

    /// Set the repeat behavior for the animation time node.
    pub fn with_repeat(mut self, repeat: AnimationRepeat) -> Self {
        self.repeat = Some(repeat);
        self
    }

    /// Set the nonzero playback speed.
    pub fn with_speed(mut self, speed: AnimationSpeed) -> Self {
        self.speed = Some(speed);
        self
    }

    /// Set the acceleration fraction.
    pub fn with_acceleration(mut self, acceleration: AnimationProgress) -> Self {
        self.acceleration = Some(acceleration);
        self
    }

    /// Set the deceleration fraction.
    pub fn with_deceleration(mut self, deceleration: AnimationProgress) -> Self {
        self.deceleration = Some(deceleration);
        self
    }

    /// Set whether this node is displayed in animation user interfaces.
    pub fn with_display(mut self, display: bool) -> Self {
        self.display = Some(display);
        self
    }

    /// Set the total duration for repeated playback.
    pub fn with_repeat_duration(mut self, duration: impl Into<Duration>) -> Self {
        self.repeat_duration = Some(duration.into());
        self
    }

    /// Set synchronization with the containing time group.
    pub fn with_sync_behavior(mut self, behavior: AnimationSyncBehavior) -> Self {
        self.sync_behavior = Some(behavior);
        self
    }

    /// Mark or unmark this node as an after-effect.
    pub fn with_after_effect(mut self, after_effect: bool) -> Self {
        self.after_effect = Some(after_effect);
        self
    }

    /// Set a normalized-time warp filter.
    pub fn with_time_filter(mut self, filter: AnimationTimeFilter) -> Self {
        self.time_filter = Some(filter);
        self
    }

    /// Put this effect in a shape-triggered sequence using PowerPoint's filter.
    pub fn with_interactive_trigger(mut self, trigger_shape_id: u32) -> Self {
        self.sequence_context = AnimationSequenceContext::Interactive {
            trigger_shape_id,
            event_filter: Some(AnimationEventFilter::CancelBubble),
        };
        self
    }

    /// Set the structural sequence context explicitly.
    pub fn with_sequence_context(mut self, context: AnimationSequenceContext) -> Self {
        self.sequence_context = context;
        self
    }

    /// Associate this effect cTn with a build-list timing group.
    pub fn with_group_id(mut self, group_id: impl Into<AnimationGroupId>) -> Self {
        self.group_id = Some(group_id.into());
        self
    }
}

/// A trigger event on an ordered timing condition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationConditionEvent {
    OnBegin,
    OnEnd,
    Begin,
    End,
    OnClick,
    OnDoubleClick,
    OnMouseOver,
    OnMouseOut,
    OnNext,
    OnPrevious,
    OnStopAudio,
}

impl AnimationConditionEvent {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "onBegin" => Ok(Self::OnBegin),
            "onEnd" => Ok(Self::OnEnd),
            "begin" => Ok(Self::Begin),
            "end" => Ok(Self::End),
            "onClick" => Ok(Self::OnClick),
            "onDblClick" => Ok(Self::OnDoubleClick),
            "onMouseOver" => Ok(Self::OnMouseOver),
            "onMouseOut" => Ok(Self::OnMouseOut),
            "onNext" => Ok(Self::OnNext),
            "onPrev" => Ok(Self::OnPrevious),
            "onStopAudio" => Ok(Self::OnStopAudio),
            _ => Err(invalid("invalid animation condition event")),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::OnBegin => "onBegin",
            Self::OnEnd => "onEnd",
            Self::Begin => "begin",
            Self::End => "end",
            Self::OnClick => "onClick",
            Self::OnDoubleClick => "onDblClick",
            Self::OnMouseOver => "onMouseOver",
            Self::OnMouseOut => "onMouseOut",
            Self::OnNext => "onNext",
            Self::OnPrevious => "onPrev",
            Self::OnStopAudio => "onStopAudio",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationRuntimeTrigger {
    First,
    Last,
    All,
}
impl AnimationRuntimeTrigger {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "first" => Ok(Self::First),
            "last" => Ok(Self::Last),
            "all" => Ok(Self::All),
            _ => Err(invalid("invalid animation runtime trigger")),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::First => "first",
            Self::Last => "last",
            Self::All => "all",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationConditionTarget {
    Shape(u32),
    Slide,
    TimeNode(u32),
    Runtime(AnimationRuntimeTrigger),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationTimeCondition {
    pub event: Option<AnimationConditionEvent>,
    pub delay: Duration,
    pub target: Option<AnimationConditionTarget>,
}
impl Default for AnimationTimeCondition {
    fn default() -> Self {
        Self {
            event: None,
            delay: Duration::Finite(0),
            target: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationPresetClass {
    Entrance,
    Exit,
    Emphasis,
    MotionPath,
    Verb,
    MediaCall,
}
impl AnimationPresetClass {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "entr" => Ok(Self::Entrance),
            "exit" => Ok(Self::Exit),
            "emph" => Ok(Self::Emphasis),
            "path" => Ok(Self::MotionPath),
            "verb" => Ok(Self::Verb),
            "mediacall" => Ok(Self::MediaCall),
            _ => Err(invalid("invalid animation preset class")),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::Entrance => "entr",
            Self::Exit => "exit",
            Self::Emphasis => "emph",
            Self::MotionPath => "path",
            Self::Verb => "verb",
            Self::MediaCall => "mediacall",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationPresetTimeNode {
    pub preset_id: u32,
    pub class: AnimationPresetClass,
    pub subtype: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnimationTimeNodeType {
    ClickEffect,
    WithEffect,
    AfterEffect,
    MainSequence,
    InteractiveSequence,
    ClickParallel,
    WithGroup,
    AfterGroup,
    TimingRoot,
}
impl AnimationTimeNodeType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "clickEffect" => Ok(Self::ClickEffect),
            "withEffect" => Ok(Self::WithEffect),
            "afterEffect" => Ok(Self::AfterEffect),
            "mainSeq" => Ok(Self::MainSequence),
            "interactiveSeq" => Ok(Self::InteractiveSequence),
            "clickPar" => Ok(Self::ClickParallel),
            "withGroup" => Ok(Self::WithGroup),
            "afterGroup" => Ok(Self::AfterGroup),
            "tmRoot" => Ok(Self::TimingRoot),
            _ => Err(invalid("invalid animation time-node type")),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::ClickEffect => "clickEffect",
            Self::WithEffect => "withEffect",
            Self::AfterEffect => "afterEffect",
            Self::MainSequence => "mainSeq",
            Self::InteractiveSequence => "interactiveSeq",
            Self::ClickParallel => "clickPar",
            Self::WithGroup => "withGroup",
            Self::AfterGroup => "afterGroup",
            Self::TimingRoot => "tmRoot",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationNextAction {
    #[default]
    None,
    Seek,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationPreviousAction {
    #[default]
    None,
    SkipTimed,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationTimingNodeKind {
    Parallel,
    Sequence {
        concurrent: bool,
        next_action: AnimationNextAction,
        previous_action: AnimationPreviousAction,
    },
    Exclusive,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationTimingChild {
    Node(AnimationTimingNode),
    Opaque(Box<str>),
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationCommonTimeNode {
    /// Optional schema time-node identifier.
    pub id: Option<u32>,
    pub duration: Option<Duration>,
    pub node_type: Option<AnimationTimeNodeType>,
    pub preset: Option<AnimationPresetTimeNode>,
    pub start_conditions: Vec<AnimationTimeCondition>,
    pub end_conditions: Vec<AnimationTimeCondition>,
    pub children: Vec<AnimationTimingChild>,
    pub sub_nodes: Vec<AnimationTimingChild>,
    pub opaque_children: Vec<Box<str>>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnimationTimingNode {
    pub kind: AnimationTimingNodeKind,
    pub common: AnimationCommonTimeNode,
    pub opaque_children: Vec<Box<str>>,
}
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct AnimationTimingTree {
    pub roots: Vec<AnimationTimingChild>,
    pub opaque_children: Vec<Box<str>>,
    source_xml: Option<Box<str>>,
    source_roots: Option<Box<[AnimationTimingChild]>>,
    source_opaque_children: Option<Box<[Box<str>]>>,
}
impl AnimationTimingTree {
    pub fn parse(xml: &str) -> Result<Self> {
        check_xml_size(xml.len())?;
        let processed = crate::common::mce::process_str(xml)?;
        parse_recursive_timing_tree(&processed)
    }
    pub fn to_xml(&self) -> String {
        if let (Some(xml), Some(roots), Some(opaque)) = (
            &self.source_xml,
            &self.source_roots,
            &self.source_opaque_children,
        ) && self.roots.as_slice() == roots.as_ref()
            && self.opaque_children.as_slice() == opaque.as_ref()
        {
            return xml.to_string();
        }
        let mut xml = String::from("<p:timing><p:tnLst>");
        for child in &self.roots {
            write_timing_child(&mut xml, child);
        }
        xml.push_str("</p:tnLst>");
        for child in &self.opaque_children {
            xml.push_str(child);
        }
        xml.push_str("</p:timing>");
        xml
    }
}

/// Animation sequence for a slide.
#[derive(Debug, Clone, Default)]
pub struct AnimationSequence {
    /// List of animations in order
    pub animations: Vec<Animation>,
    /// Typed paragraph entries from the slide build list.
    pub paragraph_builds: Vec<AnimationParagraphBuild>,
    /// Typed OLE diagram entries from the slide build list.
    pub diagram_builds: Vec<AnimationDiagramBuild>,
    /// Typed chart and SmartArt entries from the slide build list.
    pub graphic_builds: Vec<AnimationGraphicBuild>,
    /// Typed embedded OLE chart entries from the slide build list.
    pub ole_chart_builds: Vec<AnimationOleChartBuild>,
    pub timing_tree: Option<AnimationTimingTree>,
    source_timing_xml: Option<Box<str>>,
    source_animations: Option<Box<[Animation]>>,
    source_paragraph_builds: Option<Box<[AnimationParagraphBuild]>>,
    source_diagram_builds: Option<Box<[AnimationDiagramBuild]>>,
    source_graphic_builds: Option<Box<[AnimationGraphicBuild]>>,
    source_ole_chart_builds: Option<Box<[AnimationOleChartBuild]>>,
    source_timing_tree: Option<Box<AnimationTimingTree>>,
}

impl PartialEq for AnimationSequence {
    fn eq(&self, other: &Self) -> bool {
        self.animations == other.animations
            && self.paragraph_builds == other.paragraph_builds
            && self.diagram_builds == other.diagram_builds
            && self.graphic_builds == other.graphic_builds
            && self.ole_chart_builds == other.ole_chart_builds
    }
}

impl Eq for AnimationSequence {}

impl AnimationSequence {
    /// Create a new empty animation sequence.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add an animation to the sequence.
    pub fn add(&mut self, mut animation: Animation) {
        animation.order = u32::try_from(self.animations.len() + 1).unwrap_or(u32::MAX);
        self.animations.push(animation);
    }

    /// Add a paragraph build to the slide build list.
    pub fn add_paragraph_build(&mut self, build: AnimationParagraphBuild) {
        self.paragraph_builds.push(build);
    }

    /// Add an OLE diagram build to the slide build list.
    pub fn add_diagram_build(&mut self, build: AnimationDiagramBuild) {
        self.diagram_builds.push(build);
    }

    /// Add a chart or SmartArt build to the slide build list.
    pub fn add_graphic_build(&mut self, build: AnimationGraphicBuild) {
        self.graphic_builds.push(build);
    }

    /// Add an embedded OLE chart build to the slide build list.
    pub fn add_ole_chart_build(&mut self, build: AnimationOleChartBuild) {
        self.ole_chart_builds.push(build);
    }

    /// Get the number of animations.
    pub fn len(&self) -> usize {
        self.animations.len()
    }

    /// Check if the sequence is empty.
    pub fn is_empty(&self) -> bool {
        self.animations.is_empty()
    }

    /// Return the preserved source timing subtree, when this sequence was parsed.
    pub fn preserved_timing_xml(&self) -> Option<&str> {
        self.source_timing_xml.as_deref()
    }

    pub fn timing_tree(&self) -> Option<&AnimationTimingTree> {
        self.timing_tree.as_ref()
    }

    /// Parse timing XML from a slide.
    pub fn parse_timing_xml(xml: &str) -> Result<Self> {
        check_xml_size(xml.len())?;
        let xml = crate::common::mce::process_str(xml)?;
        check_xml_size(xml.len())?;
        parse_processed_timing(xml.as_bytes(), false)
    }

    pub(crate) fn parse_slide_xml(xml: &[u8]) -> Result<Self> {
        check_xml_size(xml.len())?;
        parse_processed_timing(xml, true)
    }

    /// Parse a slide timing tree and strictly validate build targets against its OPC package.
    ///
    /// Unlike the XML-only parser, this resolves chart, SmartArt, and OLE relationship IDs,
    /// requires internal existing target parts with matching relationship/content types, and
    /// never reads or executes embedded target bytes.
    pub fn from_package_slide(
        package: &litchi_opc::OpcPackage,
        slide_part_name: &litchi_opc::PackURI,
    ) -> Result<Self> {
        crate::pptx::animation_relationships::parse_package_slide(package, slide_part_name)
    }

    /// Generate timing XML for a slide.
    pub fn to_xml(&self) -> String {
        if let (
            Some(xml),
            Some(source),
            Some(source_builds),
            Some(source_diagram_builds),
            Some(source_graphic_builds),
            Some(source_ole_chart_builds),
            source_timing_tree,
        ) = (
            &self.source_timing_xml,
            &self.source_animations,
            &self.source_paragraph_builds,
            &self.source_diagram_builds,
            &self.source_graphic_builds,
            &self.source_ole_chart_builds,
            &self.source_timing_tree,
        ) && self.animations.as_slice() == source.as_ref()
            && self.paragraph_builds.as_slice() == source_builds.as_ref()
            && self.diagram_builds.as_slice() == source_diagram_builds.as_ref()
            && self.graphic_builds.as_slice() == source_graphic_builds.as_ref()
            && self.ole_chart_builds.as_slice() == source_ole_chart_builds.as_ref()
            && self.timing_tree.as_ref() == source_timing_tree.as_deref()
        {
            return xml.to_string();
        }
        if self.animations.as_slice() == self.source_animations.as_deref().unwrap_or_default()
            && self.paragraph_builds.as_slice()
                == self.source_paragraph_builds.as_deref().unwrap_or_default()
            && self.diagram_builds.as_slice()
                == self.source_diagram_builds.as_deref().unwrap_or_default()
            && self.graphic_builds.as_slice()
                == self.source_graphic_builds.as_deref().unwrap_or_default()
            && self.ole_chart_builds.as_slice()
                == self.source_ole_chart_builds.as_deref().unwrap_or_default()
            && let Some(timing_tree) = &self.timing_tree
        {
            return timing_tree.to_xml();
        }
        if self.is_empty() {
            return String::new();
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str("<p:timing>");
        xml.push_str("<p:tnLst>");
        xml.push_str(r#"<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">"#);
        xml.push_str(r#"<p:childTnLst><p:seq concurrent="1" nextAc="seek">"#);
        xml.push_str(r#"<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>"#);

        let mut tn_id = 3u32;
        for anim in self
            .animations
            .iter()
            .filter(|anim| anim.sequence_context == AnimationSequenceContext::Main)
        {
            write_animation_xml(&mut xml, anim, &mut tn_id, None);
        }

        xml.push_str("</p:childTnLst></p:cTn>");
        xml.push_str(r#"<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>"#);
        xml.push_str(r#"<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>"#);
        xml.push_str("</p:seq>");

        let mut contexts = Vec::<&AnimationSequenceContext>::new();
        for animation in &self.animations {
            if animation.sequence_context != AnimationSequenceContext::Main
                && !contexts.contains(&&animation.sequence_context)
            {
                contexts.push(&animation.sequence_context);
            }
        }
        for context in contexts {
            let AnimationSequenceContext::Interactive {
                trigger_shape_id,
                event_filter,
            } = context
            else {
                continue;
            };
            xml.push_str(r#"<p:seq concurrent="1" nextAc="seek"><p:cTn"#);
            xml.push_str(&format!(
                r#" id="{}" dur="indefinite" restart="whenNotActive" nodeType="interactiveSeq""#,
                tn_id
            ));
            tn_id += 1;
            if let Some(event_filter) = event_filter {
                xml.push_str(&format!(r#" evtFilter="{}""#, event_filter.as_str()));
            }
            xml.push_str("><p:childTnLst>");
            for animation in self
                .animations
                .iter()
                .filter(|animation| &animation.sequence_context == context)
            {
                write_animation_xml(&mut xml, animation, &mut tn_id, Some(*trigger_shape_id));
            }
            xml.push_str("</p:childTnLst></p:cTn></p:seq>");
        }

        xml.push_str("</p:childTnLst></p:cTn></p:par>");
        xml.push_str("</p:tnLst>");
        if !self.paragraph_builds.is_empty()
            || !self.diagram_builds.is_empty()
            || !self.graphic_builds.is_empty()
            || !self.ole_chart_builds.is_empty()
        {
            xml.push_str("<p:bldLst>");
            for build in &self.paragraph_builds {
                xml.push_str(&format!(
                    r#"<p:bldP spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != AnimationParagraphBuildType::Whole {
                    xml.push_str(&format!(r#" build="{}""#, build.build_type.as_str()));
                }
                if build.build_level != 1 {
                    xml.push_str(&format!(r#" bldLvl="{}""#, build.build_level));
                }
                if build.animate_background {
                    xml.push_str(r#" animBg="1""#);
                }
                if !build.auto_update_animate_background {
                    xml.push_str(r#" autoUpdateAnimBg="0""#);
                }
                if build.reverse {
                    xml.push_str(r#" rev="1""#);
                }
                if build.auto_advance != Duration::Indefinite {
                    xml.push_str(&format!(
                        r#" advAuto="{}""#,
                        build.auto_advance.write_value()
                    ));
                }
                if build.templates.is_empty() {
                    xml.push_str("/>");
                } else {
                    xml.push_str("><p:tmplLst>");
                    for template in &build.templates {
                        xml.push_str("<p:tmpl");
                        if template.level != 0 {
                            xml.push_str(&format!(r#" lvl="{}""#, template.level));
                        }
                        xml.push_str("><p:tnLst>");
                        xml.push_str(template.time_node.as_xml());
                        xml.push_str("</p:tnLst></p:tmpl>");
                    }
                    xml.push_str("</p:tmplLst></p:bldP>");
                }
            }
            for build in &self.diagram_builds {
                xml.push_str(&format!(
                    r#"<p:bldDgm spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != AnimationDiagramBuildType::Whole {
                    xml.push_str(&format!(r#" bld="{}""#, build.build_type.as_str()));
                }
                xml.push_str("/>");
            }
            for build in &self.graphic_builds {
                xml.push_str(&format!(
                    r#"<p:bldGraphic spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                xml.push('>');
                match build.mode {
                    AnimationGraphicBuildMode::AsOne => xml.push_str("<p:bldAsOne/>"),
                    AnimationGraphicBuildMode::Diagram {
                        build_type,
                        reverse,
                    } => {
                        xml.push_str("<p:bldSub><a:bldDgm");
                        if build_type != AnimationGraphicDiagramBuildType::AllAtOnce {
                            xml.push_str(&format!(r#" bld="{}""#, build_type.as_str()));
                        }
                        if reverse {
                            xml.push_str(r#" rev="1""#);
                        }
                        xml.push_str("/></p:bldSub>");
                    },
                    AnimationGraphicBuildMode::Chart {
                        build_type,
                        animate_background,
                    } => {
                        xml.push_str("<p:bldSub><a:bldChart");
                        if build_type != AnimationGraphicChartBuildType::AllAtOnce {
                            xml.push_str(&format!(r#" bld="{}""#, build_type.as_str()));
                        }
                        if !animate_background {
                            xml.push_str(r#" animBg="0""#);
                        }
                        xml.push_str("/></p:bldSub>");
                    },
                }
                xml.push_str("</p:bldGraphic>");
            }
            for build in &self.ole_chart_builds {
                xml.push_str(&format!(
                    r#"<p:bldOleChart spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != AnimationOleChartBuildType::AllAtOnce {
                    xml.push_str(&format!(r#" bld="{}""#, build.build_type.as_str()));
                }
                if !build.animate_background {
                    xml.push_str(r#" animBg="0""#);
                }
                xml.push_str("/>");
            }
            xml.push_str("</p:bldLst>");
        }
        xml.push_str("</p:timing>");

        xml
    }

    pub(crate) fn to_xml_for_slide(&self, valid_targets: &HashSet<u32>) -> Result<String> {
        if self.len() > MAX_ANIMATIONS {
            return Err(invalid("slide animation count exceeds safety limit"));
        }
        if self.paragraph_builds.len()
            + self.diagram_builds.len()
            + self.graphic_builds.len()
            + self.ole_chart_builds.len()
            > MAX_ANIMATION_BUILDS
        {
            return Err(invalid("slide animation build count exceeds safety limit"));
        }
        let animation_groups: HashSet<_> = self
            .animations
            .iter()
            .filter_map(|animation| animation.group_id)
            .collect();
        let mut build_groups = HashSet::new();
        let mut build_pairs = HashSet::new();
        for build in &self.paragraph_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "paragraph build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !build_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate paragraph build shape/group pair"));
            }
            if build.build_type != AnimationParagraphBuildType::Paragraph && build.build_level != 1
            {
                return Err(invalid(
                    "non-default paragraph build level requires build type p",
                ));
            }
            if build.reverse && build.build_type != AnimationParagraphBuildType::Paragraph {
                return Err(invalid("reverse paragraph order requires build type p"));
            }
            if build.templates.len() > MAX_PARAGRAPH_TEMPLATES {
                return Err(invalid("paragraph template count exceeds PowerPoint limit"));
            }
            let mut levels = HashSet::new();
            for template in &build.templates {
                if template.level > 9 {
                    return Err(invalid("paragraph template level exceeds PowerPoint limit"));
                }
                if !levels.insert(template.level) {
                    return Err(invalid("duplicate paragraph template level"));
                }
            }
            if build.build_type == AnimationParagraphBuildType::Whole && build.templates.len() > 1 {
                return Err(invalid(
                    "whole paragraph builds support exactly one template effect",
                ));
            }
            build_groups.insert(build.group_id);
        }
        let mut diagram_pairs = HashSet::new();
        for build in &self.diagram_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "diagram build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !diagram_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate diagram build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        let mut graphic_pairs = HashSet::new();
        for build in &self.graphic_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "graphical-object build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !graphic_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate graphical-object build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        let mut ole_chart_pairs = HashSet::new();
        for build in &self.ole_chart_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "OLE chart build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !ole_chart_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate OLE chart build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        if animation_groups != build_groups {
            return Err(invalid(
                "animation cTn group IDs and paragraph build group IDs do not match",
            ));
        }
        for animation in &self.animations {
            if animation.shape_id == 0 || !valid_targets.contains(&animation.shape_id) {
                return Err(invalid(format!(
                    "animation target {} is not a supported shape on the current slide",
                    animation.shape_id
                )));
            }
            if let AnimationSequenceContext::Interactive {
                trigger_shape_id, ..
            } = &animation.sequence_context
                && (*trigger_shape_id == 0 || !valid_targets.contains(trigger_shape_id))
            {
                return Err(invalid(format!(
                    "interactive animation trigger {} is not a supported shape on the current slide",
                    trigger_shape_id
                )));
            }
            if animation.delay > MAX_TIMING_MILLISECONDS {
                return Err(invalid("animation delay exceeds the supported OOXML limit"));
            }
            if let Duration::Finite(duration) = animation.duration
                && duration > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation duration exceeds the supported OOXML limit",
                ));
            }
            if let Some(direction) = &animation.direction
                && direction_subtype(&animation.effect, direction).is_none()
            {
                return Err(invalid(
                    "animation direction is not supported for this animation effect",
                ));
            }
            if let Some(AnimationRepeat::Finite(repeat)) = animation.repeat
                && repeat > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation repeat count exceeds the supported OOXML limit",
                ));
            }
            if let Some(Duration::Finite(repeat_duration)) = animation.repeat_duration
                && repeat_duration > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation repeat duration exceeds the supported OOXML limit",
                ));
            }
            if let Some(time_filter) = &animation.time_filter
                && time_filter.write_value().len() > MAX_TIME_FILTER_BYTES
            {
                return Err(invalid("animation time filter exceeds safety limit"));
            }
        }
        Ok(self.to_xml())
    }
}

fn write_animation_xml(
    xml: &mut String,
    anim: &Animation,
    tn_id: &mut u32,
    interactive_trigger: Option<u32>,
) {
    xml.push_str(&format!(
        r#"<p:par><p:cTn id="{}" fill="hold"><p:stCondLst>"#,
        *tn_id
    ));
    *tn_id += 1;
    if anim.trigger == AnimationTrigger::OnClick {
        if let Some(trigger_shape_id) = interactive_trigger {
            xml.push_str(&format!(
                r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{}"/></p:tgtEl></p:cond>"#,
                trigger_shape_id
            ));
        } else {
            xml.push_str(r#"<p:cond delay="indefinite"/>"#);
        }
    } else {
        xml.push_str(r#"<p:cond delay="0"/>"#);
    }
    xml.push_str("</p:stCondLst><p:childTnLst><p:par>");
    xml.push_str(&format!(
        r#"<p:cTn id="{}" fill="hold"><p:stCondLst><p:cond delay="{}"/></p:stCondLst>"#,
        *tn_id, anim.delay
    ));
    *tn_id += 1;

    xml.push_str("<p:childTnLst><p:par>");
    let node_type = match anim.trigger {
        AnimationTrigger::OnClick => "clickEffect",
        AnimationTrigger::WithPrevious => "withEffect",
        AnimationTrigger::AfterPrevious => "afterEffect",
    };
    let preset_subtype = anim
        .direction
        .as_ref()
        .and_then(|direction| direction_subtype(&anim.effect, direction))
        .unwrap_or(0);
    xml.push_str(&format!(
        r#"<p:cTn id="{}" presetID="{}" presetClass="{}" presetSubtype="{}""#,
        *tn_id,
        anim.effect.preset_id(),
        anim.effect.preset_class(),
        preset_subtype
    ));
    if let Some(fill) = anim.fill {
        xml.push_str(&format!(r#" fill="{}""#, fill.as_str()));
    }
    if let Some(restart) = anim.restart {
        xml.push_str(&format!(r#" restart="{}""#, restart.as_str()));
    }
    if anim.auto_reverse {
        xml.push_str(r#" autoRev="1""#);
    }
    if let Some(repeat) = anim.repeat {
        xml.push_str(&format!(r#" repeatCount="{}""#, repeat.write_value()));
    }
    if let Some(speed) = anim.speed {
        xml.push_str(&format!(r#" spd="{}""#, speed.thousandths_percent()));
    }
    if let Some(acceleration) = anim.acceleration {
        xml.push_str(&format!(
            r#" accel="{}""#,
            acceleration.thousandths_percent()
        ));
    }
    if let Some(deceleration) = anim.deceleration {
        xml.push_str(&format!(
            r#" decel="{}""#,
            deceleration.thousandths_percent()
        ));
    }
    if let Some(display) = anim.display {
        xml.push_str(if display {
            r#" display="1""#
        } else {
            r#" display="0""#
        });
    }
    if let Some(repeat_duration) = anim.repeat_duration {
        xml.push_str(&format!(
            r#" repeatDur="{}""#,
            repeat_duration.write_value()
        ));
    }
    if let Some(sync_behavior) = anim.sync_behavior {
        xml.push_str(&format!(r#" syncBehavior="{}""#, sync_behavior.as_str()));
    }
    if let Some(after_effect) = anim.after_effect {
        xml.push_str(if after_effect {
            r#" afterEffect="1""#
        } else {
            r#" afterEffect="0""#
        });
    }
    if let Some(time_filter) = &anim.time_filter {
        xml.push_str(&format!(r#" tmFilter="{}""#, time_filter.write_value()));
    }
    if let Some(group_id) = anim.group_id {
        xml.push_str(&format!(r#" grpId="{}""#, group_id.value()));
    }
    xml.push_str(&format!(
        r#" nodeType="{}" dur="{}">"#,
        node_type,
        anim.duration.write_value()
    ));
    *tn_id += 1;

    xml.push_str("<p:childTnLst>");
    xml.push_str(&format!(r#"<p:set><p:cBhvr><p:cTn id="{}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>"#, *tn_id));
    *tn_id += 1;
    xml.push_str(&format!(
        r#"<p:tgtEl><p:spTgt spid="{}"/></p:tgtEl>"#,
        anim.shape_id
    ));
    xml.push_str(r#"<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>"#);
    xml.push_str("</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>");
}

#[derive(Clone, Copy)]
enum TimingValue {
    Indefinite,
    Milliseconds(u32),
}

struct TimeNodeFrame {
    depth: usize,
    start_delay: Option<TimingValue>,
    start_on_click: bool,
    start_target: Option<u32>,
    interactive_event_filter: Option<Option<AnimationEventFilter>>,
}

struct PendingAnimation {
    depth: usize,
    animation: Animation,
    target: Option<u32>,
    target_element_depth: Option<usize>,
}

struct PendingParagraphTemplate {
    depth: usize,
    build_index: usize,
    level: u8,
    time_list_depth: Option<usize>,
    saw_time_list: bool,
    root_depth: Option<usize>,
    root_start: Option<usize>,
    root_range: Option<Range<usize>>,
}

struct PendingGraphicBuild {
    depth: usize,
    shape_id: u32,
    group_id: AnimationGroupId,
    ui_expand: bool,
    sub_build_depth: Option<usize>,
    mode: Option<AnimationGraphicBuildMode>,
}

struct TimingParser {
    sequence: AnimationSequence,
    shape_ids: HashSet<u32>,
    time_nodes: Vec<TimeNodeFrame>,
    pending: Vec<PendingAnimation>,
    timing_depth: Option<usize>,
    start_conditions_depth: Vec<usize>,
    condition_depth: Vec<usize>,
    condition_target_depth: Option<usize>,
    build_list_depth: Option<usize>,
    saw_build_list: bool,
    timing_group_ids: HashSet<AnimationGroupId>,
    build_group_ids: HashSet<AnimationGroupId>,
    build_pairs: HashSet<(u8, u32, AnimationGroupId)>,
    paragraph_build_depth: Option<usize>,
    paragraph_build_index: Option<usize>,
    template_list_depth: Option<usize>,
    template_levels: HashSet<u8>,
    pending_template: Option<PendingParagraphTemplate>,
    template_ranges: Vec<(usize, u8, Range<usize>)>,
    diagram_build_depth: Option<usize>,
    ole_chart_build_depth: Option<usize>,
    pending_graphic_build: Option<PendingGraphicBuild>,
    graphic_frame_depth: Option<usize>,
    graphic_depth: Option<usize>,
    graphic_data_depth: Option<usize>,
    graphic_frame_shape_id: Option<u32>,
    graphic_frame_has_ole_object: bool,
    graphic_frame_has_ole_chart: bool,
    graphic_frame_has_chart: bool,
    graphic_frame_has_diagram: bool,
    ole_diagram_shape_ids: HashSet<u32>,
    ole_chart_shape_ids: HashSet<u32>,
    chart_shape_ids: HashSet<u32>,
    graphical_diagram_shape_ids: HashSet<u32>,
    saw_timing: bool,
    require_valid_targets: bool,
    timing_start: Option<usize>,
    timing_range: Option<Range<usize>>,
}

impl TimingParser {
    fn new(require_valid_targets: bool) -> Self {
        Self {
            sequence: AnimationSequence::new(),
            shape_ids: HashSet::new(),
            time_nodes: Vec::new(),
            pending: Vec::new(),
            timing_depth: None,
            start_conditions_depth: Vec::new(),
            condition_depth: Vec::new(),
            condition_target_depth: None,
            build_list_depth: None,
            saw_build_list: false,
            timing_group_ids: HashSet::new(),
            build_group_ids: HashSet::new(),
            build_pairs: HashSet::new(),
            paragraph_build_depth: None,
            paragraph_build_index: None,
            template_list_depth: None,
            template_levels: HashSet::new(),
            pending_template: None,
            template_ranges: Vec::new(),
            diagram_build_depth: None,
            ole_chart_build_depth: None,
            pending_graphic_build: None,
            graphic_frame_depth: None,
            graphic_depth: None,
            graphic_data_depth: None,
            graphic_frame_shape_id: None,
            graphic_frame_has_ole_object: false,
            graphic_frame_has_ole_chart: false,
            graphic_frame_has_chart: false,
            graphic_frame_has_diagram: false,
            ole_diagram_shape_ids: HashSet::new(),
            ole_chart_shape_ids: HashSet::new(),
            chart_shape_ids: HashSet::new(),
            graphical_diagram_shape_ids: HashSet::new(),
            saw_timing: false,
            require_valid_targets,
            timing_start: None,
            timing_range: None,
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        depth: usize,
        empty: bool,
        event_start: usize,
        event_end: usize,
    ) -> Result<()> {
        check_attribute_count(element)?;

        if self.require_valid_targets
            && is_presentationml_name(namespace, element.name(), b"graphicFrame")
            && !empty
        {
            if self.graphic_frame_depth.is_some() {
                return Err(invalid("nested graphic frames are not supported"));
            }
            self.graphic_frame_depth = Some(depth);
            self.graphic_depth = None;
            self.graphic_data_depth = None;
            self.graphic_frame_shape_id = None;
            self.graphic_frame_has_ole_object = false;
            self.graphic_frame_has_ole_chart = false;
            self.graphic_frame_has_chart = false;
            self.graphic_frame_has_diagram = false;
        }

        if self.require_valid_targets
            && is_presentationml_name(namespace, element.name(), b"cNvPr")
            && let Some(value) = attribute(element, b"id", decoder)?
        {
            let id = parse_shape_id(&value)?;
            if self.graphic_frame_depth.is_some() {
                if self.graphic_frame_shape_id.is_none() {
                    self.graphic_frame_shape_id = Some(id);
                    if !self.shape_ids.insert(id) {
                        return Err(invalid("duplicate shape ID in slide"));
                    }
                }
            } else if !self.shape_ids.insert(id) {
                return Err(invalid("duplicate shape ID in slide"));
            }
        }

        if self.require_valid_targets
            && self
                .graphic_data_depth
                .is_some_and(|data_depth| depth == data_depth + 1)
            && is_presentationml_name(namespace, element.name(), b"oleObj")
        {
            if self.graphic_frame_has_ole_object {
                return Err(invalid("graphic frame has multiple direct OLE objects"));
            }
            self.graphic_frame_has_ole_object = true;
            self.graphic_frame_has_ole_chart = attribute(element, b"progId", decoder)?
                .as_deref()
                .map(is_known_ole_chart_program_id)
                .unwrap_or(true);
        }

        if self.require_valid_targets
            && self
                .graphic_frame_depth
                .is_some_and(|frame_depth| depth == frame_depth + 1)
            && is_drawingml_name(namespace, element.name(), b"graphic")
        {
            if self.graphic_depth.is_some() {
                return Err(invalid("graphic frame has multiple direct graphic hosts"));
            }
            if !empty {
                self.graphic_depth = Some(depth);
            }
        }

        if self.require_valid_targets
            && self
                .graphic_depth
                .is_some_and(|graphic_depth| depth == graphic_depth + 1)
            && is_drawingml_name(namespace, element.name(), b"graphicData")
        {
            if self.graphic_data_depth.is_some() {
                return Err(invalid(
                    "graphic host has multiple direct graphic-data elements",
                ));
            }
            if !empty {
                self.graphic_data_depth = Some(depth);
            }
        }

        if self.require_valid_targets
            && self
                .graphic_data_depth
                .is_some_and(|data_depth| depth == data_depth + 1)
        {
            if is_chartml_name(namespace, element.name(), b"chart") {
                if self.graphic_frame_has_chart || self.graphic_frame_has_diagram {
                    return Err(invalid(
                        "graphic frame has duplicate or ambiguous subtype markers",
                    ));
                }
                self.graphic_frame_has_chart = true;
            }
            if is_namespace_name(namespace, element.name(), DIAGRAM_NS, b"relIds") {
                if self.graphic_frame_has_chart || self.graphic_frame_has_diagram {
                    return Err(invalid(
                        "graphic frame has duplicate or ambiguous subtype markers",
                    ));
                }
                self.graphic_frame_has_diagram = true;
            }
        }

        if is_presentationml_name(namespace, element.name(), b"timing") {
            if self.saw_timing {
                return Err(invalid("slide contains multiple timing trees"));
            }
            self.saw_timing = true;
            if empty {
                self.timing_range = Some(event_start..event_end);
            } else {
                self.timing_depth = Some(depth);
                self.timing_start = Some(event_start);
            }
            return Ok(());
        }

        let Some(timing_depth) = self.timing_depth else {
            return Ok(());
        };
        if depth <= timing_depth {
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"bldLst") {
            if self.saw_build_list {
                return Err(invalid("timing tree contains multiple build lists"));
            }
            self.saw_build_list = true;
            if !empty {
                self.build_list_depth = Some(depth);
            }
            return Ok(());
        }

        if self.diagram_build_depth.is_some() {
            return Err(invalid(
                "diagram build elements cannot contain child elements",
            ));
        }
        if self.ole_chart_build_depth.is_some() {
            return Err(invalid(
                "OLE chart build elements cannot contain child elements",
            ));
        }

        if let Some(pending) = self.pending_graphic_build.as_mut() {
            if depth == pending.depth + 1 {
                if pending.mode.is_some() || pending.sub_build_depth.is_some() {
                    return Err(invalid(
                        "graphical-object build has multiple content choices",
                    ));
                }
                if is_presentationml_name(namespace, element.name(), b"bldAsOne") {
                    if !empty {
                        return Err(invalid("graphical-object build-as-one must be empty"));
                    }
                    pending.mode = Some(AnimationGraphicBuildMode::AsOne);
                    return Ok(());
                }
                if is_presentationml_name(namespace, element.name(), b"bldSub") {
                    if empty {
                        return Err(invalid(
                            "graphical-object sub-build is missing its build type",
                        ));
                    }
                    pending.sub_build_depth = Some(depth);
                    return Ok(());
                }
                return Err(invalid("graphical-object build has invalid child content"));
            }
            if pending
                .sub_build_depth
                .is_some_and(|sub_depth| depth == sub_depth + 1)
            {
                if pending.mode.is_some() {
                    return Err(invalid(
                        "graphical-object sub-build has multiple build types",
                    ));
                }
                if !empty {
                    return Err(invalid("graphical-object DrawingML build must be empty"));
                }
                if is_drawingml_name(namespace, element.name(), b"bldDgm") {
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| AnimationGraphicDiagramBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let reverse = attribute(element, b"rev", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    pending.mode = Some(AnimationGraphicBuildMode::Diagram {
                        build_type,
                        reverse,
                    });
                    return Ok(());
                }
                if is_drawingml_name(namespace, element.name(), b"bldChart") {
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| AnimationGraphicChartBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(true);
                    pending.mode = Some(AnimationGraphicBuildMode::Chart {
                        build_type,
                        animate_background,
                    });
                    return Ok(());
                }
                return Err(invalid(
                    "graphical-object sub-build has invalid DrawingML content",
                ));
            }
            return Err(invalid("graphical-object build has invalid nested content"));
        }

        if let Some(pending) = self.pending_template.as_mut() {
            if let Some(root_depth) = pending.root_depth {
                if depth > root_depth {
                    return Ok(());
                }
            }
            if let Some(time_list_depth) = pending.time_list_depth {
                if depth == time_list_depth + 1 {
                    if !is_presentationml_name(namespace, element.name(), b"par") {
                        return Err(invalid(
                            "paragraph template time list must contain a par node",
                        ));
                    }
                    if pending.root_start.is_some() || pending.root_range.is_some() {
                        return Err(invalid(
                            "paragraph template time list has multiple root nodes",
                        ));
                    }
                    if empty {
                        return Err(invalid("paragraph template par node cannot be empty"));
                    }
                    pending.root_depth = Some(depth);
                    pending.root_start = Some(event_start);
                    return Ok(());
                }
                return Err(invalid(
                    "paragraph template time list has invalid content order",
                ));
            }
            if depth == pending.depth + 1
                && is_presentationml_name(namespace, element.name(), b"tnLst")
            {
                if pending.saw_time_list {
                    return Err(invalid("paragraph template has multiple time lists"));
                }
                if empty {
                    return Err(invalid("paragraph template time list cannot be empty"));
                }
                pending.saw_time_list = true;
                pending.time_list_depth = Some(depth);
                return Ok(());
            }
            return Err(invalid("paragraph template has invalid child content"));
        }

        if let Some(template_list_depth) = self.template_list_depth {
            if depth != template_list_depth + 1
                || !is_presentationml_name(namespace, element.name(), b"tmpl")
            {
                return Err(invalid("paragraph template list has invalid child content"));
            }
            if self.template_levels.len() >= MAX_PARAGRAPH_TEMPLATES {
                return Err(invalid("paragraph template count exceeds PowerPoint limit"));
            }
            let level = attribute(element, b"lvl", decoder)?
                .map(|value| {
                    value
                        .parse::<u8>()
                        .map_err(|_| invalid("invalid paragraph template level"))
                })
                .transpose()?
                .unwrap_or(0);
            if level > 9 {
                return Err(invalid("paragraph template level exceeds PowerPoint limit"));
            }
            if !self.template_levels.insert(level) {
                return Err(invalid("duplicate paragraph template level"));
            }
            if empty {
                return Err(invalid("paragraph template is missing its time list"));
            }
            let build_index = self
                .paragraph_build_index
                .ok_or_else(|| invalid("paragraph template has no containing build"))?;
            self.pending_template = Some(PendingParagraphTemplate {
                depth,
                build_index,
                level,
                time_list_depth: None,
                saw_time_list: false,
                root_depth: None,
                root_start: None,
                root_range: None,
            });
            return Ok(());
        }

        if let Some(paragraph_build_depth) = self.paragraph_build_depth {
            if depth != paragraph_build_depth + 1
                || !is_presentationml_name(namespace, element.name(), b"tmplLst")
            {
                return Err(invalid("paragraph build has invalid child content"));
            }
            if self.template_list_depth.is_some() {
                return Err(invalid("paragraph build has multiple template lists"));
            }
            if !empty {
                self.template_list_depth = Some(depth);
                self.template_levels.clear();
            }
            return Ok(());
        }

        if self
            .build_list_depth
            .is_some_and(|list_depth| depth == list_depth + 1)
        {
            let kind = if is_presentationml_name(namespace, element.name(), b"bldP") {
                Some(1)
            } else if is_presentationml_name(namespace, element.name(), b"bldDgm") {
                Some(2)
            } else if is_presentationml_name(namespace, element.name(), b"bldGraphic") {
                Some(3)
            } else if is_presentationml_name(namespace, element.name(), b"bldOleChart") {
                Some(4)
            } else {
                None
            };
            if let Some(kind) = kind {
                if self.build_pairs.len() >= MAX_ANIMATION_BUILDS {
                    return Err(invalid("animation build count exceeds safety limit"));
                }
                let shape_id = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation build is missing spid"))
                    .and_then(|value| parse_shape_id(&value))?;
                let group_id = attribute(element, b"grpId", decoder)?
                    .ok_or_else(|| invalid("animation build is missing grpId"))
                    .and_then(|value| parse_group_id(&value))?;
                if !self.build_pairs.insert((kind, shape_id, group_id)) {
                    return Err(invalid("duplicate animation build shape/group pair"));
                }
                self.build_group_ids.insert(group_id);
                if kind == 1 {
                    let build_type = attribute(element, b"build", decoder)?
                        .map(|value| AnimationParagraphBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_level_attribute = attribute(element, b"bldLvl", decoder)?;
                    let build_level = build_level_attribute
                        .as_deref()
                        .map(|value| {
                            value
                                .parse::<u32>()
                                .map_err(|_| invalid("invalid paragraph build level"))
                        })
                        .transpose()?
                        .unwrap_or(1);
                    if build_level_attribute.is_some()
                        && build_type != AnimationParagraphBuildType::Paragraph
                    {
                        return Err(invalid(
                            "bldLvl is only supported when paragraph build type is p",
                        ));
                    }
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let auto_update_animate_background =
                        attribute(element, b"autoUpdateAnimBg", decoder)?
                            .map(|value| parse_xml_bool(&value))
                            .transpose()?
                            .unwrap_or(true);
                    let reverse_attribute = attribute(element, b"rev", decoder)?;
                    let reverse = reverse_attribute
                        .as_deref()
                        .map(parse_xml_bool)
                        .transpose()?
                        .unwrap_or(false);
                    if reverse_attribute.is_some()
                        && build_type != AnimationParagraphBuildType::Paragraph
                    {
                        return Err(invalid(
                            "rev is only supported when paragraph build type is p",
                        ));
                    }
                    let auto_advance = attribute(element, b"advAuto", decoder)?
                        .map(|value| parse_build_auto_advance(&value))
                        .transpose()?
                        .unwrap_or(Duration::Indefinite);
                    self.sequence
                        .paragraph_builds
                        .push(AnimationParagraphBuild {
                            shape_id,
                            group_id,
                            ui_expand,
                            build_type,
                            build_level,
                            animate_background,
                            auto_update_animate_background,
                            reverse,
                            auto_advance,
                            templates: Vec::new(),
                        });
                    if !empty {
                        self.paragraph_build_depth = Some(depth);
                        self.paragraph_build_index = Some(self.sequence.paragraph_builds.len() - 1);
                    }
                } else if kind == 2 {
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| AnimationDiagramBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    self.sequence.diagram_builds.push(AnimationDiagramBuild {
                        shape_id,
                        group_id,
                        ui_expand,
                        build_type,
                    });
                    if !empty {
                        self.diagram_build_depth = Some(depth);
                    }
                } else if kind == 3 {
                    if empty {
                        return Err(invalid(
                            "graphical-object build is missing its content choice",
                        ));
                    }
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    self.pending_graphic_build = Some(PendingGraphicBuild {
                        depth,
                        shape_id,
                        group_id,
                        ui_expand,
                        sub_build_depth: None,
                        mode: None,
                    });
                } else if kind == 4 {
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| AnimationOleChartBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(true);
                    self.sequence.ole_chart_builds.push(AnimationOleChartBuild {
                        shape_id,
                        group_id,
                        ui_expand,
                        build_type,
                        animate_background,
                    });
                    if !empty {
                        self.ole_chart_build_depth = Some(depth);
                    }
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"stCondLst") {
            if !empty {
                self.start_conditions_depth.push(depth);
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"cond")
            && !self.start_conditions_depth.is_empty()
        {
            let value = attribute(element, b"delay", decoder)?
                .map(|value| parse_timing_value(&value))
                .transpose()?
                .unwrap_or(TimingValue::Milliseconds(0));
            let frame = self
                .time_nodes
                .last_mut()
                .ok_or_else(|| invalid("animation condition has no containing time node"))?;
            if frame.start_delay.is_none() {
                frame.start_delay = Some(value);
                frame.start_on_click =
                    attribute(element, b"evt", decoder)?.as_deref() == Some("onClick");
            }
            if !empty {
                self.condition_depth.push(depth);
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"tgtEl") {
            if !self.condition_depth.is_empty() {
                if self.condition_target_depth.replace(depth).is_some() {
                    return Err(invalid("animation condition has multiple target elements"));
                }
                return Ok(());
            }
            if let Some(pending) = self.pending.last_mut() {
                if pending.target_element_depth.is_some() {
                    return Err(invalid("animation has multiple target elements"));
                }
                if !empty {
                    pending.target_element_depth = Some(depth);
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"spTgt") {
            if self.condition_target_depth.is_some() {
                let value = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation condition shape target is missing spid"))?;
                let id = parse_shape_id(&value)?;
                let frame = self
                    .time_nodes
                    .last_mut()
                    .ok_or_else(|| invalid("animation condition has no containing time node"))?;
                if frame.start_target.replace(id).is_some() {
                    return Err(invalid("animation condition has multiple shape targets"));
                }
                return Ok(());
            }
            if let Some(pending) = self.pending.last_mut()
                && pending.target_element_depth.is_some()
            {
                let value = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation shape target is missing spid"))?;
                let id = parse_shape_id(&value)?;
                if pending.target.replace(id).is_some() {
                    return Err(invalid("animation has multiple shape targets"));
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"cTn") {
            let node_type = attribute(element, b"nodeType", decoder)?;
            let event_filter = attribute(element, b"evtFilter", decoder)?;
            let group_id = attribute(element, b"grpId", decoder)?
                .map(|value| parse_group_id(&value))
                .transpose()?;
            if let Some(group_id) = group_id {
                self.timing_group_ids.insert(group_id);
            }
            let is_interactive = node_type.as_deref() == Some("interactiveSeq");
            if event_filter.is_some() && !is_interactive {
                return Err(invalid(
                    "animation event filter is only valid on an interactive sequence",
                ));
            }
            let interactive_event_filter = if is_interactive {
                Some(
                    event_filter
                        .map(|value| AnimationEventFilter::parse(&value))
                        .transpose()?,
                )
            } else {
                None
            };
            let preset_id = attribute(element, b"presetID", decoder)?;
            if let Some(preset_id) = preset_id {
                if is_interactive {
                    return Err(invalid(
                        "interactive sequence cannot also be a preset effect",
                    ));
                }
                if self.sequence.len() >= MAX_ANIMATIONS {
                    return Err(invalid("slide animation count exceeds safety limit"));
                }
                let preset_id = preset_id
                    .parse::<u32>()
                    .map_err(|_| invalid("invalid animation preset ID"))?;
                let preset_class = attribute(element, b"presetClass", decoder)?
                    .unwrap_or_else(|| "entr".to_string());
                if !matches!(
                    preset_class.as_str(),
                    "entr" | "exit" | "emph" | "path" | "verb" | "mediacall"
                ) {
                    return Err(invalid("invalid animation preset class"));
                }
                let preset_subtype = attribute(element, b"presetSubtype", decoder)?
                    .map(|value| {
                        value
                            .parse::<u32>()
                            .map_err(|_| invalid("invalid animation preset subtype"))
                    })
                    .transpose()?
                    .unwrap_or(0);
                let duration = match attribute(element, b"dur", decoder)? {
                    Some(value) => match parse_timing_value(&value)? {
                        TimingValue::Milliseconds(value) => Duration::Finite(value),
                        TimingValue::Indefinite => Duration::Indefinite,
                    },
                    None => Duration::Finite(0),
                };
                let fill = attribute(element, b"fill", decoder)?
                    .map(|value| AnimationFill::parse(&value))
                    .transpose()?;
                let restart = attribute(element, b"restart", decoder)?
                    .map(|value| AnimationRestart::parse(&value))
                    .transpose()?;
                let auto_reverse = attribute(element, b"autoRev", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?
                    .unwrap_or(false);
                let repeat = attribute(element, b"repeatCount", decoder)?
                    .map(|value| {
                        Ok::<AnimationRepeat, OoxmlError>(match parse_timing_value(&value)? {
                            TimingValue::Milliseconds(value) => AnimationRepeat::Finite(value),
                            TimingValue::Indefinite => AnimationRepeat::Indefinite,
                        })
                    })
                    .transpose()?;
                let speed = attribute(element, b"spd", decoder)?
                    .map(|value| {
                        let value = value
                            .parse::<i32>()
                            .map_err(|_| invalid("invalid animation speed percentage"))?;
                        AnimationSpeed::new(value)
                    })
                    .transpose()?;
                let acceleration = attribute(element, b"accel", decoder)?
                    .map(|value| parse_progress(&value, "acceleration"))
                    .transpose()?;
                let deceleration = attribute(element, b"decel", decoder)?
                    .map(|value| parse_progress(&value, "deceleration"))
                    .transpose()?;
                let display = attribute(element, b"display", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?;
                let repeat_duration = attribute(element, b"repeatDur", decoder)?
                    .map(|value| {
                        Ok::<Duration, OoxmlError>(match parse_timing_value(&value)? {
                            TimingValue::Milliseconds(value) => Duration::Finite(value),
                            TimingValue::Indefinite => Duration::Indefinite,
                        })
                    })
                    .transpose()?;
                let sync_behavior = attribute(element, b"syncBehavior", decoder)?
                    .map(|value| AnimationSyncBehavior::parse(&value))
                    .transpose()?;
                let after_effect = attribute(element, b"afterEffect", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?;
                let time_filter = attribute(element, b"tmFilter", decoder)?
                    .map(|value| AnimationTimeFilter::parse(&value))
                    .transpose()?;
                let sequence_context = parse_sequence_context(&self.time_nodes)?;
                let trigger = trigger(node_type.as_deref(), &self.time_nodes);
                let delay = self
                    .time_nodes
                    .iter()
                    .rev()
                    .find_map(|node| match node.start_delay {
                        Some(TimingValue::Milliseconds(value)) => Some(value),
                        _ => None,
                    })
                    .unwrap_or(0);
                let order = u32::try_from(self.sequence.len() + 1)
                    .map_err(|_| invalid("animation order exceeds u32"))?;
                let effect = AnimationEffect::from_preset_parts(&preset_class, preset_id);
                self.pending.push(PendingAnimation {
                    depth,
                    animation: Animation {
                        shape_id: 0,
                        direction: direction_from_subtype(&effect, preset_subtype),
                        effect,
                        trigger,
                        duration,
                        delay,
                        fill,
                        restart,
                        auto_reverse,
                        repeat,
                        speed,
                        acceleration,
                        deceleration,
                        display,
                        repeat_duration,
                        sync_behavior,
                        after_effect,
                        time_filter,
                        sequence_context,
                        group_id,
                        order,
                    },
                    target: None,
                    target_element_depth: None,
                });
                if empty {
                    return Err(invalid("preset animation has no shape target"));
                }
            }
            if !empty {
                self.time_nodes.push(TimeNodeFrame {
                    depth,
                    start_delay: None,
                    start_on_click: false,
                    start_target: None,
                    interactive_event_filter,
                });
            }
        }

        Ok(())
    }

    fn end(
        &mut self,
        namespace: &ResolveResult<'_>,
        name: quick_xml::name::QName<'_>,
        depth: usize,
        event_end: usize,
    ) -> Result<()> {
        if self.require_valid_targets
            && self.graphic_data_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphicData")
        {
            self.graphic_data_depth = None;
        }
        if self.require_valid_targets
            && self.graphic_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphic")
        {
            if self.graphic_data_depth.is_some() {
                return Err(invalid("graphic frame has an incomplete graphic-data host"));
            }
            self.graphic_depth = None;
        }
        if self.require_valid_targets
            && self.graphic_frame_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"graphicFrame")
        {
            if self.graphic_frame_has_ole_object
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.ole_diagram_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_ole_chart
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.ole_chart_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_chart
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.chart_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_diagram
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.graphical_diagram_shape_ids.insert(shape_id);
            }
            self.graphic_frame_depth = None;
            self.graphic_depth = None;
            self.graphic_data_depth = None;
            self.graphic_frame_shape_id = None;
            self.graphic_frame_has_ole_object = false;
            self.graphic_frame_has_ole_chart = false;
            self.graphic_frame_has_chart = false;
            self.graphic_frame_has_diagram = false;
        }

        if self.diagram_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldDgm")
        {
            self.diagram_build_depth = None;
            return Ok(());
        }
        if self.ole_chart_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldOleChart")
        {
            self.ole_chart_build_depth = None;
            return Ok(());
        }
        if let Some(pending) = self.pending_graphic_build.as_mut() {
            if pending.sub_build_depth == Some(depth)
                && is_presentationml_name(namespace, name, b"bldSub")
            {
                if pending.mode.is_none() {
                    return Err(invalid(
                        "graphical-object sub-build is missing its build type",
                    ));
                }
                pending.sub_build_depth = None;
                return Ok(());
            }
            if pending.depth == depth && is_presentationml_name(namespace, name, b"bldGraphic") {
                let pending = self
                    .pending_graphic_build
                    .take()
                    .expect("pending graphical-object build checked above");
                if pending.sub_build_depth.is_some() {
                    return Err(invalid(
                        "graphical-object build has an incomplete sub-build",
                    ));
                }
                let mode = pending.mode.ok_or_else(|| {
                    invalid("graphical-object build is missing its content choice")
                })?;
                self.sequence.graphic_builds.push(AnimationGraphicBuild {
                    shape_id: pending.shape_id,
                    group_id: pending.group_id,
                    ui_expand: pending.ui_expand,
                    mode,
                });
                return Ok(());
            }
        }
        if let Some(pending) = self.pending_template.as_mut() {
            if let Some(root_depth) = pending.root_depth {
                if depth > root_depth {
                    return Ok(());
                }
                if depth == root_depth && is_presentationml_name(namespace, name, b"par") {
                    let start = pending
                        .root_start
                        .take()
                        .ok_or_else(|| invalid("paragraph template root offset is missing"))?;
                    pending.root_range = Some(start..event_end);
                    pending.root_depth = None;
                    return Ok(());
                }
            }
            if pending.time_list_depth == Some(depth)
                && is_presentationml_name(namespace, name, b"tnLst")
            {
                if pending.root_range.is_none() {
                    return Err(invalid("paragraph template time list has no root par node"));
                }
                pending.time_list_depth = None;
                return Ok(());
            }
            if pending.depth == depth && is_presentationml_name(namespace, name, b"tmpl") {
                let pending = self
                    .pending_template
                    .take()
                    .expect("pending template checked above");
                if pending.time_list_depth.is_some() || !pending.saw_time_list {
                    return Err(invalid("paragraph template has an incomplete time list"));
                }
                let range = pending
                    .root_range
                    .ok_or_else(|| invalid("paragraph template has no root time node"))?;
                self.template_ranges
                    .push((pending.build_index, pending.level, range));
                return Ok(());
            }
            return Ok(());
        }

        if self.template_list_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"tmplLst")
        {
            self.template_list_depth = None;
            self.template_levels.clear();
            return Ok(());
        }

        if self.paragraph_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldP")
        {
            self.paragraph_build_depth = None;
            self.paragraph_build_index = None;
            return Ok(());
        }

        if is_presentationml_name(namespace, name, b"tgtEl") {
            if self.condition_target_depth == Some(depth) {
                self.condition_target_depth = None;
            }
            if let Some(pending) = self.pending.last_mut()
                && pending.target_element_depth == Some(depth)
            {
                pending.target_element_depth = None;
            }
        }

        if is_presentationml_name(namespace, name, b"cond")
            && self.condition_depth.last() == Some(&depth)
        {
            self.condition_depth.pop();
        }

        if is_presentationml_name(namespace, name, b"cTn") {
            if self
                .pending
                .last()
                .is_some_and(|pending| pending.depth == depth)
            {
                let mut pending = self.pending.pop().expect("pending animation checked above");
                pending.animation.shape_id = pending
                    .target
                    .ok_or_else(|| invalid("preset animation has no shape target"))?;
                self.sequence.add(pending.animation);
            }
            let frame = self
                .time_nodes
                .pop()
                .ok_or_else(|| invalid("unbalanced animation time node"))?;
            if frame.depth != depth {
                return Err(invalid("unbalanced animation time-node depth"));
            }
        }

        if is_presentationml_name(namespace, name, b"stCondLst")
            && self.start_conditions_depth.last() == Some(&depth)
        {
            self.start_conditions_depth.pop();
        }
        if is_presentationml_name(namespace, name, b"bldLst")
            && self.build_list_depth == Some(depth)
        {
            self.build_list_depth = None;
        }
        if is_presentationml_name(namespace, name, b"timing") && self.timing_depth == Some(depth) {
            self.timing_depth = None;
            let start = self
                .timing_start
                .take()
                .ok_or_else(|| invalid("timing subtree start offset is missing"))?;
            self.timing_range = Some(start..event_end);
        }
        Ok(())
    }

    fn finish(mut self, xml: &[u8]) -> Result<AnimationSequence> {
        if !self.pending.is_empty()
            || !self.time_nodes.is_empty()
            || self.timing_depth.is_some()
            || self.build_list_depth.is_some()
            || self.paragraph_build_depth.is_some()
            || self.template_list_depth.is_some()
            || self.pending_template.is_some()
            || self.diagram_build_depth.is_some()
            || self.ole_chart_build_depth.is_some()
            || self.pending_graphic_build.is_some()
            || self.graphic_frame_depth.is_some()
            || self.graphic_depth.is_some()
            || self.graphic_data_depth.is_some()
        {
            return Err(invalid("incomplete animation timing tree"));
        }
        if self.timing_group_ids != self.build_group_ids {
            return Err(invalid(
                "animation cTn group IDs and build-list group IDs do not match",
            ));
        }
        for (build_index, level, range) in self.template_ranges {
            let raw = xml
                .get(range)
                .ok_or_else(|| invalid("paragraph template range is outside slide XML"))?;
            if raw.len() > MAX_TEMPLATE_TIME_NODE_BYTES {
                return Err(invalid("paragraph template time node exceeds safety limit"));
            }
            let raw = std::str::from_utf8(raw)
                .map_err(|_| invalid("paragraph template time node is not UTF-8"))?;
            let build = self
                .sequence
                .paragraph_builds
                .get_mut(build_index)
                .ok_or_else(|| invalid("paragraph template build index is invalid"))?;
            build.templates.push(AnimationParagraphTemplate {
                level,
                time_node: AnimationTemplateTimeNode::parse(raw)?,
            });
        }
        for build in &self.sequence.paragraph_builds {
            if build.build_type == AnimationParagraphBuildType::Whole && build.templates.len() > 1 {
                return Err(invalid(
                    "whole paragraph builds support exactly one template effect",
                ));
            }
        }
        if self.require_valid_targets {
            for animation in &self.sequence.animations {
                if !self.shape_ids.contains(&animation.shape_id) {
                    return Err(invalid(format!(
                        "animation target {} is not a shape on the current slide",
                        animation.shape_id
                    )));
                }
                if let AnimationSequenceContext::Interactive {
                    trigger_shape_id, ..
                } = &animation.sequence_context
                    && !self.shape_ids.contains(trigger_shape_id)
                {
                    return Err(invalid(format!(
                        "interactive animation trigger {} is not a shape on the current slide",
                        trigger_shape_id
                    )));
                }
            }
            for (_, shape_id, _) in &self.build_pairs {
                if !self.shape_ids.contains(shape_id) {
                    return Err(invalid(format!(
                        "animation build target {} is not a shape on the current slide",
                        shape_id
                    )));
                }
            }
            for (kind, shape_id, _) in &self.build_pairs {
                if *kind == 2 && !self.ole_diagram_shape_ids.contains(shape_id) {
                    return Err(invalid(format!(
                        "diagram build target {} is not an OLE graphic-frame shape",
                        shape_id
                    )));
                }
            }
            for build in &self.sequence.graphic_builds {
                let valid = match build.mode {
                    AnimationGraphicBuildMode::AsOne => {
                        self.chart_shape_ids.contains(&build.shape_id)
                            || self.graphical_diagram_shape_ids.contains(&build.shape_id)
                    },
                    AnimationGraphicBuildMode::Diagram { .. } => {
                        self.graphical_diagram_shape_ids.contains(&build.shape_id)
                    },
                    AnimationGraphicBuildMode::Chart { .. } => {
                        self.chart_shape_ids.contains(&build.shape_id)
                    },
                };
                if !valid {
                    return Err(invalid(format!(
                        "graphical-object build target {} does not match its chart/diagram build type",
                        build.shape_id
                    )));
                }
            }
            for build in &self.sequence.ole_chart_builds {
                if !self.ole_chart_shape_ids.contains(&build.shape_id) {
                    return Err(invalid(format!(
                        "OLE chart build target {} is not an embedded chart graphic-frame shape",
                        build.shape_id
                    )));
                }
            }
        }
        if let Some(range) = self.timing_range {
            let raw = xml
                .get(range)
                .ok_or_else(|| invalid("timing subtree range is outside slide XML"))?;
            if raw.len() > MAX_PRESERVED_TIMING_BYTES {
                return Err(invalid("preserved timing subtree exceeds safety limit"));
            }
            let raw =
                std::str::from_utf8(raw).map_err(|_| invalid("timing subtree is not UTF-8"))?;
            self.sequence.source_animations =
                Some(self.sequence.animations.clone().into_boxed_slice());
            self.sequence.source_paragraph_builds =
                Some(self.sequence.paragraph_builds.clone().into_boxed_slice());
            self.sequence.source_diagram_builds =
                Some(self.sequence.diagram_builds.clone().into_boxed_slice());
            self.sequence.source_graphic_builds =
                Some(self.sequence.graphic_builds.clone().into_boxed_slice());
            self.sequence.source_ole_chart_builds =
                Some(self.sequence.ole_chart_builds.clone().into_boxed_slice());
            let slide_xml =
                std::str::from_utf8(xml).map_err(|_| invalid("slide timing XML is not UTF-8"))?;
            let timing_tree = parse_recursive_timing_tree(slide_xml)?;
            self.sequence.source_timing_tree = Some(Box::new(timing_tree.clone()));
            self.sequence.timing_tree = Some(timing_tree);
            self.sequence.source_timing_xml = Some(raw.to_string().into_boxed_str());
        }
        Ok(self.sequence)
    }
}

fn parse_sequence_context(time_nodes: &[TimeNodeFrame]) -> Result<AnimationSequenceContext> {
    let Some((interactive_index, event_filter)) = time_nodes
        .iter()
        .enumerate()
        .rev()
        .find_map(|(index, frame)| frame.interactive_event_filter.map(|filter| (index, filter)))
    else {
        return Ok(AnimationSequenceContext::Main);
    };
    let trigger_shape_id = time_nodes[interactive_index..]
        .iter()
        .find_map(|frame| frame.start_on_click.then_some(frame.start_target).flatten())
        .ok_or_else(|| {
            invalid("interactive animation sequence lacks a shape-targeted onClick condition")
        })?;
    Ok(AnimationSequenceContext::Interactive {
        trigger_shape_id,
        event_filter,
    })
}

struct RecursiveNodeFrame {
    depth: usize,
    sub_node: bool,
    node: AnimationTimingNode,
}

fn parse_recursive_timing_tree(xml: &str) -> Result<AnimationTimingTree> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut count = 0usize;
    let mut timing_depth = None;
    let mut timing_start = None;
    let mut frames = Vec::<RecursiveNodeFrame>::new();
    let mut roots = Vec::new();
    let mut child_lists = Vec::<(usize, bool)>::new();
    let mut condition_lists = Vec::<(usize, bool)>::new();
    let mut condition: Option<(usize, bool, AnimationTimeCondition)> = None;
    let mut source_range = None;
    loop {
        let event_start = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                depth += 1;
                count += 1;
                if depth > MAX_TIMING_DEPTH || count > MAX_TIMING_NODES {
                    return Err(invalid("animation timing tree exceeds safety limit"));
                }
                check_attribute_count(element)?;
                let empty = matches!(event, Event::Empty(_));
                if timing_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"timing")
                {
                    timing_depth = Some(depth);
                    timing_start = Some(event_start);
                } else if timing_depth.is_some() {
                    let local = element.local_name();
                    if is_presentationml_name(&namespace, element.name(), b"par")
                        || is_presentationml_name(&namespace, element.name(), b"seq")
                        || is_presentationml_name(&namespace, element.name(), b"excl")
                    {
                        if empty {
                            return Err(invalid("animation time container cannot be empty"));
                        }
                        let kind = if local.as_ref() == b"seq" {
                            let concurrent = attribute(element, b"concurrent", reader.decoder())?
                                .map(|v| parse_xml_bool(&v))
                                .transpose()?
                                .unwrap_or(false);
                            let next_action =
                                match attribute(element, b"nextAc", reader.decoder())?
                                    .as_deref()
                                    .unwrap_or("none")
                                {
                                    "none" => AnimationNextAction::None,
                                    "seek" => AnimationNextAction::Seek,
                                    _ => return Err(invalid("invalid animation next action")),
                                };
                            let previous_action =
                                match attribute(element, b"prevAc", reader.decoder())?
                                    .as_deref()
                                    .unwrap_or("none")
                                {
                                    "none" => AnimationPreviousAction::None,
                                    "skipTimed" => AnimationPreviousAction::SkipTimed,
                                    _ => return Err(invalid("invalid animation previous action")),
                                };
                            AnimationTimingNodeKind::Sequence {
                                concurrent,
                                next_action,
                                previous_action,
                            }
                        } else if local.as_ref() == b"excl" {
                            AnimationTimingNodeKind::Exclusive
                        } else {
                            AnimationTimingNodeKind::Parallel
                        };
                        frames.push(RecursiveNodeFrame {
                            depth,
                            sub_node: child_lists.last().is_some_and(|(_, sub)| *sub),
                            node: AnimationTimingNode {
                                kind,
                                common: AnimationCommonTimeNode {
                                id: None,
                                    duration: None,
                                    node_type: None,
                                    preset: None,
                                    start_conditions: Vec::new(),
                                    end_conditions: Vec::new(),
                                    children: Vec::new(),
                                    sub_nodes: Vec::new(),
                                    opaque_children: Vec::new(),
                                },
                                opaque_children: Vec::new(),
                            },
                        });
                    } else if is_presentationml_name(&namespace, element.name(), b"cTn")
                        && frames
                            .last()
                            .is_some_and(|frame| depth == frame.depth + 1)
                    {
                        let frame = frames
                            .last_mut()
                            .ok_or_else(|| invalid("common time node has no container"))?;
                        frame.node.common.id = attribute(element, b"id", reader.decoder())?
                            .map(|value| {
                                value
                                    .parse::<u32>()
                                    .map_err(|_| invalid("invalid common time-node ID"))
                            })
                            .transpose()?;
                        frame.node.common.duration = attribute(element, b"dur", reader.decoder())?
                            .map(|v| parse_timing_value(&v))
                            .transpose()?
                            .map(|v| match v {
                                TimingValue::Indefinite => Duration::Indefinite,
                                TimingValue::Milliseconds(ms) => Duration::Finite(ms),
                            });
                        frame.node.common.node_type =
                            attribute(element, b"nodeType", reader.decoder())?
                                .map(|v| AnimationTimeNodeType::parse(&v))
                                .transpose()?;
                        if let Some(value) = attribute(element, b"presetID", reader.decoder())? {
                            let preset_id = value
                                .parse::<u32>()
                                .map_err(|_| invalid("invalid animation preset ID"))?;
                            let class = AnimationPresetClass::parse(
                                attribute(element, b"presetClass", reader.decoder())?
                                    .as_deref()
                                    .unwrap_or("entr"),
                            )?;
                            let subtype = attribute(element, b"presetSubtype", reader.decoder())?
                                .map(|v| {
                                    v.parse::<u32>()
                                        .map_err(|_| invalid("invalid animation preset subtype"))
                                })
                                .transpose()?;
                            frame.node.common.preset = Some(AnimationPresetTimeNode {
                                preset_id,
                                class,
                                subtype,
                            });
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"childTnLst")
                        || is_presentationml_name(&namespace, element.name(), b"subTnLst")
                    {
                        if !empty {
                            child_lists.push((depth, local.as_ref() == b"subTnLst"));
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"stCondLst")
                        || is_presentationml_name(&namespace, element.name(), b"endCondLst")
                    {
                        if !empty {
                            condition_lists.push((depth, local.as_ref() == b"stCondLst"));
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"cond")
                        && !condition_lists.is_empty()
                    {
                        let delay = attribute(element, b"delay", reader.decoder())?
                            .map(|v| parse_timing_value(&v))
                            .transpose()?
                            .unwrap_or(TimingValue::Milliseconds(0));
                        let delay = match delay {
                            TimingValue::Indefinite => Duration::Indefinite,
                            TimingValue::Milliseconds(ms) => Duration::Finite(ms),
                        };
                        let current = AnimationTimeCondition {
                            event: attribute(element, b"evt", reader.decoder())?
                                .map(|v| AnimationConditionEvent::parse(&v))
                                .transpose()?,
                            delay,
                            target: None,
                        };
                        let start = condition_lists.last().expect("checked above").1;
                        if empty {
                            let common = &mut frames
                                .last_mut()
                                .ok_or_else(|| invalid("condition has no common time node"))?
                                .node
                                .common;
                            if start {
                                common.start_conditions.push(current)
                            } else {
                                common.end_conditions.push(current)
                            }
                        } else if condition.replace((depth, start, current)).is_some() {
                            return Err(invalid("nested animation conditions are invalid"));
                        }
                    } else if let Some((_, _, current)) = condition.as_mut() {
                        if is_presentationml_name(&namespace, element.name(), b"spTgt") {
                            current.target = Some(AnimationConditionTarget::Shape(parse_shape_id(
                                &attribute(element, b"spid", reader.decoder())?.ok_or_else(
                                    || invalid("condition shape target is missing its ID"),
                                )?,
                            )?));
                        } else if is_presentationml_name(&namespace, element.name(), b"sldTgt") {
                            current.target = Some(AnimationConditionTarget::Slide);
                        } else if is_presentationml_name(&namespace, element.name(), b"tn") {
                            let id = attribute(element, b"val", reader.decoder())?
                                .ok_or_else(|| {
                                    invalid("condition time-node target is missing its ID")
                                })?
                                .parse::<u32>()
                                .map_err(|_| invalid("invalid condition time-node ID"))?;
                            current.target = Some(AnimationConditionTarget::TimeNode(id));
                        } else if is_presentationml_name(&namespace, element.name(), b"rtn") {
                            current.target = Some(AnimationConditionTarget::Runtime(
                                AnimationRuntimeTrigger::parse(
                                    &attribute(element, b"val", reader.decoder())?.ok_or_else(
                                        || invalid("runtime condition target is missing its value"),
                                    )?,
                                )?,
                            ));
                        }
                    }
                }
                if empty {
                    depth -= 1;
                }
            },
            Event::End(name) => {
                if condition.as_ref().is_some_and(|(d, _, _)| *d == depth)
                    && is_presentationml_name(&namespace, name.name(), b"cond")
                {
                    let (_, start, value) = condition.take().expect("checked above");
                    let common = &mut frames
                        .last_mut()
                        .ok_or_else(|| invalid("condition has no common time node"))?
                        .node
                        .common;
                    if start {
                        common.start_conditions.push(value)
                    } else {
                        common.end_conditions.push(value)
                    }
                }
                if condition_lists.last().is_some_and(|(d, _)| *d == depth) {
                    condition_lists.pop();
                }
                if child_lists.last().is_some_and(|(d, _)| *d == depth) {
                    child_lists.pop();
                }
                if frames.last().is_some_and(|frame| frame.depth == depth) {
                    let frame = frames.pop().expect("checked above");
                    let child = AnimationTimingChild::Node(frame.node);
                    if let Some(parent) = frames.last_mut() {
                        if frame.sub_node {
                            parent.node.common.sub_nodes.push(child)
                        } else {
                            parent.node.common.children.push(child)
                        }
                    } else {
                        roots.push(child);
                    }
                }
                if timing_depth == Some(depth)
                    && is_presentationml_name(&namespace, name.name(), b"timing")
                {
                    source_range =
                        Some(timing_start.expect("timing start set")..reader.buffer_position());
                    timing_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced animation timing XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if timing_depth.is_some() || !frames.is_empty() || condition.is_some() {
        return Err(invalid("incomplete recursive animation timing tree"));
    }
    let range = source_range.ok_or_else(|| invalid("animation timing subtree is missing"))?;
    let range = (range.start as usize)..(range.end as usize);
    let source = xml
        .get(range)
        .ok_or_else(|| invalid("animation timing range is invalid"))?
        .to_string()
        .into_boxed_str();
    let mut tree = AnimationTimingTree {
        roots,
        opaque_children: Vec::new(),
        source_xml: Some(source),
        source_roots: None,
        source_opaque_children: None,
    };
    tree.source_roots = Some(tree.roots.clone().into_boxed_slice());
    tree.source_opaque_children = Some(tree.opaque_children.clone().into_boxed_slice());
    Ok(tree)
}

fn write_timing_child(xml: &mut String, child: &AnimationTimingChild) {
    let AnimationTimingChild::Node(node) = child else {
        if let AnimationTimingChild::Opaque(raw) = child {
            xml.push_str(raw);
        }
        return;
    };
    match node.kind {
        AnimationTimingNodeKind::Parallel => xml.push_str("<p:par>"),
        AnimationTimingNodeKind::Exclusive => xml.push_str("<p:excl>"),
        AnimationTimingNodeKind::Sequence {
            concurrent,
            next_action,
            previous_action,
        } => {
            xml.push_str("<p:seq");
            if concurrent {
                xml.push_str(" concurrent=\"1\"");
            }
            if next_action == AnimationNextAction::Seek {
                xml.push_str(" nextAc=\"seek\"");
            }
            if previous_action == AnimationPreviousAction::SkipTimed {
                xml.push_str(" prevAc=\"skipTimed\"");
            }
            xml.push('>');
        },
    }
    let common = &node.common;
    xml.push_str("<p:cTn");
    if let Some(id) = common.id {
        xml.push_str(&format!(" id=\"{id}\""));
    }
    if let Some(duration) = common.duration {
        xml.push_str(&format!(" dur=\"{}\"", duration.write_value()));
    }
    if let Some(node_type) = common.node_type {
        xml.push_str(&format!(" nodeType=\"{}\"", node_type.as_str()));
    }
    if let Some(preset) = &common.preset {
        xml.push_str(&format!(
            " presetID=\"{}\" presetClass=\"{}\"",
            preset.preset_id,
            preset.class.as_str()
        ));
        if let Some(subtype) = preset.subtype {
            xml.push_str(&format!(" presetSubtype=\"{}\"", subtype));
        }
    }
    xml.push('>');
    write_condition_list(xml, "stCondLst", &common.start_conditions);
    write_condition_list(xml, "endCondLst", &common.end_conditions);
    if !common.children.is_empty() {
        xml.push_str("<p:childTnLst>");
        for child in &common.children {
            write_timing_child(xml, child);
        }
        xml.push_str("</p:childTnLst>");
    }
    if !common.sub_nodes.is_empty() {
        xml.push_str("<p:subTnLst>");
        for child in &common.sub_nodes {
            write_timing_child(xml, child);
        }
        xml.push_str("</p:subTnLst>");
    }
    for raw in &common.opaque_children {
        xml.push_str(raw);
    }
    xml.push_str("</p:cTn>");
    for raw in &node.opaque_children {
        xml.push_str(raw);
    }
    match node.kind {
        AnimationTimingNodeKind::Parallel => xml.push_str("</p:par>"),
        AnimationTimingNodeKind::Exclusive => xml.push_str("</p:excl>"),
        AnimationTimingNodeKind::Sequence { .. } => xml.push_str("</p:seq>"),
    }
}

fn write_condition_list(xml: &mut String, name: &str, conditions: &[AnimationTimeCondition]) {
    if conditions.is_empty() {
        return;
    }
    xml.push_str(&format!("<p:{name}>"));
    for condition in conditions {
        xml.push_str("<p:cond");
        if let Some(event) = condition.event {
            xml.push_str(&format!(" evt=\"{}\"", event.as_str()));
        }
        xml.push_str(&format!(" delay=\"{}\"", condition.delay.write_value()));
        match condition.target {
            None => xml.push_str("/>"),
            Some(AnimationConditionTarget::Shape(id)) => xml.push_str(&format!(
                "><p:tgtEl><p:spTgt spid=\"{id}\"/></p:tgtEl></p:cond>"
            )),
            Some(AnimationConditionTarget::Slide) => {
                xml.push_str("><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>")
            },
            Some(AnimationConditionTarget::TimeNode(id)) => {
                xml.push_str(&format!("><p:tn val=\"{id}\"/></p:cond>"))
            },
            Some(AnimationConditionTarget::Runtime(value)) => {
                xml.push_str(&format!("><p:rtn val=\"{}\"/></p:cond>", value.as_str()))
            },
        }
    }
    xml.push_str(&format!("</p:{name}>"));
}

fn parse_group_id(value: &str) -> Result<AnimationGroupId> {
    value
        .parse::<u32>()
        .map(AnimationGroupId::new)
        .map_err(|_| invalid("invalid unsigned animation group ID"))
}

fn parse_build_auto_advance(value: &str) -> Result<Duration> {
    if value == "indefinite" {
        return Ok(Duration::Indefinite);
    }
    value
        .parse::<u32>()
        .map(Duration::Finite)
        .map_err(|_| invalid("invalid paragraph build auto-advance time"))
}

fn validate_template_time_node(xml: &str) -> Result<()> {
    if xml.len() > MAX_TEMPLATE_TIME_NODE_BYTES {
        return Err(invalid("paragraph template time node exceeds safety limit"));
    }
    let wrapped = format!(
        r#"<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{xml}</root>"#
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    let mut saw_par = false;
    let mut saw_ctn = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                nodes += 1;
                if depth > MAX_TIMING_DEPTH || nodes > MAX_TIMING_NODES {
                    return Err(invalid("paragraph template time node exceeds safety limit"));
                }
                check_attribute_count(&element)?;
                if depth == 2 {
                    if saw_par || !is_presentationml_name(&namespace, element.name(), b"par") {
                        return Err(invalid(
                            "paragraph template must contain exactly one par root",
                        ));
                    }
                    saw_par = true;
                } else if depth == 3 {
                    if saw_ctn || !is_presentationml_name(&namespace, element.name(), b"cTn") {
                        return Err(invalid(
                            "paragraph template par must contain exactly one cTn",
                        ));
                    }
                    saw_ctn = true;
                }
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_TIMING_NODES {
                    return Err(invalid("paragraph template time node exceeds safety limit"));
                }
                check_attribute_count(&element)?;
                let element_depth = depth + 1;
                if element_depth == 2 {
                    return Err(invalid("paragraph template par node cannot be empty"));
                }
                if element_depth == 3 {
                    if saw_ctn || !is_presentationml_name(&namespace, element.name(), b"cTn") {
                        return Err(invalid(
                            "paragraph template par must contain exactly one cTn",
                        ));
                    }
                    saw_ctn = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced paragraph template XML"))?;
            },
            Event::Text(text) => {
                text_bytes = text_bytes
                    .checked_add(text.len())
                    .ok_or_else(|| invalid("paragraph template text size overflows"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("paragraph template text exceeds safety limit"));
                }
                if depth <= 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("paragraph template has text outside its par root"));
                }
            },
            Event::CData(text) => {
                text_bytes = text_bytes
                    .checked_add(text.len())
                    .ok_or_else(|| invalid("paragraph template text size overflows"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("paragraph template text exceeds safety limit"));
                }
                if depth <= 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("paragraph template has text outside its par root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "active XML constructs are not allowed in paragraph templates",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !saw_par || !saw_ctn {
        return Err(invalid("incomplete paragraph template time node"));
    }
    Ok(())
}

fn parse_processed_timing(xml: &[u8], require_valid_targets: bool) -> Result<AnimationSequence> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut parser = TimingParser::new(require_valid_targets);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("animation XML offset does not fit usize"))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("animation XML offset does not fit usize"))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("animation XML node counter overflow"))?;
        if nodes > MAX_TIMING_NODES {
            return Err(invalid("animation XML node count exceeds safety limit"));
        }
        match event {
            Event::Start(element) => {
                parser.start(
                    &namespace,
                    &element,
                    decoder,
                    depth,
                    false,
                    event_start,
                    event_end,
                )?;
                depth += 1;
                if depth > MAX_TIMING_DEPTH {
                    return Err(invalid("animation XML depth exceeds safety limit"));
                }
            },
            Event::Empty(element) => parser.start(
                &namespace,
                &element,
                decoder,
                depth,
                true,
                event_start,
                event_end,
            )?,
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("animation XML has an unmatched end element"))?;
                parser.end(&namespace, element.name(), depth, event_end)?;
            },
            Event::Text(text) => {
                text_bytes = text_bytes
                    .checked_add(text.as_ref().len())
                    .ok_or_else(|| invalid("animation XML text counter overflow"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("animation XML text exceeds safety limit"));
                }
            },
            Event::CData(text) => {
                text_bytes = text_bytes
                    .checked_add(text.as_ref().len())
                    .ok_or_else(|| invalid("animation XML text counter overflow"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("animation XML text exceeds safety limit"));
                }
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in animation XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("incomplete animation XML"));
    }
    parser.finish(xml)
}

fn direction_subtype(effect: &AnimationEffect, direction: &AnimationDirection) -> Option<u32> {
    match effect {
        AnimationEffect::FlyIn | AnimationEffect::Wipe => Some(match direction {
            AnimationDirection::Up => 1,
            AnimationDirection::Right => 2,
            AnimationDirection::UpRight => 3,
            AnimationDirection::Down => 4,
            AnimationDirection::DownRight => 6,
            AnimationDirection::Left => 8,
            AnimationDirection::UpLeft => 9,
            AnimationDirection::DownLeft => 12,
            _ => return None,
        }),
        AnimationEffect::Split => Some(match direction {
            AnimationDirection::VerticalIn => 21,
            AnimationDirection::HorizontalIn => 26,
            AnimationDirection::VerticalOut => 37,
            AnimationDirection::HorizontalOut => 42,
            _ => return None,
        }),
        AnimationEffect::Zoom => Some(match direction {
            AnimationDirection::In => 16,
            AnimationDirection::Out => 32,
            AnimationDirection::OutFromScreenCenter => 36,
            AnimationDirection::InSlightly => 272,
            AnimationDirection::OutSlightly => 288,
            AnimationDirection::InFromScreenCenter => 528,
            _ => return None,
        }),
        _ => None,
    }
}

fn direction_from_subtype(effect: &AnimationEffect, subtype: u32) -> Option<AnimationDirection> {
    match effect {
        AnimationEffect::FlyIn | AnimationEffect::Wipe => match subtype {
            1 => Some(AnimationDirection::Up),
            2 => Some(AnimationDirection::Right),
            3 => Some(AnimationDirection::UpRight),
            4 => Some(AnimationDirection::Down),
            6 => Some(AnimationDirection::DownRight),
            8 => Some(AnimationDirection::Left),
            9 => Some(AnimationDirection::UpLeft),
            12 => Some(AnimationDirection::DownLeft),
            _ => None,
        },
        AnimationEffect::Split => match subtype {
            21 => Some(AnimationDirection::VerticalIn),
            26 => Some(AnimationDirection::HorizontalIn),
            37 => Some(AnimationDirection::VerticalOut),
            42 => Some(AnimationDirection::HorizontalOut),
            _ => None,
        },
        AnimationEffect::Zoom => match subtype {
            16 => Some(AnimationDirection::In),
            32 => Some(AnimationDirection::Out),
            36 => Some(AnimationDirection::OutFromScreenCenter),
            272 => Some(AnimationDirection::InSlightly),
            288 => Some(AnimationDirection::OutSlightly),
            528 => Some(AnimationDirection::InFromScreenCenter),
            _ => None,
        },
        _ => None,
    }
}

fn trigger(node_type: Option<&str>, ancestors: &[TimeNodeFrame]) -> AnimationTrigger {
    match node_type {
        Some("withEffect" | "withGroup") => AnimationTrigger::WithPrevious,
        Some("afterEffect" | "afterGroup") => AnimationTrigger::AfterPrevious,
        Some("clickEffect" | "clickPar") if ancestors.iter().any(|node| node.start_on_click) => {
            AnimationTrigger::OnClick
        },
        Some("clickEffect" | "clickPar") => {
            match ancestors.iter().find_map(|node| node.start_delay) {
                Some(TimingValue::Milliseconds(_)) => AnimationTrigger::WithPrevious,
                _ => AnimationTrigger::OnClick,
            }
        },
        _ => match ancestors.iter().find_map(|node| node.start_delay) {
            Some(TimingValue::Indefinite) => AnimationTrigger::OnClick,
            Some(TimingValue::Milliseconds(_)) => AnimationTrigger::WithPrevious,
            None => AnimationTrigger::OnClick,
        },
    }
}

fn parse_timing_value(value: &str) -> Result<TimingValue> {
    if value == "indefinite" {
        return Ok(TimingValue::Indefinite);
    }
    let value = value
        .parse::<u32>()
        .map_err(|_| invalid("invalid animation timing value"))?;
    if value > MAX_TIMING_MILLISECONDS {
        return Err(invalid(
            "animation timing value exceeds the supported OOXML limit",
        ));
    }
    Ok(TimingValue::Milliseconds(value))
}

fn parse_xml_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid animation boolean value")),
    }
}

fn parse_progress(value: &str, name: &str) -> Result<AnimationProgress> {
    let value = value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid animation {name} percentage")))?;
    AnimationProgress::new(value)
}

fn parse_shape_id(value: &str) -> Result<u32> {
    let id = value
        .parse::<u32>()
        .map_err(|_| invalid("invalid animation shape target ID"))?;
    if id == 0 {
        return Err(invalid("animation shape target ID must be nonzero"));
    }
    Ok(id)
}

fn attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<String>> {
    crate::common::xml::unqualified_attribute_value(element, name, decoder)
}

fn check_attribute_count(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        count += 1;
        if count > MAX_TIMING_ATTRIBUTES {
            return Err(invalid(
                "animation XML attribute count exceeds safety limit",
            ));
        }
    }
    Ok(())
}

fn check_xml_size(size: usize) -> Result<()> {
    if size > MAX_TIMING_XML_BYTES {
        Err(invalid("animation XML exceeds safety limit"))
    } else {
        Ok(())
    }
}

fn is_namespace_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected_namespace)
        && name.local_name().as_ref() == expected_local_name
}

fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    is_namespace_name(namespace, name, DRAWINGML_NS, local)
        || is_namespace_name(namespace, name, DRAWINGML_STRICT_NS, local)
}

fn is_chartml_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    is_namespace_name(namespace, name, CHART_NS, local)
        || is_namespace_name(namespace, name, CHART_STRICT_NS, local)
}

fn is_known_ole_chart_program_id(value: &str) -> bool {
    value == "Excel.Chart"
        || value.starts_with("Excel.Chart.")
        || value == "MSGraph.Chart"
        || value.starts_with("MSGraph.Chart.")
}

#[cfg(test)]
mod recursive_timing_tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    #[test]
    fn preserves_nested_presets_and_ordered_conditions() {
        let xml = format!(
            r#"<p:timing xmlns:p="{NS}"><p:tnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="1" nodeType="interactiveSeq" presetID="10" presetClass="entr"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cond><p:cond evt="onEnd" delay="25"><p:tn val="9"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="2" presetID="11" presetClass="emph"/></p:par></p:childTnLst></p:cTn><p:extLst><p:ext uri="opaque"><x:data xmlns:x="urn:test"/></p:ext></p:extLst></p:seq></p:tnLst></p:timing>"#
        );
        let tree = AnimationTimingTree::parse(&xml).expect("nested timing parses");
        assert_eq!(tree.to_xml(), xml);
        let AnimationTimingChild::Node(root) = &tree.roots[0] else {
            panic!("typed root")
        };
        assert!(matches!(
            root.kind,
            AnimationTimingNodeKind::Sequence {
                concurrent: true,
                next_action: AnimationNextAction::Seek,
                ..
            }
        ));
        assert_eq!(root.common.start_conditions.len(), 2);
        assert!(matches!(
            root.common.start_conditions[0].target,
            Some(AnimationConditionTarget::Shape(7))
        ));
        assert!(matches!(
            root.common.start_conditions[1].target,
            Some(AnimationConditionTarget::TimeNode(9))
        ));
        assert!(root.common.preset.is_some());
        let AnimationTimingChild::Node(child) = &root.common.children[0] else {
            panic!("typed child")
        };
        assert_eq!(
            child.common.preset.as_ref().map(|preset| preset.preset_id),
            Some(11)
        );
    }

    #[test]
    fn rejects_malformed_common_time_node_id() {
        let xml = format!(
            r#"<p:timing xmlns:p="{NS}"><p:tnLst><p:par><p:cTn id="not-a-number"/></p:par></p:tnLst></p:timing>"#
        );
        assert!(AnimationTimingTree::parse(&xml).is_err());
    }

    #[test]
    fn rejects_excessive_recursive_depth() {
        let mut xml = format!(r#"<p:timing xmlns:p="{NS}"><p:tnLst>"#);
        for id in 1..=MAX_TIMING_DEPTH + 1 {
            xml.push_str(&format!("<p:par><p:cTn id=\"{id}\"><p:childTnLst>"));
        }
        for _ in 1..=MAX_TIMING_DEPTH + 1 {
            xml.push_str("</p:childTnLst></p:cTn></p:par>");
        }
        xml.push_str("</p:tnLst></p:timing>");
        assert!(AnimationTimingTree::parse(&xml).is_err());
    }

    #[test]
    fn rejects_excessive_node_count() {
        let mut xml = format!(r#"<p:timing xmlns:p="{NS}"><p:tnLst>"#);
        for id in 1..=MAX_TIMING_NODES + 1 {
            xml.push_str(&format!("<p:par><p:cTn id=\"{id}\"/></p:par>"));
        }
        xml.push_str("</p:tnLst></p:timing>");
        assert!(AnimationTimingTree::parse(&xml).is_err());
    }
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn slide(timing: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="{P}"><p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="A"/></p:nvSpPr></p:sp>
                <p:pic><p:nvPicPr><p:cNvPr id="4" name="B"/></p:nvPicPr></p:pic>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="C"/></p:nvGraphicFramePr></p:graphicFrame>
            </p:spTree></p:cSld>{timing}</p:sld>"#
        )
    }

    fn effect(
        shape: &str,
        preset: u32,
        class: &str,
        node_type: &str,
        trigger_delay: &str,
        delay: u32,
        duration: u32,
    ) -> String {
        format!(
            r#"<p:par><p:cTn><p:stCondLst><p:cond delay="{trigger_delay}"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn presetID="{preset}" presetClass="{class}" presetSubtype="0" nodeType="{node_type}" dur="{duration}">
            <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst>
            </p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"#
        )
    }

    fn interactive_effect(trigger_shape: &str, event_filter: &str) -> String {
        let triggered = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replacen(
                r#"<p:cond delay="indefinite"/>"#,
                &format!(
                    r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_shape}"/></p:tgtEl></p:cond>"#
                ),
                1,
            );
        format!(
            r#"<p:seq><p:cTn nodeType="interactiveSeq" evtFilter="{event_filter}"><p:childTnLst>{triggered}</p:childTnLst></p:cTn></p:seq>"#
        )
    }

    fn grouped_timing(shape_id: &str, group_id: &str) -> String {
        let grouped = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldP spid="{shape_id}" grpId="{group_id}"/></p:bldLst></p:timing>"#
        )
    }

    fn diagram_timing(shape_id: &str, group_id: &str, attributes: &str) -> String {
        let grouped = effect("5", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldDgm spid="{shape_id}" grpId="{group_id}"{attributes}/></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_ole(timing: &str) -> String {
        slide(timing).replace(
            r#"</p:nvGraphicFramePr></p:graphicFrame>"#,
            r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj/></a:graphicData></a:graphic></p:graphicFrame>"#,
        )
    }

    fn ole_chart_timing(shape_id: &str, group_id: &str, attributes: &str) -> String {
        let grouped = effect(shape_id, 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldOleChart spid="{shape_id}" grpId="{group_id}"{attributes}/></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_ole_chart(timing: &str) -> String {
        slide_with_ole(timing).replace("<p:oleObj/>", r#"<p:oleObj progId="MSGraph.Chart.8"/>"#)
    }

    fn graphic_timing(shape_id: &str, group_id: &str, attributes: &str, content: &str) -> String {
        let grouped = effect(shape_id, 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldGraphic spid="{shape_id}" grpId="{group_id}"{attributes}>{content}</p:bldGraphic></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_graphic_hosts(timing: &str) -> String {
        slide(timing)
            .replace(
                r#"<p:sld xmlns:p=""#,
                r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p=""#,
            )
            .replace(
                r#"</p:nvGraphicFramePr></p:graphicFrame>"#,
                r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="SmartArt"/></p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphicData></a:graphic></p:graphicFrame>"#,
            )
    }

    #[test]
    fn test_animation_effect_preset() {
        assert_eq!(AnimationEffect::Fade.preset_class(), "entr");
        assert_eq!(AnimationEffect::Fade.preset_id(), 10);
        assert_eq!(AnimationEffect::from_preset("fade"), AnimationEffect::Fade);
    }

    #[test]
    fn test_animation_sequence() {
        let mut seq = AnimationSequence::new();
        seq.add(Animation::new(1, AnimationEffect::Fade).with_duration_ms(1000));
        seq.add(
            Animation::new(2, AnimationEffect::FlyIn).with_trigger(AnimationTrigger::AfterPrevious),
        );

        assert_eq!(seq.len(), 2);
        assert!(!seq.to_xml().is_empty());
    }

    #[test]
    fn parses_typed_timing_metadata_from_slide() {
        let timing = format!(
            "<p:timing><p:tnLst>{}{}{}</p:tnLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 125, 750),
            effect("4", 42, "entr", "withEffect", "0", 20, 600),
            effect("5", 8, "emph", "afterEffect", "0", 40, 900),
        );
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.len(), 3);
        assert_eq!(sequence.animations[0].shape_id, 3);
        assert_eq!(sequence.animations[0].effect, AnimationEffect::Fade);
        assert_eq!(sequence.animations[0].trigger, AnimationTrigger::OnClick);
        assert_eq!(sequence.animations[0].duration, Duration::Finite(750));
        assert_eq!(sequence.animations[0].delay, 125);
        assert_eq!(sequence.animations[1].effect, AnimationEffect::FloatIn);
        assert_eq!(
            sequence.animations[1].trigger,
            AnimationTrigger::WithPrevious
        );
        assert_eq!(sequence.animations[2].effect, AnimationEffect::Spin);
        assert_eq!(
            sequence.animations[2].trigger,
            AnimationTrigger::AfterPrevious
        );
        assert_eq!(sequence.animations[2].order, 3);
    }

    #[test]
    fn rejects_malformed_missing_duplicate_spoofed_and_off_slide_targets() {
        let cases = [
            effect("0", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("nope", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("99", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("<p:spTgt spid=\"3\"/>", ""),
        ];
        for effect in cases {
            let timing = format!("<p:timing><p:tnLst>{effect}</p:tnLst></p:timing>");
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }

        let duplicate = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            "<p:spTgt spid=\"3\"/>",
            "<p:spTgt spid=\"3\"/><p:spTgt spid=\"4\"/>",
        );
        let timing = format!("<p:timing><p:tnLst>{duplicate}</p:tnLst></p:timing>");
        assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());

        let spoofed = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            "<p:spTgt spid=\"3\"/>",
            "<x:spTgt xmlns:x=\"urn:foreign\" spid=\"3\"/>",
        );
        let timing = format!("<p:timing><p:tnLst>{spoofed}</p:tnLst></p:timing>");
        assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
    }

    #[test]
    fn rejects_excessive_timing_depth() {
        let nested = format!(
            "<p:timing>{}{}</p:timing>",
            "<p:par>".repeat(MAX_TIMING_DEPTH + 1),
            "</p:par>".repeat(MAX_TIMING_DEPTH + 1)
        );
        assert!(AnimationSequence::parse_slide_xml(slide(&nested).as_bytes()).is_err());
    }

    #[test]
    fn preserves_indefinite_duration() {
        let timing = format!(
            "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("dur=\"500\"", "dur=\"indefinite\"")
        );
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.animations[0].duration, Duration::Indefinite);
    }

    #[test]
    fn preserves_unsupported_timing_subtrees_until_typed_data_changes() {
        let timing = format!(
            "<p:timing><p:tnLst>{}</p:tnLst><p:extLst><p:ext uri=\"urn:test\"><x:opaque xmlns:x=\"urn:opaque\" value=\"kept\"><![CDATA[raw]]></x:opaque></p:ext></p:extLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
        );
        let mut sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.preserved_timing_xml(), Some(timing.as_str()));
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains("dur=\"750\""));
        assert!(!canonical.contains("x:opaque"));
    }

    #[test]
    fn parses_directional_preset_subtypes() {
        let timing = format!(
            "<p:timing><p:tnLst>{}{}{}{}</p:tnLst></p:timing>",
            effect("3", 2, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"3\""),
            effect("4", 22, "entr", "withEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"12\""),
            effect("5", 16, "entr", "afterEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"26\""),
            effect("3", 23, "entr", "withEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"288\"")
        );
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.animations[0].direction,
            Some(AnimationDirection::UpRight)
        );
        assert_eq!(
            sequence.animations[1].direction,
            Some(AnimationDirection::DownLeft)
        );
        assert_eq!(
            sequence.animations[2].direction,
            Some(AnimationDirection::HorizontalIn)
        );
        assert_eq!(
            sequence.animations[3].direction,
            Some(AnimationDirection::OutSlightly)
        );
    }

    #[test]
    fn parses_common_time_node_playback_controls() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " fill=\"freeze\" restart=\"whenNotActive\" autoRev=\"1\" repeatCount=\"3500\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(animation.fill, Some(AnimationFill::Freeze));
        assert_eq!(animation.restart, Some(AnimationRestart::WhenNotActive));
        assert!(animation.auto_reverse);
        assert_eq!(animation.repeat, Some(AnimationRepeat::Finite(3500)));
    }

    #[test]
    fn rejects_invalid_playback_control_values() {
        for attribute in [
            "fill=\"sticky\"",
            "restart=\"sometimes\"",
            "autoRev=\"yes\"",
            "repeatCount=\"2147483626\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_common_time_node_progression_controls() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " spd=\"-50000\" accel=\"25000\" decel=\"10000\" display=\"0\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(
            animation.speed.map(AnimationSpeed::thousandths_percent),
            Some(-50000)
        );
        assert_eq!(
            animation
                .acceleration
                .map(AnimationProgress::thousandths_percent),
            Some(25000)
        );
        assert_eq!(
            animation
                .deceleration
                .map(AnimationProgress::thousandths_percent),
            Some(10000)
        );
        assert_eq!(animation.display, Some(false));
    }

    #[test]
    fn rejects_invalid_progression_control_values() {
        for attribute in [
            "spd=\"0\"",
            "spd=\"2147483648\"",
            "accel=\"100001\"",
            "accel=\"-1\"",
            "decel=\"100001\"",
            "display=\"visible\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_repeat_duration_sync_and_after_effect() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " repeatDur=\"indefinite\" syncBehavior=\"locked\" afterEffect=\"1\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(animation.repeat_duration, Some(Duration::Indefinite));
        assert_eq!(animation.sync_behavior, Some(AnimationSyncBehavior::Locked));
        assert_eq!(animation.after_effect, Some(true));
    }

    #[test]
    fn rejects_invalid_repeat_duration_sync_and_after_effect() {
        for attribute in [
            "repeatDur=\"2147483626\"",
            "repeatDur=\"forever\"",
            "syncBehavior=\"slippery\"",
            "afterEffect=\"yes\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_exact_normalized_time_filter_pairs() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            " nodeType=\"clickEffect\"",
            " tmFilter=\"0.0,0.0; 0.25,0.07; 0.50,0.2; 1.0,1.0\" nodeType=\"clickEffect\"",
        );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let points = sequence.animations[0]
            .time_filter
            .as_ref()
            .unwrap()
            .points();
        assert_eq!(points.len(), 4);
        assert_eq!(
            (
                points[1].local_time.numerator(),
                points[1].local_time.scale()
            ),
            (25, 100)
        );
        assert_eq!(
            (
                points[1].warped_time.numerator(),
                points[1].warped_time.scale()
            ),
            (7, 100)
        );
    }

    #[test]
    fn rejects_malformed_out_of_range_or_unordered_time_filters() {
        for filter in [
            "",
            "0",
            "0,0,0",
            "-0.1,0",
            "0,1.0001",
            "0.5,0;0.5,1",
            "0.75,0;0.25,1",
            "0.1234567890123456789,0",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" tmFilter=\"{filter}\" nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_and_canonicalizes_contextual_cancel_bubble_filter() {
        let interactive = interactive_effect("4", "cancelBubble");
        let timing = format!("<p:timing><p:tnLst>{interactive}</p:tnLst></p:timing>");
        let mut sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.animations[0].sequence_context,
            AnimationSequenceContext::Interactive {
                trigger_shape_id: 4,
                event_filter: Some(AnimationEventFilter::CancelBubble),
            }
        );
        assert_eq!(sequence.preserved_timing_xml(), Some(timing.as_str()));
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"nodeType="interactiveSeq" evtFilter="cancelBubble""#));
        assert!(
            canonical.contains(r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"/>"#)
        );
        let reparsed = AnimationSequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_event_filter_outside_proven_triggered_sequence_context() {
        let on_effect = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            r#" evtFilter="cancelBubble" nodeType="clickEffect""#,
        );
        let cases = [
            format!("<p:timing><p:tnLst>{on_effect}</p:tnLst></p:timing>"),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "bubble")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble")
                    .replace("evt=\"onClick\"", "evt=\"onNext\"")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble")
                    .replace(r#"<p:tgtEl><p:spTgt spid="4"/></p:tgtEl>"#, "")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("99", "cancelBubble")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble").replace(
                    r#"<p:spTgt spid="4"/>"#,
                    r#"<p:spTgt spid="4"/><p:spTgt spid="5"/>"#,
                )
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble").replace(
                    r#"<p:spTgt spid="4"/>"#,
                    r#"<x:spTgt xmlns:x="urn:foreign" spid="4"/>"#,
                )
            ),
        ];
        for timing in cases {
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn writes_typed_interactive_context_and_validates_trigger_shape() {
        let mut sequence = AnimationSequence::new();
        sequence.add(
            Animation::new(3, AnimationEffect::Fade)
                .with_interactive_trigger(4)
                .with_trigger(AnimationTrigger::OnClick),
        );
        let targets = HashSet::from([3, 4]);
        let xml = sequence.to_xml_for_slide(&targets).unwrap();
        assert!(xml.contains(r#"nodeType="interactiveSeq" evtFilter="cancelBubble""#));
        let parsed = AnimationSequence::parse_slide_xml(slide(&xml).as_bytes()).unwrap();
        assert_eq!(parsed, sequence);

        sequence.animations[0].sequence_context = AnimationSequenceContext::Interactive {
            trigger_shape_id: 99,
            event_filter: Some(AnimationEventFilter::CancelBubble),
        };
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_paragraph_build_group_references() {
        let timing = grouped_timing("3", "42");
        let mut sequence = AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.animations[0].group_id,
            Some(AnimationGroupId::new(42))
        );
        assert_eq!(
            sequence.paragraph_builds,
            vec![AnimationParagraphBuild::new(3, AnimationGroupId::new(42))]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"grpId="42" nodeType="clickEffect""#));
        assert!(canonical.contains(r#"<p:bldLst><p:bldP spid="3" grpId="42"/></p:bldLst>"#));
        let reparsed = AnimationSequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_malformed_dangling_duplicate_or_off_slide_build_groups() {
        let cases = [
            grouped_timing("3", "-1"),
            grouped_timing("3", "4294967296"),
            grouped_timing("99", "42"),
            grouped_timing("3", "42").replace(
                r#" grpId="42" nodeType="clickEffect""#,
                r#" nodeType="clickEffect""#,
            ),
            grouped_timing("3", "42").replace(r#"<p:bldP spid="3" grpId="42"/>"#, ""),
            grouped_timing("3", "42")
                .replace(r#"<p:bldP spid="3" grpId="42"/>"#, r#"<p:bldP spid="3"/>"#),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<p:bldP grpId="42"/>"#,
            ),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<p:bldP spid="3" grpId="42"/><p:bldP spid="3" grpId="42"/>"#,
            ),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<x:bldP xmlns:x="urn:foreign" spid="3" grpId="42"/>"#,
            ),
        ];
        for timing in cases {
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_build_group_membership_and_targets() {
        let targets = HashSet::from([3]);
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(3, AnimationEffect::Fade).with_group_id(7));
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        sequence.add_paragraph_build(AnimationParagraphBuild::new(3, AnimationGroupId::new(7)));
        assert!(sequence.to_xml_for_slide(&targets).is_ok());

        sequence.paragraph_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_bldp_schema_defaults_and_powerpoint_auto_advance_semantics() {
        let sequence =
            AnimationSequence::parse_slide_xml(slide(&grouped_timing("3", "42")).as_bytes())
                .unwrap();
        let build = &sequence.paragraph_builds[0];
        assert!(!build.ui_expand);
        assert_eq!(build.build_type, AnimationParagraphBuildType::Whole);
        assert_eq!(build.build_level, 1);
        assert!(!build.animate_background);
        assert!(build.auto_update_animate_background);
        assert!(!build.reverse);
        assert_eq!(build.auto_advance, Duration::Indefinite);
        assert_eq!(build.powerpoint_auto_advance_milliseconds(), 0);
    }

    #[test]
    fn round_trips_complete_typed_bldp_optional_attribute_grammar() {
        let configured = grouped_timing("3", "42").replace(
            r#"<p:bldP spid="3" grpId="42"/>"#,
            r#"<p:bldP spid="3" grpId="42" uiExpand="true" build="p" bldLvl="3" animBg="1" autoUpdateAnimBg="false" rev="true" advAuto="4294967295"/>"#,
        );
        let mut sequence =
            AnimationSequence::parse_slide_xml(slide(&configured).as_bytes()).unwrap();
        let build = &sequence.paragraph_builds[0];
        assert!(build.ui_expand);
        assert_eq!(build.build_type, AnimationParagraphBuildType::Paragraph);
        assert_eq!(build.build_level, 3);
        assert!(build.animate_background);
        assert!(!build.auto_update_animate_background);
        assert!(build.reverse);
        assert_eq!(build.auto_advance, Duration::Finite(u32::MAX));

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldP spid="3" grpId="42" uiExpand="1" build="p" bldLvl="3" animBg="1" autoUpdateAnimBg="0" rev="1" advAuto="4294967295"/>"#
        ));
        let reparsed = AnimationSequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_invalid_bldp_optional_attributes_and_cross_field_combinations() {
        for attributes in [
            r#"build="paragraph""#,
            r#"uiExpand="yes""#,
            r#"build="whole" bldLvl="2""#,
            r#"build="whole" rev="1""#,
            r#"bldLvl="-1" build="p""#,
            r#"bldLvl="4294967296" build="p""#,
            r#"animBg="sometimes""#,
            r#"autoUpdateAnimBg="sometimes""#,
            r#"rev="sometimes" build="p""#,
            r#"advAuto="-1""#,
            r#"advAuto="4294967296""#,
            r#"advAuto="forever""#,
        ] {
            let timing = grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                &format!(r#"<p:bldP spid="3" grpId="42" {attributes}/>"#),
            );
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_paragraph_build_cross_field_constraints() {
        let targets = HashSet::from([3]);
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(3, AnimationEffect::Fade).with_group_id(7));
        sequence.add_paragraph_build(
            AnimationParagraphBuild::new(3, AnimationGroupId::new(7)).with_build_level(2),
        );
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        sequence.paragraph_builds[0] = sequence.paragraph_builds[0]
            .clone()
            .with_build_type(AnimationParagraphBuildType::Paragraph)
            .with_reverse(true)
            .with_auto_advance(250u32);
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
    }

    #[test]
    fn parses_preserves_and_canonicalizes_complete_paragraph_template_lists() {
        let first = r#"<p:par><p:cTn id="80" dur="500"/></p:par>"#;
        let second = r#"<p:par><p:cTn id="81" dur="indefinite"><p:childTnLst/></p:cTn></p:par>"#;
        let configured = grouped_timing("3", "42").replace(
            r#"<p:bldP spid="3" grpId="42"/>"#,
            &format!(r#"<p:bldP spid="3" grpId="42" build="p"><p:tmplLst><p:tmpl><p:tnLst>{first}</p:tnLst></p:tmpl><p:tmpl lvl="2"><p:tnLst>{second}</p:tnLst></p:tmpl></p:tmplLst></p:bldP>"#),
        );
        let mut sequence =
            AnimationSequence::parse_slide_xml(slide(&configured).as_bytes()).unwrap();
        let templates = &sequence.paragraph_builds[0].templates;
        assert_eq!(templates.len(), 2);
        assert_eq!(templates[0].level, 0);
        assert_eq!(templates[0].time_node.as_xml(), first);
        assert_eq!(templates[1].level, 2);
        assert_eq!(templates[1].time_node.as_xml(), second);
        assert_eq!(sequence.to_xml(), configured);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(&format!(r#"<p:tmplLst><p:tmpl><p:tnLst>{first}</p:tnLst></p:tmpl><p:tmpl lvl="2"><p:tnLst>{second}</p:tnLst></p:tmpl></p:tmplLst>"#)));
        let reparsed = AnimationSequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_invalid_paragraph_template_cardinality_levels_order_and_namespaces() {
        let par = r#"<p:par><p:cTn id="80"/></p:par>"#;
        let template = |level: &str, body: &str| {
            format!(r#"<p:tmpl{level}><p:tnLst>{body}</p:tnLst></p:tmpl>"#)
        };
        let lists = [
            template("", ""),
            r#"<p:tmpl/>"#.to_string(),
            format!(r#"<p:tmpl><p:tnLst>{par}{par}</p:tnLst></p:tmpl>"#),
            r#"<p:tmpl><p:tnLst><p:par><p:cTn/><p:cTn/></p:par></p:tnLst></p:tmpl>"#.to_string(),
            format!(r#"<p:tmpl><p:tnLst><p:seq><p:cTn/></p:seq></p:tnLst></p:tmpl>"#),
            format!(
                r#"<p:tmpl><p:tnLst><x:par xmlns:x="urn:foreign"><x:cTn/></x:par></p:tnLst></p:tmpl>"#
            ),
            format!(r#"<p:tmpl><p:tnLst>{par}</p:tnLst><p:tnLst>{par}</p:tnLst></p:tmpl>"#),
            format!(r#"<p:tmpl lvl="10"><p:tnLst>{par}</p:tnLst></p:tmpl>"#),
            format!(r#"<p:tmpl lvl="nope"><p:tnLst>{par}</p:tnLst></p:tmpl>"#),
            format!(r#"{}{}"#, template("", par), template("", par)),
            (0..10)
                .map(|level| template(&format!(r#" lvl="{level}""#), par))
                .collect::<String>(),
        ];
        for list in lists {
            let timing = grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                &format!(r#"<p:bldP spid="3" grpId="42" build="p"><p:tmplLst>{list}</p:tmplLst></p:bldP>"#),
            );
            assert!(AnimationSequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_paragraph_template_fragments_and_constraints() {
        for xml in [
            "",
            r#"<p:seq><p:cTn/></p:seq>"#,
            r#"<p:par/>"#,
            r#"<p:par><p:cTn/><p:cTn/></p:par>"#,
            r#"<x:par xmlns:x="urn:foreign"><x:cTn/></x:par>"#,
            r#"<!DOCTYPE x><p:par><p:cTn/></p:par>"#,
        ] {
            assert!(AnimationTemplateTimeNode::parse(xml).is_err());
        }

        let node = AnimationTemplateTimeNode::parse(r#"<p:par><p:cTn id="90"/></p:par>"#).unwrap();
        assert!(AnimationParagraphTemplate::new(10, node.clone()).is_err());
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(3, AnimationEffect::Fade).with_group_id(7));
        let duplicate = AnimationParagraphTemplate::new(1, node.clone()).unwrap();
        sequence.add_paragraph_build(
            AnimationParagraphBuild::new(3, AnimationGroupId::new(7))
                .with_build_type(AnimationParagraphBuildType::Paragraph)
                .with_template(duplicate.clone())
                .with_template(duplicate),
        );
        assert!(sequence.to_xml_for_slide(&HashSet::from([3])).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_diagram_build_grammar() {
        let timing = diagram_timing("5", "77", r#" uiExpand="true" bld="ccwOut""#);
        let mut sequence =
            AnimationSequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.diagram_builds,
            vec![AnimationDiagramBuild {
                shape_id: 5,
                group_id: AnimationGroupId::new(77),
                ui_expand: true,
                build_type: AnimationDiagramBuildType::CounterClockwiseOut,
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"<p:bldDgm spid="5" grpId="77" uiExpand="1" bld="ccwOut"/>"#));
        let reparsed =
            AnimationSequence::parse_slide_xml(slide_with_ole(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_every_diagram_build_enum_and_schema_defaults() {
        for token in [
            "whole",
            "depthByNode",
            "depthByBranch",
            "breadthByNode",
            "breadthByLvl",
            "cw",
            "cwIn",
            "cwOut",
            "ccw",
            "ccwIn",
            "ccwOut",
            "inByRing",
            "outByRing",
            "up",
            "down",
            "allAtOnce",
            "cust",
        ] {
            let timing = diagram_timing("5", "77", &format!(r#" bld="{token}""#));
            assert!(AnimationSequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).is_ok());
        }
        let timing = diagram_timing("5", "77", "");
        let sequence =
            AnimationSequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.diagram_builds[0].build_type,
            AnimationDiagramBuildType::Whole
        );
        assert!(!sequence.diagram_builds[0].ui_expand);
    }

    #[test]
    fn rejects_invalid_diagram_builds_and_non_ole_or_spoofed_targets() {
        let cases = [
            slide_with_ole(&diagram_timing("5", "77", r#" bld="sideways""#)),
            slide_with_ole(&diagram_timing("5", "77", r#" uiExpand="yes""#)),
            slide_with_ole(&diagram_timing("99", "77", "")),
            slide(&diagram_timing("5", "77", "")),
            slide_with_ole(&diagram_timing("3", "77", "")),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5" grpId="77"/><p:bldDgm spid="5" grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5" grpId="77"><p:extLst/></p:bldDgm>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<x:bldDgm xmlns:x="urn:foreign" spid="5" grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", ""))
                .replace(r#"<p:oleObj/>"#, r#"<x:oleObj xmlns:x="urn:foreign"/>"#),
        ];
        for xml in cases {
            assert!(AnimationSequence::parse_slide_xml(xml.as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_diagram_build_groups_targets_and_duplicates() {
        let targets = HashSet::from([5]);
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(5, AnimationEffect::Fade).with_group_id(77));
        sequence.add_diagram_build(
            AnimationDiagramBuild::new(5, AnimationGroupId::new(77))
                .with_ui_expand(true)
                .with_build_type(AnimationDiagramBuildType::BreadthByLevel),
        );
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.diagram_builds.push(sequence.diagram_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.diagram_builds.pop();
        sequence.diagram_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_graphic_chart_build() {
        let timing = graphic_timing(
            "5",
            "88",
            r#" uiExpand="true""#,
            r#"<p:bldSub><a:bldChart bld="seriesEl" animBg="false"/></p:bldSub>"#,
        );
        let mut sequence =
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes())
                .unwrap();
        assert_eq!(
            sequence.graphic_builds,
            vec![AnimationGraphicBuild {
                shape_id: 5,
                group_id: AnimationGroupId::new(88),
                ui_expand: true,
                mode: AnimationGraphicBuildMode::Chart {
                    build_type: AnimationGraphicChartBuildType::SeriesElement,
                    animate_background: false,
                },
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldGraphic spid="5" grpId="88" uiExpand="1"><p:bldSub><a:bldChart bld="seriesEl" animBg="0"/></p:bldSub></p:bldGraphic>"#
        ));
        let reparsed =
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&canonical).as_bytes())
                .unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_all_graphic_build_modes_tokens_and_schema_defaults() {
        let as_one = graphic_timing("5", "88", "", "<p:bldAsOne/>");
        let sequence =
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&as_one).as_bytes())
                .unwrap();
        assert_eq!(
            sequence.graphic_builds[0].mode,
            AnimationGraphicBuildMode::AsOne
        );

        for token in ["allAtOnce", "one", "lvlOne", "lvlAtOnce"] {
            let timing = graphic_timing(
                "6",
                "88",
                "",
                &format!(r#"<p:bldSub><a:bldDgm bld="{token}"/></p:bldSub>"#),
            );
            assert!(
                AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes())
                    .is_ok()
            );
        }
        for token in ["allAtOnce", "series", "category", "seriesEl", "categoryEl"] {
            let timing = graphic_timing(
                "5",
                "88",
                "",
                &format!(r#"<p:bldSub><a:bldChart bld="{token}"/></p:bldSub>"#),
            );
            assert!(
                AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes())
                    .is_ok()
            );
        }

        let diagram = graphic_timing("6", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>");
        let sequence =
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&diagram).as_bytes())
                .unwrap();
        assert_eq!(
            sequence.graphic_builds[0].mode,
            AnimationGraphicBuildMode::Diagram {
                build_type: AnimationGraphicDiagramBuildType::AllAtOnce,
                reverse: false,
            }
        );
        let chart = graphic_timing("5", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>");
        let sequence =
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&chart).as_bytes())
                .unwrap();
        assert_eq!(
            sequence.graphic_builds[0].mode,
            AnimationGraphicBuildMode::Chart {
                build_type: AnimationGraphicChartBuildType::AllAtOnce,
                animate_background: true,
            }
        );
    }

    #[test]
    fn rejects_hostile_graphic_build_grammar_namespaces_and_host_mismatches() {
        let valid_chart = graphic_timing("5", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>");
        let cases = [
            graphic_timing("5", "88", "", ""),
            graphic_timing("5", "88", "", "<p:bldSub/>"),
            graphic_timing("5", "88", "", "<p:bldAsOne><p:extLst/></p:bldAsOne>"),
            graphic_timing("5", "88", "", "<p:bldAsOne/><p:bldAsOne/>"),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart/><a:bldChart/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart><a:ext/></a:bldChart></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><x:bldChart xmlns:x=\"urn:foreign\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<x:bldSub xmlns:x=\"urn:foreign\"><a:bldChart/></x:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart bld=\"rows\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldDgm bld=\"rows\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart animBg=\"yes\"/></p:bldSub>",
            ),
            graphic_timing("5", "88", r#" uiExpand="yes""#, "<p:bldAsOne/>"),
            graphic_timing("6", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>"),
            graphic_timing("5", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>"),
            graphic_timing("3", "88", "", "<p:bldAsOne/>"),
            valid_chart.replace(r#" spid="5" grpId="88""#, r#" spid="5""#),
            valid_chart.replace(r#" spid="5" grpId="88""#, r#" grpId="88""#),
        ];
        for timing in cases {
            assert!(
                AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes())
                    .is_err()
            );
        }

        let spoofed_host = slide_with_graphic_hosts(&valid_chart).replace(
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
            r#"<x:chart xmlns:x="urn:foreign"/>"#,
        );
        assert!(AnimationSequence::parse_slide_xml(spoofed_host.as_bytes()).is_err());

        let chart_marker =
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let marker_elsewhere = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, "")
            .replace(
                r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
                &format!(r#"</p:nvGraphicFramePr>{chart_marker}<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#),
            );
        assert!(AnimationSequence::parse_slide_xml(marker_elsewhere.as_bytes()).is_err());

        let nested_marker = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, &format!(r#"<a:ext>{chart_marker}</a:ext>"#));
        assert!(AnimationSequence::parse_slide_xml(nested_marker.as_bytes()).is_err());

        let duplicate_marker = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, &format!(r#"{chart_marker}{chart_marker}"#));
        assert!(AnimationSequence::parse_slide_xml(duplicate_marker.as_bytes()).is_err());

        let ambiguous_marker = slide_with_graphic_hosts(&valid_chart).replace(
            chart_marker,
            &format!(r#"{chart_marker}<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#),
        );
        assert!(AnimationSequence::parse_slide_xml(ambiguous_marker.as_bytes()).is_err());

        let valid_diagram = graphic_timing("6", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>");
        let diagram_marker =
            r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#;
        let nested_diagram_marker = slide_with_graphic_hosts(&valid_diagram).replace(
            diagram_marker,
            &format!(r#"<a:ext>{diagram_marker}</a:ext>"#),
        );
        assert!(AnimationSequence::parse_slide_xml(nested_diagram_marker.as_bytes()).is_err());

        let duplicate = valid_chart.replace(
            "</p:bldLst>",
            r#"<p:bldGraphic spid="5" grpId="88"><p:bldAsOne/></p:bldGraphic></p:bldLst>"#,
        );
        assert!(
            AnimationSequence::parse_slide_xml(slide_with_graphic_hosts(&duplicate).as_bytes())
                .is_err()
        );
    }

    #[test]
    fn validates_programmatic_graphic_builds_and_combined_build_boundary() {
        let targets = HashSet::from([5]);
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(5, AnimationEffect::Fade).with_group_id(88));
        sequence.add_graphic_build(AnimationGraphicBuild::chart(5, AnimationGroupId::new(88)));
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.graphic_builds.push(sequence.graphic_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.graphic_builds.pop();
        sequence.graphic_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        let mut oversized = AnimationSequence::new();
        oversized.graphic_builds = vec![
            AnimationGraphicBuild::as_one(5, AnimationGroupId::new(88));
            MAX_ANIMATION_BUILDS + 1
        ];
        assert!(oversized.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_ole_chart_build_grammar() {
        let timing = ole_chart_timing(
            "5",
            "91",
            r#" uiExpand="true" bld="categoryEl" animBg="false""#,
        );
        let mut sequence =
            AnimationSequence::parse_slide_xml(slide_with_ole_chart(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.ole_chart_builds,
            vec![AnimationOleChartBuild {
                shape_id: 5,
                group_id: AnimationGroupId::new(91),
                ui_expand: true,
                build_type: AnimationOleChartBuildType::CategoryElement,
                animate_background: false,
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldOleChart spid="5" grpId="91" uiExpand="1" bld="categoryEl" animBg="0"/>"#
        ));
        let reparsed =
            AnimationSequence::parse_slide_xml(slide_with_ole_chart(&canonical).as_bytes())
                .unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_every_ole_chart_build_token_and_schema_defaults() {
        for token in ["allAtOnce", "series", "category", "seriesEl", "categoryEl"] {
            let timing = ole_chart_timing("5", "91", &format!(r#" bld="{token}""#));
            assert!(
                AnimationSequence::parse_slide_xml(slide_with_ole_chart(&timing).as_bytes())
                    .is_ok()
            );
        }
        let timing = ole_chart_timing("5", "91", "");
        let sequence =
            AnimationSequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.ole_chart_builds[0].build_type,
            AnimationOleChartBuildType::AllAtOnce
        );
        assert!(!sequence.ole_chart_builds[0].ui_expand);
        assert!(sequence.ole_chart_builds[0].animate_background);
    }

    #[test]
    fn rejects_hostile_invalid_and_non_chart_ole_builds() {
        let valid = ole_chart_timing("5", "91", "");
        let cases = [
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" bld="rows""#)),
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" uiExpand="yes""#)),
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" animBg="yes""#)),
            slide_with_ole_chart(&ole_chart_timing("99", "91", "")),
            slide(&valid),
            slide_with_ole_chart(&ole_chart_timing("3", "91", "")),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart spid="5"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart grpId="91"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart spid="5" grpId="91"><p:extLst/></p:bldOleChart>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<x:bldOleChart xmlns:x="urn:foreign" spid="5" grpId="91"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<p:oleObj progId="Word.Document.12"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<x:oleObj xmlns:x="urn:foreign" progId="MSGraph.Chart.8"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj progId="MSGraph.Chart.8"/></a:graphicData>"#,
                r#"<p:oleObj progId="MSGraph.Chart.8"/><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<p:oleObj progId="MSGraph.Chart.8"/><p:oleObj progId="MSGraph.Chart.8"/>"#,
            ),
        ];
        for xml in cases {
            assert!(AnimationSequence::parse_slide_xml(xml.as_bytes()).is_err());
        }

        let duplicate = slide_with_ole_chart(&valid).replace(
            "</p:bldLst>",
            r#"<p:bldOleChart spid="5" grpId="91"/></p:bldLst>"#,
        );
        assert!(AnimationSequence::parse_slide_xml(duplicate.as_bytes()).is_err());
    }

    #[test]
    fn validates_programmatic_ole_chart_builds_and_combined_boundary() {
        let targets = HashSet::from([5]);
        let mut sequence = AnimationSequence::new();
        sequence.add(Animation::new(5, AnimationEffect::Fade).with_group_id(91));
        sequence.add_ole_chart_build(
            AnimationOleChartBuild::new(5, AnimationGroupId::new(91))
                .with_ui_expand(true)
                .with_build_type(AnimationOleChartBuildType::Series)
                .with_animate_background(false),
        );
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.ole_chart_builds.push(sequence.ole_chart_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.ole_chart_builds.pop();
        sequence.ole_chart_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        let mut oversized = AnimationSequence::new();
        oversized.ole_chart_builds = vec![
            AnimationOleChartBuild::new(5, AnimationGroupId::new(91));
            MAX_ANIMATION_BUILDS + 1
        ];
        assert!(oversized.to_xml_for_slide(&targets).is_err());
    }
}
