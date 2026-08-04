use super::codec::{MAX_NORMALIZED_TIME_DECIMALS, MAX_TIME_FILTER_BYTES, MAX_TIME_FILTER_POINTS};
use super::invalid;
use crate::Result;
/// EffectInstance effect type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Effect {
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

impl Effect {
    /// Parse from preset ID.
    pub fn from_preset_id(id: u32) -> Self {
        match id {
            1 => Effect::Appear,
            2 => Effect::FlyIn,
            6 => Effect::GrowShrink,
            8 => Effect::Spin,
            10 => Effect::Fade,
            16 => Effect::Split,
            22 => Effect::Wipe,
            23 => Effect::Zoom,
            24 => Effect::Bounce,
            42 => Effect::FloatIn,
            _ => Effect::Custom(format!("preset_{}", id)),
        }
    }

    pub(super) fn from_preset_parts(class: &str, id: u32) -> Self {
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
            "entr" | "appear" => Effect::Appear,
            "fade" => Effect::Fade,
            "fly" | "flyin" => Effect::FlyIn,
            "float" | "floatin" => Effect::FloatIn,
            "split" => Effect::Split,
            "wipe" => Effect::Wipe,
            "zoom" => Effect::Zoom,
            "bounce" => Effect::Bounce,
            "spin" => Effect::Spin,
            "grow" | "growshrink" => Effect::GrowShrink,
            other => Effect::Custom(other.to_string()),
        }
    }

    /// Get the preset ID for this effect.
    /// These are defined in ECMA-376 Part 1.
    pub fn preset_id(&self) -> u32 {
        match self {
            Effect::Appear => 1,
            Effect::FlyIn => 2,
            Effect::FloatIn => 42,
            Effect::Split => 16,
            Effect::Fade => 10,
            Effect::Wipe => 22,
            Effect::Zoom => 23,
            Effect::Bounce => 24,
            Effect::Spin => 8,       // Spin is emphasis, but using ID 8
            Effect::GrowShrink => 6, // GrowShrink is emphasis
            Effect::Custom(value) => value
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
            Effect::Appear => "entr",
            Effect::FlyIn => "entr",
            Effect::FloatIn => "entr",
            Effect::Split => "entr",
            Effect::Fade => "entr",
            Effect::Wipe => "entr",
            Effect::Zoom => "entr",
            Effect::Bounce => "entr",
            // Emphasis effects
            Effect::Spin => "emph",
            Effect::GrowShrink => "emph",
            // Default to entrance
            Effect::Custom(value) => value
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

/// EffectInstance trigger type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Trigger {
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
pub struct GroupId(u32);

impl GroupId {
    /// Construct an OOXML timing group identifier.
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the encoded unsigned group identifier.
    pub const fn value(self) -> u32 {
        self.0
    }
}

impl From<u32> for GroupId {
    fn from(value: u32) -> Self {
        Self::new(value)
    }
}

/// Paragraph build mode from `ST_TLParaBuildType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ParagraphBuildType {
    AllAtOnce,
    Paragraph,
    Custom,
    /// Schema default: build the text shape as a whole.
    #[default]
    Whole,
}

impl ParagraphBuildType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "p" => Ok(Self::Paragraph),
            "cust" => Ok(Self::Custom),
            "whole" => Ok(Self::Whole),
            _ => Err(invalid("invalid paragraph build type")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
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
pub enum DiagramBuildType {
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

impl DiagramBuildType {
    pub(super) fn parse(value: &str) -> Result<Self> {
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

    pub(super) const fn as_str(self) -> &'static str {
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
pub struct DiagramBuild {
    /// OLE graphic-frame shape receiving the build.
    pub shape_id: u32,
    /// Timing group referenced by the build.
    pub group_id: GroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    /// Diagram build mode. Defaults to `Whole`.
    pub build_type: DiagramBuildType,
}

impl DiagramBuild {
    /// Construct a diagram build using schema defaults.
    pub const fn new(shape_id: u32, group_id: GroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: DiagramBuildType::Whole,
        }
    }

    /// Set whether this build appears expanded in animation UIs.
    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }

    /// Set the diagram build mode.
    pub fn with_build_type(mut self, build_type: DiagramBuildType) -> Self {
        self.build_type = build_type;
        self
    }
}

/// DrawingML diagram build mode used inside `p:bldGraphic/p:bldSub`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum GraphicDiagramBuildType {
    /// Schema default: animate all diagram content at once.
    #[default]
    AllAtOnce,
    One,
    LevelOne,
    LevelAtOnce,
}

impl GraphicDiagramBuildType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "one" => Ok(Self::One),
            "lvlOne" => Ok(Self::LevelOne),
            "lvlAtOnce" => Ok(Self::LevelAtOnce),
            _ => Err(invalid("invalid graphical-object diagram build type")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
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
pub enum GraphicChartBuildType {
    /// Schema default: animate all chart content at once.
    #[default]
    AllAtOnce,
    Series,
    Category,
    SeriesElement,
    CategoryElement,
}

impl GraphicChartBuildType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid graphical-object chart build type")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
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
pub enum GraphicBuildMode {
    /// Animate the chart or diagram as one graphical object.
    AsOne,
    /// Animate SmartArt/diagram sub-elements.
    Diagram {
        build_type: GraphicDiagramBuildType,
        reverse: bool,
    },
    /// Animate chart sub-elements.
    Chart {
        build_type: GraphicChartBuildType,
        animate_background: bool,
    },
}

/// Build information for a chart or SmartArt graphical frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GraphicBuild {
    pub shape_id: u32,
    pub group_id: GroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    pub mode: GraphicBuildMode,
}

impl GraphicBuild {
    pub const fn new(shape_id: u32, group_id: GroupId, mode: GraphicBuildMode) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            mode,
        }
    }

    pub const fn as_one(shape_id: u32, group_id: GroupId) -> Self {
        Self::new(shape_id, group_id, GraphicBuildMode::AsOne)
    }

    pub const fn diagram(shape_id: u32, group_id: GroupId) -> Self {
        Self::new(
            shape_id,
            group_id,
            GraphicBuildMode::Diagram {
                build_type: GraphicDiagramBuildType::AllAtOnce,
                reverse: false,
            },
        )
    }

    pub const fn chart(shape_id: u32, group_id: GroupId) -> Self {
        Self::new(
            shape_id,
            group_id,
            GraphicBuildMode::Chart {
                build_type: GraphicChartBuildType::AllAtOnce,
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
pub enum OleChartBuildType {
    #[default]
    AllAtOnce,
    Series,
    Category,
    SeriesElement,
    CategoryElement,
}

impl OleChartBuildType {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid OLE chart build type")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
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
pub struct OleChartBuild {
    pub shape_id: u32,
    pub group_id: GroupId,
    /// Whether the build is expanded in animation UIs. Defaults to `false`.
    pub ui_expand: bool,
    /// Chart build mode. Defaults to `AllAtOnce`.
    pub build_type: OleChartBuildType,
    /// Whether the chart background participates. Defaults to `true`.
    pub animate_background: bool,
}

impl OleChartBuild {
    pub const fn new(shape_id: u32, group_id: GroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: OleChartBuildType::AllAtOnce,
            animate_background: true,
        }
    }

    pub fn with_ui_expand(mut self, expanded: bool) -> Self {
        self.ui_expand = expanded;
        self
    }

    pub fn with_build_type(mut self, build_type: OleChartBuildType) -> Self {
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
pub struct TemplateTimeNode {
    pub(super) xml: Box<str>,
}

/// Template effects for one paragraph level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTemplate {
    /// PowerPoint paragraph level in the inclusive range `0..=9`.
    pub level: u8,
    /// Required single root time node.
    pub time_node: TemplateTimeNode,
}

impl ParagraphTemplate {
    /// Construct a paragraph template with a PowerPoint-supported level.
    pub fn new(level: u8, time_node: TemplateTimeNode) -> Result<Self> {
        if level > 9 {
            return Err(invalid("paragraph template level exceeds PowerPoint limit"));
        }
        Ok(Self { level, time_node })
    }
}

/// A paragraph build associated with a text shape and timing group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphBuild {
    /// Shape whose paragraphs participate in the build.
    pub shape_id: u32,
    /// Timing group referenced by the build.
    pub group_id: GroupId,
    /// Whether the build is expanded in the animation UI. Defaults to `false`.
    pub ui_expand: bool,
    /// Paragraph build mode. Defaults to `Whole`.
    pub build_type: ParagraphBuildType,
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
    pub templates: Vec<ParagraphTemplate>,
}

impl ParagraphBuild {
    /// Construct a paragraph build reference.
    pub fn new(shape_id: u32, group_id: GroupId) -> Self {
        Self {
            shape_id,
            group_id,
            ui_expand: false,
            build_type: ParagraphBuildType::Whole,
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
    pub fn with_build_type(mut self, build_type: ParagraphBuildType) -> Self {
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
    pub fn with_template(mut self, template: ParagraphTemplate) -> Self {
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
pub enum EventFilter {
    /// Prevent the trigger event from bubbling beyond the interactive sequence.
    CancelBubble,
}

impl EventFilter {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "cancelBubble" => Ok(Self::CancelBubble),
            _ => Err(invalid("invalid PowerPoint animation event filter")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::CancelBubble => "cancelBubble",
        }
    }
}

/// Structural sequence containing an animation effect.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum SequenceContext {
    /// The slide's ordinary click sequence.
    #[default]
    Main,
    /// A sequence activated by clicking a shape on the slide.
    Interactive {
        /// Shape whose click activates or advances the sequence.
        trigger_shape_id: u32,
        /// Optional PowerPoint event-bubbling filter on the `interactiveSeq` cTn.
        event_filter: Option<EventFilter>,
    },
}

/// EffectInstance direction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Direction {
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
pub enum Fill {
    Remove,
    Freeze,
    Hold,
    Transition,
}

impl Fill {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "remove" => Ok(Self::Remove),
            "freeze" => Ok(Self::Freeze),
            "hold" => Ok(Self::Hold),
            "transition" => Ok(Self::Transition),
            _ => Err(invalid("invalid animation fill behavior")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
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
pub enum Restart {
    Always,
    WhenNotActive,
    Never,
}

impl Restart {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "always" => Ok(Self::Always),
            "whenNotActive" => Ok(Self::WhenNotActive),
            "never" => Ok(Self::Never),
            _ => Err(invalid("invalid animation restart behavior")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::Always => "always",
            Self::WhenNotActive => "whenNotActive",
            Self::Never => "never",
        }
    }
}

/// Repeat count for a time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Repeat {
    /// Count in OOXML thousandths, where `1000` means one iteration.
    Finite(u32),
    Indefinite,
}

/// Nonzero playback speed in thousandths of a percent.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Speed(i32);

impl Speed {
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
pub struct MotionFraction(u32);

impl MotionFraction {
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
pub enum SyncBehavior {
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
        while scale > 1 && numerator.is_multiple_of(10) {
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
pub struct TimePoint {
    pub local_time: NormalizedTime,
    pub warped_time: NormalizedTime,
}

impl TimePoint {
    pub const fn new(local_time: NormalizedTime, warped_time: NormalizedTime) -> Self {
        Self {
            local_time,
            warped_time,
        }
    }
}

/// Bounded piecewise time-warp filter for a common time node.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeFilter {
    pub(super) points: Box<[TimePoint]>,
}

impl TimeFilter {
    /// Construct a filter whose local times are strictly increasing.
    pub fn new(points: Vec<TimePoint>) -> Result<Self> {
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
    pub fn points(&self) -> &[TimePoint] {
        &self.points
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
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
            points.push(TimePoint::new(
                NormalizedTime::parse(local.trim())?,
                NormalizedTime::parse(warped.trim())?,
            ));
        }
        Self::new(points)
    }

    pub(super) fn write_value(&self) -> String {
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

impl SyncBehavior {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "canSlip" => Ok(Self::CanSlip),
            "locked" => Ok(Self::Locked),
            "none" => Ok(Self::None),
            _ => Err(invalid("invalid animation synchronization behavior")),
        }
    }

    pub(super) const fn as_str(self) -> &'static str {
        match self {
            Self::CanSlip => "canSlip",
            Self::Locked => "locked",
            Self::None => "none",
        }
    }
}

impl Repeat {
    pub(super) fn write_value(self) -> String {
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

    pub(super) fn write_value(self) -> String {
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
pub struct EffectInstance {
    /// Target shape ID
    pub shape_id: u32,
    /// EffectInstance effect
    pub effect: Effect,
    /// Trigger type
    pub trigger: Trigger,
    /// Duration in milliseconds
    pub duration: Duration,
    /// Delay before starting (ms)
    pub delay: u32,
    /// Direction (for directional effects)
    pub direction: Option<Direction>,
    /// Property state retained after the animation becomes inactive.
    pub fill: Option<Fill>,
    /// Policy for restarting this time node.
    pub restart: Option<Restart>,
    /// Whether the animation runs backward after reaching its end.
    pub auto_reverse: bool,
    /// Optional repeat count.
    pub repeat: Option<Repeat>,
    /// Optional nonzero playback speed.
    pub speed: Option<Speed>,
    /// Optional acceleration fraction.
    pub acceleration: Option<MotionFraction>,
    /// Optional deceleration fraction.
    pub deceleration: Option<MotionFraction>,
    /// Whether this time node is visible in the animation user interface.
    pub display: Option<bool>,
    /// Optional total duration for repeated playback.
    pub repeat_duration: Option<Duration>,
    /// Optional synchronization policy with the containing time group.
    pub sync_behavior: Option<SyncBehavior>,
    /// Whether this node is an after-effect.
    pub after_effect: Option<bool>,
    /// Optional normalized-time warp filter.
    pub time_filter: Option<TimeFilter>,
    /// Main-sequence or shape-triggered interactive-sequence context.
    pub sequence_context: SequenceContext,
    /// Optional build-list group containing this effect time node.
    pub group_id: Option<GroupId>,
    /// Sequence order (1-based)
    pub order: u32,
}

impl EffectInstance {
    /// Create a new animation.
    pub fn new(shape_id: u32, effect: Effect) -> Self {
        Self {
            shape_id,
            effect,
            trigger: Trigger::OnClick,
            duration: Duration::Finite(500),
            delay: 0,
            direction: None,
            fill: Some(Fill::Hold),
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
            sequence_context: SequenceContext::Main,
            group_id: None,
            order: 1,
        }
    }

    /// Set the trigger type.
    pub fn with_trigger(mut self, trigger: Trigger) -> Self {
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
    pub fn with_direction(mut self, direction: Direction) -> Self {
        self.direction = Some(direction);
        self
    }

    /// Set the fill behavior for the animation time node.
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set the restart behavior for the animation time node.
    pub fn with_restart(mut self, restart: Restart) -> Self {
        self.restart = Some(restart);
        self
    }

    /// Enable or disable automatic reversal.
    pub fn with_auto_reverse(mut self, auto_reverse: bool) -> Self {
        self.auto_reverse = auto_reverse;
        self
    }

    /// Set the repeat behavior for the animation time node.
    pub fn with_repeat(mut self, repeat: Repeat) -> Self {
        self.repeat = Some(repeat);
        self
    }

    /// Set the nonzero playback speed.
    pub fn with_speed(mut self, speed: Speed) -> Self {
        self.speed = Some(speed);
        self
    }

    /// Set the acceleration fraction.
    pub fn with_acceleration(mut self, acceleration: MotionFraction) -> Self {
        self.acceleration = Some(acceleration);
        self
    }

    /// Set the deceleration fraction.
    pub fn with_deceleration(mut self, deceleration: MotionFraction) -> Self {
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
    pub fn with_sync_behavior(mut self, behavior: SyncBehavior) -> Self {
        self.sync_behavior = Some(behavior);
        self
    }

    /// Mark or unmark this node as an after-effect.
    pub fn with_after_effect(mut self, after_effect: bool) -> Self {
        self.after_effect = Some(after_effect);
        self
    }

    /// Set a normalized-time warp filter.
    pub fn with_time_filter(mut self, filter: TimeFilter) -> Self {
        self.time_filter = Some(filter);
        self
    }

    /// Put this effect in a shape-triggered sequence using PowerPoint's filter.
    pub fn with_interactive_trigger(mut self, trigger_shape_id: u32) -> Self {
        self.sequence_context = SequenceContext::Interactive {
            trigger_shape_id,
            event_filter: Some(EventFilter::CancelBubble),
        };
        self
    }

    /// Set the structural sequence context explicitly.
    pub fn with_sequence_context(mut self, context: SequenceContext) -> Self {
        self.sequence_context = context;
        self
    }

    /// Associate this effect cTn with a build-list timing group.
    pub fn with_group_id(mut self, group_id: impl Into<GroupId>) -> Self {
        self.group_id = Some(group_id.into());
        self
    }
}

/// A trigger event on an ordered timing condition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionEvent {
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

impl ConditionEvent {
    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) fn as_str(self) -> &'static str {
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
pub enum RuntimeTrigger {
    First,
    Last,
    All,
}
impl RuntimeTrigger {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "first" => Ok(Self::First),
            "last" => Ok(Self::Last),
            "all" => Ok(Self::All),
            _ => Err(invalid("invalid animation runtime trigger")),
        }
    }
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::First => "first",
            Self::Last => "last",
            Self::All => "all",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConditionTarget {
    Shape(u32),
    Slide,
    TimeNode(u32),
    Runtime(RuntimeTrigger),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeCondition {
    pub event: Option<ConditionEvent>,
    pub delay: Duration,
    pub target: Option<ConditionTarget>,
}
impl Default for TimeCondition {
    fn default() -> Self {
        Self {
            event: None,
            delay: Duration::Finite(0),
            target: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresetClass {
    Entrance,
    Exit,
    Emphasis,
    MotionPath,
    Verb,
    MediaCall,
}
impl PresetClass {
    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) fn as_str(self) -> &'static str {
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
pub struct PresetTimeNode {
    pub preset_id: u32,
    pub class: PresetClass,
    pub subtype: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeNodeType {
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
impl TimeNodeType {
    pub(super) fn parse(value: &str) -> Result<Self> {
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
    pub(super) fn as_str(self) -> &'static str {
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
pub enum NextAction {
    #[default]
    None,
    Seek,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PreviousAction {
    #[default]
    None,
    SkipTimed,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimingNodeKind {
    Parallel,
    Sequence {
        concurrent: bool,
        next_action: NextAction,
        previous_action: PreviousAction,
    },
    Exclusive,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimingChild {
    Node(TimingNode),
    Opaque(Box<str>),
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommonTimeNode {
    /// Optional schema time-node identifier.
    pub id: Option<u32>,
    pub duration: Option<Duration>,
    pub node_type: Option<TimeNodeType>,
    pub preset: Option<PresetTimeNode>,
    pub start_conditions: Vec<TimeCondition>,
    pub end_conditions: Vec<TimeCondition>,
    pub children: Vec<TimingChild>,
    pub sub_nodes: Vec<TimingChild>,
    pub opaque_children: Vec<Box<str>>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimingNode {
    pub kind: TimingNodeKind,
    pub common: CommonTimeNode,
    pub opaque_children: Vec<Box<str>>,
}
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TimingTree {
    pub roots: Vec<TimingChild>,
    pub opaque_children: Vec<Box<str>>,
    pub(super) source_xml: Option<Box<str>>,
    pub(super) source_roots: Option<Box<[TimingChild]>>,
    pub(super) source_opaque_children: Option<Box<[Box<str>]>>,
}
/// EffectInstance sequence for a slide.
#[derive(Debug, Clone, Default)]
pub struct Sequence {
    /// List of animations in order
    pub animations: Vec<EffectInstance>,
    /// Typed paragraph entries from the slide build list.
    pub paragraph_builds: Vec<ParagraphBuild>,
    /// Typed OLE diagram entries from the slide build list.
    pub diagram_builds: Vec<DiagramBuild>,
    /// Typed chart and SmartArt entries from the slide build list.
    pub graphic_builds: Vec<GraphicBuild>,
    /// Typed embedded OLE chart entries from the slide build list.
    pub ole_chart_builds: Vec<OleChartBuild>,
    pub timing_tree: Option<TimingTree>,
    pub(super) source_timing_xml: Option<Box<str>>,
    pub(super) source_animations: Option<Box<[EffectInstance]>>,
    pub(super) source_paragraph_builds: Option<Box<[ParagraphBuild]>>,
    pub(super) source_diagram_builds: Option<Box<[DiagramBuild]>>,
    pub(super) source_graphic_builds: Option<Box<[GraphicBuild]>>,
    pub(super) source_ole_chart_builds: Option<Box<[OleChartBuild]>>,
    pub(super) source_timing_tree: Option<Box<TimingTree>>,
}

impl PartialEq for Sequence {
    fn eq(&self, other: &Self) -> bool {
        self.animations == other.animations
            && self.paragraph_builds == other.paragraph_builds
            && self.diagram_builds == other.diagram_builds
            && self.graphic_builds == other.graphic_builds
            && self.ole_chart_builds == other.ole_chart_builds
    }
}

impl Eq for Sequence {}

impl Sequence {
    /// Create a new empty animation sequence.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add an animation to the sequence.
    pub fn add(&mut self, mut animation: EffectInstance) {
        animation.order = u32::try_from(self.animations.len() + 1).unwrap_or(u32::MAX);
        self.animations.push(animation);
    }

    /// Add a paragraph build to the slide build list.
    pub fn add_paragraph_build(&mut self, build: ParagraphBuild) {
        self.paragraph_builds.push(build);
    }

    /// Add an OLE diagram build to the slide build list.
    pub fn add_diagram_build(&mut self, build: DiagramBuild) {
        self.diagram_builds.push(build);
    }

    /// Add a chart or SmartArt build to the slide build list.
    pub fn add_graphic_build(&mut self, build: GraphicBuild) {
        self.graphic_builds.push(build);
    }

    /// Add an embedded OLE chart build to the slide build list.
    pub fn add_ole_chart_build(&mut self, build: OleChartBuild) {
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

    pub fn timing_tree(&self) -> Option<&TimingTree> {
        self.timing_tree.as_ref()
    }
}
