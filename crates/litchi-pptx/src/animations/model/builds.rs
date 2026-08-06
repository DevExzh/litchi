//! Public build-list models for text, diagrams, charts, and OLE charts.

use super::{Duration, GroupId};

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
    pub(in crate::animations) xml: Box<str>,
}

/// Template effects for one paragraph level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTemplate {
    /// PowerPoint paragraph level in the inclusive range `0..=9`.
    pub level: u8,
    /// Required single root time node.
    pub time_node: TemplateTimeNode,
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
