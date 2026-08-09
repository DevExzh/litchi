//! `PowerPoint` 2002 build-list records.

use super::effects::BuildLevel;
use super::time::ExtendedTimeNode;

/// High-level build list information used by animation authoring APIs.
#[derive(Debug, Clone, PartialEq)]
pub struct BuildInfo {
    /// Individual build items.
    pub builds: Vec<BuildLevel>,
}

impl BuildInfo {
    /// Create a new empty build info.
    #[must_use]
    pub fn new() -> Self {
        Self { builds: Vec::new() }
    }

    /// Add a build item.
    pub fn add_build(&mut self, build: BuildLevel) {
        self.builds.push(build);
    }
}

/// Exact `PowerPoint` 2002 build-list record for a slide.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct BuildList {
    /// Paragraph, chart, and diagram build subcontainers in file order.
    pub builds: Vec<BuildListEntry>,
}

/// `PowerPoint` 2002 animation metadata stored in a slide's `___PPT10` tag.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SlideAnimationExtension {
    /// Optional root animation timing tree.
    pub time_node: Option<ExtendedTimeNode>,
    /// Optional shape build list.
    pub build_list: Option<BuildList>,
    /// Optional slide-level flags.
    pub slide_flags: Option<Flags>,
    /// Optional slide creation time as 100-nanosecond ticks since 1601-01-01 UTC.
    pub creation_time_filetime: Option<u64>,
    /// Optional hash of the slide's shape-animation information.
    pub animation_hash: Option<u32>,
    /// Optional inert reference to a linked slide in an associated document.
    pub linked_slide: Option<crate::animation::linked_slide::LinkedSlide>,
    /// Inert linked-shape references in file order.
    pub linked_shapes: Vec<crate::animation::linked_slide::LinkedShape>,
}

impl SlideAnimationExtension {
    /// Return the `PowerPoint` 2002 animation hash as its normative typed atom.
    pub fn animation_hash_atom(&self) -> Option<crate::animation::hash::Hash10> {
        self.animation_hash.map(crate::animation::hash::Hash10::new)
    }

    /// Set or clear the animation hash through the normative typed atom.
    pub fn set_animation_hash_atom(&mut self, atom: Option<crate::animation::hash::Hash10>) {
        self.animation_hash = atom.map(crate::animation::hash::Hash10::hash);
    }
}

/// `PowerPoint` 10 slide-level flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Flags {
    /// Raw flags word. Only bits 0 and 1 are defined by MS-PPT.
    pub raw: u32,
    /// Whether an otherwise unused main or title master slide is preserved.
    pub preserve_master: bool,
    /// Whether this slide overrides animations inherited from its master.
    pub override_master_animation: bool,
}

impl BuildList {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_build(&mut self, build: BuildListEntry) {
        self.builds.push(build);
    }
}

impl Default for BuildInfo {
    fn default() -> Self {
        Self::new()
    }
}

/// Shared build kind stored in a `PowerPoint` 2002 `BuildAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BuildKind {
    Paragraph,
    Chart,
    Diagram,
}

impl BuildKind {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Paragraph),
            2 => Some(Self::Chart),
            3 => Some(Self::Diagram),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::Paragraph => 1,
            Self::Chart => 2,
            Self::Diagram => 3,
        }
    }
}

/// Exact fields shared by paragraph, chart, and diagram build containers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BuildAtom {
    pub build_id: u32,
    pub shape_id_ref: u32,
    pub expanded: bool,
    pub ui_expanded: bool,
}

/// Paragraph build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphBuildType {
    AllAtOnce,
    BuildByNthLevel,
    CustomBuild,
    AsAWhole,
}

impl ParagraphBuildType {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::AllAtOnce),
            1 => Some(Self::BuildByNthLevel),
            2 => Some(Self::CustomBuild),
            3 => Some(Self::AsAWhole),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AllAtOnce => 0,
            Self::BuildByNthLevel => 1,
            Self::CustomBuild => 2,
            Self::AsAWhole => 3,
        }
    }
}

/// Exact paragraph-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the bool fields mirror the independent flag bits of the fixed MS-PPT `ParagraphBuildAtom` layout, so they cannot be merged into enums without losing the bit-level mapping"
)]
pub struct ParagraphBuildAtom {
    pub build_type: ParagraphBuildType,
    pub build_level: u32,
    pub animate_background: bool,
    pub reverse: bool,
    pub user_set_animate_background: bool,
    pub automatic: bool,
    pub delay_time_ms: u32,
}

/// Template effect for one paragraph level.
#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphBuildLevel {
    /// Paragraph level, at most 9.
    pub level: u32,
    /// Inert extended time node retained without executing actions.
    pub time_node: ExtendedTimeNode,
}

/// Text paragraph build subcontainer.
#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphBuild {
    pub atom: BuildAtom,
    pub paragraph: ParagraphBuildAtom,
    pub levels: Vec<ParagraphBuildLevel>,
}

/// Chart build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartBuildType {
    AsOneObject,
    BySeries,
    ByCategory,
    ByElementInSeries,
    ByElementInCategory,
    Custom,
}

impl ChartBuildType {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::AsOneObject),
            1 => Some(Self::BySeries),
            2 => Some(Self::ByCategory),
            3 => Some(Self::ByElementInSeries),
            4 => Some(Self::ByElementInCategory),
            5 => Some(Self::Custom),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AsOneObject => 0,
            Self::BySeries => 1,
            Self::ByCategory => 2,
            Self::ByElementInSeries => 3,
            Self::ByElementInCategory => 4,
            Self::Custom => 5,
        }
    }
}

/// Chart-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartBuildAtom {
    pub build_type: ChartBuildType,
    pub animate_background: bool,
}

/// Chart build subcontainer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartBuild {
    pub atom: BuildAtom,
    pub chart: ChartBuildAtom,
}

/// Diagram build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramBuildType {
    AsOneObject,
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
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0x00 => Some(Self::AsOneObject),
            0x01 => Some(Self::DepthByNode),
            0x02 => Some(Self::DepthByBranch),
            0x03 => Some(Self::BreadthByNode),
            0x04 => Some(Self::BreadthByLevel),
            0x05 => Some(Self::Clockwise),
            0x06 => Some(Self::ClockwiseIn),
            0x07 => Some(Self::ClockwiseOut),
            0x08 => Some(Self::CounterClockwise),
            0x09 => Some(Self::CounterClockwiseIn),
            0x0A => Some(Self::CounterClockwiseOut),
            0x0B => Some(Self::InByRing),
            0x0C => Some(Self::OutByRing),
            0x0D => Some(Self::Up),
            0x0E => Some(Self::Down),
            0x0F => Some(Self::AllAtOnce),
            0x10 => Some(Self::Custom),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AsOneObject => 0x00,
            Self::DepthByNode => 0x01,
            Self::DepthByBranch => 0x02,
            Self::BreadthByNode => 0x03,
            Self::BreadthByLevel => 0x04,
            Self::Clockwise => 0x05,
            Self::ClockwiseIn => 0x06,
            Self::ClockwiseOut => 0x07,
            Self::CounterClockwise => 0x08,
            Self::CounterClockwiseIn => 0x09,
            Self::CounterClockwiseOut => 0x0A,
            Self::InByRing => 0x0B,
            Self::OutByRing => 0x0C,
            Self::Up => 0x0D,
            Self::Down => 0x0E,
            Self::AllAtOnce => 0x0F,
            Self::Custom => 0x10,
        }
    }
}

/// Diagram-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramBuildAtom {
    pub build_type: DiagramBuildType,
}

/// Diagram build subcontainer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramBuild {
    pub atom: BuildAtom,
    pub diagram: DiagramBuildAtom,
}

/// One spec-defined `PowerPoint` 2002 build-list child.
#[derive(Debug, Clone, PartialEq)]
pub enum BuildListEntry {
    Paragraph(ParagraphBuild),
    Chart(ChartBuild),
    Diagram(DiagramBuild),
}
