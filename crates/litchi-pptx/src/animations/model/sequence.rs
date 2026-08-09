//! Mutable slide-level animation sequence state.

use super::{
    DiagramBuild, EffectInstance, GraphicBuild, OleChartBuild, ParagraphBuild, TimingTree,
};

/// `EffectInstance` sequence for a slide.
#[derive(Debug, Clone, Default)]
pub struct Sequence {
    /// List of animations in order
    pub animations: Vec<EffectInstance>,
    /// Typed paragraph entries from the slide build list.
    pub paragraph_builds: Vec<ParagraphBuild>,
    /// Typed OLE diagram entries from the slide build list.
    pub diagram_builds: Vec<DiagramBuild>,
    /// Typed chart and `SmartArt` entries from the slide build list.
    pub graphic_builds: Vec<GraphicBuild>,
    /// Typed embedded OLE chart entries from the slide build list.
    pub ole_chart_builds: Vec<OleChartBuild>,
    pub timing_tree: Option<TimingTree>,
    pub(in crate::animations) source_timing_xml: Option<Box<str>>,
    pub(in crate::animations) source_animations: Option<Box<[EffectInstance]>>,
    pub(in crate::animations) source_paragraph_builds: Option<Box<[ParagraphBuild]>>,
    pub(in crate::animations) source_diagram_builds: Option<Box<[DiagramBuild]>>,
    pub(in crate::animations) source_graphic_builds: Option<Box<[GraphicBuild]>>,
    pub(in crate::animations) source_ole_chart_builds: Option<Box<[OleChartBuild]>>,
    pub(in crate::animations) source_timing_tree: Option<Box<TimingTree>>,
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
    #[must_use]
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

    /// Add a chart or `SmartArt` build to the slide build list.
    pub fn add_graphic_build(&mut self, build: GraphicBuild) {
        self.graphic_builds.push(build);
    }

    /// Add an embedded OLE chart build to the slide build list.
    pub fn add_ole_chart_build(&mut self, build: OleChartBuild) {
        self.ole_chart_builds.push(build);
    }

    /// Get the number of animations.
    #[must_use]
    pub fn len(&self) -> usize {
        self.animations.len()
    }

    /// Check if the sequence is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.animations.is_empty()
    }

    /// Return the preserved source timing subtree, when this sequence was parsed.
    #[must_use]
    pub fn preserved_timing_xml(&self) -> Option<&str> {
        self.source_timing_xml.as_deref()
    }

    #[must_use]
    pub fn timing_tree(&self) -> Option<&TimingTree> {
        self.timing_tree.as_ref()
    }
}
