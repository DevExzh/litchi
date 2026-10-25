//! Layered animation support for `PowerPoint` presentations.
//!
//! The owner is split by responsibility: `model` contains the typed timing
//! vocabulary, `codec` contains bounded `PresentationML` XML parsing and
//! writing, `package` validates package relationships, and `tests` keeps
//! the conformance and resource-limit coverage beside the owner.

use crate::Error;

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::MAX_TIMING_MILLISECONDS;
pub use model::{
    CommonTimeNode, ConditionEvent, ConditionTarget, DiagramBuild, DiagramBuildType, Direction,
    Duration, Effect, EffectInstance, EventFilter, Fill, GraphicBuild, GraphicBuildMode,
    GraphicChartBuildType, GraphicDiagramBuildType, GroupId, MotionFraction, NextAction,
    NormalizedTime, OleChartBuild, OleChartBuildType, ParagraphBuild, ParagraphBuildType,
    ParagraphTemplate, PresetClass, PresetTimeNode, PreviousAction, Repeat, Restart,
    RuntimeTrigger, Sequence, SequenceContext, Speed, SyncBehavior, TemplateTimeNode,
    TimeCondition, TimeFilter, TimeNodeType, TimePoint, TimingChild, TimingNode, TimingNodeKind,
    TimingTree, Trigger,
};
pub(crate) use package::parse_package_slide_with_limits;
pub use package::{parse_package_slide, parse_package_slide_with_tolerance};

/// Resource policy for raw PresentationML timing XML.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_nodes: usize,
    pub max_text_bytes: usize,
    pub max_template_bytes: usize,
    pub max_attributes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_input_bytes: crate::animations::codec::validation::MAX_TIMING_XML_BYTES,
        max_output_bytes: crate::animations::codec::validation::MAX_TIMING_XML_BYTES,
        max_depth: crate::animations::codec::validation::MAX_TIMING_DEPTH,
        max_nodes: crate::animations::codec::validation::MAX_TIMING_NODES,
        max_text_bytes: crate::animations::codec::validation::MAX_TIMING_TEXT_BYTES,
        max_template_bytes: crate::animations::codec::validation::MAX_TEMPLATE_TIME_NODE_BYTES,
        max_attributes: crate::animations::codec::validation::MAX_TIMING_ATTRIBUTES,
    };

    #[must_use]
    pub const fn from_source(source: litchi_opc::SourceLimits) -> Self {
        Self::DEFAULT.constrained_by_source(source)
    }

    /// Intersect a selected semantic policy with the physical source ceiling.
    #[must_use]
    pub const fn constrained_by_source(self, source: litchi_opc::SourceLimits) -> Self {
        let part = source.max_part_bytes().min(usize::MAX as u64) as usize;
        Self {
            max_input_bytes: self.max_input_bytes.min(part),
            max_output_bytes: self.max_output_bytes.min(part),
            max_depth: self.max_depth.min(source.max_depth()),
            max_nodes: self.max_nodes.min(source.max_events()),
            max_text_bytes: self.max_text_bytes.min(part),
            max_template_bytes: self.max_template_bytes.min(part),
            max_attributes: self.max_attributes.min(source.max_xml_attribute_bytes()),
            ..self
        }
    }
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
