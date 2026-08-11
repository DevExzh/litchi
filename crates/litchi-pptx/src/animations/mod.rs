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
pub use package::parse_package_slide;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
