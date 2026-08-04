//! Layered animation support for PowerPoint presentations.
//!
//! The owner is split by responsibility: [`model`] contains the typed timing
//! vocabulary, [`codec`] contains bounded PresentationML XML parsing and
//! writing, [`package`] validates package relationships, and [`tests`] keeps
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

// Compatibility aliases retain the pre-layering public surface while the
// canonical names stay concise inside the `animations` context.
pub type Animation = EffectInstance;
pub type AnimationProgress = MotionFraction;
pub type AnimationEffect = Effect;
pub type AnimationTrigger = Trigger;
pub type AnimationGroupId = GroupId;
pub type AnimationParagraphBuildType = ParagraphBuildType;
pub type AnimationDiagramBuildType = DiagramBuildType;
pub type AnimationDiagramBuild = DiagramBuild;
pub type AnimationGraphicDiagramBuildType = GraphicDiagramBuildType;
pub type AnimationGraphicChartBuildType = GraphicChartBuildType;
pub type AnimationGraphicBuildMode = GraphicBuildMode;
pub type AnimationGraphicBuild = GraphicBuild;
pub type AnimationOleChartBuildType = OleChartBuildType;
pub type AnimationOleChartBuild = OleChartBuild;
pub type AnimationTemplateTimeNode = TemplateTimeNode;
pub type AnimationParagraphTemplate = ParagraphTemplate;
pub type AnimationParagraphBuild = ParagraphBuild;
pub type AnimationEventFilter = EventFilter;
pub type AnimationSequenceContext = SequenceContext;
pub type AnimationDirection = Direction;
pub type AnimationFill = Fill;
pub type AnimationRestart = Restart;
pub type AnimationRepeat = Repeat;
pub type AnimationSpeed = Speed;
pub type AnimationSyncBehavior = SyncBehavior;
pub type AnimationTimePoint = TimePoint;
pub type AnimationTimeFilter = TimeFilter;
pub type AnimationConditionEvent = ConditionEvent;
pub type AnimationRuntimeTrigger = RuntimeTrigger;
pub type AnimationConditionTarget = ConditionTarget;
pub type AnimationTimeCondition = TimeCondition;
pub type AnimationPresetClass = PresetClass;
pub type AnimationPresetTimeNode = PresetTimeNode;
pub type AnimationTimeNodeType = TimeNodeType;
pub type AnimationNextAction = NextAction;
pub type AnimationPreviousAction = PreviousAction;
pub type AnimationTimingNodeKind = TimingNodeKind;
pub type AnimationTimingChild = TimingChild;
pub type AnimationCommonTimeNode = CommonTimeNode;
pub type AnimationTimingNode = TimingNode;
pub type AnimationTimingTree = TimingTree;
pub type AnimationSequence = Sequence;
