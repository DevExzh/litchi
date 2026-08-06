//! Typed animation model facade.
//!
//! The public surface is organized by semantic ownership while wire parsing
//! and invariant enforcement remain private implementation layers. Keeping
//! these layers here lets the surrounding codec consume the same ergonomic
//! types without exposing format plumbing in the API.

mod builds;
mod codec;
mod effects;
mod sequence;
mod timing;
mod validation;
mod values;

pub use builds::{
    DiagramBuild, DiagramBuildType, GraphicBuild, GraphicBuildMode, GraphicChartBuildType,
    GraphicDiagramBuildType, OleChartBuild, OleChartBuildType, ParagraphBuild, ParagraphBuildType,
    ParagraphTemplate, TemplateTimeNode,
};
pub use effects::{
    Direction, Effect, EffectInstance, EventFilter, GroupId, SequenceContext, Trigger,
};
pub use sequence::Sequence;
pub use timing::{
    CommonTimeNode, ConditionEvent, ConditionTarget, NextAction, PresetClass, PresetTimeNode,
    PreviousAction, RuntimeTrigger, TimeCondition, TimeNodeType, TimingChild, TimingNode,
    TimingNodeKind, TimingTree,
};
pub use values::{
    Duration, Fill, MotionFraction, NormalizedTime, Repeat, Restart, Speed, SyncBehavior,
    TimeFilter, TimePoint,
};
