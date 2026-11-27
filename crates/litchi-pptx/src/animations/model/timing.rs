//! Structural timing-tree and condition models.

use super::{Duration, MotionFraction};

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeTrigger {
    First,
    Last,
    All,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConditionTarget {
    Shape(u32),
    /// A named bookmark on an audio or video shape.
    ///
    /// This is the `p14:bmkTgt` extension described by MS-PPTX 2.2.2.
    /// The shape ID remains the ordinary timing target; the bookmark name is
    /// inert text and is checked against the bounded XML grammar.
    MediaBookmark {
        shape_id: u32,
        name: String,
    },
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
    /// Optional Office 2010 UI bounce fraction (`p14:presetBounceEnd`).
    ///
    /// This is distinct from the behavior-level `bounceEnd` attributes. It
    /// describes the UI preset on this common time node.
    pub preset_bounce_end: Option<MotionFraction>,
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
    pub(in crate::animations) source_xml: Option<Box<str>>,
    pub(in crate::animations) source_roots: Option<Box<[TimingChild]>>,
    pub(in crate::animations) source_opaque_children: Option<Box<[Box<str>]>>,
}
