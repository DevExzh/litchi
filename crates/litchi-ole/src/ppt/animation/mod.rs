//! PPT animation support.
//!
//! This module provides structures and functions for parsing and writing
//! PowerPoint binary animation records, including:
//! - Basic and advanced animation effects
//! - Motion paths
//! - Interactive triggers
//! - Sound support
//! - Build animations (chart, diagram, paragraph)

pub mod motion_path;
pub mod parser;
pub mod sound;
pub mod triggers;
pub mod types;
pub mod writer;

pub use motion_path::{MotionPath, MotionPathBuilder, MotionPathType, PathCommand, PathEditMode};
pub use parser::{
    parse_animation_info, parse_animation_info_atom, parse_build_list, parse_extended_time_node,
    parse_slide_animation_extension, parse_time_node_atom,
};
pub use sound::{AnimationSound, BuiltinSound, SoundType};
pub use triggers::{
    AnimationCondition, BeginCondition, EndCondition, InteractiveTrigger, IterationType,
    NextCondition, PreviousCondition, RepeatBehavior,
};
pub use types::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildAtom, BuildInfo,
    BuildLevel, BuildList, BuildListEntry, BuildType, ChartBuild, ChartBuildAtom, ChartBuildType,
    DiagramBuild, DiagramBuildAtom, DiagramBuildType, EffectDirection, EffectSpeed,
    ExtendedTimeNode, FillMode, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect, ParagraphBuild, ParagraphBuildAtom, ParagraphBuildLevel,
    ParagraphBuildType, RestartMode, ShapeAnimation, SlideAnimationExtension, TimeNodeAtom,
    TimeNodeContainer, TimeNodeFill, TimeNodeKind, TimeNodeRestart, TimeNodeType,
};
pub use writer::{
    write_animation_info, write_animation_info_atom, write_build_list, write_extended_time_node,
    write_time_node_atom,
};
