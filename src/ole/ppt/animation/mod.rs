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
pub use parser::{parse_animation_info, parse_build_list};
pub use sound::{AnimationSound, BuiltinSound, SoundType};
pub use triggers::{
    AnimationCondition, BeginCondition, EndCondition, InteractiveTrigger, IterationType,
    NextCondition, PreviousCondition, RepeatBehavior,
};
pub use types::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildInfo, BuildLevel,
    BuildType, EffectDirection, EffectSpeed, FillMode, RestartMode, TimeNodeContainer,
    TimeNodeType,
};
pub use writer::{write_animation_info, write_build_list};
