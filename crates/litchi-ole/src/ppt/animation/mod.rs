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
    parse_slide_animation_extension, parse_time_animate_behavior, parse_time_animate_behavior_atom,
    parse_time_animation_value_atom, parse_time_animation_value_list, parse_time_behavior,
    parse_time_behavior_atom, parse_time_behavior_property_list, parse_time_color_behavior,
    parse_time_color_behavior_atom, parse_time_command_behavior, parse_time_command_behavior_atom,
    parse_time_condition, parse_time_condition_atom, parse_time_effect_behavior,
    parse_time_effect_behavior_atom, parse_time_iterate_data, parse_time_modifier,
    parse_time_motion_behavior, parse_time_motion_behavior_atom, parse_time_node_atom,
    parse_time_node_property_list, parse_time_rotation_behavior, parse_time_rotation_behavior_atom,
    parse_time_scale_behavior, parse_time_scale_behavior_atom, parse_time_sequence_data,
    parse_time_set_behavior, parse_time_set_behavior_atom, parse_time_sub_effect,
    parse_time_visual_element,
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
    ParagraphBuildType, RestartMode, ShapeAnimation, SlideAnimationExtension, TimeAnimateBehavior,
    TimeAnimateBehaviorAtom, TimeAnimateCalculationMode, TimeAnimateColor, TimeAnimateColorBy,
    TimeAnimateValueType, TimeAnimationValue, TimeAnimationValueList, TimeBehavior,
    TimeBehaviorAdditive, TimeBehaviorAtom, TimeBehaviorProperty, TimeBehaviorPropertyList,
    TimeColorBehavior, TimeColorBehaviorAtom, TimeColorDirection, TimeColorModel,
    TimeCommandBehavior, TimeCommandBehaviorAtom, TimeCommandBehaviorType, TimeCondition,
    TimeConditionAtom, TimeConditionType, TimeEffectBehavior, TimeEffectBehaviorAtom,
    TimeEffectFilter, TimeEffectNodeType, TimeEffectTransition, TimeEffectType, TimeIterateData,
    TimeIterateDirection, TimeIterateIntervalType, TimeIterateType, TimeMasterRelation,
    TimeModifier, TimeMotionBehavior, TimeMotionBehaviorAtom, TimeMotionOrigin, TimeNodeAtom,
    TimeNodeBehavior, TimeNodeContainer, TimeNodeFill, TimeNodeKind, TimeNodeProperty,
    TimeNodePropertyList, TimeNodeRestart, TimeNodeType, TimePropertyListContext,
    TimeRotationBehavior, TimeRotationBehaviorAtom, TimeRotationDirection, TimeScaleBehavior,
    TimeScaleBehaviorAtom, TimeSequenceData, TimeSequenceNextAction, TimeSequencePreviousAction,
    TimeSetBehavior, TimeSetBehaviorAtom, TimeSubEffect, TimeSubEffectBehavior, TimeTriggerEvent,
    TimeTriggerObject, TimeVariantValue, TimeVisualElement, TimeVisualElementKind,
};
pub use writer::{
    write_animation_info, write_animation_info_atom, write_build_list, write_extended_time_node,
    write_time_animate_behavior, write_time_animate_behavior_atom, write_time_animation_value_atom,
    write_time_animation_value_list, write_time_behavior, write_time_behavior_atom,
    write_time_behavior_property_list, write_time_color_behavior, write_time_color_behavior_atom,
    write_time_command_behavior, write_time_command_behavior_atom, write_time_condition,
    write_time_condition_atom, write_time_effect_behavior, write_time_effect_behavior_atom,
    write_time_iterate_data, write_time_modifier, write_time_motion_behavior,
    write_time_motion_behavior_atom, write_time_node_atom, write_time_node_property_list,
    write_time_rotation_behavior, write_time_rotation_behavior_atom, write_time_scale_behavior,
    write_time_scale_behavior_atom, write_time_sequence_data, write_time_set_behavior,
    write_time_set_behavior_atom, write_time_sub_effect, write_time_visual_element,
};
