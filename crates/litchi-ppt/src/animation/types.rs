//! Animation data types.

mod animation;
mod build;
mod effects;
mod time;
mod validation;

pub use animation::{
    AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect, ShapeAnimation,
};
pub(crate) use build::BuildKind;
pub use build::{
    BuildAtom, BuildInfo, BuildList, BuildListEntry, ChartBuild, ChartBuildAtom, ChartBuildType,
    DiagramBuild, DiagramBuildAtom, DiagramBuildType, Flags, ParagraphBuild, ParagraphBuildAtom,
    ParagraphBuildLevel, ParagraphBuildType, SlideAnimationExtension,
};
pub use effects::{
    AfterEffect, AnimationEffect, AnimationTrigger, BuildLevel, BuildType, EffectDirection,
    EffectSpeed, FillMode, RestartMode, TimeNodeContainer, TimeNodeType,
};
pub(crate) use time::has_valid_time_effect_properties;
pub use time::{
    ExtendedTimeNode, TimeAnimateBehavior, TimeAnimateBehaviorAtom, TimeAnimateCalculationMode,
    TimeAnimateColor, TimeAnimateColorBy, TimeAnimateValueType, TimeAnimationValue,
    TimeAnimationValueList, TimeBehavior, TimeBehaviorAdditive, TimeBehaviorAtom,
    TimeBehaviorProperty, TimeBehaviorPropertyList, TimeColorBehavior, TimeColorBehaviorAtom,
    TimeColorDirection, TimeColorModel, TimeCommandBehavior, TimeCommandBehaviorAtom,
    TimeCommandBehaviorType, TimeCondition, TimeConditionAtom, TimeConditionType,
    TimeEffectBehavior, TimeEffectBehaviorAtom, TimeEffectFilter, TimeEffectNodeType,
    TimeEffectTransition, TimeEffectType, TimeIterateData, TimeIterateDirection,
    TimeIterateIntervalType, TimeIterateType, TimeMasterRelation, TimeModifier, TimeMotionBehavior,
    TimeMotionBehaviorAtom, TimeMotionOrigin, TimeNodeAtom, TimeNodeBehavior, TimeNodeFill,
    TimeNodeKind, TimeNodeProperty, TimeNodePropertyList, TimeNodeRestart, TimePropertyListContext,
    TimeRotationBehavior, TimeRotationBehaviorAtom, TimeRotationDirection, TimeScaleBehavior,
    TimeScaleBehaviorAtom, TimeSequenceData, TimeSequenceNextAction, TimeSequencePreviousAction,
    TimeSetBehavior, TimeSetBehaviorAtom, TimeSubEffect, TimeSubEffectBehavior, TimeTriggerEvent,
    TimeTriggerObject, TimeVariantValue, TimeVisualElement, TimeVisualElementKind,
};
pub(crate) use validation::{
    is_valid_animation_attribute_name, is_valid_motion_path, is_valid_runtime_context,
    is_valid_time_animate_value, is_valid_time_filter, is_valid_time_formula,
    is_valid_time_points_types, is_valid_time_set_value, time_animation_attribute_value_type,
    time_set_attribute_value_type,
};
