//! Shared fixtures and imports for animation parser tests.
pub(super) use crate::animation::triggers::IterationType;
pub(super) use crate::animation::types::{
    AfterEffect, AnimationInfo, BuildAtom, BuildInfo, BuildList, BuildListEntry, ChartBuild,
    ChartBuildAtom, ChartBuildType, DiagramBuild, DiagramBuildAtom, DiagramBuildType,
    ExtendedTimeNode, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect, ParagraphBuild, ParagraphBuildAtom, ParagraphBuildLevel,
    ParagraphBuildType, TimeAnimateBehavior, TimeAnimateBehaviorAtom, TimeAnimateCalculationMode,
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
    is_valid_time_set_value, time_set_attribute_value_type,
};
pub(super) use crate::animation::{
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
pub(super) use crate::consts::RecordType;
pub(super) use crate::records::Record;

pub(super) fn sample_legacy_atom() -> LegacyAnimationAtom {
    LegacyAnimationAtom {
        dim_color: 0x0011_2233,
        reverse: true,
        automatic: true,
        has_sound: true,
        stop_sound: true,
        play: true,
        synchronous: true,
        hide_while_not_playing: true,
        animate_background: true,
        sound_id_ref: 42,
        delay_time_ms: 750,
        order_id: -2,
        slide_count: 3,
        build_type: LegacyAnimationBuild::Level3,
        effect: LegacyAnimationEffect::Fly,
        effect_direction: 0x1C,
        after_effect: AfterEffect::HideOnNextClick,
        text_build_sub_effect: LegacyTextBuildSubEffect::ByCharacter,
        ole_verb: 2,
    }
}

pub(super) fn empty_time_node() -> ExtendedTimeNode {
    ExtendedTimeNode::default()
}

pub(super) fn simple_condition(condition_type: TimeConditionType) -> TimeCondition {
    TimeCondition {
        condition_type,
        atom: TimeConditionAtom {
            trigger_object: TimeTriggerObject::None,
            trigger_event: TimeTriggerEvent::None,
            target_id: 0,
            delay_ms: 0,
        },
        visual_target: None,
    }
}

pub(super) fn sample_set_node_behavior() -> TimeNodeBehavior {
    TimeNodeBehavior::Set(TimeSetBehavior {
        atom: TimeSetBehaviorAtom {
            to_used: true,
            value_type: Some(TimeAnimateValueType::Number),
        },
        to: Some("hidden".to_string()),
        behavior: TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: true,
            },
            attribute_names: Some(vec!["style.visibility".to_string()]),
            properties: None,
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 23,
                data1: 0,
                data2: 0,
            },
        },
    })
}

pub(super) fn sample_build_list() -> BuildList {
    BuildList {
        builds: vec![
            BuildListEntry::Paragraph(ParagraphBuild {
                atom: BuildAtom {
                    build_id: 10,
                    shape_id_ref: 100,
                    expanded: true,
                    ui_expanded: false,
                },
                paragraph: ParagraphBuildAtom {
                    build_type: ParagraphBuildType::AsAWhole,
                    build_level: 4,
                    animate_background: true,
                    reverse: true,
                    user_set_animate_background: true,
                    automatic: true,
                    delay_time_ms: 750,
                },
                levels: vec![ParagraphBuildLevel {
                    level: 0,
                    time_node: empty_time_node(),
                }],
            }),
            BuildListEntry::Chart(ChartBuild {
                atom: BuildAtom {
                    build_id: 11,
                    shape_id_ref: 101,
                    expanded: false,
                    ui_expanded: true,
                },
                chart: ChartBuildAtom {
                    build_type: ChartBuildType::ByElementInCategory,
                    animate_background: true,
                },
            }),
            BuildListEntry::Diagram(DiagramBuild {
                atom: BuildAtom {
                    build_id: 12,
                    shape_id_ref: 102,
                    expanded: true,
                    ui_expanded: true,
                },
                diagram: DiagramBuildAtom {
                    build_type: DiagramBuildType::CounterClockwiseOut,
                },
            }),
        ],
    }
}
