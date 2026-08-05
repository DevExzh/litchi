//! Focused parser round-trip and malformed-record tests.

use super::*;
use crate::animation::triggers::IterationType;
use crate::animation::types::{
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
use crate::animation::{
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
use crate::consts::RecordType;
use crate::records::Record;

fn sample_legacy_atom() -> LegacyAnimationAtom {
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

fn empty_time_node() -> ExtendedTimeNode {
    ExtendedTimeNode::default()
}

fn simple_condition(condition_type: TimeConditionType) -> TimeCondition {
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

fn sample_set_node_behavior() -> TimeNodeBehavior {
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

fn sample_build_list() -> BuildList {
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

#[test]
fn round_trips_exact_animation_info_atoms_and_containers() {
    let atom = sample_legacy_atom();
    let bytes = write_animation_info_atom(&atom).unwrap();
    assert_eq!(bytes.len(), 36);
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_animation_info_atom(&record).unwrap(), atom);

    let mut info = AnimationInfo::new();
    info.legacy_atom = Some(atom.clone());
    let (container, sound_ref) = write_animation_info(&info).unwrap();
    assert_eq!(sound_ref, 42);
    let (record, consumed) = Record::parse(&container, 0).unwrap();
    assert_eq!(consumed, container.len());
    let parsed = parse_animation_info(&record).unwrap();
    assert_eq!(parsed.legacy_atom, Some(atom));
    assert_eq!(parsed.animation_count(), 1);
    assert_eq!(parsed.after_effect_color, Some(0x0011_2233));
    assert_eq!(parsed.iteration, IterationType::ByLetter);
}

#[test]
fn rejects_malformed_animation_info_atoms() {
    let valid = write_animation_info_atom(&sample_legacy_atom()).unwrap();
    let mutations: &[(usize, u8)] = &[
        (12, 0x02), // invalid bool2 value
        (28, 0xFF), // invalid build type
        (29, 0x0F), // undefined effect
        (30, 0xFF), // invalid direction for Fly
        (31, 0x04), // invalid after effect
        (32, 0x03), // invalid text subdivision
    ];
    for &(offset, value) in mutations {
        let mut bytes = valid.clone();
        bytes[offset] = value;
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert!(
            parse_animation_info_atom(&record).is_err(),
            "accepted mutation at byte {offset}"
        );
    }

    let mut short = valid;
    short[4..8].copy_from_slice(&27u32.to_le_bytes());
    let (record, _) = Record::parse(&short, 0).unwrap();
    assert!(parse_animation_info_atom(&record).is_err());
}

#[test]
fn round_trips_exact_time_node_atoms_and_envelopes() {
    let atom = TimeNodeAtom {
        fill: Some(TimeNodeFill::ResetWhenInactiveLegacy),
        restart: Some(TimeNodeRestart::NeverLegacy),
        node_type: Some(TimeNodeKind::Behavior),
        duration_ms: Some(-1),
    };
    let bytes = write_time_node_atom(&atom);
    assert_eq!(bytes.len(), 40);
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_node_atom(&record).unwrap(), atom);

    let behavior_child = ExtendedTimeNode {
        atom: atom.clone(),
        behavior: Some(sample_set_node_behavior()),
        iterate_data: Some(TimeIterateData {
            interval: Some(100),
            iterate_type: Some(TimeIterateType::ByWord),
            direction: Some(TimeIterateDirection::Forward),
            interval_type: Some(TimeIterateIntervalType::Milliseconds),
        }),
        ..ExtendedTimeNode::default()
    };
    let media_child = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Media),
            ..TimeNodeAtom::default()
        },
        visual_target: Some(TimeVisualElement::Sound {
            kind: TimeVisualElementKind::Audio,
            sound_id_ref: 7,
        }),
        ..ExtendedTimeNode::default()
    };
    let TimeNodeBehavior::Set(subeffect_set) = sample_set_node_behavior() else {
        unreachable!();
    };
    let sub_effect = TimeSubEffect {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Behavior),
            ..TimeNodeAtom::default()
        },
        properties: Some(TimeNodePropertyList {
            properties: vec![TimeNodeProperty::MasterRelation(
                TimeMasterRelation::StartWithMaster,
            )],
        }),
        behavior: Some(TimeSubEffectBehavior::Set(subeffect_set)),
        visual_target: None,
        begin_conditions: vec![simple_condition(TimeConditionType::Begin)],
        end_conditions: vec![simple_condition(TimeConditionType::End)],
        modifiers: vec![TimeModifier::RepeatCount(2)],
    };
    let node = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Sequential),
            ..atom
        },
        properties: Some(TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::EffectType(TimeEffectType::Entrance),
                TimeNodeProperty::EffectNodeType(TimeEffectNodeType::ClickEffect),
            ],
        }),
        sequence_data: Some(TimeSequenceData {
            concurrent: Some(true),
            next_action: Some(TimeSequenceNextAction::SeekToNaturalEnd),
            previous_action: Some(TimeSequencePreviousAction::SkipTimedChildren),
        }),
        begin_conditions: vec![
            simple_condition(TimeConditionType::Begin),
            simple_condition(TimeConditionType::Next),
        ],
        end_conditions: vec![
            simple_condition(TimeConditionType::End),
            simple_condition(TimeConditionType::Previous),
        ],
        end_sync_condition: Some(simple_condition(TimeConditionType::EndSync)),
        modifiers: vec![TimeModifier::Speed(100), TimeModifier::AutomaticReverse(1)],
        sub_effects: vec![sub_effect],
        children: vec![behavior_child, media_child],
        ..ExtendedTimeNode::default()
    };
    let bytes = write_extended_time_node(&node).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(
        record
            .children
            .iter()
            .map(|child| child.record_type)
            .collect::<Vec<_>>(),
        vec![
            RecordType::TimeNode,
            RecordType::TimePropertyList,
            RecordType::TimeSequenceData,
            RecordType::TimeConditionContainer,
            RecordType::TimeConditionContainer,
            RecordType::TimeConditionContainer,
            RecordType::TimeConditionContainer,
            RecordType::TimeConditionContainer,
            RecordType::TimeModifier,
            RecordType::TimeModifier,
            RecordType::TimeSubEffectContainer,
            RecordType::ExtTimeNode,
            RecordType::ExtTimeNode,
        ]
    );
    assert_eq!(parse_extended_time_node(&record).unwrap(), node);
}

#[test]
fn rejects_malformed_time_node_atoms() {
    let default = write_time_node_atom(&TimeNodeAtom::default());
    let mutations: &[(usize, u32)] = &[
        (12, 1), // restart value without fRestartProperty
        (16, 1), // type value without fGroupingTypeProperty
        (20, 1), // fill value without fFillProperty
        (32, 1), // duration without fDurationProperty
    ];
    for &(offset, value) in mutations {
        let mut bytes = default.clone();
        bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert!(parse_time_node_atom(&record).is_err());
    }

    let mut invalid_enum = default;
    invalid_enum[20..24].copy_from_slice(&5u32.to_le_bytes());
    invalid_enum[36..40].copy_from_slice(&1u32.to_le_bytes());
    let (record, _) = Record::parse(&invalid_enum, 0).unwrap();
    assert!(parse_time_node_atom(&record).is_err());
}

#[test]
fn rejects_invalid_extended_time_node_structure() {
    let behavior_node = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Behavior),
            ..TimeNodeAtom::default()
        },
        behavior: Some(sample_set_node_behavior()),
        ..ExtendedTimeNode::default()
    };
    let mut invalid = behavior_node.clone();
    invalid.atom.node_type = Some(TimeNodeKind::Parallel);
    assert!(write_extended_time_node(&invalid).is_err());

    invalid = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Behavior),
            ..TimeNodeAtom::default()
        },
        visual_target: Some(TimeVisualElement::Page),
        ..ExtendedTimeNode::default()
    };
    assert!(write_extended_time_node(&invalid).is_err());

    invalid = ExtendedTimeNode {
        sequence_data: Some(TimeSequenceData {
            concurrent: None,
            next_action: None,
            previous_action: None,
        }),
        ..ExtendedTimeNode::default()
    };
    assert!(write_extended_time_node(&invalid).is_err());

    invalid = ExtendedTimeNode {
        begin_conditions: vec![simple_condition(TimeConditionType::Next)],
        ..ExtendedTimeNode::default()
    };
    assert!(write_extended_time_node(&invalid).is_err());

    invalid = ExtendedTimeNode {
        end_sync_condition: Some(simple_condition(TimeConditionType::End)),
        ..ExtendedTimeNode::default()
    };
    assert!(write_extended_time_node(&invalid).is_err());

    let bytes = write_extended_time_node(&behavior_node).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children[0].data[8..12].copy_from_slice(&0u32.to_le_bytes());
    assert!(parse_extended_time_node(&record).is_err());

    let behavior_record = record.children[1].clone();
    record.children[0].data[8..12].copy_from_slice(&3u32.to_le_bytes());
    let TimeNodeBehavior::Set(set) = behavior_node.behavior.as_ref().unwrap() else {
        unreachable!();
    };
    let behavior_bytes = write_time_set_behavior(set).unwrap();
    record.data.extend(&behavior_bytes);
    record.data_length += behavior_bytes.len() as u32;
    record.children.push(behavior_record);
    assert!(parse_extended_time_node(&record).is_err());

    let ordered = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Sequential),
            ..TimeNodeAtom::default()
        },
        sequence_data: Some(TimeSequenceData {
            concurrent: None,
            next_action: None,
            previous_action: None,
        }),
        begin_conditions: vec![simple_condition(TimeConditionType::Begin)],
        ..ExtendedTimeNode::default()
    };
    let bytes = write_extended_time_node(&ordered).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children.swap(1, 2);
    assert!(parse_extended_time_node(&record).is_err());
}

#[test]
fn round_trips_and_validates_subordinate_effects() {
    assert_eq!(RecordType::TimeSubEffectContainer.as_u16(), 0xF145);
    let media = TimeSubEffect {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Media),
            duration_ms: Some(500),
            ..TimeNodeAtom::default()
        },
        properties: Some(TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::MasterRelation(TimeMasterRelation::DoNotStart),
                TimeNodeProperty::MediaMute(true),
            ],
        }),
        behavior: None,
        visual_target: Some(TimeVisualElement::Sound {
            kind: TimeVisualElementKind::Audio,
            sound_id_ref: 9,
        }),
        begin_conditions: vec![simple_condition(TimeConditionType::Begin)],
        end_conditions: vec![simple_condition(TimeConditionType::End)],
        modifiers: vec![TimeModifier::RepeatDuration(1_000)],
    };
    let bytes = write_time_sub_effect(&media).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_sub_effect(&record).unwrap(), media);

    let mut invalid = media.clone();
    invalid.atom.node_type = None;
    assert!(write_time_sub_effect(&invalid).is_err());

    invalid = media.clone();
    let TimeNodeBehavior::Set(set) = sample_set_node_behavior() else {
        unreachable!();
    };
    invalid.behavior = Some(TimeSubEffectBehavior::Set(set));
    assert!(write_time_sub_effect(&invalid).is_err());

    invalid = media.clone();
    invalid.begin_conditions[0].condition_type = TimeConditionType::Next;
    assert!(write_time_sub_effect(&invalid).is_err());

    let mut record = record;
    record.children.swap(2, 3);
    assert!(parse_time_sub_effect(&record).is_err());
}

#[test]
fn supports_every_time_node_atom_enum_value_and_ignores_reserved_fields() {
    let fills = [
        TimeNodeFill::HoldUntilParentEnds,
        TimeNodeFill::ResetWhenInactive,
        TimeNodeFill::HoldUntilNext,
        TimeNodeFill::HoldUntilParentEndsLegacy,
        TimeNodeFill::ResetWhenInactiveLegacy,
    ];
    let restarts = [
        TimeNodeRestart::Never,
        TimeNodeRestart::Always,
        TimeNodeRestart::WhenNotActive,
        TimeNodeRestart::NeverLegacy,
    ];
    let kinds = [
        TimeNodeKind::Parallel,
        TimeNodeKind::Sequential,
        TimeNodeKind::Behavior,
        TimeNodeKind::Media,
    ];
    for fill in fills {
        for restart in restarts {
            for node_type in kinds {
                let expected = TimeNodeAtom {
                    fill: Some(fill),
                    restart: Some(restart),
                    node_type: Some(node_type),
                    duration_ms: Some(i32::MIN),
                };
                let mut bytes = write_time_node_atom(&expected);
                bytes[8..12].copy_from_slice(&u32::MAX.to_le_bytes());
                bytes[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
                bytes[28] = 0xFF;
                let flags = u32::from_le_bytes(bytes[36..40].try_into().unwrap()) | 0xFFFF_FFE0;
                bytes[36..40].copy_from_slice(&flags.to_le_bytes());
                let (record, _) = Record::parse(&bytes, 0).unwrap();
                assert_eq!(parse_time_node_atom(&record).unwrap(), expected);
            }
        }
    }
}

#[test]
fn round_trips_all_time_node_property_variants() {
    assert_eq!(RecordType::TimePropertyList.as_u16(), 0xF13D);
    assert_eq!(RecordType::TimeVariant.as_u16(), 0xF142);
    let root = TimeNodePropertyList {
        properties: vec![
            TimeNodeProperty::DisplayHidden(true),
            TimeNodeProperty::EffectId(2),
            TimeNodeProperty::EffectDirection(2),
            TimeNodeProperty::EffectType(TimeEffectType::Entrance),
            TimeNodeProperty::AfterEffect(true),
            TimeNodeProperty::SlideCount(3),
            TimeNodeProperty::TimeFilter("0.0,0.5;1.0,1.0".to_string()),
            TimeNodeProperty::EventFilter("cancelBubble".to_string()),
            TimeNodeProperty::HideWhenStopped(false),
            TimeNodeProperty::GroupId(9),
            TimeNodeProperty::EffectNodeType(TimeEffectNodeType::InteractiveSequence),
            TimeNodeProperty::PlaceholderNode(true),
            TimeNodeProperty::MediaVolume(100_000.0),
            TimeNodeProperty::MediaMute(true),
            TimeNodeProperty::ZoomToFullScreen(false),
        ],
    };
    let bytes = write_time_node_property_list(&root, TimePropertyListContext::TimeNode).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(
        parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).unwrap(),
        root
    );

    let subeffect = TimeNodePropertyList {
        properties: vec![
            TimeNodeProperty::DisplayHidden(false),
            TimeNodeProperty::MasterRelation(TimeMasterRelation::StartWithMaster),
            TimeNodeProperty::SubType,
            TimeNodeProperty::AfterEffect(false),
            TimeNodeProperty::PlaceholderNode(false),
            TimeNodeProperty::MediaVolume(0.0),
            TimeNodeProperty::MediaMute(false),
        ],
    };
    let bytes =
        write_time_node_property_list(&subeffect, TimePropertyListContext::SubEffect).unwrap();
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(
        parse_time_node_property_list(&record, TimePropertyListContext::SubEffect).unwrap(),
        subeffect
    );
}

#[test]
fn validates_effect_ids_and_category_specific_directions() {
    let list = |effect_type: TimeEffectType, effect_id: i32, direction: Option<i32>| {
        TimeNodePropertyList {
            properties: [
                Some(TimeNodeProperty::EffectType(effect_type)),
                Some(TimeNodeProperty::EffectId(effect_id)),
                direction.map(TimeNodeProperty::EffectDirection),
            ]
            .into_iter()
            .flatten()
            .collect(),
        }
    };
    for (effect_type, effect_id, direction) in [
        (TimeEffectType::Entrance, 0x3A, None),
        (TimeEffectType::Exit, 2, Some(12)),
        (TimeEffectType::Entrance, 3, Some(5)),
        (TimeEffectType::Entrance, 4, Some(0x20)),
        (TimeEffectType::Entrance, 0x0C, Some(8)),
        (TimeEffectType::Entrance, 0x10, Some(0x2A)),
        (TimeEffectType::Entrance, 0x11, Some(10)),
        (TimeEffectType::Entrance, 0x12, Some(9)),
        (TimeEffectType::Entrance, 0x15, Some(8)),
        (TimeEffectType::Entrance, 0x17, Some(0x210)),
        (TimeEffectType::Emphasis, 0x24, None),
        (TimeEffectType::Emphasis, 1, Some(10)),
        (TimeEffectType::Emphasis, 3, Some(6)),
        (TimeEffectType::Emphasis, 4, Some(2)),
        (TimeEffectType::Emphasis, 5, Some(7)),
        (TimeEffectType::MotionPath, 0x40, Some(i32::MIN)),
        (TimeEffectType::MediaCommand, 3, Some(i32::MAX)),
    ] {
        let expected = list(effect_type, effect_id, direction);
        let bytes =
            write_time_node_property_list(&expected, TimePropertyListContext::TimeNode).unwrap();
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(
            parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).unwrap(),
            expected
        );
    }

    let invalid = [
        list(TimeEffectType::Entrance, -1, None),
        list(TimeEffectType::Entrance, 0x3B, None),
        list(TimeEffectType::Emphasis, 0x25, None),
        list(TimeEffectType::MotionPath, 0x41, None),
        list(TimeEffectType::MediaCommand, 4, None),
        list(TimeEffectType::ActionVerb, 0, None),
        list(TimeEffectType::Entrance, 2, Some(5)),
        list(TimeEffectType::Entrance, 3, Some(1)),
        list(TimeEffectType::Entrance, 0x10, Some(0x20)),
        list(TimeEffectType::Entrance, 0x17, Some(0x211)),
        list(TimeEffectType::Emphasis, 1, Some(3)),
        list(TimeEffectType::Emphasis, 3, Some(2)),
        list(TimeEffectType::Emphasis, 5, Some(8)),
        TimeNodePropertyList {
            properties: vec![TimeNodeProperty::EffectId(1)],
        },
        TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::EffectType(TimeEffectType::Entrance),
                TimeNodeProperty::EffectDirection(1),
            ],
        },
    ];
    for invalid in invalid {
        assert!(
            write_time_node_property_list(&invalid, TimePropertyListContext::TimeNode).is_err()
        );
    }

    let valid = list(TimeEffectType::Entrance, 2, Some(2));
    let bytes = write_time_node_property_list(&valid, TimePropertyListContext::TimeNode).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    let direction = record
        .children
        .iter_mut()
        .find(|child| child.instance == 0x0A)
        .unwrap();
    direction.data[1..5].copy_from_slice(&5i32.to_le_bytes());
    assert!(parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).is_err());
}

#[test]
fn rejects_invalid_time_node_property_lists() {
    let duplicate = TimeNodePropertyList {
        properties: vec![
            TimeNodeProperty::MediaMute(false),
            TimeNodeProperty::MediaMute(true),
        ],
    };
    assert!(write_time_node_property_list(&duplicate, TimePropertyListContext::TimeNode).is_err());
    let wrong_context = TimeNodePropertyList {
        properties: vec![TimeNodeProperty::MasterRelation(
            TimeMasterRelation::DoNotStart,
        )],
    };
    assert!(
        write_time_node_property_list(&wrong_context, TimePropertyListContext::TimeNode).is_err()
    );
    for invalid in [
        TimeNodeProperty::MediaVolume(f32::NAN),
        TimeNodeProperty::MediaVolume(100_001.0),
        TimeNodeProperty::TimeFilter("0.0,2.0".to_string()),
        TimeNodeProperty::EventFilter("other".to_string()),
    ] {
        let list = TimeNodePropertyList {
            properties: vec![invalid],
        };
        assert!(write_time_node_property_list(&list, TimePropertyListContext::TimeNode).is_err());
    }

    let valid = TimeNodePropertyList {
        properties: vec![TimeNodeProperty::MediaMute(true)],
    };
    let bytes = write_time_node_property_list(&valid, TimePropertyListContext::TimeNode).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children[0].data[0] = 1;
    assert!(parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).is_err());
}

#[test]
fn round_trips_shared_time_behaviors_and_all_properties() {
    assert_eq!(RecordType::TimeBehaviorContainer.as_u16(), 0xF12A);
    assert_eq!(RecordType::TimeBehavior.as_u16(), 0xF133);
    assert_eq!(RecordType::TimeClientVisualElement.as_u16(), 0xF13C);
    assert_eq!(RecordType::TimeVariantList.as_u16(), 0xF13E);
    let properties = TimeBehaviorPropertyList {
        properties: vec![
            TimeBehaviorProperty::UnknownPropertyList("vendor.extension".to_string()),
            TimeBehaviorProperty::RuntimeContext("GTE  PPT 12.0;PpT;".to_string()),
            TimeBehaviorProperty::MotionPathEditRelative(true),
            TimeBehaviorProperty::ColorModel(TimeColorModel::Hsl),
            TimeBehaviorProperty::ColorDirection(TimeColorDirection::CounterClockwise),
            TimeBehaviorProperty::Override,
            TimeBehaviorProperty::PathEditRotationAngle(90.0),
            TimeBehaviorProperty::PathEditRotationX(-0.5),
            TimeBehaviorProperty::PathEditRotationY(1.25),
            TimeBehaviorProperty::PointsTypes("AaFfTtSs".to_string()),
        ],
    };
    let behavior = TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: Some(TimeBehaviorAdditive::Add),
            attribute_names_used: true,
        },
        attribute_names: Some(vec![
            "style.opacity".to_string(),
            "style.rotation".to_string(),
        ]),
        properties: Some(properties.clone()),
        target: TimeVisualElement::Shape {
            kind: TimeVisualElementKind::TextRange,
            shape_id_ref: 0xC03,
            data1: 0,
            data2: 12,
        },
    };

    let atom_bytes = write_time_behavior_atom(&behavior.atom);
    let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
    assert_eq!(
        parse_time_behavior_atom(&atom_record).unwrap(),
        behavior.atom
    );

    let property_bytes = write_time_behavior_property_list(&properties).unwrap();
    let (property_record, _) = Record::parse(&property_bytes, 0).unwrap();
    assert_eq!(
        parse_time_behavior_property_list(&property_record).unwrap(),
        properties
    );

    let bytes = write_time_behavior(&behavior).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_behavior(&record).unwrap(), behavior);
}

#[test]
fn round_trips_all_time_visual_element_forms() {
    let targets = [
        TimeVisualElement::Page,
        TimeVisualElement::Sound {
            kind: TimeVisualElementKind::Audio,
            sound_id_ref: 42,
        },
        TimeVisualElement::Shape {
            kind: TimeVisualElementKind::ShapeOnly,
            shape_id_ref: 100,
            data1: -7,
            data2: 9,
        },
        TimeVisualElement::Chart {
            shape_id_ref: 101,
            build_type: ChartBuildType::ByElementInSeries,
            element_index: -1,
        },
    ];
    for target in targets {
        let bytes = write_time_visual_element(&target).unwrap();
        let (record, consumed) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_visual_element(&record).unwrap(), target);
    }
}

#[test]
fn rejects_malformed_shared_time_behaviors() {
    let mut atom = write_time_behavior_atom(&TimeBehaviorAtom {
        additive: None,
        attribute_names_used: false,
    });
    atom[12..16].copy_from_slice(&1u32.to_le_bytes());
    let (atom_record, _) = Record::parse(&atom, 0).unwrap();
    assert!(parse_time_behavior_atom(&atom_record).is_err());

    for property in [
        TimeBehaviorProperty::RuntimeContext("ppt 1.".to_string()),
        TimeBehaviorProperty::PointsTypes("A?".to_string()),
    ] {
        let list = TimeBehaviorPropertyList {
            properties: vec![property],
        };
        assert!(write_time_behavior_property_list(&list).is_err());
    }
    let duplicate = TimeBehaviorPropertyList {
        properties: vec![
            TimeBehaviorProperty::Override,
            TimeBehaviorProperty::Override,
        ],
    };
    assert!(write_time_behavior_property_list(&duplicate).is_err());
    let valid = TimeBehaviorPropertyList {
        properties: vec![TimeBehaviorProperty::Override],
    };
    let bytes = write_time_behavior_property_list(&valid).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children[0].data_length += 1;
    assert!(parse_time_behavior_property_list(&record).is_err());
    assert!(
        write_time_visual_element(&TimeVisualElement::Shape {
            kind: TimeVisualElementKind::ChartElement,
            shape_id_ref: 1,
            data1: 0,
            data2: 0,
        })
        .is_err()
    );
    assert!(
        write_time_visual_element(&TimeVisualElement::Chart {
            shape_id_ref: 1,
            build_type: ChartBuildType::AsOneObject,
            element_index: -2,
        })
        .is_err()
    );

    let sound = TimeVisualElement::Sound {
        kind: TimeVisualElementKind::Audio,
        sound_id_ref: 42,
    };
    let bytes = write_time_visual_element(&sound).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children[0].data[12..16].copy_from_slice(&0u32.to_le_bytes());
    assert!(parse_time_visual_element(&record).is_err());
}

#[test]
fn round_trips_color_behaviors_and_color_models() {
    assert_eq!(RecordType::TimeColorBehaviorContainer.as_u16(), 0xF12C);
    assert_eq!(RecordType::TimeColorBehavior.as_u16(), 0xF135);

    for by in [
        TimeAnimateColorBy::Rgb {
            red: -255,
            green: 0,
            blue: 255,
        },
        TimeAnimateColorBy::Hsl {
            hue: 120,
            saturation: -40,
            luminance: 15,
        },
        TimeAnimateColorBy::Scheme(7),
    ] {
        let expected = TimeColorBehaviorAtom {
            by: Some(by),
            from: Some(TimeAnimateColor::Rgb {
                red: 1,
                green: 2,
                blue: 255,
            }),
            to: Some(TimeAnimateColor::Scheme(3)),
            color_space_used: true,
            direction_used: true,
        };
        let bytes = write_time_color_behavior_atom(&expected).unwrap();
        let (record, consumed) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_color_behavior_atom(&record).unwrap(), expected);
    }

    let expected = TimeColorBehavior {
        atom: TimeColorBehaviorAtom {
            by: Some(TimeAnimateColorBy::Hsl {
                hue: 45,
                saturation: 20,
                luminance: -10,
            }),
            from: None,
            to: Some(TimeAnimateColor::Rgb {
                red: 0x11,
                green: 0x22,
                blue: 0x33,
            }),
            color_space_used: true,
            direction_used: true,
        },
        behavior: TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Override),
                attribute_names_used: true,
            },
            attribute_names: Some(vec!["fill.color".to_string()]),
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![
                    TimeBehaviorProperty::RuntimeContext("ppt".to_string()),
                    TimeBehaviorProperty::ColorModel(TimeColorModel::Hsl),
                    TimeBehaviorProperty::ColorDirection(TimeColorDirection::CounterClockwise),
                ],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 17,
                data1: 0,
                data2: 0,
            },
        },
    };
    let bytes = write_time_color_behavior(&expected).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_color_behavior(&record).unwrap(), expected);
}

#[test]
fn rejects_malformed_color_behaviors() {
    for atom in [
        TimeColorBehaviorAtom {
            by: Some(TimeAnimateColorBy::Rgb {
                red: 256,
                green: 0,
                blue: 0,
            }),
            from: None,
            to: None,
            color_space_used: false,
            direction_used: false,
        },
        TimeColorBehaviorAtom {
            by: None,
            from: Some(TimeAnimateColor::Rgb {
                red: 0,
                green: 0,
                blue: 0,
            }),
            to: None,
            color_space_used: false,
            direction_used: false,
        },
        TimeColorBehaviorAtom {
            by: Some(TimeAnimateColorBy::Scheme(8)),
            from: None,
            to: None,
            color_space_used: false,
            direction_used: false,
        },
        TimeColorBehaviorAtom {
            by: None,
            from: None,
            to: Some(TimeAnimateColor::Rgb {
                red: 256,
                green: 0,
                blue: 0,
            }),
            color_space_used: false,
            direction_used: false,
        },
    ] {
        assert!(write_time_color_behavior_atom(&atom).is_err());
    }

    let valid_atom = TimeColorBehaviorAtom {
        by: Some(TimeAnimateColorBy::Rgb {
            red: 1,
            green: 2,
            blue: 3,
        }),
        from: None,
        to: None,
        color_space_used: false,
        direction_used: false,
    };
    let mut bytes = write_time_color_behavior_atom(&valid_atom).unwrap();
    bytes[12..16].copy_from_slice(&3u32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_color_behavior_atom(&record).is_err());

    let common = |name: &str, properties: Vec<TimeBehaviorProperty>| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: true,
        },
        attribute_names: Some(vec![name.to_string()]),
        properties: Some(TimeBehaviorPropertyList { properties }),
        target: TimeVisualElement::Page,
    };
    for invalid in [
        TimeColorBehavior {
            atom: TimeColorBehaviorAtom {
                color_space_used: true,
                ..valid_atom.clone()
            },
            behavior: common("fill.color", vec![]),
        },
        TimeColorBehavior {
            atom: valid_atom.clone(),
            behavior: common("style.opacity", vec![]),
        },
        TimeColorBehavior {
            atom: valid_atom,
            behavior: common(
                "fill.color",
                vec![TimeBehaviorProperty::MotionPathEditRelative(true)],
            ),
        },
    ] {
        assert!(write_time_color_behavior(&invalid).is_err());
    }
}

#[test]
fn round_trips_all_image_effect_filters() {
    assert_eq!(RecordType::TimeEffectBehaviorContainer.as_u16(), 0xF12D);
    assert_eq!(RecordType::TimeEffectBehavior.as_u16(), 0xF136);
    let filters = [
        TimeEffectFilter::BlindsHorizontal,
        TimeEffectFilter::BlindsVertical,
        TimeEffectFilter::BoxIn,
        TimeEffectFilter::BoxOut,
        TimeEffectFilter::CheckerboardAcross,
        TimeEffectFilter::CheckerboardDown,
        TimeEffectFilter::CircleIn,
        TimeEffectFilter::CircleOut,
        TimeEffectFilter::DiamondIn,
        TimeEffectFilter::DiamondOut,
        TimeEffectFilter::Dissolve,
        TimeEffectFilter::Fade,
        TimeEffectFilter::PlusIn,
        TimeEffectFilter::PlusOut,
        TimeEffectFilter::BarnInVertical,
        TimeEffectFilter::BarnInHorizontal,
        TimeEffectFilter::BarnOutVertical,
        TimeEffectFilter::BarnOutHorizontal,
        TimeEffectFilter::RandomBarHorizontal,
        TimeEffectFilter::RandomBarVertical,
        TimeEffectFilter::StripsDownLeft,
        TimeEffectFilter::StripsUpLeft,
        TimeEffectFilter::StripsDownRight,
        TimeEffectFilter::StripsUpRight,
        TimeEffectFilter::Wedge,
        TimeEffectFilter::Wheel1,
        TimeEffectFilter::Wheel2,
        TimeEffectFilter::Wheel3,
        TimeEffectFilter::Wheel4,
        TimeEffectFilter::Wheel8,
        TimeEffectFilter::WipeRight,
        TimeEffectFilter::WipeLeft,
        TimeEffectFilter::WipeUp,
        TimeEffectFilter::WipeDown,
    ];
    let common = || TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: Some(TimeBehaviorAdditive::Override),
            attribute_names_used: false,
        },
        attribute_names: Some(vec!["ignored".to_string()]),
        properties: Some(TimeBehaviorPropertyList {
            properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
        }),
        target: TimeVisualElement::Shape {
            kind: TimeVisualElementKind::Shape,
            shape_id_ref: 21,
            data1: 0,
            data2: 0,
        },
    };
    for filter in filters {
        assert_eq!(TimeEffectFilter::parse(filter.as_str()), Some(filter));
        let expected = TimeEffectBehavior {
            atom: TimeEffectBehaviorAtom {
                transition: Some(TimeEffectTransition::Out),
                filter_used: true,
                progress_used: true,
                runtime_context_used: true,
            },
            filter: Some(filter),
            progress: Some(0.625),
            runtime_context: Some("GTE PPT 10.0;PpT;".to_string()),
            behavior: common(),
        };
        let bytes = write_time_effect_behavior(&expected).unwrap();
        let (record, consumed) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_effect_behavior(&record).unwrap(), expected);
    }

    for transition in [
        None,
        Some(TimeEffectTransition::In),
        Some(TimeEffectTransition::Out),
    ] {
        let expected = TimeEffectBehaviorAtom {
            transition,
            filter_used: false,
            progress_used: false,
            runtime_context_used: false,
        };
        let bytes = write_time_effect_behavior_atom(&expected);
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_effect_behavior_atom(&record).unwrap(), expected);
    }
}

#[test]
fn rejects_malformed_image_effect_behaviors() {
    let common = || TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: false,
        },
        attribute_names: None,
        properties: None,
        target: TimeVisualElement::Page,
    };
    let valid = TimeEffectBehavior {
        atom: TimeEffectBehaviorAtom {
            transition: None,
            filter_used: true,
            progress_used: true,
            runtime_context_used: true,
        },
        filter: Some(TimeEffectFilter::Fade),
        progress: Some(0.5),
        runtime_context: Some("ppt".to_string()),
        behavior: common(),
    };
    for invalid in [
        TimeEffectBehavior {
            filter: None,
            ..valid.clone()
        },
        TimeEffectBehavior {
            progress: None,
            ..valid.clone()
        },
        TimeEffectBehavior {
            runtime_context: None,
            ..valid.clone()
        },
        TimeEffectBehavior {
            progress: Some(-0.01),
            ..valid.clone()
        },
        TimeEffectBehavior {
            progress: Some(f32::NAN),
            ..valid.clone()
        },
        TimeEffectBehavior {
            runtime_context: Some("ppt 1.".to_string()),
            ..valid.clone()
        },
        TimeEffectBehavior {
            behavior: TimeBehavior {
                properties: Some(TimeBehaviorPropertyList {
                    properties: vec![TimeBehaviorProperty::ColorModel(TimeColorModel::Rgb)],
                }),
                ..common()
            },
            ..valid.clone()
        },
    ] {
        assert!(write_time_effect_behavior(&invalid).is_err());
    }

    let mut bytes = write_time_effect_behavior_atom(&TimeEffectBehaviorAtom {
        transition: None,
        filter_used: false,
        progress_used: false,
        runtime_context_used: false,
    });
    bytes[12..16].copy_from_slice(&1u32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_effect_behavior_atom(&record).is_err());

    let bytes = write_time_effect_behavior(&valid).unwrap();
    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children[1].data = vec![3, b'n', 0, b'o', 0, b'p', 0, b'e', 0];
    record.children[1].data_length = 9;
    assert!(parse_time_effect_behavior(&record).is_err());

    let (mut record, _) = Record::parse(&bytes, 0).unwrap();
    record.children.swap(1, 2);
    assert!(parse_time_effect_behavior(&record).is_err());
}

#[test]
fn round_trips_motion_behaviors_and_formula_paths() {
    assert_eq!(RecordType::TimeMotionBehaviorContainer.as_u16(), 0xF12E);
    assert_eq!(RecordType::TimeMotionBehavior.as_u16(), 0xF137);
    let path = "M 0 0 L 1.0 (ppt_x+$) C 0 0.5 (sin(pi)) 1 1 (max(#ppt_y,0.25)) Z E ignored";
    let expected = TimeMotionBehavior {
        atom: TimeMotionBehaviorAtom {
            by: Some((0.25, -0.5)),
            from: Some((0.0, 0.0)),
            to: Some((1.0, 1.0)),
            origin: Some(TimeMotionOrigin::ObjectCenter),
            path_used: true,
            edit_rotation_used: true,
            points_types_used: true,
        },
        path: Some(path.to_string()),
        reserved: Some(-7),
        behavior: TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Add),
                attribute_names_used: true,
            },
            attribute_names: Some(vec!["ppt_x".to_string(), "ppt_y".to_string()]),
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![
                    TimeBehaviorProperty::MotionPathEditRelative(true),
                    TimeBehaviorProperty::PathEditRotationAngle(45.0),
                    TimeBehaviorProperty::PathEditRotationX(0.5),
                    TimeBehaviorProperty::PathEditRotationY(0.5),
                    TimeBehaviorProperty::PointsTypes("AaFfTtSs".to_string()),
                ],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 22,
                data1: 0,
                data2: 0,
            },
        },
    };
    let bytes = write_time_motion_behavior(&expected).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_motion_behavior(&record).unwrap(), expected);

    for origin in [
        None,
        Some(TimeMotionOrigin::Slide),
        Some(TimeMotionOrigin::SlideLegacy),
        Some(TimeMotionOrigin::ObjectCenter),
    ] {
        let expected = TimeMotionBehaviorAtom {
            by: None,
            from: None,
            to: None,
            origin,
            path_used: false,
            edit_rotation_used: false,
            points_types_used: false,
        };
        let bytes = write_time_motion_behavior_atom(&expected).unwrap();
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_motion_behavior_atom(&record).unwrap(), expected);
    }
}

#[test]
fn rejects_malformed_motion_behaviors_and_paths() {
    let common = || TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: false,
        },
        attribute_names: None,
        properties: None,
        target: TimeVisualElement::Page,
    };
    let atom = TimeMotionBehaviorAtom {
        by: Some((1.0, 1.0)),
        from: None,
        to: None,
        origin: None,
        path_used: true,
        edit_rotation_used: false,
        points_types_used: false,
    };
    for path in [
        "",
        "Q 0 0",
        "M -1 0",
        "M .5 0",
        "M 1. 0",
        "M (unknown) 0",
        "M (max(1,2,3)) 0",
        "M (sin( 1)) 0",
        "C 0 0 1 1",
        "M 0 0 X",
    ] {
        let invalid = TimeMotionBehavior {
            atom: atom.clone(),
            path: Some(path.to_string()),
            reserved: None,
            behavior: common(),
        };
        assert!(write_time_motion_behavior(&invalid).is_err(), "{path}");
    }

    let mut invalid_atom = atom.clone();
    invalid_atom.by = None;
    invalid_atom.from = Some((0.0, 0.0));
    assert!(write_time_motion_behavior_atom(&invalid_atom).is_err());
    let mut bytes = write_time_motion_behavior_atom(&TimeMotionBehaviorAtom {
        path_used: false,
        ..atom.clone()
    })
    .unwrap();
    bytes[36..40].copy_from_slice(&0u32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_motion_behavior_atom(&record).is_err());

    for invalid in [
        TimeMotionBehavior {
            atom: atom.clone(),
            path: None,
            reserved: None,
            behavior: common(),
        },
        TimeMotionBehavior {
            atom: TimeMotionBehaviorAtom {
                edit_rotation_used: true,
                ..atom.clone()
            },
            path: Some("M 0 0".to_string()),
            reserved: None,
            behavior: common(),
        },
        TimeMotionBehavior {
            atom: TimeMotionBehaviorAtom {
                points_types_used: true,
                ..atom.clone()
            },
            path: Some("M 0 0".to_string()),
            reserved: None,
            behavior: common(),
        },
        TimeMotionBehavior {
            atom: atom.clone(),
            path: Some("M 0 0".to_string()),
            reserved: None,
            behavior: TimeBehavior {
                properties: Some(TimeBehaviorPropertyList {
                    properties: vec![TimeBehaviorProperty::ColorDirection(
                        TimeColorDirection::Clockwise,
                    )],
                }),
                ..common()
            },
        },
        TimeMotionBehavior {
            atom,
            path: Some("M 0 0".to_string()),
            reserved: None,
            behavior: TimeBehavior {
                atom: TimeBehaviorAtom {
                    additive: None,
                    attribute_names_used: true,
                },
                attribute_names: Some(vec!["ppt_x".into(), "ppt_y".into(), "ppt_w".into()]),
                ..common()
            },
        },
    ] {
        assert!(write_time_motion_behavior(&invalid).is_err());
    }
}

#[test]
fn round_trips_rotation_and_scale_behaviors() {
    assert_eq!(RecordType::TimeRotationBehaviorContainer.as_u16(), 0xF12F);
    assert_eq!(RecordType::TimeScaleBehaviorContainer.as_u16(), 0xF130);
    assert_eq!(RecordType::TimeRotationBehavior.as_u16(), 0xF138);
    assert_eq!(RecordType::TimeScaleBehavior.as_u16(), 0xF139);
    let common = |attribute_names: Option<Vec<String>>, used| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: Some(TimeBehaviorAdditive::Override),
            attribute_names_used: used,
        },
        attribute_names,
        properties: Some(TimeBehaviorPropertyList {
            properties: vec![TimeBehaviorProperty::RuntimeContext("ppt 12".to_string())],
        }),
        target: TimeVisualElement::Shape {
            kind: TimeVisualElementKind::Shape,
            shape_id_ref: 7,
            data1: 0,
            data2: 0,
        },
    };
    let rotation = TimeRotationBehavior {
        atom: TimeRotationBehaviorAtom {
            by_degrees: Some(45.0),
            from_degrees: Some(-15.0),
            to_degrees: Some(180.0),
            direction: Some(TimeRotationDirection::CounterClockwise),
        },
        behavior: common(Some(vec!["ppt_r".to_string()]), true),
    };
    let atom_bytes = write_time_rotation_behavior_atom(&rotation.atom).unwrap();
    let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
    assert_eq!(
        parse_time_rotation_behavior_atom(&atom_record).unwrap(),
        rotation.atom
    );
    let bytes = write_time_rotation_behavior(&rotation).unwrap();
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(parse_time_rotation_behavior(&record).unwrap(), rotation);

    let scale = TimeScaleBehavior {
        atom: TimeScaleBehaviorAtom {
            by_percent: Some((10.0, 20.0)),
            from_percent: Some((80.0, 90.0)),
            to_percent: Some((120.0, 130.0)),
            zoom_contents: Some(false),
        },
        behavior: common(Some(vec!["ignored".to_string()]), false),
    };
    let atom_bytes = write_time_scale_behavior_atom(&scale.atom).unwrap();
    let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
    assert_eq!(
        parse_time_scale_behavior_atom(&atom_record).unwrap(),
        scale.atom
    );
    let bytes = write_time_scale_behavior(&scale).unwrap();
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(parse_time_scale_behavior(&record).unwrap(), scale);
}

#[test]
fn rejects_malformed_rotation_and_scale_behaviors() {
    let invalid_rotation = TimeRotationBehaviorAtom {
        by_degrees: None,
        from_degrees: Some(1.0),
        to_degrees: None,
        direction: None,
    };
    assert!(write_time_rotation_behavior_atom(&invalid_rotation).is_err());
    let invalid_scale = TimeScaleBehaviorAtom {
        by_percent: None,
        from_percent: Some((1.0, 1.0)),
        to_percent: None,
        zoom_contents: None,
    };
    assert!(write_time_scale_behavior_atom(&invalid_scale).is_err());

    let mut bytes = write_time_rotation_behavior_atom(&TimeRotationBehaviorAtom {
        by_degrees: None,
        from_degrees: None,
        to_degrees: None,
        direction: None,
    })
    .unwrap();
    bytes[20..24].copy_from_slice(&0f32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_rotation_behavior_atom(&record).is_err());

    let mut bytes = write_time_scale_behavior_atom(&TimeScaleBehaviorAtom {
        by_percent: None,
        from_percent: None,
        to_percent: None,
        zoom_contents: None,
    })
    .unwrap();
    bytes[36] = 0;
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_scale_behavior_atom(&record).is_err());
}

#[test]
fn round_trips_generic_animate_behaviors_and_keyframes() {
    assert_eq!(RecordType::TimeAnimateBehaviorContainer.as_u16(), 0xF12B);
    assert_eq!(RecordType::TimeAnimateBehavior.as_u16(), 0xF134);
    assert_eq!(RecordType::TimeAnimationValueList.as_u16(), 0xF13F);
    assert_eq!(RecordType::TimeAnimationValue.as_u16(), 0xF143);
    let common = |attribute: &str| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: Some(TimeBehaviorAdditive::Override),
            attribute_names_used: true,
        },
        attribute_names: Some(vec![attribute.to_string()]),
        properties: Some(TimeBehaviorPropertyList {
            properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
        }),
        target: TimeVisualElement::Shape {
            kind: TimeVisualElementKind::Shape,
            shape_id_ref: 24,
            data1: 0,
            data2: 0,
        },
    };
    let values = TimeAnimationValueList {
        entries: vec![
            TimeAnimationValue {
                time: -1000,
                value: Some(TimeVariantValue::Boolean(true)),
                formula: None,
            },
            TimeAnimationValue {
                time: 333,
                value: Some(TimeVariantValue::Integer(-2)),
                formula: Some("max($,#ppt_y)".to_string()),
            },
            TimeAnimationValue {
                time: 667,
                value: Some(TimeVariantValue::Float(1.25)),
                formula: None,
            },
            TimeAnimationValue {
                time: 1000,
                value: Some(TimeVariantValue::String("2".to_string())),
                formula: None,
            },
        ],
    };
    let expected = TimeAnimateBehavior {
        atom: TimeAnimateBehaviorAtom {
            calculation_mode: Some(TimeAnimateCalculationMode::Formula),
            by_used: true,
            from_used: true,
            to_used: true,
            animation_values_used: true,
            value_type: None,
        },
        values: Some(values.clone()),
        by: Some("1".to_string()),
        from: Some("0".to_string()),
        to: Some("2".to_string()),
        behavior: common("ppt_x"),
    };
    let bytes = write_time_animate_behavior(&expected).unwrap();
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_time_animate_behavior(&record).unwrap(), expected);

    let bytes = write_time_animation_value_list(&values).unwrap();
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(parse_time_animation_value_list(&record).unwrap(), values);

    for (attribute, value_type, value) in [
        ("image", TimeAnimateValueType::String, "arbitrary 👋"),
        ("fill.color", TimeAnimateValueType::Color, "#A0b1C2"),
    ] {
        let expected = TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                calculation_mode: Some(TimeAnimateCalculationMode::Discrete),
                by_used: true,
                from_used: false,
                to_used: true,
                animation_values_used: false,
                value_type: Some(value_type),
            },
            values: None,
            by: Some(value.to_string()),
            from: None,
            to: Some(value.to_string()),
            behavior: common(attribute),
        };
        let bytes = write_time_animate_behavior(&expected).unwrap();
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_animate_behavior(&record).unwrap(), expected);
    }

    for mode in [
        None,
        Some(TimeAnimateCalculationMode::Discrete),
        Some(TimeAnimateCalculationMode::Linear),
        Some(TimeAnimateCalculationMode::Formula),
    ] {
        for value_type in [
            None,
            Some(TimeAnimateValueType::String),
            Some(TimeAnimateValueType::Number),
            Some(TimeAnimateValueType::Color),
        ] {
            let expected = TimeAnimateBehaviorAtom {
                calculation_mode: mode,
                by_used: false,
                from_used: false,
                to_used: false,
                animation_values_used: false,
                value_type,
            };
            let bytes = write_time_animate_behavior_atom(&expected);
            let (record, _) = Record::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_animate_behavior_atom(&record).unwrap(), expected);
        }
    }
}

#[test]
fn rejects_malformed_generic_animate_behaviors() {
    let common = |attribute: &str| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: true,
        },
        attribute_names: Some(vec![attribute.to_string()]),
        properties: None,
        target: TimeVisualElement::Page,
    };
    let valid = TimeAnimateBehavior {
        atom: TimeAnimateBehaviorAtom {
            calculation_mode: None,
            by_used: true,
            from_used: false,
            to_used: false,
            animation_values_used: false,
            value_type: None,
        },
        values: None,
        by: Some("1".to_string()),
        from: None,
        to: None,
        behavior: common("ppt_x"),
    };
    for invalid in [
        TimeAnimateBehavior {
            by: None,
            ..valid.clone()
        },
        TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                animation_values_used: true,
                ..valid.atom.clone()
            },
            ..valid.clone()
        },
        TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                by_used: false,
                ..valid.atom.clone()
            },
            by: None,
            from: Some("0".to_string()),
            ..valid.clone()
        },
        TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                value_type: Some(TimeAnimateValueType::Color),
                ..valid.atom.clone()
            },
            ..valid.clone()
        },
        TimeAnimateBehavior {
            by: Some("invalid".to_string()),
            ..valid.clone()
        },
        TimeAnimateBehavior {
            behavior: common("unsupported.attribute"),
            ..valid.clone()
        },
        TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                calculation_mode: Some(TimeAnimateCalculationMode::Formula),
                ..valid.atom.clone()
            },
            ..valid.clone()
        },
    ] {
        assert!(write_time_animate_behavior(&invalid).is_err());
    }

    for time in [-1001, 1001] {
        assert!(write_time_animation_value_atom(time).is_err());
    }
    let invalid_list = TimeAnimationValueList {
        entries: vec![TimeAnimationValue {
            time: 0,
            value: None,
            formula: Some("unknown+1".to_string()),
        }],
    };
    assert!(write_time_animation_value_list(&invalid_list).is_err());

    let mut atom = write_time_animate_behavior_atom(&TimeAnimateBehaviorAtom {
        calculation_mode: None,
        by_used: false,
        from_used: false,
        to_used: false,
        animation_values_used: false,
        value_type: None,
    });
    atom[8..12].copy_from_slice(&0u32.to_le_bytes());
    let (record, _) = Record::parse(&atom, 0).unwrap();
    assert!(parse_time_animate_behavior_atom(&record).is_err());
}

#[test]
fn round_trips_set_behaviors_for_all_value_categories() {
    assert_eq!(RecordType::TimeSetBehaviorContainer.as_u16(), 0xF131);
    assert_eq!(RecordType::TimeSetBehavior.as_u16(), 0xF13A);
    let common = |attribute: &str| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: Some(TimeBehaviorAdditive::Override),
            attribute_names_used: true,
        },
        attribute_names: Some(vec![attribute.to_string()]),
        properties: Some(TimeBehaviorPropertyList {
            properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
        }),
        target: TimeVisualElement::Shape {
            kind: TimeVisualElementKind::Shape,
            shape_id_ref: 23,
            data1: 0,
            data2: 0,
        },
    };
    let cases = [
        ("style.visibility", TimeAnimateValueType::Number, "hidden"),
        ("style.fontWeight", TimeAnimateValueType::Number, "bold"),
        ("fill.type", TimeAnimateValueType::Number, "gradientRadial"),
        (
            "stroke.dashstyle",
            TimeAnimateValueType::Number,
            "longDashDotDot",
        ),
        (
            "stroke.startArrow",
            TimeAnimateValueType::Number,
            "doublechevron",
        ),
        (
            "extrusion.render",
            TimeAnimateValueType::Number,
            "boundingcube",
        ),
        (
            "ppt_x",
            TimeAnimateValueType::Number,
            "(max($,#ppt_y)+1.5e2)",
        ),
        ("shadow.matrix.ytoy", TimeAnimateValueType::Number, "-.5"),
        (
            "extrusion.rotationcenter.z",
            TimeAnimateValueType::Number,
            "1-e2",
        ),
        ("ppt_c", TimeAnimateValueType::Color, "#00aF7C"),
        ("extrusion.color", TimeAnimateValueType::Color, "#AABBCC"),
    ];
    for (attribute, value_type, value) in cases {
        let expected = TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: (value_type != TimeAnimateValueType::Number).then_some(value_type),
            },
            to: Some(value.to_string()),
            behavior: common(attribute),
        };
        let bytes = write_time_set_behavior(&expected).unwrap();
        let (record, consumed) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_set_behavior(&record).unwrap(), expected);
    }

    for value_type in [
        None,
        Some(TimeAnimateValueType::String),
        Some(TimeAnimateValueType::Number),
        Some(TimeAnimateValueType::Color),
    ] {
        let expected = TimeSetBehaviorAtom {
            to_used: false,
            value_type,
        };
        let bytes = write_time_set_behavior_atom(&expected);
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_set_behavior_atom(&record).unwrap(), expected);
    }
}

#[test]
fn validates_set_presets_numbers_formulas_and_colors() {
    let presets = [
        ("style.visibility", "visible"),
        ("style.fontStyle", "italic"),
        ("style.textEffectEmboss", "emboss"),
        ("style.textShadow", "auto"),
        ("style.textTransform", "super"),
        ("style.textDecorationUnderline", "true"),
        ("style.textEffectOutline", "false"),
        ("style.textDecorationLineThrough", "true"),
        ("imageData.grayscale", "false"),
        ("fill.on", "t"),
        ("fill.method", "sigma"),
        ("stroke.on", "f"),
        ("stroke.linestyle", "thickBetweenThin"),
        ("stroke.filltype", "frame"),
        ("stroke.endArrow", "chevron"),
        ("stroke.startArrowWidth", "narrow"),
        ("stroke.startArrowLength", "long"),
        ("stroke.endArrowWidth", "wide"),
        ("stroke.endArrowLength", "short"),
        ("shadow.on", "true"),
        ("shadow.type", "perspective"),
        ("skew.on", "false"),
        ("extrusion.on", "true"),
        ("extrusion.type", "parallel"),
        ("extrusion.plane", "yz"),
        ("extrusion.lockrotationcenter", "false"),
        ("extrusion.autorotationcenter", "true"),
        ("extrusion.colormode", "false"),
    ];
    for (attribute, value) in presets {
        assert_eq!(
            time_set_attribute_value_type(attribute),
            Some(TimeAnimateValueType::Number)
        );
        assert!(is_valid_time_set_value(attribute, value));
    }
    for value in ["0", "-1", "1.", ".5", "-.5", "1e2", "1-e2", "(sqrt(4))"] {
        assert!(is_valid_time_set_value("ppt_x", value), "{value}");
    }
    for attribute in [
        "ppt_c",
        "fillcolor",
        "style.color",
        "imageData.chromakey",
        "fill.color",
        "fill.color2",
        "stroke.color",
        "stroke.color2",
        "shadow.color",
        "shadow.color2",
        "extrusion.color",
    ] {
        assert_eq!(
            time_set_attribute_value_type(attribute),
            Some(TimeAnimateValueType::Color)
        );
        assert!(is_valid_time_set_value(attribute, "#123abc"));
    }
}

#[test]
fn rejects_malformed_set_behaviors() {
    let common = |attribute_names: Option<Vec<String>>, used| TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: used,
        },
        attribute_names,
        properties: None,
        target: TimeVisualElement::Page,
    };
    let valid = TimeSetBehavior {
        atom: TimeSetBehaviorAtom {
            to_used: true,
            value_type: None,
        },
        to: Some("visible".to_string()),
        behavior: common(Some(vec!["style.visibility".to_string()]), true),
    };
    for invalid in [
        TimeSetBehavior {
            to: None,
            ..valid.clone()
        },
        TimeSetBehavior {
            to: Some("opaque".to_string()),
            ..valid.clone()
        },
        TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: Some(TimeAnimateValueType::Color),
            },
            ..valid.clone()
        },
        TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: Some(TimeAnimateValueType::String),
            },
            ..valid.clone()
        },
        TimeSetBehavior {
            behavior: common(None, false),
            ..valid.clone()
        },
        TimeSetBehavior {
            behavior: common(Some(vec!["image".to_string()]), true),
            ..valid.clone()
        },
        TimeSetBehavior {
            behavior: common(Some(vec!["ppt_x".to_string(), "ppt_y".to_string()]), true),
            ..valid.clone()
        },
    ] {
        assert!(write_time_set_behavior(&invalid).is_err());
    }
    for value in ["", "-", ".", "1-", "1e-2", "(unknown+1)"] {
        let invalid = TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: None,
            },
            to: Some(value.to_string()),
            behavior: common(Some(vec!["ppt_x".to_string()]), true),
        };
        assert!(write_time_set_behavior(&invalid).is_err(), "{value}");
    }
    for value in ["123456", "#12345", "#12345g", "#1234567"] {
        let invalid = TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: Some(TimeAnimateValueType::Color),
            },
            to: Some(value.to_string()),
            behavior: common(Some(vec!["fill.color".to_string()]), true),
        };
        assert!(write_time_set_behavior(&invalid).is_err(), "{value}");
    }

    let mut atom = write_time_set_behavior_atom(&TimeSetBehaviorAtom {
        to_used: false,
        value_type: None,
    });
    atom[12..16].copy_from_slice(&2u32.to_le_bytes());
    let (record, _) = Record::parse(&atom, 0).unwrap();
    assert!(parse_time_set_behavior_atom(&record).is_err());
}

#[test]
fn round_trips_and_validates_command_behaviors() {
    assert_eq!(RecordType::TimeCommandBehaviorContainer.as_u16(), 0xF132);
    assert_eq!(RecordType::TimeCommandBehavior.as_u16(), 0xF13B);
    let common = || TimeBehavior {
        atom: TimeBehaviorAtom {
            additive: None,
            attribute_names_used: false,
        },
        attribute_names: None,
        properties: None,
        target: TimeVisualElement::Sound {
            kind: TimeVisualElementKind::Audio,
            sound_id_ref: 4,
        },
    };
    for (command_type, command) in [
        (Some(TimeCommandBehaviorType::Event), "onstopaudio"),
        (Some(TimeCommandBehaviorType::Call), "playFrom(1.25)"),
        (Some(TimeCommandBehaviorType::OleVerb), "-2"),
        (None, "togglePause"),
    ] {
        let expected = TimeCommandBehavior {
            atom: TimeCommandBehaviorAtom {
                command_type,
                command_used: true,
            },
            command: Some(command.to_string()),
            behavior: common(),
        };
        let atom_bytes = write_time_command_behavior_atom(&expected.atom);
        let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
        assert_eq!(
            parse_time_command_behavior_atom(&atom_record).unwrap(),
            expected.atom
        );
        let bytes = write_time_command_behavior(&expected).unwrap();
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_command_behavior(&record).unwrap(), expected);
    }

    for (command_type, command) in [
        (TimeCommandBehaviorType::Event, "stop"),
        (TimeCommandBehaviorType::Call, "playFrom(-1)"),
        (TimeCommandBehaviorType::OleVerb, "verb"),
    ] {
        let invalid = TimeCommandBehavior {
            atom: TimeCommandBehaviorAtom {
                command_type: Some(command_type),
                command_used: true,
            },
            command: Some(command.to_string()),
            behavior: common(),
        };
        assert!(write_time_command_behavior(&invalid).is_err());
    }

    let mut atom = write_time_command_behavior_atom(&TimeCommandBehaviorAtom {
        command_type: None,
        command_used: false,
    });
    atom[12..16].copy_from_slice(&0u32.to_le_bytes());
    let (record, _) = Record::parse(&atom, 0).unwrap();
    assert!(parse_time_command_behavior_atom(&record).is_err());
}

#[test]
fn round_trips_iterate_and_sequence_data_atoms() {
    assert_eq!(RecordType::TimeIterateData.as_u16(), 0xF140);
    assert_eq!(RecordType::TimeSequenceData.as_u16(), 0xF141);
    for iterate_type in [
        TimeIterateType::AllAtOnce,
        TimeIterateType::ByWord,
        TimeIterateType::ByLetter,
    ] {
        for direction in [
            TimeIterateDirection::Backward,
            TimeIterateDirection::Forward,
        ] {
            for interval_type in [
                TimeIterateIntervalType::Milliseconds,
                TimeIterateIntervalType::TenthsOfAPercent,
            ] {
                let expected = TimeIterateData {
                    interval: Some(250),
                    iterate_type: Some(iterate_type),
                    direction: Some(direction),
                    interval_type: Some(interval_type),
                };
                let bytes = write_time_iterate_data(&expected);
                let (record, _) = Record::parse(&bytes, 0).unwrap();
                assert_eq!(parse_time_iterate_data(&record).unwrap(), expected);
            }
        }
    }
    let expected = TimeSequenceData {
        concurrent: Some(true),
        next_action: Some(TimeSequenceNextAction::SeekToNaturalEnd),
        previous_action: Some(TimeSequencePreviousAction::SkipTimedChildren),
    };
    let bytes = write_time_sequence_data(&expected);
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(parse_time_sequence_data(&record).unwrap(), expected);

    let mut bytes = write_time_iterate_data(&TimeIterateData {
        interval: None,
        iterate_type: None,
        direction: None,
        interval_type: None,
    });
    bytes[8] = 1;
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_iterate_data(&record).is_err());

    let mut bytes = write_time_sequence_data(&TimeSequenceData {
        concurrent: Some(false),
        next_action: None,
        previous_action: None,
    });
    bytes[8] = 2;
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_sequence_data(&record).is_err());
}

#[test]
fn round_trips_time_conditions_and_modifiers() {
    assert_eq!(RecordType::TimeConditionContainer.as_u16(), 0xF125);
    assert_eq!(RecordType::TimeCondition.as_u16(), 0xF128);
    assert_eq!(RecordType::TimeModifier.as_u16(), 0xF129);
    let condition_types = [
        TimeConditionType::None,
        TimeConditionType::Begin,
        TimeConditionType::End,
        TimeConditionType::Next,
        TimeConditionType::Previous,
        TimeConditionType::EndSync,
    ];
    for condition_type in condition_types {
        let expected = TimeCondition {
            condition_type,
            atom: TimeConditionAtom {
                trigger_object: TimeTriggerObject::VisualElement,
                trigger_event: TimeTriggerEvent::MouseClick,
                target_id: 0,
                delay_ms: -1,
            },
            visual_target: Some(TimeVisualElement::Page),
        };
        let bytes = write_time_condition(&expected).unwrap();
        let (record, consumed) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_condition(&record).unwrap(), expected);
    }

    for trigger_event in [
        TimeTriggerEvent::None,
        TimeTriggerEvent::OnBegin,
        TimeTriggerEvent::TimeNodeStart,
        TimeTriggerEvent::TimeNodeEnd,
        TimeTriggerEvent::MouseClick,
        TimeTriggerEvent::MouseOver,
        TimeTriggerEvent::OnNext,
        TimeTriggerEvent::OnPrevious,
        TimeTriggerEvent::StopAudio,
    ] {
        let atom = TimeConditionAtom {
            trigger_object: TimeTriggerObject::RuntimeNodeReference,
            trigger_event,
            target_id: 2,
            delay_ms: i32::MIN,
        };
        let bytes = write_time_condition_atom(&atom).unwrap();
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_condition_atom(&record).unwrap(), atom);
    }

    let modifiers = [
        TimeModifier::RepeatCount(1),
        TimeModifier::RepeatDuration(2),
        TimeModifier::Speed(3),
        TimeModifier::Accelerate(4),
        TimeModifier::Decelerate(5),
        TimeModifier::AutomaticReverse(u32::MAX),
    ];
    for modifier in modifiers {
        let bytes = write_time_modifier(&modifier);
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_modifier(&record).unwrap(), modifier);
    }
}

#[test]
fn rejects_malformed_time_conditions_and_modifiers() {
    let missing_target = TimeCondition {
        condition_type: TimeConditionType::Begin,
        atom: TimeConditionAtom {
            trigger_object: TimeTriggerObject::VisualElement,
            trigger_event: TimeTriggerEvent::OnBegin,
            target_id: 0,
            delay_ms: 0,
        },
        visual_target: None,
    };
    assert!(write_time_condition(&missing_target).is_err());
    let bad_runtime = TimeConditionAtom {
        trigger_object: TimeTriggerObject::RuntimeNodeReference,
        trigger_event: TimeTriggerEvent::TimeNodeStart,
        target_id: 1,
        delay_ms: 0,
    };
    assert!(write_time_condition_atom(&bad_runtime).is_err());

    let mut bytes = write_time_condition_atom(&TimeConditionAtom {
        trigger_object: TimeTriggerObject::None,
        trigger_event: TimeTriggerEvent::None,
        target_id: 0,
        delay_ms: 0,
    })
    .unwrap();
    bytes[12..16].copy_from_slice(&2u32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_condition_atom(&record).is_err());

    let mut bytes = write_time_modifier(&TimeModifier::Speed(100));
    bytes[8..12].copy_from_slice(&6u32.to_le_bytes());
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    assert!(parse_time_modifier(&record).is_err());
}

#[test]
fn parses_slide_animation_extensions_and_rejects_duplicates_or_truncation() {
    fn atom(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend(0u16.to_le_bytes());
        bytes.extend(record_type.to_le_bytes());
        bytes.extend(u32::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend(payload);
        bytes
    }

    let node = ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(TimeNodeKind::Sequential),
            ..TimeNodeAtom::default()
        },
        ..ExtendedTimeNode::default()
    };
    let build_list = BuildList::new();
    let mut data = Vec::new();
    data.extend(0u16.to_le_bytes());
    data.extend(0x7777u16.to_le_bytes());
    data.extend(0u32.to_le_bytes());
    data.extend(write_extended_time_node(&node).unwrap());
    let build_bytes = write_build_list(&build_list).unwrap();
    data.extend(&build_bytes);
    let flags_bytes = atom(12010, &0xffff_ffffu32.to_le_bytes());
    let time_bytes = atom(12011, &0x0123_4567_89ab_cdefu64.to_le_bytes());
    let hash_bytes = atom(0x2b00, &0x89ab_cdefu32.to_le_bytes());
    data.extend(&flags_bytes);
    data.extend(&time_bytes);
    data.extend(&hash_bytes);

    let parsed = parse_slide_animation_extension(&data).unwrap();
    assert_eq!(parsed.time_node, Some(node));
    assert_eq!(parsed.build_list, Some(build_list));
    let flags = parsed.slide_flags.unwrap();
    assert_eq!(flags.raw, 0xffff_ffff);
    assert!(flags.preserve_master);
    assert!(flags.override_master_animation);
    assert_eq!(parsed.creation_time_filetime, Some(0x0123_4567_89ab_cdef));
    assert_eq!(parsed.animation_hash, Some(0x89ab_cdef));

    let mut duplicate = data.clone();
    duplicate.extend(build_bytes);
    assert!(parse_slide_animation_extension(&duplicate).is_err());
    for atom in [&flags_bytes, &time_bytes, &hash_bytes] {
        let mut duplicate = data.clone();
        duplicate.extend(atom);
        assert!(parse_slide_animation_extension(&duplicate).is_err());
    }

    for malformed in [
        atom(12010, &[0; 3]),
        atom(12011, &[0; 7]),
        atom(0x2b00, &[0; 3]),
    ] {
        assert!(parse_slide_animation_extension(&malformed).is_err());
    }

    let mut truncated = data;
    truncated.pop();
    assert!(parse_slide_animation_extension(&truncated).is_err());
}

#[test]
fn round_trips_exact_powerpoint_2002_build_lists() {
    assert_eq!(RecordType::BuildList.as_u16(), 0x2B02);
    assert_eq!(RecordType::LevelInfoAtom.as_u16(), 0x2B0A);
    let bytes = write_build_list(&sample_build_list()).unwrap();
    assert_eq!(RecordType::TimeNode.as_u16(), 0xF127);
    assert_eq!(bytes.len(), 216);
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(record.record_type, RecordType::BuildList);
    assert_eq!(record.children.len(), 3);

    let parsed = parse_build_list(&record).unwrap();
    assert_eq!(parsed.builds.len(), 3);
    let BuildListEntry::Paragraph(paragraph) = &parsed.builds[0] else {
        panic!("expected paragraph build");
    };
    assert_eq!(paragraph.atom.build_id, 10);
    assert_eq!(paragraph.atom.shape_id_ref, 100);
    assert_eq!(paragraph.paragraph.build_type, ParagraphBuildType::AsAWhole);
    assert_eq!(paragraph.paragraph.delay_time_ms, 750);
    assert_eq!(paragraph.levels.len(), 1);
    assert_eq!(paragraph.levels[0].level, 0);
    assert_eq!(paragraph.levels[0].time_node.atom, TimeNodeAtom::default());

    let BuildListEntry::Chart(chart) = &parsed.builds[1] else {
        panic!("expected chart build");
    };
    assert_eq!(chart.chart.build_type, ChartBuildType::ByElementInCategory);
    assert!(chart.chart.animate_background);

    let BuildListEntry::Diagram(diagram) = &parsed.builds[2] else {
        panic!("expected diagram build");
    };
    assert_eq!(
        diagram.diagram.build_type,
        DiagramBuildType::CounterClockwiseOut
    );
}

#[test]
fn rejects_malformed_powerpoint_2002_build_lists() {
    let bytes = write_build_list(&sample_build_list()).unwrap();
    let (valid, _) = Record::parse(&bytes, 0).unwrap();

    let mut truncated = bytes.clone();
    let claimed_length = u32::from_le_bytes(truncated[4..8].try_into().unwrap()) + 1;
    truncated[4..8].copy_from_slice(&claimed_length.to_le_bytes());
    let (truncated, _) = Record::parse(&truncated, 0).unwrap();
    assert_eq!(truncated.data_length, claimed_length);
    assert!(parse_build_list(&truncated).is_err());

    let mut malformed = Vec::new();
    let mut wrong_header = valid.clone();
    wrong_header.version = 0;
    malformed.push(wrong_header);

    let mut wrong_bool = valid.clone();
    wrong_bool.children[1].children[1].data[4] = 2;
    malformed.push(wrong_bool);

    let mut wrong_kind = valid.clone();
    wrong_kind.children[2].children[0].data[0..4].copy_from_slice(&1u32.to_le_bytes());
    malformed.push(wrong_kind);

    let mut wrong_level = valid.clone();
    wrong_level.children[0].children[2].data[0..4].copy_from_slice(&10u32.to_le_bytes());
    malformed.push(wrong_level);

    let mut duplicate = valid.clone();
    duplicate.children[2].children[0].data[4..8].copy_from_slice(&11u32.to_le_bytes());
    duplicate.children[2].children[0].data[8..12].copy_from_slice(&101u32.to_le_bytes());
    malformed.push(duplicate);

    let mut wrong_order = valid.clone();
    wrong_order.children[1].children.swap(0, 1);
    malformed.push(wrong_order);

    for record in malformed {
        assert!(parse_build_list(&record).is_err());
    }
}

#[test]
fn test_animation_info_default() {
    let info = AnimationInfo::default();
    assert!(!info.has_animations());
    assert_eq!(info.animation_count(), 0);
}

#[test]
fn test_build_info_default() {
    let build_info = BuildInfo::default();
    assert!(build_info.builds.is_empty());
}
