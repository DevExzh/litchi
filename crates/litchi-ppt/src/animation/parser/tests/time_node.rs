#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Time-node containers, properties, conditions, modifiers, and extensions.
use super::super::*;
use super::support::*;

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
    let node_bytes = write_extended_time_node(&node).unwrap();
    let (node_record, node_consumed) = Record::parse(&node_bytes, 0).unwrap();
    assert_eq!(node_consumed, node_bytes.len());
    assert_eq!(
        node_record
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
    assert_eq!(parse_extended_time_node(&node_record).unwrap(), node);
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
    record.data_length += u32::try_from(behavior_bytes.len()).unwrap();
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
    let ordered_bytes = write_extended_time_node(&ordered).unwrap();
    let (mut ordered_record, _) = Record::parse(&ordered_bytes, 0).unwrap();
    ordered_record.children.swap(1, 2);
    assert!(parse_extended_time_node(&ordered_record).is_err());
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

    let mut swapped_record = record;
    swapped_record.children.swap(2, 3);
    assert!(parse_time_sub_effect(&swapped_record).is_err());
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
fn decodes_grouping_type_property_flag_from_time_node_atom() {
    let mut bytes = write_time_node_atom(&TimeNodeAtom::default());
    bytes[16..20].copy_from_slice(&TimeNodeKind::Sequential.as_u32().to_le_bytes());
    bytes[36..40].copy_from_slice(&0x08u32.to_le_bytes());
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(
        parse_time_node_atom(&record).unwrap().node_type,
        Some(TimeNodeKind::Sequential)
    );
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
    let subeffect_bytes =
        write_time_node_property_list(&subeffect, TimePropertyListContext::SubEffect).unwrap();
    let (subeffect_record, _) = Record::parse(&subeffect_bytes, 0).unwrap();
    assert_eq!(
        parse_time_node_property_list(&subeffect_record, TimePropertyListContext::SubEffect)
            .unwrap(),
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
    for invalid_list in invalid {
        assert!(
            write_time_node_property_list(&invalid_list, TimePropertyListContext::TimeNode)
                .is_err()
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

    let mut modifier_bytes = write_time_modifier(&TimeModifier::Speed(100));
    modifier_bytes[8..12].copy_from_slice(&6u32.to_le_bytes());
    let (modifier_record, _) = Record::parse(&modifier_bytes, 0).unwrap();
    assert!(parse_time_modifier(&modifier_record).is_err());
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
        let mut duplicate_atom = data.clone();
        duplicate_atom.extend(atom);
        assert!(parse_slide_animation_extension(&duplicate_atom).is_err());
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
