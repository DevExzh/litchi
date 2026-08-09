#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Shared and specialized animation behavior round trips and validation.
use super::super::*;
use crate::animation::types::{is_valid_time_set_value, time_set_attribute_value_type};
use crate::animation::*;
use crate::consts::RecordType;
use crate::records::Record;

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
    let sound_bytes = write_time_visual_element(&sound).unwrap();
    let (mut sound_record, _) = Record::parse(&sound_bytes, 0).unwrap();
    sound_record.children[0].data[12..16].copy_from_slice(&0u32.to_le_bytes());
    assert!(parse_time_visual_element(&sound_record).is_err());
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

    let mut atom_bytes = write_time_effect_behavior_atom(&TimeEffectBehaviorAtom {
        transition: None,
        filter_used: false,
        progress_used: false,
        runtime_context_used: false,
    });
    atom_bytes[12..16].copy_from_slice(&1u32.to_le_bytes());
    let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
    assert!(parse_time_effect_behavior_atom(&atom_record).is_err());

    let bytes = write_time_effect_behavior(&valid).unwrap();
    let (mut bad_filter_record, _) = Record::parse(&bytes, 0).unwrap();
    bad_filter_record.children[1].data = vec![3, b'n', 0, b'o', 0, b'p', 0, b'e', 0];
    bad_filter_record.children[1].data_length = 9;
    assert!(parse_time_effect_behavior(&bad_filter_record).is_err());

    let (mut swapped_record, _) = Record::parse(&bytes, 0).unwrap();
    swapped_record.children.swap(1, 2);
    assert!(parse_time_effect_behavior(&swapped_record).is_err());
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
        let motion_atom = TimeMotionBehaviorAtom {
            by: None,
            from: None,
            to: None,
            origin,
            path_used: false,
            edit_rotation_used: false,
            points_types_used: false,
        };
        let atom_bytes = write_time_motion_behavior_atom(&motion_atom).unwrap();
        let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
        assert_eq!(
            parse_time_motion_behavior_atom(&atom_record).unwrap(),
            motion_atom
        );
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
    let scale_atom_bytes = write_time_scale_behavior_atom(&scale.atom).unwrap();
    let (scale_atom_record, _) = Record::parse(&scale_atom_bytes, 0).unwrap();
    assert_eq!(
        parse_time_scale_behavior_atom(&scale_atom_record).unwrap(),
        scale.atom
    );
    let scale_bytes = write_time_scale_behavior(&scale).unwrap();
    let (scale_record, _) = Record::parse(&scale_bytes, 0).unwrap();
    assert_eq!(parse_time_scale_behavior(&scale_record).unwrap(), scale);
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

    let mut rot_bytes = write_time_rotation_behavior_atom(&TimeRotationBehaviorAtom {
        by_degrees: None,
        from_degrees: None,
        to_degrees: None,
        direction: None,
    })
    .unwrap();
    rot_bytes[20..24].copy_from_slice(&0f32.to_le_bytes());
    let (rot_record, _) = Record::parse(&rot_bytes, 0).unwrap();
    assert!(parse_time_rotation_behavior_atom(&rot_record).is_err());

    let mut scale_bytes = write_time_scale_behavior_atom(&TimeScaleBehaviorAtom {
        by_percent: None,
        from_percent: None,
        to_percent: None,
        zoom_contents: None,
    })
    .unwrap();
    scale_bytes[36] = 0;
    let (scale_record, _) = Record::parse(&scale_bytes, 0).unwrap();
    assert!(parse_time_scale_behavior_atom(&scale_record).is_err());
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

    let list_bytes = write_time_animation_value_list(&values).unwrap();
    let (list_record, _) = Record::parse(&list_bytes, 0).unwrap();
    assert_eq!(
        parse_time_animation_value_list(&list_record).unwrap(),
        values
    );

    for (attribute, value_type, value) in [
        ("image", TimeAnimateValueType::String, "arbitrary 👋"),
        ("fill.color", TimeAnimateValueType::Color, "#A0b1C2"),
    ] {
        let discrete_expected = TimeAnimateBehavior {
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
        let discrete_bytes = write_time_animate_behavior(&discrete_expected).unwrap();
        let (discrete_record, _) = Record::parse(&discrete_bytes, 0).unwrap();
        assert_eq!(
            parse_time_animate_behavior(&discrete_record).unwrap(),
            discrete_expected
        );
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
            let animate_atom = TimeAnimateBehaviorAtom {
                calculation_mode: mode,
                by_used: false,
                from_used: false,
                to_used: false,
                animation_values_used: false,
                value_type,
            };
            let atom_bytes = write_time_animate_behavior_atom(&animate_atom);
            let (atom_record, _) = Record::parse(&atom_bytes, 0).unwrap();
            assert_eq!(
                parse_time_animate_behavior_atom(&atom_record).unwrap(),
                animate_atom
            );
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

    let mut bad_iterate_bytes = write_time_iterate_data(&TimeIterateData {
        interval: None,
        iterate_type: None,
        direction: None,
        interval_type: None,
    });
    bad_iterate_bytes[8] = 1;
    let (bad_iterate_record, _) = Record::parse(&bad_iterate_bytes, 0).unwrap();
    assert!(parse_time_iterate_data(&bad_iterate_record).is_err());

    let mut bad_sequence_bytes = write_time_sequence_data(&TimeSequenceData {
        concurrent: Some(false),
        next_action: None,
        previous_action: None,
    });
    bad_sequence_bytes[8] = 2;
    let (bad_sequence_record, _) = Record::parse(&bad_sequence_bytes, 0).unwrap();
    assert!(parse_time_sequence_data(&bad_sequence_record).is_err());
}
