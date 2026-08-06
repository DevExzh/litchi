//! Record-level parsing for shared PowerPoint animation behaviors.

use super::super::support::{read_f32, read_i32, read_u32, require_atom, require_container};
use super::model::{
    parse_animate_color, parse_animate_color_by, parse_generic_time_variant,
    parse_time_string_list, parse_time_variant_bool, parse_time_variant_f32,
    parse_time_variant_i32, parse_time_variant_string, validate_time_formula,
};
use super::validation::{
    validate_animate_behavior, validate_color_behavior, validate_effect_behavior,
    validate_motion_behavior,
};
use crate::animation::types::{
    ChartBuildType, TimeAnimateBehavior, TimeAnimateBehaviorAtom, TimeAnimateCalculationMode,
    TimeAnimateValueType, TimeAnimationValue, TimeAnimationValueList, TimeBehavior,
    TimeBehaviorAdditive, TimeBehaviorAtom, TimeBehaviorProperty, TimeBehaviorPropertyList,
    TimeColorBehavior, TimeColorBehaviorAtom, TimeColorDirection, TimeColorModel,
    TimeEffectBehavior, TimeEffectBehaviorAtom, TimeEffectFilter, TimeEffectTransition,
    TimeMotionBehavior, TimeMotionBehaviorAtom, TimeMotionOrigin, TimeVisualElement,
    TimeVisualElementKind, is_valid_runtime_context, is_valid_time_points_types,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse the common behavior information shared by all extended animation behaviors.
pub fn parse_time_behavior(record: &Record) -> Result<TimeBehavior> {
    require_container(
        record,
        RecordType::TimeBehaviorContainer,
        0,
        "TimeBehaviorContainer",
    )?;
    let atom_record = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("TimeBehaviorContainer has no atom".to_string()))?;
    let atom = parse_time_behavior_atom(atom_record)?;
    let mut index = 1;
    let attribute_names = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == RecordType::TimeVariantList)
    {
        let names = parse_time_string_list(&record.children[index])?;
        index += 1;
        Some(names)
    } else {
        None
    };
    let properties = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == RecordType::TimePropertyList)
    {
        let properties = parse_time_behavior_property_list(&record.children[index])?;
        index += 1;
        Some(properties)
    } else {
        None
    };
    let target = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("TimeBehaviorContainer has no target".to_string()))
        .and_then(parse_time_visual_element)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "TimeBehaviorContainer has invalid child order or extra children".to_string(),
        ));
    }
    Ok(TimeBehavior {
        atom,
        attribute_names,
        properties,
        target,
    })
}

/// Parse an exact 16-byte `TimeBehaviorAtom` payload.
pub fn parse_time_behavior_atom(record: &Record) -> Result<TimeBehaviorAtom> {
    require_atom(record, RecordType::TimeBehavior, 0, 16, "TimeBehaviorAtom")?;
    let flags = read_u32(&record.data, 0);
    let additive_value = read_u32(&record.data, 4);
    let additive = if flags & 0x01 != 0 {
        Some(match additive_value {
            0 => TimeBehaviorAdditive::Override,
            1 => TimeBehaviorAdditive::Add,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid TimeBehavior additive mode {value}"
                )));
            },
        })
    } else if additive_value == 0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "TimeBehavior additive mode must be zero when not explicitly set".to_string(),
        ));
    };
    if read_u32(&record.data, 8) != 0 || read_u32(&record.data, 12) != 0 {
        return Err(Error::InvalidFormat(
            "TimeBehavior accumulation and transform modes must be zero".to_string(),
        ));
    }
    Ok(TimeBehaviorAtom {
        additive,
        attribute_names_used: flags & 0x04 != 0,
    })
}

/// Parse an exact generic property-animation behavior container.
pub fn parse_time_animate_behavior(record: &Record) -> Result<TimeAnimateBehavior> {
    require_container(
        record,
        RecordType::TimeAnimateBehaviorContainer,
        0,
        "TimeAnimateBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("animate behavior has no atom".to_string()))
        .and_then(parse_time_animate_behavior_atom)?;
    let mut index = 1;
    let values = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == RecordType::TimeAnimationValueList)
    {
        let values = parse_time_animation_value_list(&record.children[index])?;
        index += 1;
        Some(values)
    } else {
        None
    };
    let mut take_string = |instance| -> Result<Option<String>> {
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == instance
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Ok(Some(value))
        } else {
            Ok(None)
        }
    };
    let by = take_string(1)?;
    let from = take_string(2)?;
    let to = take_string(3)?;
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("animate behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "animate behavior has invalid child order or extra children".to_string(),
        ));
    }
    let animate = TimeAnimateBehavior {
        atom,
        values,
        by,
        from,
        to,
        behavior,
    };
    validate_animate_behavior(&animate)?;
    Ok(animate)
}

/// Parse an exact 12-byte `TimeAnimateBehaviorAtom` payload.
pub fn parse_time_animate_behavior_atom(record: &Record) -> Result<TimeAnimateBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeAnimateBehavior,
        0,
        12,
        "TimeAnimateBehaviorAtom",
    )?;
    let mode_value = read_u32(&record.data, 0);
    let flags = read_u32(&record.data, 4);
    let calculation_mode = if flags & 0x08 != 0 {
        Some(match mode_value {
            0 => TimeAnimateCalculationMode::Discrete,
            1 => TimeAnimateCalculationMode::Linear,
            2 => TimeAnimateCalculationMode::Formula,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid animate calculation mode {value}"
                )));
            },
        })
    } else if mode_value == 1 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "animate calculation mode must be linear when unused".to_string(),
        ));
    };
    let type_value = read_u32(&record.data, 8);
    let value_type = if flags & 0x20 != 0 {
        Some(match type_value {
            0 => TimeAnimateValueType::String,
            1 => TimeAnimateValueType::Number,
            2 => TimeAnimateValueType::Color,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid animate value type {value}"
                )));
            },
        })
    } else if type_value == 1 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "animate value type must be numeric when unused".to_string(),
        ));
    };
    Ok(TimeAnimateBehaviorAtom {
        calculation_mode,
        by_used: flags & 0x01 != 0,
        from_used: flags & 0x02 != 0,
        to_used: flags & 0x04 != 0,
        animation_values_used: flags & 0x10 != 0,
        value_type,
    })
}

/// Parse a generic animation keyframe list.
pub fn parse_time_animation_value_list(record: &Record) -> Result<TimeAnimationValueList> {
    require_container(
        record,
        RecordType::TimeAnimationValueList,
        0,
        "TimeAnimationValueListContainer",
    )?;
    let mut entries = Vec::new();
    let mut index = 0;
    while index < record.children.len() {
        let time = parse_time_animation_value_atom(&record.children[index])?;
        index += 1;
        let value = if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 0
        }) {
            let value = parse_generic_time_variant(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
        let formula = if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 1
        }) {
            let formula = parse_time_variant_string(&record.children[index])?;
            validate_time_formula(&formula)?;
            index += 1;
            Some(formula)
        } else {
            None
        };
        entries.push(TimeAnimationValue {
            time,
            value,
            formula,
        });
    }
    Ok(TimeAnimationValueList { entries })
}

/// Parse an exact 4-byte `TimeAnimationValueAtom` payload.
pub fn parse_time_animation_value_atom(record: &Record) -> Result<i32> {
    require_atom(
        record,
        RecordType::TimeAnimationValue,
        0,
        4,
        "TimeAnimationValueAtom",
    )?;
    let time = read_i32(&record.data, 0);
    if time != -1000 && !(0..=1000).contains(&time) {
        return Err(Error::InvalidFormat(
            "animation keyframe time is out of range".to_string(),
        ));
    }
    Ok(time)
}

/// Parse an exact color behavior container.
pub fn parse_time_color_behavior(record: &Record) -> Result<TimeColorBehavior> {
    require_container(
        record,
        RecordType::TimeColorBehaviorContainer,
        0,
        "TimeColorBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(Error::InvalidFormat(
            "TimeColorBehaviorContainer requires an atom and common behavior".to_string(),
        ));
    }
    let atom = parse_time_color_behavior_atom(&record.children[0])?;
    let behavior = parse_time_behavior(&record.children[1])?;
    validate_color_behavior(&atom, &behavior)?;
    Ok(TimeColorBehavior { atom, behavior })
}

/// Parse an exact 52-byte `TimeColorBehaviorAtom` payload.
pub fn parse_time_color_behavior_atom(record: &Record) -> Result<TimeColorBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeColorBehavior,
        0,
        52,
        "TimeColorBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by = (flags & 0x01 != 0)
        .then(|| parse_animate_color_by(&record.data[4..20]))
        .transpose()?;
    let from = (flags & 0x02 != 0)
        .then(|| parse_animate_color(&record.data[20..36]))
        .transpose()?;
    let to = (flags & 0x04 != 0)
        .then(|| parse_animate_color(&record.data[36..52]))
        .transpose()?;
    if from.is_some() && by.is_none() && to.is_none() {
        return Err(Error::InvalidFormat(
            "color from value requires a by or to value".to_string(),
        ));
    }
    Ok(TimeColorBehaviorAtom {
        by,
        from,
        to,
        color_space_used: flags & 0x08 != 0,
        direction_used: flags & 0x10 != 0,
    })
}

/// Parse an exact image-effect behavior container.
pub fn parse_time_effect_behavior(record: &Record) -> Result<TimeEffectBehavior> {
    require_container(
        record,
        RecordType::TimeEffectBehaviorContainer,
        0,
        "TimeEffectBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("effect behavior has no atom".to_string()))
        .and_then(parse_time_effect_behavior_atom)?;
    let mut index = 1;
    let filter =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 1
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(TimeEffectFilter::parse(&value).ok_or_else(|| {
                Error::InvalidFormat(format!("invalid image-effect filter {value}"))
            })?)
        } else {
            None
        };
    let progress =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 2
        }) {
            let value = parse_time_variant_f32(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let runtime_context =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 3
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("effect behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "effect behavior has invalid child order or extra children".to_string(),
        ));
    }
    let effect = TimeEffectBehavior {
        atom,
        filter,
        progress,
        runtime_context,
        behavior,
    };
    validate_effect_behavior(&effect)?;
    Ok(effect)
}

/// Parse an exact 8-byte `TimeEffectBehaviorAtom` payload.
pub fn parse_time_effect_behavior_atom(record: &Record) -> Result<TimeEffectBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeEffectBehavior,
        0,
        8,
        "TimeEffectBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let value = read_u32(&record.data, 4);
    let transition = if flags & 0x01 != 0 {
        Some(match value {
            0 => TimeEffectTransition::In,
            1 => TimeEffectTransition::Out,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid image-effect transition {value}"
                )));
            },
        })
    } else if value == 0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "image-effect transition must be transition-in when unused".to_string(),
        ));
    };
    Ok(TimeEffectBehaviorAtom {
        transition,
        filter_used: flags & 0x02 != 0,
        progress_used: flags & 0x04 != 0,
        runtime_context_used: flags & 0x08 != 0,
    })
}

/// Parse an exact motion-path behavior container.
pub fn parse_time_motion_behavior(record: &Record) -> Result<TimeMotionBehavior> {
    require_container(
        record,
        RecordType::TimeMotionBehaviorContainer,
        0,
        "TimeMotionBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("motion behavior has no atom".to_string()))
        .and_then(parse_time_motion_behavior_atom)?;
    let mut index = 1;
    let path =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 1
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let reserved =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 2
        }) {
            let value = parse_time_variant_i32(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("motion behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "motion behavior has invalid child order or extra children".to_string(),
        ));
    }
    let motion = TimeMotionBehavior {
        atom,
        path,
        reserved,
        behavior,
    };
    validate_motion_behavior(&motion)?;
    Ok(motion)
}

/// Parse an exact 32-byte `TimeMotionBehaviorAtom` payload.
pub fn parse_time_motion_behavior_atom(record: &Record) -> Result<TimeMotionBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeMotionBehavior,
        0,
        32,
        "TimeMotionBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by = (flags & 0x01 != 0).then(|| (read_f32(&record.data, 4), read_f32(&record.data, 8)));
    let from =
        (flags & 0x02 != 0).then(|| (read_f32(&record.data, 12), read_f32(&record.data, 16)));
    let to = (flags & 0x04 != 0).then(|| (read_f32(&record.data, 20), read_f32(&record.data, 24)));
    if from.is_some() && by.is_none() && to.is_none() {
        return Err(Error::InvalidFormat(
            "motion from values require by or to values".to_string(),
        ));
    }
    let origin_value = read_u32(&record.data, 28);
    let origin = if flags & 0x08 != 0 {
        Some(match origin_value {
            0 => TimeMotionOrigin::Slide,
            1 => TimeMotionOrigin::SlideLegacy,
            2 => TimeMotionOrigin::ObjectCenter,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid motion origin {value}"
                )));
            },
        })
    } else if origin_value == 2 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "motion origin must be object center when unused".to_string(),
        ));
    };
    Ok(TimeMotionBehaviorAtom {
        by,
        from,
        to,
        origin,
        path_used: flags & 0x10 != 0,
        edit_rotation_used: flags & 0x40 != 0,
        points_types_used: flags & 0x80 != 0,
    })
}

/// Parse a `TimePropertyList4TimeBehavior` record.
pub fn parse_time_behavior_property_list(record: &Record) -> Result<TimeBehaviorPropertyList> {
    require_container(
        record,
        RecordType::TimePropertyList,
        0,
        "TimePropertyList4TimeBehavior",
    )?;
    let mut seen = std::collections::HashSet::with_capacity(record.children.len());
    let mut properties = Vec::with_capacity(record.children.len());
    for child in &record.children {
        if child.record_type != RecordType::TimeVariant || child.version != 0 {
            return Err(Error::InvalidFormat(
                "invalid TimePropertyList4TimeBehavior child".to_string(),
            ));
        }
        if !seen.insert(child.instance) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time behavior property {:#X}",
                child.instance
            )));
        }
        properties.push(parse_time_behavior_property(child)?);
    }
    Ok(TimeBehaviorPropertyList { properties })
}

fn parse_time_behavior_property(record: &Record) -> Result<TimeBehaviorProperty> {
    let property = match record.instance {
        0x01 => TimeBehaviorProperty::UnknownPropertyList(parse_time_variant_string(record)?),
        0x02 => {
            let value = parse_time_variant_string(record)?;
            if !is_valid_runtime_context(&value) {
                return Err(Error::InvalidFormat(
                    "invalid time runtime context".to_string(),
                ));
            }
            TimeBehaviorProperty::RuntimeContext(value)
        },
        0x03 => TimeBehaviorProperty::MotionPathEditRelative(parse_time_variant_bool(record)?),
        0x04 => TimeBehaviorProperty::ColorModel(match parse_time_variant_i32(record)? {
            0 => TimeColorModel::Rgb,
            1 => TimeColorModel::Hsl,
            2 => TimeColorModel::Scheme,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid time color model {value}"
                )));
            },
        }),
        0x05 => TimeBehaviorProperty::ColorDirection(match parse_time_variant_i32(record)? {
            0 => TimeColorDirection::Clockwise,
            1 => TimeColorDirection::CounterClockwise,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid time color direction {value}"
                )));
            },
        }),
        0x06 => match parse_time_variant_i32(record)? {
            1 => TimeBehaviorProperty::Override,
            _ => {
                return Err(Error::InvalidFormat(
                    "invalid time behavior override".to_string(),
                ));
            },
        },
        0x07 => TimeBehaviorProperty::PathEditRotationAngle(parse_time_variant_f32(record)?),
        0x08 => TimeBehaviorProperty::PathEditRotationX(parse_time_variant_f32(record)?),
        0x09 => TimeBehaviorProperty::PathEditRotationY(parse_time_variant_f32(record)?),
        0x0A => {
            let value = parse_time_variant_string(record)?;
            if !is_valid_time_points_types(&value) {
                return Err(Error::InvalidFormat(
                    "invalid time path point types".to_string(),
                ));
            }
            TimeBehaviorProperty::PointsTypes(value)
        },
        id => {
            return Err(Error::InvalidFormat(format!(
                "unknown time behavior property {id:#X}"
            )));
        },
    };
    Ok(property)
}

/// Parse a `ClientVisualElementContainer` animation target.
pub fn parse_time_visual_element(record: &Record) -> Result<TimeVisualElement> {
    require_container(
        record,
        RecordType::TimeClientVisualElement,
        0,
        "ClientVisualElementContainer",
    )?;
    if record.children.len() != 1 {
        return Err(Error::InvalidFormat(
            "ClientVisualElementContainer requires exactly one atom".to_string(),
        ));
    }
    let atom = &record.children[0];
    if atom.record_type == RecordType::VisualPageAtom {
        require_atom(atom, RecordType::VisualPageAtom, 0, 4, "VisualPageAtom")?;
        if read_u32(&atom.data, 0) != TimeVisualElementKind::Page.as_u32() {
            return Err(Error::InvalidFormat(
                "VisualPageAtom has a non-page target type".to_string(),
            ));
        }
        return Ok(TimeVisualElement::Page);
    }
    require_atom(
        atom,
        RecordType::VisualShapeAtom,
        0,
        20,
        "VisualShapeOrSoundAtom",
    )?;
    let kind = TimeVisualElementKind::parse(read_u32(&atom.data, 0))
        .ok_or_else(|| Error::InvalidFormat("invalid visual element target type".to_string()))?;
    if kind == TimeVisualElementKind::Page {
        return Err(Error::InvalidFormat(
            "VisualShapeOrSoundAtom cannot target a page".to_string(),
        ));
    }
    match read_u32(&atom.data, 4) {
        1 if kind == TimeVisualElementKind::ChartElement => {
            let build_type = ChartBuildType::parse(read_u32(&atom.data, 12)).ok_or_else(|| {
                Error::InvalidFormat("invalid chart target build type".to_string())
            })?;
            let element_index = read_i32(&atom.data, 16);
            if element_index < -1 {
                return Err(Error::InvalidFormat(
                    "chart target element index must be at least -1".to_string(),
                ));
            }
            Ok(TimeVisualElement::Chart {
                shape_id_ref: read_u32(&atom.data, 8),
                build_type,
                element_index,
            })
        },
        1 => Ok(TimeVisualElement::Shape {
            kind,
            shape_id_ref: read_u32(&atom.data, 8),
            data1: read_i32(&atom.data, 12),
            data2: read_i32(&atom.data, 16),
        }),
        2 => {
            if read_u32(&atom.data, 12) != u32::MAX || read_u32(&atom.data, 16) != u32::MAX {
                return Err(Error::InvalidFormat(
                    "VisualSoundAtom reserved data must be -1".to_string(),
                ));
            }
            Ok(TimeVisualElement::Sound {
                kind,
                sound_id_ref: read_u32(&atom.data, 8),
            })
        },
        value => Err(Error::InvalidFormat(format!(
            "invalid visual element reference type {value}"
        ))),
    }
}
