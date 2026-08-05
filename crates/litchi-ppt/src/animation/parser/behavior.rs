//! Common, property, color, effect, motion, and target behavior records.

use super::support::{parse_bool1, read_f32, read_i32, read_u32, require_atom, require_container};
use super::timeline::validate_basic_behavior_properties;
use crate::animation::types::{
    ChartBuildType, TimeAnimateBehavior, TimeAnimateBehaviorAtom, TimeAnimateCalculationMode,
    TimeAnimateColor, TimeAnimateColorBy, TimeAnimateValueType, TimeAnimationValue,
    TimeAnimationValueList, TimeBehavior, TimeBehaviorAdditive, TimeBehaviorAtom,
    TimeBehaviorProperty, TimeBehaviorPropertyList, TimeColorBehavior, TimeColorBehaviorAtom,
    TimeColorDirection, TimeColorModel, TimeEffectBehavior, TimeEffectBehaviorAtom,
    TimeEffectFilter, TimeEffectTransition, TimeMotionBehavior, TimeMotionBehaviorAtom,
    TimeMotionOrigin, TimeVariantValue, TimeVisualElement, TimeVisualElementKind,
    is_valid_animation_attribute_name, is_valid_motion_path, is_valid_runtime_context,
    is_valid_time_animate_value, is_valid_time_formula, is_valid_time_points_types,
    time_animation_attribute_value_type,
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
            if !is_valid_time_formula(&formula) {
                return Err(Error::InvalidFormat(
                    "invalid animation keyframe formula".to_string(),
                ));
            }
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

fn parse_generic_time_variant(record: &Record) -> Result<TimeVariantValue> {
    match record.data.first() {
        Some(0) => parse_time_variant_bool(record).map(TimeVariantValue::Boolean),
        Some(1) => parse_time_variant_i32(record).map(TimeVariantValue::Integer),
        Some(2) => parse_time_variant_f32(record).map(TimeVariantValue::Float),
        Some(3) => parse_time_variant_string(record).map(TimeVariantValue::String),
        _ => Err(Error::InvalidFormat(
            "invalid animation keyframe value type".to_string(),
        )),
    }
}

fn validate_animate_behavior(animate: &TimeAnimateBehavior) -> Result<()> {
    for (used, value, field) in [
        (animate.atom.by_used, animate.by.as_ref(), "by"),
        (animate.atom.from_used, animate.from.as_ref(), "from"),
        (animate.atom.to_used, animate.to.as_ref(), "to"),
    ] {
        if used && value.is_none() {
            return Err(Error::InvalidFormat(format!(
                "animate {field}-use flag requires a value"
            )));
        }
    }
    if animate.atom.animation_values_used && animate.values.is_none() {
        return Err(Error::InvalidFormat(
            "animate-values-use flag requires a keyframe list".to_string(),
        ));
    }
    if animate.from.is_some() && animate.by.is_none() && animate.to.is_none() {
        return Err(Error::InvalidFormat(
            "animate from value requires a by or to value".to_string(),
        ));
    }
    if !animate.behavior.atom.attribute_names_used {
        return Err(Error::InvalidFormat(
            "animate behavior requires an explicit attribute name".to_string(),
        ));
    }
    let attribute = match animate.behavior.attribute_names.as_deref() {
        Some([attribute]) => attribute.as_str(),
        _ => {
            return Err(Error::InvalidFormat(
                "animate behavior requires exactly one attribute name".to_string(),
            ));
        },
    };
    let expected_type = time_animation_attribute_value_type(attribute).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "unsupported animate behavior attribute {attribute}"
        ))
    })?;
    let actual_type = animate
        .atom
        .value_type
        .unwrap_or(TimeAnimateValueType::Number);
    if actual_type != expected_type {
        return Err(Error::InvalidFormat(
            "animate value type does not match its attribute".to_string(),
        ));
    }
    if [&animate.by, &animate.from, &animate.to]
        .into_iter()
        .flatten()
        .any(|value| !is_valid_time_animate_value(attribute, actual_type, value))
    {
        return Err(Error::InvalidFormat(
            "animate value is invalid for its attribute".to_string(),
        ));
    }
    if animate.atom.calculation_mode == Some(TimeAnimateCalculationMode::Formula)
        && !animate
            .values
            .as_ref()
            .is_some_and(|list| list.entries.iter().any(|entry| entry.formula.is_some()))
    {
        return Err(Error::InvalidFormat(
            "formula calculation mode requires a keyframe formula".to_string(),
        ));
    }
    validate_basic_behavior_properties(&animate.behavior)
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

fn parse_animate_color_by(data: &[u8]) -> Result<TimeAnimateColorBy> {
    match read_u32(data, 0) {
        0 | 1 => {
            let values = (read_i32(data, 4), read_i32(data, 8), read_i32(data, 12));
            if [values.0, values.1, values.2]
                .iter()
                .any(|value| !(-255..=255).contains(value))
            {
                return Err(Error::InvalidFormat(
                    "color offset component is out of range".to_string(),
                ));
            }
            if read_u32(data, 0) == 0 {
                Ok(TimeAnimateColorBy::Rgb {
                    red: values.0,
                    green: values.1,
                    blue: values.2,
                })
            } else {
                Ok(TimeAnimateColorBy::Hsl {
                    hue: values.0,
                    saturation: values.1,
                    luminance: values.2,
                })
            }
        },
        2 => parse_scheme_color(data).map(TimeAnimateColorBy::Scheme),
        model => Err(Error::InvalidFormat(format!(
            "invalid color-by model {model}"
        ))),
    }
}

fn parse_animate_color(data: &[u8]) -> Result<TimeAnimateColor> {
    match read_u32(data, 0) {
        0 => {
            let (red, green, blue) = (read_u32(data, 4), read_u32(data, 8), read_u32(data, 12));
            if red > 255 || green > 255 || blue > 255 {
                return Err(Error::InvalidFormat(
                    "RGB color component is out of range".to_string(),
                ));
            }
            Ok(TimeAnimateColor::Rgb { red, green, blue })
        },
        2 => parse_scheme_color(data).map(TimeAnimateColor::Scheme),
        model => Err(Error::InvalidFormat(format!(
            "invalid absolute color model {model}"
        ))),
    }
}

fn parse_scheme_color(data: &[u8]) -> Result<u32> {
    let index = read_u32(data, 4);
    if index > 7 {
        return Err(Error::InvalidFormat(
            "scheme color index is out of range".to_string(),
        ));
    }
    Ok(index)
}

fn validate_color_behavior(atom: &TimeColorBehaviorAtom, behavior: &TimeBehavior) -> Result<()> {
    const NAMES: &[&str] = &[
        "ppt_c",
        "style.color",
        "imageData.chromakey",
        "fill.color",
        "fill.color2",
        "stroke.color",
        "stroke.color2",
        "shadow.color",
        "shadow.color2",
        "extrusion.color",
        "fillcolor",
    ];
    if !behavior.atom.attribute_names_used
        || !matches!(behavior.attribute_names.as_deref(), Some([name]) if NAMES.contains(&name.as_str()))
    {
        return Err(Error::InvalidFormat(
            "color behavior requires exactly one supported color attribute".to_string(),
        ));
    }
    let properties = behavior
        .properties
        .as_ref()
        .map_or(&[][..], |list| list.properties.as_slice());
    if properties.iter().any(|property| {
        matches!(
            property,
            TimeBehaviorProperty::MotionPathEditRelative(_)
                | TimeBehaviorProperty::PathEditRotationAngle(_)
                | TimeBehaviorProperty::PathEditRotationX(_)
                | TimeBehaviorProperty::PathEditRotationY(_)
                | TimeBehaviorProperty::PointsTypes(_)
        )
    }) {
        return Err(Error::InvalidFormat(
            "color behavior contains a motion-only property".to_string(),
        ));
    }
    if atom.color_space_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::ColorModel(_)))
    {
        return Err(Error::InvalidFormat(
            "color-space-used flag requires a color model property".to_string(),
        ));
    }
    if atom.direction_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::ColorDirection(_)))
    {
        return Err(Error::InvalidFormat(
            "direction-used flag requires a color direction property".to_string(),
        ));
    }
    Ok(())
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

fn validate_effect_behavior(effect: &TimeEffectBehavior) -> Result<()> {
    if effect.atom.filter_used && effect.filter.is_none() {
        return Err(Error::InvalidFormat(
            "image-effect filter-use flag requires a filter".to_string(),
        ));
    }
    if effect.atom.progress_used && effect.progress.is_none() {
        return Err(Error::InvalidFormat(
            "image-effect progress-use flag requires progress".to_string(),
        ));
    }
    if effect.atom.runtime_context_used && effect.runtime_context.is_none() {
        return Err(Error::InvalidFormat(
            "image-effect runtime-context-use flag requires a runtime context".to_string(),
        ));
    }
    if effect
        .progress
        .is_some_and(|value| !(0.0..=1.0).contains(&value))
    {
        return Err(Error::InvalidFormat(
            "image-effect progress must be between zero and one".to_string(),
        ));
    }
    if effect
        .runtime_context
        .as_deref()
        .is_some_and(|value| !is_valid_runtime_context(value))
    {
        return Err(Error::InvalidFormat(
            "invalid image-effect runtime context".to_string(),
        ));
    }
    validate_basic_behavior_properties(&effect.behavior)
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

fn validate_motion_behavior(motion: &TimeMotionBehavior) -> Result<()> {
    if motion.atom.path_used && motion.path.is_none() {
        return Err(Error::InvalidFormat(
            "motion path-use flag requires a path".to_string(),
        ));
    }
    if motion
        .path
        .as_deref()
        .is_some_and(|path| !is_valid_motion_path(path))
    {
        return Err(Error::InvalidFormat(
            "invalid motion path syntax".to_string(),
        ));
    }
    let properties = motion
        .behavior
        .properties
        .as_ref()
        .map_or(&[][..], |list| list.properties.as_slice());
    if properties.iter().any(|property| {
        matches!(
            property,
            TimeBehaviorProperty::ColorModel(_) | TimeBehaviorProperty::ColorDirection(_)
        )
    }) {
        return Err(Error::InvalidFormat(
            "motion behavior contains a color-only property".to_string(),
        ));
    }
    let has_angle = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationAngle(_)));
    let has_x = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationX(_)));
    let has_y = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationY(_)));
    if motion.atom.edit_rotation_used && !(has_angle && has_x && has_y) {
        return Err(Error::InvalidFormat(
            "motion edit-rotation flag requires angle, X, and Y properties".to_string(),
        ));
    }
    if motion.atom.points_types_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::PointsTypes(_)))
    {
        return Err(Error::InvalidFormat(
            "motion points-types flag requires a points-types property".to_string(),
        ));
    }
    if motion.behavior.atom.attribute_names_used
        && motion
            .behavior
            .attribute_names
            .as_ref()
            .is_some_and(|names| {
                names.len() > 2
                    || names
                        .iter()
                        .any(|name| !is_valid_animation_attribute_name(name))
            })
    {
        return Err(Error::InvalidFormat(
            "motion behavior requires at most two supported attribute names".to_string(),
        ));
    }
    Ok(())
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

fn parse_time_string_list(record: &Record) -> Result<Vec<String>> {
    require_container(
        record,
        RecordType::TimeVariantList,
        1,
        "TimeStringListContainer",
    )?;
    record
        .children
        .iter()
        .map(|child| {
            if child.record_type != RecordType::TimeVariant || child.version != 0 {
                return Err(Error::InvalidFormat(
                    "invalid TimeStringListContainer child".to_string(),
                ));
            }
            parse_time_variant_string(child)
        })
        .collect()
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

fn parse_time_variant_i32(record: &Record) -> Result<i32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 1 {
        return Err(Error::InvalidFormat(
            "invalid integer time variant".to_string(),
        ));
    }
    Ok(i32::from_le_bytes(
        record.data[1..5].try_into().expect("length checked"),
    ))
}

fn parse_time_variant_f32(record: &Record) -> Result<f32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 2 {
        return Err(Error::InvalidFormat(
            "invalid floating-point time variant".to_string(),
        ));
    }
    Ok(f32::from_le_bytes(
        record.data[1..5].try_into().expect("length checked"),
    ))
}

fn parse_time_variant_bool(record: &Record) -> Result<bool> {
    require_time_variant_payload(record)?;
    if record.data.len() != 2 || record.data[0] != 0 {
        return Err(Error::InvalidFormat(
            "invalid Boolean time variant".to_string(),
        ));
    }
    parse_bool1(record.data[1], "TimeVariant.boolValue")
}

pub(super) fn parse_time_variant_string(record: &Record) -> Result<String> {
    require_time_variant_payload(record)?;
    if record.data.len() % 2 != 1 || record.data.first() != Some(&3) {
        return Err(Error::InvalidFormat(
            "invalid string time variant".to_string(),
        ));
    }
    String::from_utf16(
        &record.data[1..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect::<Vec<_>>(),
    )
    .map_err(|_| Error::InvalidFormat("invalid UTF-16 time variant".to_string()))
}

pub(super) fn require_time_variant_payload(record: &Record) -> Result<()> {
    if record.data_length as usize != record.data.len() {
        return Err(Error::Corrupted(
            "truncated TimeVariant payload".to_string(),
        ));
    }
    Ok(())
}
