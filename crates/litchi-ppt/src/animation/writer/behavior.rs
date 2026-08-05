//! Common, property, color, effect, motion, and target behavior records.

use super::support::{create_record_header, wrap_record};
use super::timeline::validate_basic_behavior_properties;
use crate::animation::types::{
    TimeAnimateBehavior, TimeAnimateBehaviorAtom, TimeAnimateCalculationMode, TimeAnimateColor,
    TimeAnimateColorBy, TimeAnimateValueType, TimeAnimationValueList, TimeBehavior,
    TimeBehaviorAdditive, TimeBehaviorAtom, TimeBehaviorProperty, TimeBehaviorPropertyList,
    TimeColorBehavior, TimeColorBehaviorAtom, TimeColorDirection, TimeColorModel,
    TimeEffectBehavior, TimeEffectBehaviorAtom, TimeEffectTransition, TimeMotionBehavior,
    TimeMotionBehaviorAtom, TimeMotionOrigin, TimeVariantValue, TimeVisualElement,
    TimeVisualElementKind, is_valid_animation_attribute_name, is_valid_motion_path,
    is_valid_runtime_context, is_valid_time_animate_value, is_valid_time_formula,
    is_valid_time_points_types, time_animation_attribute_value_type,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// Serialize the common behavior information shared by extended animation behaviors.
pub fn write_time_behavior(behavior: &TimeBehavior) -> Result<Vec<u8>> {
    let mut children = write_time_behavior_atom(&behavior.atom);
    if let Some(attribute_names) = &behavior.attribute_names {
        children.extend(write_time_string_list(attribute_names)?);
    }
    if let Some(properties) = &behavior.properties {
        children.extend(write_time_behavior_property_list(properties)?);
    }
    children.extend(write_time_visual_element(&behavior.target)?);
    wrap_record(RecordType::TimeBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact 16-byte `TimeBehaviorAtom` payload.
pub fn write_time_behavior_atom(atom: &TimeBehaviorAtom) -> Vec<u8> {
    let mut data = Vec::with_capacity(16);
    let flags = u32::from(atom.additive.is_some()) | (u32::from(atom.attribute_names_used) << 2);
    data.extend(flags.to_le_bytes());
    data.extend(
        atom.additive
            .map_or(0u32, |value| match value {
                TimeBehaviorAdditive::Override => 0,
                TimeBehaviorAdditive::Add => 1,
            })
            .to_le_bytes(),
    );
    data.extend(0u32.to_le_bytes());
    data.extend(0u32.to_le_bytes());
    let mut result = create_record_header(RecordType::TimeBehavior, 0, 0, 16);
    result.extend(data);
    result
}

/// Serialize an exact generic property-animation behavior container.
pub fn write_time_animate_behavior(animate: &TimeAnimateBehavior) -> Result<Vec<u8>> {
    validate_animate_behavior(animate)?;
    let mut children = write_time_animate_behavior_atom(&animate.atom);
    if let Some(values) = &animate.values {
        children.extend(write_time_animation_value_list(values)?);
    }
    for (instance, value) in [(1, &animate.by), (2, &animate.from), (3, &animate.to)] {
        if let Some(value) = value {
            append_time_variant(&mut children, instance, encode_time_variant_string(value))?;
        }
    }
    children.extend(write_time_behavior(&animate.behavior)?);
    wrap_record(RecordType::TimeAnimateBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeAnimateBehaviorAtom`.
pub fn write_time_animate_behavior_atom(atom: &TimeAnimateBehaviorAtom) -> Vec<u8> {
    let mode = atom.calculation_mode.map_or(1u32, |mode| match mode {
        TimeAnimateCalculationMode::Discrete => 0,
        TimeAnimateCalculationMode::Linear => 1,
        TimeAnimateCalculationMode::Formula => 2,
    });
    let flags = u32::from(atom.by_used)
        | (u32::from(atom.from_used) << 1)
        | (u32::from(atom.to_used) << 2)
        | (u32::from(atom.calculation_mode.is_some()) << 3)
        | (u32::from(atom.animation_values_used) << 4)
        | (u32::from(atom.value_type.is_some()) << 5);
    let value_type = atom.value_type.map_or(1u32, |value| match value {
        TimeAnimateValueType::String => 0,
        TimeAnimateValueType::Number => 1,
        TimeAnimateValueType::Color => 2,
    });
    let mut result = create_record_header(RecordType::TimeAnimateBehavior, 0, 0, 12);
    result.extend(mode.to_le_bytes());
    result.extend(flags.to_le_bytes());
    result.extend(value_type.to_le_bytes());
    result
}

/// Serialize a generic animation keyframe list.
pub fn write_time_animation_value_list(list: &TimeAnimationValueList) -> Result<Vec<u8>> {
    let mut children = Vec::new();
    for entry in &list.entries {
        children.extend(write_time_animation_value_atom(entry.time)?);
        if let Some(value) = &entry.value {
            append_time_variant(&mut children, 0, encode_generic_time_variant(value))?;
        }
        if let Some(formula) = &entry.formula {
            if !is_valid_time_formula(formula) {
                return Err(Error::InvalidFormat(
                    "invalid animation keyframe formula".to_string(),
                ));
            }
            append_time_variant(&mut children, 1, encode_time_variant_string(formula))?;
        }
    }
    wrap_record(RecordType::TimeAnimationValueList, 0x0F, 0, children)
}

/// Serialize an exact `TimeAnimationValueAtom`.
pub fn write_time_animation_value_atom(time: i32) -> Result<Vec<u8>> {
    if time != -1000 && !(0..=1000).contains(&time) {
        return Err(Error::InvalidFormat(
            "animation keyframe time is out of range".to_string(),
        ));
    }
    let mut result = create_record_header(RecordType::TimeAnimationValue, 0, 0, 4);
    result.extend(time.to_le_bytes());
    Ok(result)
}

fn encode_generic_time_variant(value: &TimeVariantValue) -> Vec<u8> {
    match value {
        TimeVariantValue::Boolean(value) => vec![0, u8::from(*value)],
        TimeVariantValue::Integer(value) => {
            let mut data = vec![1];
            data.extend(value.to_le_bytes());
            data
        },
        TimeVariantValue::Float(value) => {
            let mut data = vec![2];
            data.extend(value.to_le_bytes());
            data
        },
        TimeVariantValue::String(value) => encode_time_variant_string(value),
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

/// Serialize an exact color behavior container.
pub fn write_time_color_behavior(behavior: &TimeColorBehavior) -> Result<Vec<u8>> {
    validate_color_behavior(&behavior.atom, &behavior.behavior)?;
    let mut children = write_time_color_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(RecordType::TimeColorBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeColorBehaviorAtom`.
pub fn write_time_color_behavior_atom(atom: &TimeColorBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from.is_some() && atom.by.is_none() && atom.to.is_none() {
        return Err(Error::InvalidFormat(
            "color from value requires a by or to value".to_string(),
        ));
    }
    let flags = u32::from(atom.by.is_some())
        | (u32::from(atom.from.is_some()) << 1)
        | (u32::from(atom.to.is_some()) << 2)
        | (u32::from(atom.color_space_used) << 3)
        | (u32::from(atom.direction_used) << 4);
    let mut data = Vec::with_capacity(52);
    data.extend(flags.to_le_bytes());
    data.extend(match &atom.by {
        Some(color) => encode_animate_color_by(color)?,
        None => [0; 16],
    });
    data.extend(match &atom.from {
        Some(color) => encode_animate_color(color)?,
        None => [0; 16],
    });
    data.extend(match &atom.to {
        Some(color) => encode_animate_color(color)?,
        None => [0; 16],
    });
    let mut result = create_record_header(RecordType::TimeColorBehavior, 0, 0, 52);
    result.extend(data);
    Ok(result)
}

fn encode_animate_color_by(color: &TimeAnimateColorBy) -> Result<[u8; 16]> {
    let (model, values) = match color {
        TimeAnimateColorBy::Rgb { red, green, blue } => (0u32, [*red, *green, *blue]),
        TimeAnimateColorBy::Hsl {
            hue,
            saturation,
            luminance,
        } => (1, [*hue, *saturation, *luminance]),
        TimeAnimateColorBy::Scheme(index) => return encode_scheme_color(*index),
    };
    if values.iter().any(|value| !(-255..=255).contains(value)) {
        return Err(Error::InvalidFormat(
            "color offset component is out of range".to_string(),
        ));
    }
    let mut data = [0; 16];
    data[0..4].copy_from_slice(&model.to_le_bytes());
    for (index, value) in values.into_iter().enumerate() {
        data[4 + index * 4..8 + index * 4].copy_from_slice(&value.to_le_bytes());
    }
    Ok(data)
}

fn encode_animate_color(color: &TimeAnimateColor) -> Result<[u8; 16]> {
    match color {
        TimeAnimateColor::Scheme(index) => encode_scheme_color(*index),
        TimeAnimateColor::Rgb { red, green, blue } => {
            if *red > 255 || *green > 255 || *blue > 255 {
                return Err(Error::InvalidFormat(
                    "RGB color component is out of range".to_string(),
                ));
            }
            let mut data = [0; 16];
            for (index, value) in [*red, *green, *blue].into_iter().enumerate() {
                data[4 + index * 4..8 + index * 4].copy_from_slice(&value.to_le_bytes());
            }
            Ok(data)
        },
    }
}

fn encode_scheme_color(index: u32) -> Result<[u8; 16]> {
    if index > 7 {
        return Err(Error::InvalidFormat(
            "scheme color index is out of range".to_string(),
        ));
    }
    let mut data = [0; 16];
    data[0..4].copy_from_slice(&2u32.to_le_bytes());
    data[4..8].copy_from_slice(&index.to_le_bytes());
    Ok(data)
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

/// Serialize an exact image-effect behavior container.
pub fn write_time_effect_behavior(effect: &TimeEffectBehavior) -> Result<Vec<u8>> {
    validate_effect_behavior(effect)?;
    let mut children = write_time_effect_behavior_atom(&effect.atom);
    if let Some(filter) = effect.filter {
        append_time_variant(
            &mut children,
            1,
            encode_time_variant_string(filter.as_str()),
        )?;
    }
    if let Some(progress) = effect.progress {
        let mut data = vec![2];
        data.extend(progress.to_le_bytes());
        append_time_variant(&mut children, 2, data)?;
    }
    if let Some(runtime_context) = &effect.runtime_context {
        append_time_variant(
            &mut children,
            3,
            encode_time_variant_string(runtime_context),
        )?;
    }
    children.extend(write_time_behavior(&effect.behavior)?);
    wrap_record(RecordType::TimeEffectBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeEffectBehaviorAtom`.
pub fn write_time_effect_behavior_atom(atom: &TimeEffectBehaviorAtom) -> Vec<u8> {
    let flags = u32::from(atom.transition.is_some())
        | (u32::from(atom.filter_used) << 1)
        | (u32::from(atom.progress_used) << 2)
        | (u32::from(atom.runtime_context_used) << 3);
    let transition = atom.transition.map_or(0u32, |value| match value {
        TimeEffectTransition::In => 0,
        TimeEffectTransition::Out => 1,
    });
    let mut result = create_record_header(RecordType::TimeEffectBehavior, 0, 0, 8);
    result.extend(flags.to_le_bytes());
    result.extend(transition.to_le_bytes());
    result
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

/// Serialize an exact motion-path behavior container.
pub fn write_time_motion_behavior(motion: &TimeMotionBehavior) -> Result<Vec<u8>> {
    validate_motion_behavior(motion)?;
    let mut children = write_time_motion_behavior_atom(&motion.atom)?;
    if let Some(path) = &motion.path {
        append_time_variant(&mut children, 1, encode_time_variant_string(path))?;
    }
    if let Some(reserved) = motion.reserved {
        let mut data = vec![1];
        data.extend(reserved.to_le_bytes());
        append_time_variant(&mut children, 2, data)?;
    }
    children.extend(write_time_behavior(&motion.behavior)?);
    wrap_record(RecordType::TimeMotionBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeMotionBehaviorAtom`.
pub fn write_time_motion_behavior_atom(atom: &TimeMotionBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from.is_some() && atom.by.is_none() && atom.to.is_none() {
        return Err(Error::InvalidFormat(
            "motion from values require by or to values".to_string(),
        ));
    }
    let flags = u32::from(atom.by.is_some())
        | (u32::from(atom.from.is_some()) << 1)
        | (u32::from(atom.to.is_some()) << 2)
        | (u32::from(atom.origin.is_some()) << 3)
        | (u32::from(atom.path_used) << 4)
        | (u32::from(atom.edit_rotation_used) << 6)
        | (u32::from(atom.points_types_used) << 7);
    let mut data = Vec::with_capacity(32);
    data.extend(flags.to_le_bytes());
    for values in [
        atom.by.unwrap_or((0.0, 0.0)),
        atom.from.unwrap_or((0.0, 0.0)),
        atom.to.unwrap_or((0.0, 0.0)),
    ] {
        data.extend(values.0.to_le_bytes());
        data.extend(values.1.to_le_bytes());
    }
    data.extend(
        atom.origin
            .map_or(2u32, |origin| match origin {
                TimeMotionOrigin::Slide => 0,
                TimeMotionOrigin::SlideLegacy => 1,
                TimeMotionOrigin::ObjectCenter => 2,
            })
            .to_le_bytes(),
    );
    let mut result = create_record_header(RecordType::TimeMotionBehavior, 0, 0, 32);
    result.extend(data);
    Ok(result)
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

pub(super) fn append_time_variant(
    children: &mut Vec<u8>,
    instance: u16,
    data: Vec<u8>,
) -> Result<()> {
    let length = u32::try_from(data.len())
        .map_err(|_| Error::InvalidFormat("time variant exceeds 4 GiB".to_string()))?;
    children.extend(create_record_header(
        RecordType::TimeVariant,
        0,
        instance,
        length,
    ));
    children.extend(data);
    Ok(())
}

/// Serialize a typed `TimePropertyList4TimeBehavior` record.
pub fn write_time_behavior_property_list(list: &TimeBehaviorPropertyList) -> Result<Vec<u8>> {
    let mut seen = std::collections::HashSet::with_capacity(list.properties.len());
    let mut children = Vec::new();
    for property in &list.properties {
        let (id, data) = encode_time_behavior_property(property)?;
        if !seen.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time behavior property {id:#X}"
            )));
        }
        let length = u32::try_from(data.len()).map_err(|_| {
            Error::InvalidFormat("time behavior property exceeds 4 GiB".to_string())
        })?;
        children.extend(create_record_header(RecordType::TimeVariant, 0, id, length));
        children.extend(data);
    }
    wrap_record(RecordType::TimePropertyList, 0x0F, 0, children)
}

fn encode_time_behavior_property(property: &TimeBehaviorProperty) -> Result<(u16, Vec<u8>)> {
    let integer = |value: i32| {
        let mut data = vec![1];
        data.extend(value.to_le_bytes());
        data
    };
    let float = |value: f32| {
        let mut data = vec![2];
        data.extend(value.to_le_bytes());
        data
    };
    let string = |value: &str| encode_time_variant_string(value);
    Ok(match property {
        TimeBehaviorProperty::UnknownPropertyList(value) => (0x01, string(value)),
        TimeBehaviorProperty::RuntimeContext(value) => {
            if !is_valid_runtime_context(value) {
                return Err(Error::InvalidFormat(
                    "invalid time runtime context".to_string(),
                ));
            }
            (0x02, string(value))
        },
        TimeBehaviorProperty::MotionPathEditRelative(value) => (0x03, vec![0, u8::from(*value)]),
        TimeBehaviorProperty::ColorModel(value) => (
            0x04,
            integer(match value {
                TimeColorModel::Rgb => 0,
                TimeColorModel::Hsl => 1,
                TimeColorModel::Scheme => 2,
            }),
        ),
        TimeBehaviorProperty::ColorDirection(value) => (
            0x05,
            integer(match value {
                TimeColorDirection::Clockwise => 0,
                TimeColorDirection::CounterClockwise => 1,
            }),
        ),
        TimeBehaviorProperty::Override => (0x06, integer(1)),
        TimeBehaviorProperty::PathEditRotationAngle(value) => (0x07, float(*value)),
        TimeBehaviorProperty::PathEditRotationX(value) => (0x08, float(*value)),
        TimeBehaviorProperty::PathEditRotationY(value) => (0x09, float(*value)),
        TimeBehaviorProperty::PointsTypes(value) => {
            if !is_valid_time_points_types(value) {
                return Err(Error::InvalidFormat(
                    "invalid time path point types".to_string(),
                ));
            }
            (0x0A, string(value))
        },
    })
}

fn write_time_string_list(names: &[String]) -> Result<Vec<u8>> {
    let mut children = Vec::new();
    for name in names {
        let data = encode_time_variant_string(name);
        let length = u32::try_from(data.len())
            .map_err(|_| Error::InvalidFormat("time attribute name exceeds 4 GiB".to_string()))?;
        children.extend(create_record_header(RecordType::TimeVariant, 0, 0, length));
        children.extend(data);
    }
    wrap_record(RecordType::TimeVariantList, 0x0F, 1, children)
}

pub(super) fn encode_time_variant_string(value: &str) -> Vec<u8> {
    let mut data = Vec::with_capacity(1 + value.len().saturating_mul(2));
    data.push(3);
    data.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
    data
}

/// Serialize a `ClientVisualElementContainer` animation target.
pub fn write_time_visual_element(target: &TimeVisualElement) -> Result<Vec<u8>> {
    let atom = match target {
        TimeVisualElement::Page => {
            let mut atom = create_record_header(RecordType::VisualPageAtom, 0, 0, 4);
            atom.extend(TimeVisualElementKind::Page.as_u32().to_le_bytes());
            atom
        },
        TimeVisualElement::Sound { kind, sound_id_ref } => {
            if *kind == TimeVisualElementKind::Page {
                return Err(Error::InvalidFormat(
                    "sound target cannot use the page element type".to_string(),
                ));
            }
            write_visual_shape_atom(*kind, 2, *sound_id_ref, u32::MAX, u32::MAX)
        },
        TimeVisualElement::Shape {
            kind,
            shape_id_ref,
            data1,
            data2,
        } => {
            if matches!(
                kind,
                TimeVisualElementKind::Page | TimeVisualElementKind::ChartElement
            ) {
                return Err(Error::InvalidFormat(
                    "general shape target has an invalid element type".to_string(),
                ));
            }
            write_visual_shape_atom(
                *kind,
                1,
                *shape_id_ref,
                u32::from_le_bytes(data1.to_le_bytes()),
                u32::from_le_bytes(data2.to_le_bytes()),
            )
        },
        TimeVisualElement::Chart {
            shape_id_ref,
            build_type,
            element_index,
        } => {
            if *element_index < -1 {
                return Err(Error::InvalidFormat(
                    "chart target element index must be at least -1".to_string(),
                ));
            }
            write_visual_shape_atom(
                TimeVisualElementKind::ChartElement,
                1,
                *shape_id_ref,
                build_type.as_u32(),
                u32::from_le_bytes(element_index.to_le_bytes()),
            )
        },
    };
    wrap_record(RecordType::TimeClientVisualElement, 0x0F, 0, atom)
}

fn write_visual_shape_atom(
    kind: TimeVisualElementKind,
    reference_type: u32,
    reference_id: u32,
    data1: u32,
    data2: u32,
) -> Vec<u8> {
    let mut atom = create_record_header(RecordType::VisualShapeAtom, 0, 0, 20);
    atom.extend(kind.as_u32().to_le_bytes());
    atom.extend(reference_type.to_le_bytes());
    atom.extend(reference_id.to_le_bytes());
    atom.extend(data1.to_le_bytes());
    atom.extend(data2.to_le_bytes());
    atom
}
