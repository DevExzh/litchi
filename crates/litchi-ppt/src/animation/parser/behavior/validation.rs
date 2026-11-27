//! Semantic validation for the typed animation behavior models.

use super::super::timeline::validate_basic_behavior_properties;
use crate::animation::types::{
    TimeAnimateBehavior, TimeAnimateCalculationMode, TimeAnimateValueType, TimeBehavior,
    TimeBehaviorProperty, TimeColorBehaviorAtom, TimeEffectBehavior, TimeMotionBehavior,
    is_valid_animation_attribute_name, is_valid_motion_path, is_valid_runtime_context,
    is_valid_time_animate_value, time_animation_attribute_value_type,
};
use crate::package::{Error, Result};

pub(super) fn validate_animate_behavior(animate: &TimeAnimateBehavior) -> Result<()> {
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

pub(super) fn validate_color_behavior(
    atom: &TimeColorBehaviorAtom,
    behavior: &TimeBehavior,
) -> Result<()> {
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

pub(super) fn validate_effect_behavior(effect: &TimeEffectBehavior) -> Result<()> {
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

pub(super) fn validate_motion_behavior(motion: &TimeMotionBehavior) -> Result<()> {
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
