//! Build and transition validation, conversion, and wire patching.

use super::*;
pub(super) fn validate_build_settings(settings: &KeynoteBuildSettings) -> Result<()> {
    for (label, value) in [
        ("delivery", settings.delivery.as_str()),
        ("animation type", settings.animation_type.as_str()),
        ("effect", settings.effect.as_str()),
    ] {
        if value.is_empty() || value.contains('\0') {
            return Err(Error::ParseError(format!(
                "Keynote build {label} must be non-empty and cannot contain NUL"
            )));
        }
    }
    if !settings.duration.is_finite() || settings.duration <= 0.0 {
        return Err(Error::ParseError(
            "Keynote build duration must be finite and greater than zero".to_owned(),
        ));
    }
    if !settings.delay.is_finite() || settings.delay < 0.0 {
        return Err(Error::ParseError(
            "Keynote build delay must be finite and non-negative".to_owned(),
        ));
    }
    if matches!(
        settings.start,
        BuildStart::OnClick | BuildStart::WithPrevious
    ) && settings.delay != 0.0
    {
        return Err(Error::ParseError(
            "Keynote On Click and With Previous builds cannot have a delay".to_owned(),
        ));
    }
    for (label, value) in [
        ("scale amount", settings.custom_parameters.scale_amount),
        (
            "travel distance",
            settings.custom_parameters.travel_distance,
        ),
    ] {
        if value.is_some_and(|value| !value.is_finite()) {
            return Err(Error::ParseError(format!(
                "Keynote custom build {label} must be finite"
            )));
        }
    }
    if is_typed_build_effect(&settings.effect) && !settings.custom_parameters.is_empty() {
        return Err(Error::ParseError(
            "Keynote typed effects cannot contain unrelated raw custom parameters".to_owned(),
        ));
    }
    if let Some(curve) = &settings.timing_curve {
        validate_timing_curve(curve)?;
        if typed_action_acceleration(settings) != Some(BuildAcceleration::Custom) {
            return Err(Error::ParseError(
                "Keynote timing curves require custom action acceleration".to_owned(),
            ));
        }
    }
    let typed_action_count = usize::from(settings.rotation.is_some())
        + usize::from(settings.scale.is_some())
        + usize::from(settings.opacity.is_some())
        + usize::from(settings.move_action.is_some())
        + usize::from(settings.emphasis.is_some());
    let typed_parameter_count = typed_action_count
        + usize::from(settings.keyboard.is_some())
        + usize::from(settings.object_effect.is_some());
    if typed_parameter_count > 1 {
        return Err(Error::ParseError(
            "Keynote build can contain only one typed effect parameter set".to_owned(),
        ));
    }
    if is_typed_action_effect(&settings.effect) && settings.animation_type != "Action" {
        return Err(Error::ParseError(
            "Keynote typed actions require animation type Action".to_owned(),
        ));
    }
    match settings.effect.as_str() {
        ROTATE_ACTION_EFFECT => {
            let rotation = settings.rotation.as_ref().ok_or_else(|| {
                Error::ParseError("Keynote Rotate is missing its typed parameters".to_owned())
            })?;
            if typed_parameter_count != 1 {
                return Err(Error::ParseError(
                    "Keynote Rotate has mismatched typed action parameters".to_owned(),
                ));
            }
            if !rotation.total_degrees.is_finite() || rotation.total_degrees <= 0.0 {
                return Err(Error::ParseError(
                    "Keynote Rotate degrees must be finite and greater than zero".to_owned(),
                ));
            }
        },
        SCALE_ACTION_EFFECT => {
            let scale = settings.scale.as_ref().ok_or_else(|| {
                Error::ParseError("Keynote Scale is missing its typed parameters".to_owned())
            })?;
            if typed_parameter_count != 1 {
                return Err(Error::ParseError(
                    "Keynote Scale has mismatched typed action parameters".to_owned(),
                ));
            }
            if !scale.scale_factor.is_finite() || scale.scale_factor <= 0.0 {
                return Err(Error::ParseError(
                    "Keynote Scale factor must be finite and greater than zero".to_owned(),
                ));
            }
        },
        OPACITY_ACTION_EFFECT => {
            let opacity = settings.opacity.as_ref().ok_or_else(|| {
                Error::ParseError("Keynote Opacity is missing its typed parameters".to_owned())
            })?;
            if typed_parameter_count != 1 {
                return Err(Error::ParseError(
                    "Keynote Opacity has mismatched typed action parameters".to_owned(),
                ));
            }
            if !opacity.opacity_percent.is_finite()
                || !(0.0..=100.0).contains(&opacity.opacity_percent)
            {
                return Err(Error::ParseError(
                    "Keynote Opacity percent must be finite and between zero and 100".to_owned(),
                ));
            }
        },
        MOVE_ACTION_EFFECT => {
            let move_action = settings.move_action.as_ref().ok_or_else(|| {
                Error::ParseError("Keynote Move is missing its typed parameters".to_owned())
            })?;
            if typed_parameter_count != 1 {
                return Err(Error::ParseError(
                    "Keynote Move has mismatched typed action parameters".to_owned(),
                ));
            }
            validate_motion_path(&move_action.path)?;
        },
        BLINK_ACTION_EFFECT | BOUNCE_ACTION_EFFECT | FLIP_ACTION_EFFECT | JIGGLE_ACTION_EFFECT
        | POP_ACTION_EFFECT | PULSE_ACTION_EFFECT => {
            let emphasis = settings.emphasis.ok_or_else(|| {
                Error::ParseError(
                    "Keynote emphasis action is missing its typed parameters".to_owned(),
                )
            })?;
            if typed_parameter_count != 1 || emphasis_effect(emphasis) != settings.effect {
                return Err(Error::ParseError(
                    "Keynote emphasis action has mismatched typed parameters".to_owned(),
                ));
            }
            match emphasis {
                KeynoteEmphasisAction::Blink { repeat_count, .. }
                | KeynoteEmphasisAction::Bounce { repeat_count, .. }
                | KeynoteEmphasisAction::Flip { repeat_count, .. }
                | KeynoteEmphasisAction::Pulse { repeat_count, .. }
                    if repeat_count == 0 =>
                {
                    return Err(Error::ParseError(
                        "Keynote emphasis repeat count must be greater than zero".to_owned(),
                    ));
                },
                KeynoteEmphasisAction::Pop { scale_percent }
                | KeynoteEmphasisAction::Pulse { scale_percent, .. }
                    if !scale_percent.is_finite() || scale_percent <= 0.0 =>
                {
                    return Err(Error::ParseError(
                        "Keynote emphasis scale percent must be finite and greater than zero"
                            .to_owned(),
                    ));
                },
                _ => {},
            }
            let expected_direction = match emphasis {
                KeynoteEmphasisAction::Flip { direction, .. } => {
                    Some(native_flip_direction(direction))
                },
                _ => None,
            };
            if settings.direction != expected_direction {
                return Err(Error::ParseError(
                    "Keynote emphasis direction does not match its typed parameters".to_owned(),
                ));
            }
        },
        KEYBOARD_BUILD_EFFECT => {
            let keyboard = settings.keyboard.ok_or_else(|| {
                Error::ParseError(
                    "Keynote Keyboard build is missing its typed parameters".to_owned(),
                )
            })?;
            if typed_parameter_count != 1
                || !matches!(settings.animation_type.as_str(), "In" | "Out")
            {
                return Err(Error::ParseError(
                    "Keynote Keyboard requires a Build In or Build Out parameter set".to_owned(),
                ));
            }
            if settings.direction != Some(native_keyboard_direction(keyboard.direction)) {
                return Err(Error::ParseError(
                    "Keynote Keyboard direction does not match its typed parameters".to_owned(),
                ));
            }
        },
        DISSOLVE_BUILD_EFFECT
        | SHIMMER_BUILD_EFFECT
        | SKID_BUILD_EFFECT
        | SWOOSH_BUILD_EFFECT
        | TRACE_BUILD_EFFECT => {
            let object_effect = settings.object_effect.ok_or_else(|| {
                Error::ParseError("Keynote object build is missing its typed parameters".to_owned())
            })?;
            if typed_parameter_count != 1
                || !matches!(settings.animation_type.as_str(), "In" | "Out")
                || object_build_effect_identifier(object_effect) != settings.effect
            {
                return Err(Error::ParseError(
                    "Keynote object effect requires its matching Build In or Build Out parameters"
                        .to_owned(),
                ));
            }
            let expected_direction = native_object_build_direction(object_effect);
            let omitted_native_default = settings.direction.is_none()
                && matches!(
                    object_effect,
                    KeynoteObjectBuildEffect::Skid {
                        direction: KeynoteHorizontalBuildDirection::LeftToRight,
                    } | KeynoteObjectBuildEffect::Trace {
                        direction: KeynoteHorizontalBuildDirection::LeftToRight,
                    }
                );
            if settings.direction != expected_direction && !omitted_native_default {
                return Err(Error::ParseError(
                    "Keynote object-build direction does not match its typed parameters".to_owned(),
                ));
            }
        },
        _ if typed_parameter_count != 0 => {
            return Err(Error::ParseError(
                "Keynote typed parameters require their matching native effect".to_owned(),
            ));
        },
        _ => {},
    }
    Ok(())
}

pub(super) fn is_typed_action_effect(effect: &str) -> bool {
    matches!(
        effect,
        ROTATE_ACTION_EFFECT | SCALE_ACTION_EFFECT | OPACITY_ACTION_EFFECT | MOVE_ACTION_EFFECT
    ) || is_emphasis_action_effect(effect)
}

pub(super) fn is_typed_build_effect(effect: &str) -> bool {
    is_typed_action_effect(effect)
        || effect == KEYBOARD_BUILD_EFFECT
        || is_object_build_effect(effect)
}

pub(super) fn is_object_build_effect(effect: &str) -> bool {
    matches!(
        effect,
        DISSOLVE_BUILD_EFFECT
            | SHIMMER_BUILD_EFFECT
            | SKID_BUILD_EFFECT
            | SWOOSH_BUILD_EFFECT
            | TRACE_BUILD_EFFECT
    )
}

pub(super) fn is_emphasis_action_effect(effect: &str) -> bool {
    matches!(
        effect,
        BLINK_ACTION_EFFECT
            | BOUNCE_ACTION_EFFECT
            | FLIP_ACTION_EFFECT
            | JIGGLE_ACTION_EFFECT
            | POP_ACTION_EFFECT
            | PULSE_ACTION_EFFECT
    )
}

pub(super) fn emphasis_effect(action: KeynoteEmphasisAction) -> &'static str {
    match action {
        KeynoteEmphasisAction::Blink { .. } => BLINK_ACTION_EFFECT,
        KeynoteEmphasisAction::Bounce { .. } => BOUNCE_ACTION_EFFECT,
        KeynoteEmphasisAction::Flip { .. } => FLIP_ACTION_EFFECT,
        KeynoteEmphasisAction::Jiggle { .. } => JIGGLE_ACTION_EFFECT,
        KeynoteEmphasisAction::Pop { .. } => POP_ACTION_EFFECT,
        KeynoteEmphasisAction::Pulse { .. } => PULSE_ACTION_EFFECT,
    }
}

pub(super) fn emphasis_decay(action: Option<KeynoteEmphasisAction>) -> Option<bool> {
    match action? {
        KeynoteEmphasisAction::Blink { fade, .. } => Some(fade),
        KeynoteEmphasisAction::Bounce { decay, .. } => Some(decay),
        _ => None,
    }
}

pub(super) fn emphasis_repeat_count(action: Option<KeynoteEmphasisAction>) -> Option<u32> {
    match action? {
        KeynoteEmphasisAction::Blink { repeat_count, .. }
        | KeynoteEmphasisAction::Bounce { repeat_count, .. }
        | KeynoteEmphasisAction::Flip { repeat_count, .. }
        | KeynoteEmphasisAction::Pulse { repeat_count, .. } => Some(repeat_count),
        _ => None,
    }
}

pub(super) fn emphasis_scale(action: Option<KeynoteEmphasisAction>) -> Option<f64> {
    match action? {
        KeynoteEmphasisAction::Pop { scale_percent }
        | KeynoteEmphasisAction::Pulse { scale_percent, .. } => Some(scale_percent),
        _ => None,
    }
}

pub(super) fn emphasis_jiggle_intensity(action: Option<KeynoteEmphasisAction>) -> Option<i32> {
    match action? {
        KeynoteEmphasisAction::Jiggle { intensity } => Some(native_jiggle_intensity(intensity)),
        _ => None,
    }
}

pub(super) fn typed_action_acceleration(
    settings: &KeynoteBuildSettings,
) -> Option<BuildAcceleration> {
    settings
        .rotation
        .as_ref()
        .map(|action| action.acceleration)
        .or_else(|| settings.scale.as_ref().map(|action| action.acceleration))
        .or_else(|| settings.opacity.as_ref().map(|action| action.acceleration))
        .or_else(|| {
            settings
                .move_action
                .as_ref()
                .map(|action| action.acceleration)
        })
}

pub(super) fn validate_motion_path(path: &KeynoteMotionPath) -> Result<()> {
    if !path.natural_width.is_finite()
        || !path.natural_height.is_finite()
        || path.natural_width < 0.0
        || path.natural_height < 0.0
        || (path.natural_width == 0.0 && path.natural_height == 0.0)
    {
        return Err(Error::ParseError(
            "Keynote Move natural size must be finite, non-negative, and non-zero".to_owned(),
        ));
    }
    if path.subpaths.is_empty() {
        return Err(Error::ParseError(
            "Keynote Move path must contain at least one subpath".to_owned(),
        ));
    }
    for (subpath_index, subpath) in path.subpaths.iter().enumerate() {
        if subpath.nodes.len() < 2 {
            return Err(Error::ParseError(format!(
                "Keynote Move subpath {subpath_index} must contain at least two nodes"
            )));
        }
        for (node_index, node) in subpath.nodes.iter().enumerate() {
            for (label, point) in [
                ("incoming control", node.in_control_point),
                ("node", node.point),
                ("outgoing control", node.out_control_point),
            ] {
                if !point.x.is_finite() || !point.y.is_finite() {
                    return Err(Error::ParseError(format!(
                        "Keynote Move subpath {subpath_index} node {node_index} {label} point must be finite"
                    )));
                }
            }
        }
    }
    let start = path.subpaths[0].nodes[0].point;
    if start.x != 0.0 || start.y != 0.0 {
        return Err(Error::ParseError(
            "Keynote Move path must start at the object's current position (0, 0)".to_owned(),
        ));
    }
    Ok(())
}

pub(super) fn validate_timing_curve(curve: &KeynoteBuildTimingCurve) -> Result<()> {
    let path = &curve.path;
    validate_motion_path(path).map_err(|error| match error {
        Error::ParseError(message) => {
            Error::ParseError(message.replace("Keynote Move", "Keynote timing curve"))
        },
        error => error,
    })?;
    if path.horizontal_flip || path.vertical_flip {
        return Err(Error::ParseError(
            "Keynote timing curves cannot be flipped".to_owned(),
        ));
    }
    if path.subpaths.len() != 1 || path.subpaths[0].closed {
        return Err(Error::ParseError(
            "Keynote timing curves require one open path".to_owned(),
        ));
    }
    let end = path.subpaths[0]
        .nodes
        .last()
        .ok_or_else(|| {
            Error::InvalidFormat(
                "Keynote timing curve validation accepted an empty path".to_owned(),
            )
        })?
        .point;
    if end.x != 1.0 || end.y != 1.0 {
        return Err(Error::ParseError(
            "Keynote timing curves must end at normalized point (1, 1)".to_owned(),
        ));
    }
    Ok(())
}

pub(super) fn native_motion_node_type(node_type: KeynoteMotionPathNodeType) -> i32 {
    use tsd::editable_bezier_path_source_archive::NodeType;
    match node_type {
        KeynoteMotionPathNodeType::Sharp => NodeType::Sharp as i32,
        KeynoteMotionPathNodeType::Bezier => NodeType::Bezier as i32,
        KeynoteMotionPathNodeType::Smooth => NodeType::Smooth as i32,
    }
}

pub(super) fn motion_node_type_from_native(value: i32) -> Option<KeynoteMotionPathNodeType> {
    use tsd::editable_bezier_path_source_archive::NodeType;
    match NodeType::try_from(value).ok()? {
        NodeType::Sharp => Some(KeynoteMotionPathNodeType::Sharp),
        NodeType::Bezier => Some(KeynoteMotionPathNodeType::Bezier),
        NodeType::Smooth => Some(KeynoteMotionPathNodeType::Smooth),
    }
}

pub(super) fn motion_point_from_native(point: &tsp::Point) -> KeynoteMotionPathPoint {
    KeynoteMotionPathPoint::new(point.x, point.y)
}

pub(super) fn native_motion_point(point: KeynoteMotionPathPoint) -> tsp::Point {
    tsp::Point {
        x: point.x,
        y: point.y,
    }
}

pub(super) fn motion_path_from_native(
    source: &tsd::PathSourceArchive,
) -> Option<KeynoteMotionPath> {
    if source.point_path_source.is_some()
        || source.scalar_path_source.is_some()
        || source.bezier_path_source.is_some()
        || source.callout_path_source.is_some()
        || source.connection_line_path_source.is_some()
    {
        return None;
    }
    let editable = source.editable_bezier_path_source.as_ref()?;
    let natural_size = editable.natural_size.as_ref()?;
    let subpaths = editable
        .subpaths
        .iter()
        .map(|subpath| {
            let nodes = subpath
                .nodes
                .iter()
                .map(|node| {
                    Some(KeynoteMotionPathNode {
                        in_control_point: motion_point_from_native(&node.in_control_point),
                        point: motion_point_from_native(&node.node_point),
                        out_control_point: motion_point_from_native(&node.out_control_point),
                        node_type: motion_node_type_from_native(node.r#type)?,
                    })
                })
                .collect::<Option<Vec<_>>>()?;
            Some(KeynoteMotionSubpath {
                nodes,
                closed: subpath.closed,
            })
        })
        .collect::<Option<Vec<_>>>()?;
    Some(KeynoteMotionPath {
        subpaths,
        natural_width: natural_size.width,
        natural_height: natural_size.height,
        horizontal_flip: source.horizontal_flip.unwrap_or(false),
        vertical_flip: source.vertical_flip.unwrap_or(false),
    })
}

pub(super) fn native_motion_path(path: &KeynoteMotionPath) -> tsd::PathSourceArchive {
    use tsd::editable_bezier_path_source_archive::{Node, Subpath};
    tsd::PathSourceArchive {
        horizontal_flip: Some(path.horizontal_flip),
        vertical_flip: Some(path.vertical_flip),
        editable_bezier_path_source: Some(tsd::EditableBezierPathSourceArchive {
            subpaths: path
                .subpaths
                .iter()
                .map(|subpath| Subpath {
                    nodes: subpath
                        .nodes
                        .iter()
                        .map(|node| Node {
                            in_control_point: native_motion_point(node.in_control_point),
                            node_point: native_motion_point(node.point),
                            out_control_point: native_motion_point(node.out_control_point),
                            r#type: native_motion_node_type(node.node_type),
                        })
                        .collect(),
                    closed: subpath.closed,
                })
                .collect(),
            natural_size: Some(tsp::Size {
                width: path.natural_width,
                height: path.natural_height,
            }),
        }),
        ..Default::default()
    }
}

pub(super) fn timing_curve_from_native(
    source: &tsd::PathSourceArchive,
) -> Option<KeynoteBuildTimingCurve> {
    let curve = KeynoteBuildTimingCurve::from_path(motion_path_from_native(source)?);
    validate_timing_curve(&curve).ok()?;
    Some(curve)
}

pub(super) fn native_rotation_direction(direction: KeynoteRotationDirection) -> i32 {
    use kn::build_attributes_archive::BuildAttributesRotationDirection;
    match direction {
        KeynoteRotationDirection::Clockwise => BuildAttributesRotationDirection::KClockwise as i32,
        KeynoteRotationDirection::Counterclockwise => {
            BuildAttributesRotationDirection::KCounterclockwise as i32
        },
    }
}

pub(super) fn native_flip_direction(direction: KeynoteFlipDirection) -> u32 {
    match direction {
        KeynoteFlipDirection::LeftToRight => 11,
        KeynoteFlipDirection::RightToLeft => 12,
    }
}

pub(super) fn flip_direction_from_native(value: u32) -> Option<KeynoteFlipDirection> {
    match value {
        11 => Some(KeynoteFlipDirection::LeftToRight),
        12 => Some(KeynoteFlipDirection::RightToLeft),
        _ => None,
    }
}

pub(super) fn native_keyboard_direction(direction: KeynoteKeyboardDirection) -> u32 {
    match direction {
        KeynoteKeyboardDirection::Forward => 111,
        KeynoteKeyboardDirection::Backward => 112,
    }
}

pub(super) fn keyboard_direction_from_native(value: u32) -> Option<KeynoteKeyboardDirection> {
    match value {
        111 => Some(KeynoteKeyboardDirection::Forward),
        112 => Some(KeynoteKeyboardDirection::Backward),
        _ => None,
    }
}

pub(super) fn object_build_effect_identifier(effect: KeynoteObjectBuildEffect) -> &'static str {
    match effect {
        KeynoteObjectBuildEffect::Dissolve => DISSOLVE_BUILD_EFFECT,
        KeynoteObjectBuildEffect::Shimmer => SHIMMER_BUILD_EFFECT,
        KeynoteObjectBuildEffect::Skid { .. } => SKID_BUILD_EFFECT,
        KeynoteObjectBuildEffect::Swoosh { .. } => SWOOSH_BUILD_EFFECT,
        KeynoteObjectBuildEffect::Trace { .. } => TRACE_BUILD_EFFECT,
    }
}

pub(super) fn native_object_build_direction(effect: KeynoteObjectBuildEffect) -> Option<u32> {
    match effect {
        KeynoteObjectBuildEffect::Dissolve
        | KeynoteObjectBuildEffect::Shimmer
        | KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::Center,
        } => None,
        KeynoteObjectBuildEffect::Skid {
            direction: KeynoteHorizontalBuildDirection::LeftToRight,
        }
        | KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::FromLeft,
        }
        | KeynoteObjectBuildEffect::Trace {
            direction: KeynoteHorizontalBuildDirection::LeftToRight,
        } => Some(11),
        KeynoteObjectBuildEffect::Skid {
            direction: KeynoteHorizontalBuildDirection::RightToLeft,
        }
        | KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::FromRight,
        }
        | KeynoteObjectBuildEffect::Trace {
            direction: KeynoteHorizontalBuildDirection::RightToLeft,
        } => Some(12),
    }
}

pub(super) fn object_build_effect_from_native(
    effect: &str,
    direction: Option<u32>,
) -> Option<KeynoteObjectBuildEffect> {
    match (effect, direction) {
        (DISSOLVE_BUILD_EFFECT, None) => Some(KeynoteObjectBuildEffect::Dissolve),
        (SHIMMER_BUILD_EFFECT, None) => Some(KeynoteObjectBuildEffect::Shimmer),
        (SKID_BUILD_EFFECT, None | Some(11)) => Some(KeynoteObjectBuildEffect::Skid {
            direction: KeynoteHorizontalBuildDirection::LeftToRight,
        }),
        (SKID_BUILD_EFFECT, Some(12)) => Some(KeynoteObjectBuildEffect::Skid {
            direction: KeynoteHorizontalBuildDirection::RightToLeft,
        }),
        (SWOOSH_BUILD_EFFECT, None) => Some(KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::Center,
        }),
        (SWOOSH_BUILD_EFFECT, Some(11)) => Some(KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::FromLeft,
        }),
        (SWOOSH_BUILD_EFFECT, Some(12)) => Some(KeynoteObjectBuildEffect::Swoosh {
            direction: KeynoteSwooshDirection::FromRight,
        }),
        (TRACE_BUILD_EFFECT, None | Some(11)) => Some(KeynoteObjectBuildEffect::Trace {
            direction: KeynoteHorizontalBuildDirection::LeftToRight,
        }),
        (TRACE_BUILD_EFFECT, Some(12)) => Some(KeynoteObjectBuildEffect::Trace {
            direction: KeynoteHorizontalBuildDirection::RightToLeft,
        }),
        _ => None,
    }
}

pub(super) fn native_jiggle_intensity(intensity: KeynoteJiggleIntensity) -> i32 {
    use kn::build_attributes_archive::ActionBuildAttributesJiggleIntensity;
    match intensity {
        KeynoteJiggleIntensity::Small => {
            ActionBuildAttributesJiggleIntensity::KJiggleIntensitySmall as i32
        },
        KeynoteJiggleIntensity::Medium => {
            ActionBuildAttributesJiggleIntensity::KJiggleIntensityMedium as i32
        },
        KeynoteJiggleIntensity::Large => {
            ActionBuildAttributesJiggleIntensity::KJiggleIntensityLarge as i32
        },
    }
}

pub(super) fn jiggle_intensity_from_native(value: i32) -> Option<KeynoteJiggleIntensity> {
    use kn::build_attributes_archive::ActionBuildAttributesJiggleIntensity;
    match ActionBuildAttributesJiggleIntensity::try_from(value).ok()? {
        ActionBuildAttributesJiggleIntensity::KJiggleIntensitySmall => {
            Some(KeynoteJiggleIntensity::Small)
        },
        ActionBuildAttributesJiggleIntensity::KJiggleIntensityMedium => {
            Some(KeynoteJiggleIntensity::Medium)
        },
        ActionBuildAttributesJiggleIntensity::KJiggleIntensityLarge => {
            Some(KeynoteJiggleIntensity::Large)
        },
    }
}

pub(super) fn rotation_direction_from_native(value: i32) -> Option<KeynoteRotationDirection> {
    use kn::build_attributes_archive::BuildAttributesRotationDirection;
    match BuildAttributesRotationDirection::try_from(value).ok()? {
        BuildAttributesRotationDirection::KClockwise => Some(KeynoteRotationDirection::Clockwise),
        BuildAttributesRotationDirection::KCounterclockwise => {
            Some(KeynoteRotationDirection::Counterclockwise)
        },
    }
}

pub(super) const fn native_build_acceleration(acceleration: BuildAcceleration) -> i32 {
    acceleration.native_value()
}

pub(super) const fn build_acceleration_from_native(value: i32) -> BuildAcceleration {
    BuildAcceleration::from_native(value)
}

#[allow(deprecated)]
pub(super) fn validate_typed_action_wire(original: &[u8], build: &kn::BuildArchive) -> Result<()> {
    let effect = build
        .attributes
        .animation_attributes
        .as_ref()
        .and_then(|attributes| attributes.effect.as_deref())
        .or(build.attributes.database_effect.as_deref())
        .unwrap_or_default();
    let has_emphasis_parameters = build.attributes.custom_action_decay.is_some()
        || build.attributes.custom_action_repeat_count.is_some()
        || build.attributes.custom_action_scale.is_some()
        || build.attributes.custom_action_jiggle_intensity.is_some();
    let has_unrelated_custom_parameters = build.attributes.custom_bounce.is_some()
        || build.attributes.custom_motion_blur.is_some()
        || build.attributes.custom_include_endpoints.is_some()
        || build.attributes.custom_shine.is_some()
        || build.attributes.custom_scale_amount.is_some()
        || build.attributes.custom_travel_distance.is_some()
        || build.attributes.custom_cursor.is_some();
    let has_unexpected_emphasis_parameters = match effect {
        BLINK_ACTION_EFFECT | BOUNCE_ACTION_EFFECT => {
            build.attributes.custom_action_scale.is_some()
                || build.attributes.custom_action_jiggle_intensity.is_some()
        },
        FLIP_ACTION_EFFECT => {
            build.attributes.custom_action_decay.is_some()
                || build.attributes.custom_action_scale.is_some()
                || build.attributes.custom_action_jiggle_intensity.is_some()
        },
        JIGGLE_ACTION_EFFECT => {
            build.attributes.custom_action_decay.is_some()
                || build.attributes.custom_action_repeat_count.is_some()
                || build.attributes.custom_action_scale.is_some()
        },
        POP_ACTION_EFFECT => {
            build.attributes.custom_action_decay.is_some()
                || build.attributes.custom_action_repeat_count.is_some()
                || build.attributes.custom_action_jiggle_intensity.is_some()
        },
        PULSE_ACTION_EFFECT => {
            build.attributes.custom_action_decay.is_some()
                || build.attributes.custom_action_jiggle_intensity.is_some()
        },
        _ => false,
    };
    match effect {
        ROTATE_ACTION_EFFECT
            if build.attributes.action_scale_size.is_some()
                || build.attributes.action_color_alpha.is_some()
                || build.attributes.action_motion_path_source.is_some()
                || build.attributes.custom_align_to_path.is_some()
                || has_emphasis_parameters
                || has_unrelated_custom_parameters =>
        {
            return Err(Error::InvalidFormat(
                "Keynote Rotate contains parameters for another typed action".to_owned(),
            ));
        },
        SCALE_ACTION_EFFECT
            if build.attributes.action_rotation_angle.is_some()
                || build.attributes.action_rotation_direction.is_some()
                || build.attributes.action_color_alpha.is_some()
                || build.attributes.action_motion_path_source.is_some()
                || build.attributes.custom_align_to_path.is_some()
                || has_emphasis_parameters
                || has_unrelated_custom_parameters =>
        {
            return Err(Error::InvalidFormat(
                "Keynote Scale contains parameters for another typed action".to_owned(),
            ));
        },
        OPACITY_ACTION_EFFECT
            if build.attributes.action_rotation_angle.is_some()
                || build.attributes.action_rotation_direction.is_some()
                || build.attributes.action_scale_size.is_some()
                || build.attributes.action_motion_path_source.is_some()
                || build.attributes.custom_align_to_path.is_some()
                || has_emphasis_parameters
                || has_unrelated_custom_parameters =>
        {
            return Err(Error::InvalidFormat(
                "Keynote Opacity contains parameters for another typed action".to_owned(),
            ));
        },
        MOVE_ACTION_EFFECT
            if build.attributes.action_rotation_angle.is_some()
                || build.attributes.action_rotation_direction.is_some()
                || build.attributes.action_scale_size.is_some()
                || build.attributes.action_color_alpha.is_some()
                || has_emphasis_parameters
                || has_unrelated_custom_parameters =>
        {
            return Err(Error::InvalidFormat(
                "Keynote Move contains parameters for another typed action".to_owned(),
            ));
        },
        BLINK_ACTION_EFFECT | BOUNCE_ACTION_EFFECT | FLIP_ACTION_EFFECT | JIGGLE_ACTION_EFFECT
        | POP_ACTION_EFFECT | PULSE_ACTION_EFFECT
            if build.attributes.action_rotation_angle.is_some()
                || build.attributes.action_rotation_direction.is_some()
                || build.attributes.action_scale_size.is_some()
                || build.attributes.action_color_alpha.is_some()
                || build.attributes.action_acceleration.is_some()
                || build.attributes.action_motion_path_source.is_some()
                || build.attributes.custom_align_to_path.is_some()
                || has_unrelated_custom_parameters
                || has_unexpected_emphasis_parameters =>
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote {effect} contains parameters for another typed action"
            )));
        },
        _ => {},
    }
    let _ = transform_length_delimited_field(original, 4, |attributes| {
        let mut attributes = attributes.to_vec();
        if effect == ROTATE_ACTION_EFFECT {
            attributes = patch_fixed64_field(
                &attributes,
                9,
                true,
                build.attributes.action_rotation_angle.map(f64::to_bits),
            )?;
            attributes = patch_varint_field(
                &attributes,
                10,
                true,
                build
                    .attributes
                    .action_rotation_direction
                    .map(|value| value as u64),
            )?;
        } else if effect == SCALE_ACTION_EFFECT {
            attributes = patch_fixed64_field(
                &attributes,
                11,
                true,
                build.attributes.action_scale_size.map(f64::to_bits),
            )?;
        } else if effect == OPACITY_ACTION_EFFECT {
            attributes = patch_fixed64_field(
                &attributes,
                12,
                true,
                build.attributes.action_color_alpha.map(f64::to_bits),
            )?;
        } else if effect == MOVE_ACTION_EFFECT {
            let source = build
                .attributes
                .action_motion_path_source
                .as_ref()
                .ok_or_else(|| {
                    Error::InvalidFormat("Keynote Move has no motion path source".to_owned())
                })?;
            attributes = transform_length_delimited_field(&attributes, 22, |path_source| {
                validate_motion_path_source_wire(path_source, source)
            })?;
            attributes = patch_varint_field(
                &attributes,
                37,
                build.attributes.custom_align_to_path.is_some(),
                build
                    .attributes
                    .custom_align_to_path
                    .map(|value| u64::from(u8::from(value))),
            )?;
        } else if effect == BLINK_ACTION_EFFECT || effect == BOUNCE_ACTION_EFFECT {
            attributes = patch_varint_field(
                &attributes,
                23,
                true,
                build
                    .attributes
                    .custom_action_decay
                    .map(|value| u64::from(u8::from(value))),
            )?;
            attributes = patch_varint_field(
                &attributes,
                24,
                true,
                build.attributes.custom_action_repeat_count.map(u64::from),
            )?;
        } else if effect == FLIP_ACTION_EFFECT {
            attributes = patch_varint_field(
                &attributes,
                24,
                true,
                build.attributes.custom_action_repeat_count.map(u64::from),
            )?;
            attributes = transform_length_delimited_field(&attributes, 18, |animation| {
                patch_varint_field(
                    animation,
                    4,
                    true,
                    build
                        .attributes
                        .animation_attributes
                        .as_ref()
                        .and_then(|animation| animation.direction)
                        .map(u64::from),
                )
            })?;
        } else if effect == JIGGLE_ACTION_EFFECT {
            attributes = patch_varint_field(
                &attributes,
                26,
                true,
                build
                    .attributes
                    .custom_action_jiggle_intensity
                    .map(|value| value as u64),
            )?;
        } else if effect == POP_ACTION_EFFECT {
            attributes = patch_fixed64_field(
                &attributes,
                25,
                true,
                build.attributes.custom_action_scale.map(f64::to_bits),
            )?;
        } else if effect == PULSE_ACTION_EFFECT {
            attributes = patch_varint_field(
                &attributes,
                24,
                true,
                build.attributes.custom_action_repeat_count.map(u64::from),
            )?;
            attributes = patch_fixed64_field(
                &attributes,
                25,
                true,
                build.attributes.custom_action_scale.map(f64::to_bits),
            )?;
        }
        if matches!(
            effect,
            ROTATE_ACTION_EFFECT | SCALE_ACTION_EFFECT | OPACITY_ACTION_EFFECT | MOVE_ACTION_EFFECT
        ) {
            attributes = patch_varint_field(
                &attributes,
                13,
                true,
                build
                    .attributes
                    .action_acceleration
                    .map(|value| value as u64),
            )?;
        }
        Ok(attributes)
    })?;
    Ok(())
}

#[allow(deprecated)]
pub(super) fn validate_keyboard_build_wire(
    original: &[u8],
    build: &kn::BuildArchive,
) -> Result<()> {
    let attributes = &build.attributes;
    if attributes.action_rotation_angle.is_some()
        || attributes.action_rotation_direction.is_some()
        || attributes.action_scale_size.is_some()
        || attributes.action_color_alpha.is_some()
        || attributes.action_acceleration.is_some()
        || attributes.action_motion_path_source.is_some()
        || attributes.custom_bounce.is_some()
        || attributes.custom_action_decay.is_some()
        || attributes.custom_action_repeat_count.is_some()
        || attributes.custom_action_scale.is_some()
        || attributes.custom_action_jiggle_intensity.is_some()
        || attributes.custom_motion_blur.is_some()
        || attributes.custom_include_endpoints.is_some()
        || attributes.custom_shine.is_some()
        || attributes.custom_scale_amount.is_some()
        || attributes.custom_travel_distance.is_some()
        || attributes.custom_align_to_path.is_some()
    {
        return Err(Error::InvalidFormat(
            "Keynote Keyboard contains parameters for another typed effect".to_owned(),
        ));
    }
    let _ = transform_length_delimited_field(original, 4, |attributes_wire| {
        let attributes_wire = patch_varint_field(
            attributes_wire,
            36,
            true,
            attributes
                .custom_cursor
                .map(|value| u64::from(u8::from(value))),
        )?;
        transform_length_delimited_field(&attributes_wire, 18, |animation_wire| {
            patch_varint_field(
                animation_wire,
                4,
                true,
                attributes
                    .animation_attributes
                    .as_ref()
                    .and_then(|animation| animation.direction)
                    .map(u64::from),
            )
        })
    })?;
    Ok(())
}

#[allow(deprecated)]
pub(super) fn validate_object_build_wire(original: &[u8], build: &kn::BuildArchive) -> Result<()> {
    let animation = build
        .attributes
        .animation_attributes
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Keynote object build has no animation".to_owned()))?;
    let effect = animation
        .effect
        .as_deref()
        .or(build.attributes.database_effect.as_deref())
        .unwrap_or_default();
    if !matches!(animation.animation_type.as_deref(), Some("In" | "Out"))
        || object_build_effect_from_native(effect, animation.direction).is_none()
        || build.attributes.action_rotation_angle.is_some()
        || build.attributes.action_rotation_direction.is_some()
        || build.attributes.action_scale_size.is_some()
        || build.attributes.action_color_alpha.is_some()
        || build.attributes.action_acceleration.is_some()
        || build.attributes.action_motion_path_source.is_some()
        || build.attributes.custom_bounce.is_some()
        || build.attributes.custom_action_decay.is_some()
        || build.attributes.custom_action_repeat_count.is_some()
        || build.attributes.custom_action_scale.is_some()
        || build.attributes.custom_action_jiggle_intensity.is_some()
        || build.attributes.custom_motion_blur.is_some()
        || build.attributes.custom_include_endpoints.is_some()
        || build.attributes.custom_shine.is_some()
        || build.attributes.custom_scale_amount.is_some()
        || build.attributes.custom_travel_distance.is_some()
        || build.attributes.custom_cursor.is_some()
        || build.attributes.custom_align_to_path.is_some()
    {
        return Err(Error::InvalidFormat(
            "Keynote object build contains parameters for another typed effect".to_owned(),
        ));
    }
    let _ = transform_length_delimited_field(original, 4, |attributes_wire| {
        transform_length_delimited_field(attributes_wire, 18, |animation_wire| {
            let mut animation_wire = patch_length_delimited_field(
                animation_wire,
                1,
                animation.animation_type.is_some(),
                animation.animation_type.as_deref().map(str::as_bytes),
            )?;
            animation_wire = patch_length_delimited_field(
                &animation_wire,
                2,
                animation.effect.is_some(),
                animation.effect.as_deref().map(str::as_bytes),
            )?;
            for (field_number, current) in [(3, animation.duration), (5, animation.delay)] {
                animation_wire = patch_fixed64_field(
                    &animation_wire,
                    field_number,
                    current.is_some(),
                    current.map(f64::to_bits),
                )?;
            }
            animation_wire = patch_varint_field(
                &animation_wire,
                4,
                animation.direction.is_some(),
                animation.direction.map(u64::from),
            )?;
            patch_varint_field(
                &animation_wire,
                6,
                animation.is_automatic.is_some(),
                animation
                    .is_automatic
                    .map(|value| u64::from(u8::from(value))),
            )
        })
    })?;
    Ok(())
}

pub(super) fn validate_custom_build_parameters_wire(
    original: &[u8],
    build: &kn::BuildArchive,
) -> Result<()> {
    let attributes = &build.attributes;
    let _ = transform_length_delimited_field(original, 4, |wire| {
        let mut wire = patch_varint_field(
            wire,
            19,
            attributes.custom_bounce.is_some(),
            attributes
                .custom_bounce
                .map(|value| u64::from(u8::from(value))),
        )?;
        wire = patch_varint_field(
            &wire,
            29,
            attributes.custom_motion_blur.is_some(),
            attributes
                .custom_motion_blur
                .map(|value| u64::from(u8::from(value))),
        )?;
        wire = patch_varint_field(
            &wire,
            30,
            attributes.custom_include_endpoints.is_some(),
            attributes
                .custom_include_endpoints
                .map(|value| u64::from(u8::from(value))),
        )?;
        wire = patch_varint_field(
            &wire,
            33,
            attributes.custom_shine.is_some(),
            attributes
                .custom_shine
                .map(|value| u64::from(u8::from(value))),
        )?;
        wire = patch_fixed64_field(
            &wire,
            34,
            attributes.custom_scale_amount.is_some(),
            attributes.custom_scale_amount.map(f64::to_bits),
        )?;
        patch_fixed64_field(
            &wire,
            35,
            attributes.custom_travel_distance.is_some(),
            attributes.custom_travel_distance.map(f64::to_bits),
        )
    })?;
    Ok(())
}

pub(super) fn validate_motion_point_wire(original: &[u8]) -> Result<Vec<u8>> {
    let point = tsp::Point::decode(original)?;
    let data = patch_fixed32_field(original, 1, true, Some(point.x.to_bits()))?;
    patch_fixed32_field(&data, 2, true, Some(point.y.to_bits()))
}

pub(super) fn validate_motion_path_source_wire(
    original: &[u8],
    source: &tsd::PathSourceArchive,
) -> Result<Vec<u8>> {
    let mut data = patch_varint_field(
        original,
        1,
        source.horizontal_flip.is_some(),
        source
            .horizontal_flip
            .map(|value| u64::from(u8::from(value))),
    )?;
    data = patch_varint_field(
        &data,
        2,
        source.vertical_flip.is_some(),
        source.vertical_flip.map(|value| u64::from(u8::from(value))),
    )?;
    let editable = source
        .editable_bezier_path_source
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Keynote Move path is not editable".to_owned()))?;
    transform_length_delimited_field(&data, 8, |editable_wire| {
        let natural_size = editable.natural_size.as_ref().ok_or_else(|| {
            Error::InvalidFormat("Keynote Move path has no natural size".to_owned())
        })?;
        let mut editable_wire = transform_length_delimited_field(editable_wire, 2, |size_wire| {
            let size = tsp::Size::decode(size_wire)?;
            let size_wire = patch_fixed32_field(size_wire, 1, true, Some(size.width.to_bits()))?;
            patch_fixed32_field(&size_wire, 2, true, Some(size.height.to_bits()))
        })?;
        editable_wire =
            transform_length_delimited_fields_at_path(&editable_wire, &[1], |subpath_wire| {
                let subpath =
                    tsd::editable_bezier_path_source_archive::Subpath::decode(subpath_wire)?;
                let mut subpath_wire = patch_varint_field(
                    subpath_wire,
                    2,
                    true,
                    Some(u64::from(u8::from(subpath.closed))),
                )?;
                subpath_wire =
                    transform_length_delimited_fields_at_path(&subpath_wire, &[1], |node_wire| {
                        let node =
                            tsd::editable_bezier_path_source_archive::Node::decode(node_wire)?;
                        let mut node_wire = transform_length_delimited_field(
                            node_wire,
                            1,
                            validate_motion_point_wire,
                        )?;
                        node_wire = transform_length_delimited_field(
                            &node_wire,
                            2,
                            validate_motion_point_wire,
                        )?;
                        node_wire = transform_length_delimited_field(
                            &node_wire,
                            3,
                            validate_motion_point_wire,
                        )?;
                        patch_varint_field(&node_wire, 4, true, Some(node.r#type as u64))
                    })?;
                Ok(subpath_wire)
            })?;
        // Keep the decoded value live so a missing size cannot be masked by an
        // unrelated same-number field with an incompatible wire type.
        let _ = natural_size;
        Ok(editable_wire)
    })
}

pub(super) fn validate_build_start_position(start: BuildStart, event_index: usize) -> Result<()> {
    match (start, event_index) {
        (BuildStart::AfterTransition, 0) | (BuildStart::OnClick, _) => Ok(()),
        (BuildStart::AfterTransition, _) => Err(Error::ParseError(
            "Keynote After Transition is only valid for the first build event".to_owned(),
        )),
        (BuildStart::WithPrevious | BuildStart::AfterPrevious, 0) => Err(Error::ParseError(
            "Keynote With Previous and After Previous require a preceding build event".to_owned(),
        )),
        (BuildStart::WithPrevious | BuildStart::AfterPrevious, _) => Ok(()),
    }
}

pub(super) fn build_start_fields(start: BuildStart) -> (bool, bool) {
    match start {
        BuildStart::OnClick => (false, true),
        BuildStart::AfterTransition | BuildStart::AfterPrevious => (true, true),
        BuildStart::WithPrevious => (true, false),
    }
}

#[allow(deprecated)]
pub(super) fn build_settings(
    build: &kn::BuildArchive,
    chunks: &[KeynoteBuildChunkInfo],
    starts_slide_events: bool,
) -> KeynoteBuildSettings {
    let animation = build.attributes.animation_attributes.as_ref();
    let animation_type = animation
        .and_then(|attributes| attributes.animation_type.clone())
        .or_else(|| build.attributes.database_animation_type.clone())
        .unwrap_or_default();
    let effect = animation
        .and_then(|attributes| attributes.effect.clone())
        .or_else(|| build.attributes.database_effect.clone())
        .unwrap_or_default();
    let direction = animation
        .and_then(|attributes| attributes.direction)
        .or(build.attributes.database_direction);
    let rotation = (effect == ROTATE_ACTION_EFFECT).then(|| {
        Some(KeynoteRotationAction {
            total_degrees: build.attributes.action_rotation_angle?,
            direction: rotation_direction_from_native(build.attributes.action_rotation_direction?)?,
            acceleration: build_acceleration_from_native(build.attributes.action_acceleration?),
        })
    });
    let scale = (effect == SCALE_ACTION_EFFECT).then(|| {
        Some(KeynoteScaleAction {
            scale_factor: build.attributes.action_scale_size?,
            acceleration: build_acceleration_from_native(build.attributes.action_acceleration?),
        })
    });
    let opacity = (effect == OPACITY_ACTION_EFFECT).then(|| {
        Some(KeynoteOpacityAction {
            opacity_percent: build.attributes.action_color_alpha?,
            acceleration: build_acceleration_from_native(build.attributes.action_acceleration?),
        })
    });
    let move_action = (effect == MOVE_ACTION_EFFECT).then(|| {
        Some(KeynoteMoveAction {
            path: motion_path_from_native(build.attributes.action_motion_path_source.as_ref()?)?,
            align_to_path: build.attributes.custom_align_to_path.unwrap_or(false),
            acceleration: build_acceleration_from_native(build.attributes.action_acceleration?),
        })
    });
    let emphasis = match effect.as_str() {
        BLINK_ACTION_EFFECT => build
            .attributes
            .custom_action_repeat_count
            .zip(build.attributes.custom_action_decay)
            .map(|(repeat_count, fade)| KeynoteEmphasisAction::Blink { repeat_count, fade }),
        BOUNCE_ACTION_EFFECT => build
            .attributes
            .custom_action_repeat_count
            .zip(build.attributes.custom_action_decay)
            .map(|(repeat_count, decay)| KeynoteEmphasisAction::Bounce {
                repeat_count,
                decay,
            }),
        FLIP_ACTION_EFFECT => build
            .attributes
            .custom_action_repeat_count
            .zip(direction.and_then(flip_direction_from_native))
            .map(|(repeat_count, direction)| KeynoteEmphasisAction::Flip {
                repeat_count,
                direction,
            }),
        JIGGLE_ACTION_EFFECT => build
            .attributes
            .custom_action_jiggle_intensity
            .and_then(jiggle_intensity_from_native)
            .map(|intensity| KeynoteEmphasisAction::Jiggle { intensity }),
        POP_ACTION_EFFECT => build
            .attributes
            .custom_action_scale
            .map(|scale_percent| KeynoteEmphasisAction::Pop { scale_percent }),
        PULSE_ACTION_EFFECT => build
            .attributes
            .custom_action_repeat_count
            .zip(build.attributes.custom_action_scale)
            .map(
                |(repeat_count, scale_percent)| KeynoteEmphasisAction::Pulse {
                    repeat_count,
                    scale_percent,
                },
            ),
        _ => None,
    };
    let keyboard =
        if effect == KEYBOARD_BUILD_EFFECT && matches!(animation_type.as_str(), "In" | "Out") {
            direction
                .and_then(keyboard_direction_from_native)
                .zip(build.attributes.custom_cursor)
                .map(|(direction, show_cursor)| KeynoteKeyboardBuild {
                    direction,
                    show_cursor,
                })
        } else {
            None
        };
    let object_effect = if matches!(animation_type.as_str(), "In" | "Out") {
        object_build_effect_from_native(&effect, direction)
    } else {
        None
    };
    let timing_curve = (is_typed_action_effect(&effect)
        && build
            .attributes
            .action_acceleration
            .map(build_acceleration_from_native)
            == Some(BuildAcceleration::Custom))
    .then(|| {
        animation
            .and_then(|attributes| attributes.custom_effect_timing_curve_1.as_ref())
            .and_then(timing_curve_from_native)
    })
    .flatten();
    KeynoteBuildSettings {
        delivery: build.delivery.clone(),
        animation_type,
        effect,
        duration: animation
            .and_then(|attributes| attributes.duration)
            .or(build.attributes.database_duration)
            .or(build.duration)
            .or_else(|| chunks.first().and_then(|chunk| chunk.duration))
            .unwrap_or_default(),
        direction,
        delay: animation
            .and_then(|attributes| attributes.delay)
            .or(build.attributes.database_delay)
            .or_else(|| chunks.first().and_then(|chunk| chunk.delay))
            .unwrap_or_default(),
        start: match (
            chunks
                .first()
                .and_then(|chunk| chunk.automatic)
                .or_else(|| animation.and_then(|attributes| attributes.is_automatic))
                .unwrap_or(false),
            chunks
                .first()
                .and_then(|chunk| chunk.referent)
                .unwrap_or(true),
            starts_slide_events,
        ) {
            (false, _, _) => BuildStart::OnClick,
            (true, false, _) => BuildStart::WithPrevious,
            (true, true, true) => BuildStart::AfterTransition,
            (true, true, false) => BuildStart::AfterPrevious,
        },
        text_delivery: build.attributes.custom_text_delivery,
        delivery_option: build.attributes.custom_delivery_option,
        event_trigger: build.attributes.event_trigger,
        rotation: rotation.flatten(),
        scale: scale.flatten(),
        opacity: opacity.flatten(),
        move_action: move_action.flatten(),
        emphasis,
        keyboard,
        object_effect,
        timing_curve,
        custom_parameters: KeynoteBuildCustomParameters {
            bounce: build.attributes.custom_bounce,
            motion_blur: build.attributes.custom_motion_blur,
            include_endpoints: build.attributes.custom_include_endpoints,
            shine: build.attributes.custom_shine,
            scale_amount: build.attributes.custom_scale_amount,
            travel_distance: build.attributes.custom_travel_distance,
        },
    }
}

pub(super) fn new_build_uuid_and_seed() -> (crate::protobuf::tsp::Uuid, u32) {
    let bytes = litchi_core::id::generate_guid_bytes();
    let mut lower = [0u8; 8];
    lower.copy_from_slice(&bytes[..8]);
    let mut upper = [0u8; 8];
    upper.copy_from_slice(&bytes[8..]);
    (
        crate::protobuf::tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        },
        u32::from_le_bytes(bytes[..4].try_into().expect("four GUID bytes")),
    )
}

#[allow(deprecated)]
pub(super) fn new_build_archive(
    drawable_object_id: u64,
    settings: &KeynoteBuildSettings,
    random_number_seed: u32,
) -> kn::BuildArchive {
    kn::BuildArchive {
        drawable: Some(crate::protobuf::tsp::Reference {
            identifier: drawable_object_id,
            ..Default::default()
        }),
        delivery: settings.delivery.clone(),
        duration: Some(0.0),
        attributes: kn::BuildAttributesArchive {
            animation_attributes: Some(kn::AnimationAttributesArchive {
                animation_type: Some(settings.animation_type.clone()),
                effect: Some(settings.effect.clone()),
                duration: Some(settings.duration),
                direction: settings.direction,
                delay: Some(settings.delay),
                custom_effect_timing_curve_1: settings
                    .timing_curve
                    .as_ref()
                    .map(|curve| native_motion_path(&curve.path)),
                random_number_seed: Some(random_number_seed),
                writing_direction_is_rtl: Some(false),
                ..Default::default()
            }),
            event_trigger: settings.event_trigger,
            chart_rotation3_d: Some(60.0),
            custom_text_delivery: settings.text_delivery,
            custom_delivery_option: settings.delivery_option,
            action_rotation_angle: settings
                .rotation
                .as_ref()
                .map(|rotation| rotation.total_degrees),
            action_rotation_direction: settings
                .rotation
                .as_ref()
                .map(|rotation| native_rotation_direction(rotation.direction)),
            action_scale_size: settings.scale.as_ref().map(|scale| scale.scale_factor),
            action_color_alpha: settings
                .opacity
                .as_ref()
                .map(|opacity| opacity.opacity_percent),
            action_motion_path_source: settings
                .move_action
                .as_ref()
                .map(|move_action| native_motion_path(&move_action.path)),
            action_acceleration: typed_action_acceleration(settings).map(native_build_acceleration),
            custom_align_to_path: settings
                .move_action
                .as_ref()
                .filter(|move_action| move_action.align_to_path)
                .map(|_| true),
            custom_action_decay: emphasis_decay(settings.emphasis),
            custom_action_repeat_count: emphasis_repeat_count(settings.emphasis),
            custom_action_scale: emphasis_scale(settings.emphasis),
            custom_action_jiggle_intensity: emphasis_jiggle_intensity(settings.emphasis),
            custom_bounce: settings.custom_parameters.bounce,
            custom_motion_blur: settings.custom_parameters.motion_blur,
            custom_include_endpoints: settings.custom_parameters.include_endpoints,
            custom_shine: settings.custom_parameters.shine,
            custom_scale_amount: settings.custom_parameters.scale_amount,
            custom_travel_distance: settings.custom_parameters.travel_distance,
            custom_cursor: settings.keyboard.map(|keyboard| keyboard.show_cursor),
            ..Default::default()
        },
        chunk_id_seed: Some(1),
    }
}

#[allow(deprecated)]
pub(super) fn new_build_chunk(
    build_object_id: u64,
    build_uuid: crate::protobuf::tsp::Uuid,
    settings: &KeynoteBuildSettings,
) -> kn::BuildChunkArchive {
    let (automatic, referent) = build_start_fields(settings.start);
    kn::BuildChunkArchive {
        build: Some(crate::protobuf::tsp::Reference {
            identifier: build_object_id,
            ..Default::default()
        }),
        index: None,
        delay: Some(settings.delay),
        duration: Some(settings.duration),
        automatic: Some(automatic),
        referent: Some(referent),
        build_chunk_identifier: Some(kn::BuildChunkIdentifierArchive {
            build_id: Some(build_uuid),
            build_chunk_id: Some(1),
        }),
        build_id: Some(build_uuid),
    }
}

pub(super) fn normalize_and_patch_build_object(
    object: &mut ArchiveObject,
    settings: &KeynoteBuildSettings,
) -> Result<()> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == BUILD_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    if indexes.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Keynote build object must have exactly one type-{BUILD_MESSAGE_TYPE} payload"
        )));
    }
    let build_index = indexes[0];
    let data = patch_build_settings_wire(object.messages[build_index].data.as_slice(), settings)?;
    object.replace_message(
        build_index,
        RawMessage {
            type_: BUILD_MESSAGE_TYPE,
            data,
        },
    )?;

    let transient = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == 0)
        .map(|(index, message)| {
            kn::BuildAttributesArchive::decode(message.data.as_slice())
                .map(|_| index)
                .map_err(Error::from)
        })
        .collect::<Result<Vec<_>>>()?;
    if !transient.is_empty() {
        if object.archive_info.should_merge != Some(true) {
            return Err(Error::InvalidFormat(
                "Keynote build has a transient payload without merge semantics".to_owned(),
            ));
        }
        for index in transient.into_iter().rev() {
            object.remove_message(index).ok_or_else(|| {
                Error::InvalidFormat("Keynote transient build payload disappeared".to_owned())
            })?;
        }
        object.archive_info.should_merge = None;
    }
    Ok(())
}

pub(super) fn patch_motion_point_wire(
    original: &[u8],
    point: KeynoteMotionPathPoint,
) -> Result<Vec<u8>> {
    let data = patch_fixed32_field(original, 1, true, Some(point.x.to_bits()))?;
    patch_fixed32_field(&data, 2, true, Some(point.y.to_bits()))
}

pub(super) fn patch_motion_node_wire(
    original: &[u8],
    node: &KeynoteMotionPathNode,
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_field(original, 1, |point| {
        patch_motion_point_wire(point, node.in_control_point)
    })?;
    let data = transform_length_delimited_field(&data, 2, |point| {
        patch_motion_point_wire(point, node.point)
    })?;
    let data = transform_length_delimited_field(&data, 3, |point| {
        patch_motion_point_wire(point, node.out_control_point)
    })?;
    patch_varint_field(
        &data,
        4,
        true,
        Some(native_motion_node_type(node.node_type) as u64),
    )
}

pub(super) fn native_motion_node(
    node: &KeynoteMotionPathNode,
) -> tsd::editable_bezier_path_source_archive::Node {
    tsd::editable_bezier_path_source_archive::Node {
        in_control_point: native_motion_point(node.in_control_point),
        node_point: native_motion_point(node.point),
        out_control_point: native_motion_point(node.out_control_point),
        r#type: native_motion_node_type(node.node_type),
    }
}

pub(super) fn patch_motion_subpath_wire(
    original: &[u8],
    subpath: &KeynoteMotionSubpath,
) -> Result<Vec<u8>> {
    let existing = repeated_length_delimited_payloads(original, 1)?;
    let replacements = subpath
        .nodes
        .iter()
        .enumerate()
        .map(|(index, node)| {
            existing.get(index).map_or_else(
                || Ok(native_motion_node(node).encode_to_vec()),
                |wire| patch_motion_node_wire(wire, node),
            )
        })
        .collect::<Result<Vec<_>>>()?;
    let data = rewrite_repeated_length_delimited_fields(original, 1, &replacements)?;
    patch_varint_field(&data, 2, true, Some(u64::from(u8::from(subpath.closed))))
}

pub(super) fn native_motion_subpath(
    subpath: &KeynoteMotionSubpath,
) -> tsd::editable_bezier_path_source_archive::Subpath {
    tsd::editable_bezier_path_source_archive::Subpath {
        nodes: subpath.nodes.iter().map(native_motion_node).collect(),
        closed: subpath.closed,
    }
}

pub(super) fn patch_editable_motion_path_wire(
    original: &[u8],
    path: &KeynoteMotionPath,
) -> Result<Vec<u8>> {
    let existing = repeated_length_delimited_payloads(original, 1)?;
    let replacements = path
        .subpaths
        .iter()
        .enumerate()
        .map(|(index, subpath)| {
            existing.get(index).map_or_else(
                || Ok(native_motion_subpath(subpath).encode_to_vec()),
                |wire| patch_motion_subpath_wire(wire, subpath),
            )
        })
        .collect::<Result<Vec<_>>>()?;
    let data = rewrite_repeated_length_delimited_fields(original, 1, &replacements)?;
    transform_length_delimited_field(&data, 2, |size| {
        let size = patch_fixed32_field(size, 1, true, Some(path.natural_width.to_bits()))?;
        patch_fixed32_field(&size, 2, true, Some(path.natural_height.to_bits()))
    })
}

pub(super) fn patch_motion_path_source_wire(
    original: &[u8],
    path: &KeynoteMotionPath,
) -> Result<Vec<u8>> {
    let source = tsd::PathSourceArchive::decode(original)?;
    let mut data = patch_varint_field(
        original,
        1,
        source.horizontal_flip.is_some(),
        Some(u64::from(u8::from(path.horizontal_flip))),
    )?;
    data = patch_varint_field(
        &data,
        2,
        source.vertical_flip.is_some(),
        Some(u64::from(u8::from(path.vertical_flip))),
    )?;
    transform_length_delimited_field(&data, 8, |editable| {
        patch_editable_motion_path_wire(editable, path)
    })
}

#[allow(deprecated)]
pub(super) fn patch_build_settings_wire(
    original: &[u8],
    settings: &KeynoteBuildSettings,
) -> Result<Vec<u8>> {
    let build = kn::BuildArchive::decode(original)?;
    let animation = build
        .attributes
        .animation_attributes
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Keynote build has no modern attributes".to_owned()))?;
    let original_effect = animation
        .effect
        .as_deref()
        .or(build.attributes.database_effect.as_deref())
        .unwrap_or_default();
    let data = patch_length_delimited_field(original, 2, true, Some(settings.delivery.as_bytes()))?;
    let data = transform_length_delimited_field(&data, 4, |attributes| {
        let mut attributes = patch_varint_field(
            attributes,
            4,
            build.attributes.event_trigger.is_some(),
            settings.event_trigger.map(u64::from),
        )?;
        attributes = patch_varint_field(
            &attributes,
            20,
            build.attributes.custom_text_delivery.is_some(),
            settings
                .text_delivery
                .map(|value| {
                    u64::try_from(value).map_err(|_| {
                        Error::ParseError("negative Keynote text-delivery value".to_owned())
                    })
                })
                .transpose()?,
        )?;
        attributes = patch_varint_field(
            &attributes,
            21,
            build.attributes.custom_delivery_option.is_some(),
            settings
                .delivery_option
                .map(|value| {
                    u64::try_from(value).map_err(|_| {
                        Error::ParseError("negative Keynote delivery-option value".to_owned())
                    })
                })
                .transpose()?,
        )?;
        let preserve_opaque_action =
            settings.animation_type == "Action" && !is_typed_action_effect(&settings.effect);
        if !preserve_opaque_action {
            attributes = patch_fixed64_field(
                &attributes,
                9,
                build.attributes.action_rotation_angle.is_some(),
                settings
                    .rotation
                    .as_ref()
                    .map(|rotation| rotation.total_degrees.to_bits()),
            )?;
            attributes = patch_varint_field(
                &attributes,
                10,
                build.attributes.action_rotation_direction.is_some(),
                settings
                    .rotation
                    .as_ref()
                    .map(|rotation| native_rotation_direction(rotation.direction) as u64),
            )?;
            attributes = patch_fixed64_field(
                &attributes,
                11,
                build.attributes.action_scale_size.is_some(),
                settings
                    .scale
                    .as_ref()
                    .map(|scale| scale.scale_factor.to_bits()),
            )?;
            attributes = patch_fixed64_field(
                &attributes,
                12,
                build.attributes.action_color_alpha.is_some(),
                settings
                    .opacity
                    .as_ref()
                    .map(|opacity| opacity.opacity_percent.to_bits()),
            )?;
            attributes = patch_varint_field(
                &attributes,
                13,
                build.attributes.action_acceleration.is_some(),
                typed_action_acceleration(settings)
                    .map(|acceleration| native_build_acceleration(acceleration) as u64),
            )?;
            if let Some(move_action) = &settings.move_action {
                if build.attributes.action_motion_path_source.is_some() {
                    attributes =
                        transform_length_delimited_field(&attributes, 22, |path_source| {
                            patch_motion_path_source_wire(path_source, &move_action.path)
                        })?;
                } else {
                    let path_source = native_motion_path(&move_action.path).encode_to_vec();
                    attributes =
                        patch_length_delimited_field(&attributes, 22, false, Some(&path_source))?;
                }
            } else {
                attributes = patch_length_delimited_field(
                    &attributes,
                    22,
                    build.attributes.action_motion_path_source.is_some(),
                    None,
                )?;
            }
            let align_to_path = settings.move_action.as_ref().and_then(|move_action| {
                if move_action.align_to_path {
                    Some(1)
                } else if build.attributes.custom_align_to_path == Some(false) {
                    Some(0)
                } else {
                    None
                }
            });
            attributes = patch_varint_field(
                &attributes,
                37,
                build.attributes.custom_align_to_path.is_some(),
                align_to_path,
            )?;
            attributes = patch_varint_field(
                &attributes,
                23,
                build.attributes.custom_action_decay.is_some(),
                emphasis_decay(settings.emphasis).map(|value| u64::from(u8::from(value))),
            )?;
            attributes = patch_varint_field(
                &attributes,
                24,
                build.attributes.custom_action_repeat_count.is_some(),
                emphasis_repeat_count(settings.emphasis).map(u64::from),
            )?;
            attributes = patch_fixed64_field(
                &attributes,
                25,
                build.attributes.custom_action_scale.is_some(),
                emphasis_scale(settings.emphasis).map(f64::to_bits),
            )?;
            attributes = patch_varint_field(
                &attributes,
                26,
                build.attributes.custom_action_jiggle_intensity.is_some(),
                emphasis_jiggle_intensity(settings.emphasis).map(|value| value as u64),
            )?;
        }
        attributes = patch_varint_field(
            &attributes,
            19,
            build.attributes.custom_bounce.is_some(),
            settings
                .custom_parameters
                .bounce
                .map(|value| u64::from(u8::from(value))),
        )?;
        attributes = patch_varint_field(
            &attributes,
            29,
            build.attributes.custom_motion_blur.is_some(),
            settings
                .custom_parameters
                .motion_blur
                .map(|value| u64::from(u8::from(value))),
        )?;
        attributes = patch_varint_field(
            &attributes,
            30,
            build.attributes.custom_include_endpoints.is_some(),
            settings
                .custom_parameters
                .include_endpoints
                .map(|value| u64::from(u8::from(value))),
        )?;
        attributes = patch_varint_field(
            &attributes,
            33,
            build.attributes.custom_shine.is_some(),
            settings
                .custom_parameters
                .shine
                .map(|value| u64::from(u8::from(value))),
        )?;
        attributes = patch_fixed64_field(
            &attributes,
            34,
            build.attributes.custom_scale_amount.is_some(),
            settings.custom_parameters.scale_amount.map(f64::to_bits),
        )?;
        attributes = patch_fixed64_field(
            &attributes,
            35,
            build.attributes.custom_travel_distance.is_some(),
            settings.custom_parameters.travel_distance.map(f64::to_bits),
        )?;
        if original_effect == KEYBOARD_BUILD_EFFECT || settings.effect == KEYBOARD_BUILD_EFFECT {
            attributes = patch_varint_field(
                &attributes,
                36,
                build.attributes.custom_cursor.is_some(),
                settings
                    .keyboard
                    .map(|keyboard| u64::from(u8::from(keyboard.show_cursor))),
            )?;
        }
        transform_length_delimited_field(&attributes, 18, |animation_data| {
            let mut animation_data = patch_length_delimited_field(
                animation_data,
                1,
                animation.animation_type.is_some(),
                Some(settings.animation_type.as_bytes()),
            )?;
            animation_data = patch_length_delimited_field(
                &animation_data,
                2,
                animation.effect.is_some(),
                Some(settings.effect.as_bytes()),
            )?;
            animation_data = patch_fixed64_field(
                &animation_data,
                3,
                animation.duration.is_some(),
                Some(settings.duration.to_bits()),
            )?;
            animation_data = patch_varint_field(
                &animation_data,
                4,
                animation.direction.is_some(),
                settings.direction.map(u64::from),
            )?;
            animation_data = patch_fixed64_field(
                &animation_data,
                5,
                animation.delay.is_some(),
                Some(settings.delay.to_bits()),
            )?;
            if typed_action_acceleration(settings) == Some(BuildAcceleration::Custom) {
                if let Some(curve) = &settings.timing_curve {
                    let replacement = native_motion_path(&curve.path).encode_to_vec();
                    animation_data = if animation
                        .custom_effect_timing_curve_1
                        .as_ref()
                        .and_then(timing_curve_from_native)
                        .is_some()
                    {
                        transform_length_delimited_field(&animation_data, 8, |path_source| {
                            patch_motion_path_source_wire(path_source, &curve.path)
                        })?
                    } else {
                        patch_length_delimited_field(
                            &animation_data,
                            8,
                            animation.custom_effect_timing_curve_1.is_some(),
                            Some(&replacement),
                        )?
                    };
                    animation_data = patch_length_delimited_field(
                        &animation_data,
                        13,
                        animation.custom_effect_timing_curve_theme_name_1.is_some(),
                        None,
                    )?;
                }
            } else {
                animation_data = patch_length_delimited_field(
                    &animation_data,
                    8,
                    animation.custom_effect_timing_curve_1.is_some(),
                    None,
                )?;
                animation_data = patch_length_delimited_field(
                    &animation_data,
                    13,
                    animation.custom_effect_timing_curve_theme_name_1.is_some(),
                    None,
                )?;
            }
            Ok(animation_data)
        })
    })?;
    let verified = kn::BuildArchive::decode(data.as_slice())?;
    let (automatic, referent) = build_start_fields(settings.start);
    let chunk = KeynoteBuildChunkInfo {
        object_id: 0,
        delay: Some(settings.delay),
        duration: Some(settings.duration),
        automatic: Some(automatic),
        referent: Some(referent),
        chunk_id: Some(1),
    };
    if build_settings(
        &verified,
        &[chunk],
        settings.start == BuildStart::AfterTransition,
    ) != *settings
    {
        return Err(Error::InvalidFormat(
            "Keynote build wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn patch_build_chunk_object(
    object: &mut ArchiveObject,
    settings: &KeynoteBuildSettings,
) -> Result<()> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == BUILD_CHUNK_MESSAGE_TYPE)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    if indexes.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Keynote build chunk must have exactly one type-{BUILD_CHUNK_MESSAGE_TYPE} payload"
        )));
    }
    let index = indexes[0];
    let original = object.messages[index].data.as_slice();
    let chunk = kn::BuildChunkArchive::decode(original)?;
    let (automatic, referent) = build_start_fields(settings.start);
    let data = patch_fixed64_field(
        original,
        3,
        chunk.delay.is_some(),
        Some(settings.delay.to_bits()),
    )?;
    let data = patch_fixed64_field(
        &data,
        4,
        chunk.duration.is_some(),
        Some(settings.duration.to_bits()),
    )?;
    let data = patch_varint_field(
        &data,
        5,
        chunk.automatic.is_some(),
        Some(u64::from(automatic)),
    )?;
    let data = patch_varint_field(
        &data,
        6,
        chunk.referent.is_some(),
        Some(u64::from(referent)),
    )?;
    let verified = kn::BuildChunkArchive::decode(data.as_slice())?;
    if verified.delay != Some(settings.delay)
        || verified.duration != Some(settings.duration)
        || verified.automatic != Some(automatic)
        || verified.referent != Some(referent)
    {
        return Err(Error::InvalidFormat(
            "Keynote build-chunk wire patch failed validation".to_owned(),
        ));
    }
    object.replace_message(
        index,
        RawMessage {
            type_: BUILD_CHUNK_MESSAGE_TYPE,
            data,
        },
    )?;
    Ok(())
}

pub(super) fn patch_slide_build_references(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    remove_builds: &[u64],
    remove_chunks: &[u64],
    additions: &[(u64, u64)],
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 5)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        }
        let index = indexes[0];
        let original = object.messages[index].data.as_slice();
        let mut data = original.to_vec();
        for (field, removed) in [(2, remove_builds), (43, remove_chunks)] {
            for identifier in removed {
                let count = repeated_length_delimited_payloads(&data, field)?
                    .into_iter()
                    .filter_map(|payload| crate::protobuf::tsp::Reference::decode(payload).ok())
                    .filter(|reference| reference.identifier == *identifier)
                    .count();
                if count != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} field {field} must contain object {identifier} exactly once"
                    )));
                }
                data = remove_repeated_length_delimited_field_where(&data, field, |payload| {
                    Ok(crate::protobuf::tsp::Reference::decode(payload)?.identifier == *identifier)
                })?;
            }
        }
        for (build_id, chunk_id) in additions {
            for (field, identifier) in [(2, *build_id), (43, *chunk_id)] {
                if repeated_length_delimited_payloads(&data, field)?
                    .into_iter()
                    .filter_map(|payload| crate::protobuf::tsp::Reference::decode(payload).ok())
                    .any(|reference| reference.identifier == identifier)
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} already references object {identifier}"
                    )));
                }
                data = append_repeated_length_delimited_field(
                    &data,
                    field,
                    &crate::protobuf::tsp::Reference {
                        identifier,
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?;
            }
        }
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        for identifier in remove_builds {
            if verified.builds.iter().any(|reference| reference.identifier == *identifier) {
                return Err(Error::InvalidFormat(
                    "Keynote slide build removal failed validation".to_owned(),
                ));
            }
        }
        for identifier in remove_chunks {
            if verified
                .build_chunks
                .iter()
                .any(|reference| reference.identifier == *identifier)
            {
                return Err(Error::InvalidFormat(
                    "Keynote slide build-chunk removal failed validation".to_owned(),
                ));
            }
        }
        object.replace_message(index, RawMessage { type_: 5, data })?;
        let info = &mut object.archive_info.message_infos[index];
        let removed = remove_builds
            .iter()
            .chain(remove_chunks)
            .copied()
            .collect::<HashSet<_>>();
        info.object_references
            .retain(|identifier| !removed.contains(identifier));
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| !removed.contains(identifier));
        }
        for (build_id, chunk_id) in additions {
            for identifier in [*build_id, *chunk_id] {
                if !info.object_references.contains(&identifier) {
                    info.object_references.push(identifier);
                }
            }
        }
        Ok(())
    })
}

pub(super) fn patch_slide_build_order_references(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    build_ids: &[u64],
    chunk_ids: &[u64],
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 5)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        }
        let index = indexes[0];
        let original = object.messages[index].data.as_slice();
        let mut data = original.to_vec();
        for (field, identifiers) in [(2, build_ids), (43, chunk_ids)] {
            let payloads = repeated_length_delimited_payloads(&data, field)?;
            if payloads.len() != identifiers.len() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_id} field {field} has {} references, expected {}",
                    payloads.len(),
                    identifiers.len()
                )));
            }
            let mut payload_by_identifier = HashMap::with_capacity(payloads.len());
            for payload in payloads {
                let identifier = crate::protobuf::tsp::Reference::decode(payload)?.identifier;
                if payload_by_identifier
                    .insert(identifier, payload.to_vec())
                    .is_some()
                {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} field {field} repeats object {identifier}"
                    )));
                }
            }
            let replacements = identifiers
                .iter()
                .map(|identifier| {
                    payload_by_identifier.remove(identifier).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Keynote slide {slide_id} field {field} does not reference object {identifier}"
                        ))
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            if !payload_by_identifier.is_empty() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_id} field {field} contains unexpected references"
                )));
            }
            data = rewrite_repeated_length_delimited_fields(&data, field, &replacements)?;
        }

        let verified = kn::SlideArchive::decode(data.as_slice())?;
        let verified_builds = verified
            .builds
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let verified_chunks = verified
            .build_chunks
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        if verified_builds != build_ids || verified_chunks != chunk_ids {
            return Err(Error::InvalidFormat(
                "Keynote slide build-order wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(index, RawMessage { type_: 5, data })?;
        Ok(())
    })
}

pub(super) fn patch_slide_build_cache(
    package: &mut IWorkPackage,
    graph: &ObjectGraph,
    node_id: u64,
    event_count: usize,
) -> Result<()> {
    let archive_name = graph.archive_name(node_id)?.to_owned();
    let count = u32::try_from(event_count)
        .map_err(|_| Error::ParseError("Keynote build event count exceeds u32".to_owned()))?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == 4)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide node {node_id} must have exactly one SlideNodeArchive payload"
            )));
        }
        let index = indexes[0];
        let original = object.messages[index].data.as_slice();
        let node = kn::SlideNodeArchive::decode(original)?;
        let mut data = patch_varint_field(
            original,
            15,
            node.build_event_count.is_some(),
            (count != 0).then_some(u64::from(count)),
        )?;
        data = patch_varint_field(
            &data,
            26,
            node.build_event_count_cache_version.is_some(),
            Some(if count == 0 { u64::from(u32::MAX) } else { 2 }),
        )?;
        data = patch_varint_field(
            &data,
            20,
            node.has_explicit_builds.is_some(),
            Some(u64::from(count != 0)),
        )?;
        data = patch_varint_field(
            &data,
            27,
            node.has_explicit_builds_cache_version.is_some(),
            Some(2),
        )?;
        let verified = kn::SlideNodeArchive::decode(data.as_slice())?;
        if verified.build_event_count != (count != 0).then_some(count)
            || verified.build_event_count_cache_version
                != Some(if count == 0 { u32::MAX } else { 2 })
            || verified.has_explicit_builds != Some(count != 0)
            || verified.has_explicit_builds_cache_version != Some(2)
        {
            return Err(Error::InvalidFormat(
                "Keynote slide build cache patch failed validation".to_owned(),
            ));
        }
        object.replace_message(index, RawMessage { type_: 4, data })?;
        Ok(())
    })
}
