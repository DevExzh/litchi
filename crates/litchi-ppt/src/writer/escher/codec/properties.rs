//! PPT shape property-family encoding (`OfficeArtFOPT`).

use litchi_odraw::write::{Property, prop_value};

use super::super::{UserShapeData, ppt_prop_value, prop_id};

/// Builds the primary property table for one PPT shape.
pub(crate) fn build_shape_properties(shape: &UserShapeData) -> Vec<Property> {
    let mut props = Vec::with_capacity(16);

    if let Some(rotation) = shape.rotation {
        props.push(Property::new(prop_id::ROTATION, rotation.cast_unsigned()));
    }
    #[allow(
        clippy::cast_possible_truncation,
        reason = "shape validation caps adjustment values at 10 entries"
    )]
    for (index, &value) in shape.adjust_values.iter().enumerate() {
        props.push(Property::new(
            prop_id::ADJUST_VALUE + index as u16,
            value.cast_unsigned(),
        ));
    }

    // Picture shapes have special handling - BLIP reference only, no fill/line.
    if let Some(picture_index) = shape.picture_index {
        props.push(Property::new(0x007F, 0x0080_0080));
        props.push(Property::new(0x4104, picture_index));
        props.push(Property::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_DISABLED,
        ));
        props.push(Property::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
        return props;
    }

    // Fill properties.
    if let Some(fill_color) = shape.fill_color {
        if let Some(fill_type) = shape.fill_type {
            props.push(Property::new(prop_id::FILL_TYPE, fill_type));
        }
        props.push(Property::new(prop_id::FILL_COLOR, fill_color));
        if let Some(back_color) = shape.fill_back_color {
            props.push(Property::new(prop_id::FILL_BACK_COLOR, back_color));
        }
        if let Some(angle) = shape.fill_angle {
            props.push(Property::new(prop_id::FILL_ANGLE, angle.cast_unsigned()));
        }
        if let Some(opacity) = shape.fill_opacity {
            props.push(Property::new(prop_id::FILL_OPACITY, opacity));
        }
        props.push(Property::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_ENABLED,
        ));
    } else {
        props.push(Property::new(prop_id::FILL_COLOR, 0x0800_0004));
        props.push(Property::new(prop_id::FILL_BACK_COLOR, 0x0800_0000));
        props.push(Property::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_DISABLED,
        ));
    }

    if let Some(blip_index) = shape.fill_blip_index {
        props.push(Property::new(prop_id::FILL_BLIP, blip_index));
    }

    // Line properties.
    if let Some(line_color) = shape.line_color {
        props.push(Property::new(prop_id::LINE_COLOR, line_color));
        if let Some(opacity) = shape.line_opacity {
            props.push(Property::new(prop_id::LINE_OPACITY, opacity));
        }
        if let Some(width) = shape.line_width {
            props.push(Property::new(prop_id::LINE_WIDTH, width.cast_unsigned()));
        }
        if let Some(style) = shape.line_style {
            props.push(Property::new(prop_id::LINE_STYLE, style));
        }
        if let Some(dash) = shape.line_dash_style {
            props.push(Property::new(prop_id::LINE_DASH_STYLE, dash));
        }
        if let Some(arrow) = shape.line_start_arrow {
            props.push(Property::new(prop_id::LINE_START_ARROW, arrow));
            props.push(Property::new(
                prop_id::LINE_START_ARROW_WIDTH,
                shape.line_start_arrow_width.unwrap_or(1),
            ));
            props.push(Property::new(
                prop_id::LINE_START_ARROW_LENGTH,
                shape.line_start_arrow_length.unwrap_or(1),
            ));
        }
        if let Some(arrow) = shape.line_end_arrow {
            props.push(Property::new(prop_id::LINE_END_ARROW, arrow));
            props.push(Property::new(
                prop_id::LINE_END_ARROW_WIDTH,
                shape.line_end_arrow_width.unwrap_or(1),
            ));
            props.push(Property::new(
                prop_id::LINE_END_ARROW_LENGTH,
                shape.line_end_arrow_length.unwrap_or(1),
            ));
        }
        if let Some(join_style) = shape.line_join_style {
            props.push(Property::new(prop_id::LINE_JOIN_STYLE, join_style));
        }
        if let Some(end_cap_style) = shape.line_end_cap_style {
            props.push(Property::new(prop_id::LINE_END_CAP_STYLE, end_cap_style));
        }
        props.push(Property::new(prop_id::LINE_STYLE_BOOL, 0x0018_0018));
    } else {
        props.push(Property::new(prop_id::LINE_COLOR, 0x0800_0001));
        props.push(Property::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
    }

    // Shadow properties.
    if shape.has_shadow {
        props.push(Property::new(
            prop_id::SHADOW_TYPE,
            shape.shadow_type.unwrap_or(0),
        ));
        props.push(Property::new(
            prop_id::SHADOW_COLOR,
            shape.shadow_color.unwrap_or(0x0800_0002),
        ));
        props.push(Property::new(
            prop_id::SHADOW_OFFSET_X,
            shape.shadow_offset_x.unwrap_or(25400).cast_unsigned(),
        ));
        props.push(Property::new(
            prop_id::SHADOW_OFFSET_Y,
            shape.shadow_offset_y.unwrap_or(25400).cast_unsigned(),
        ));
        if let Some(opacity) = shape.shadow_opacity {
            props.push(Property::new(prop_id::SHADOW_OPACITY, opacity));
        }
        props.push(Property::new(
            prop_id::SHADOW_BOOL,
            ppt_prop_value::SHADOW_STYLE_ENABLED,
        ));
    } else {
        props.push(Property::new(
            prop_id::SHADOW_COLOR,
            prop_value::SCHEME_SHADOW,
        ));
        props.push(Property::new(
            prop_id::SHADOW_BOOL,
            ppt_prop_value::SHADOW_STYLE_DISABLED,
        ));
    }

    props
}
