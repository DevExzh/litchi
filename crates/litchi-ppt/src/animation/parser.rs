//! Animation record parser façade.

mod animation;
mod behavior;
mod build;
mod support;
#[cfg(test)]
mod tests;
mod time_node;
mod timeline;

pub use super::diagram_build::parse_record as parse_diagram_build_record;
pub use animation::{parse_animation_info, parse_animation_info_atom};
pub use behavior::{
    parse_time_animate_behavior, parse_time_animate_behavior_atom, parse_time_animation_value_atom,
    parse_time_animation_value_list, parse_time_behavior, parse_time_behavior_atom,
    parse_time_behavior_property_list, parse_time_color_behavior, parse_time_color_behavior_atom,
    parse_time_effect_behavior, parse_time_effect_behavior_atom, parse_time_motion_behavior,
    parse_time_motion_behavior_atom, parse_time_visual_element,
};
pub use build::parse_build_list;
pub use time_node::{
    parse_extended_time_node, parse_slide_animation_extension, parse_time_node_atom,
    parse_time_node_property_list, parse_time_sub_effect,
};
pub use timeline::{
    parse_time_command_behavior, parse_time_command_behavior_atom, parse_time_condition,
    parse_time_condition_atom, parse_time_iterate_data, parse_time_modifier,
    parse_time_rotation_behavior, parse_time_rotation_behavior_atom, parse_time_scale_behavior,
    parse_time_scale_behavior_atom, parse_time_sequence_data, parse_time_set_behavior,
    parse_time_set_behavior_atom,
};
