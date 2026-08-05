//! Animation record writer façade.

mod animation;
mod behavior;
mod build;
mod support;
#[cfg(test)]
mod tests;
mod time_node;
mod timeline;

pub use animation::{
    write_animation_info, write_animation_info_atom, write_interactive_info_with_sound,
};
pub use behavior::{
    write_time_animate_behavior, write_time_animate_behavior_atom, write_time_animation_value_atom,
    write_time_animation_value_list, write_time_behavior, write_time_behavior_atom,
    write_time_behavior_property_list, write_time_color_behavior, write_time_color_behavior_atom,
    write_time_effect_behavior, write_time_effect_behavior_atom, write_time_motion_behavior,
    write_time_motion_behavior_atom, write_time_visual_element,
};
pub use build::write_build_list;
pub use time_node::{
    write_extended_time_node, write_time_node_atom, write_time_node_property_list,
    write_time_sub_effect,
};
pub use timeline::{
    write_time_command_behavior, write_time_command_behavior_atom, write_time_condition,
    write_time_condition_atom, write_time_iterate_data, write_time_modifier,
    write_time_rotation_behavior, write_time_rotation_behavior_atom, write_time_scale_behavior,
    write_time_scale_behavior_atom, write_time_sequence_data, write_time_set_behavior,
    write_time_set_behavior_atom,
};
