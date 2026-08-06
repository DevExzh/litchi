//! Layered parsers for shared PowerPoint animation behavior records.
//!
//! The façade keeps the historical parser entry points together while the
//! implementation is organized by responsibility: record parsing, typed
//! payload models, and semantic validation.

mod model;
mod parser;
#[cfg(test)]
mod tests;
mod validation;

pub use parser::{
    parse_time_animate_behavior, parse_time_animate_behavior_atom, parse_time_animation_value_atom,
    parse_time_animation_value_list, parse_time_behavior, parse_time_behavior_atom,
    parse_time_behavior_property_list, parse_time_color_behavior, parse_time_color_behavior_atom,
    parse_time_effect_behavior, parse_time_effect_behavior_atom, parse_time_motion_behavior,
    parse_time_motion_behavior_atom, parse_time_visual_element,
};

pub(super) use model::{parse_time_variant_string, require_time_variant_payload};
