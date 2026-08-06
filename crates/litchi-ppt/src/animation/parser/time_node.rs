//! Layered parsing façade for PowerPoint timing records.
//!
//! The implementation is separated into semantic models, record parsers,
//! property/extension decoders, and validation while retaining the historical
//! parser entry points.

mod atom;
mod extension;
mod model;
mod parser;
mod properties;
mod validation;

pub use atom::parse_time_node_atom;
pub use extension::parse_slide_animation_extension;
pub use parser::{parse_extended_time_node, parse_time_sub_effect};
pub use properties::parse_time_node_property_list;
