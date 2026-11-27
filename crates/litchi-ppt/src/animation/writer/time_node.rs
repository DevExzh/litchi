//! Layered writer façade for `PowerPoint` 2002 timing records.
//!
//! The semantic view, validation rules, and record codecs are private to this
//! owner. The historical writer entry points remain available through the
//! surrounding animation writer façade.

mod codec;
mod model;
mod properties;
mod validation;

pub use codec::{write_extended_time_node, write_time_node_atom, write_time_sub_effect};
pub use properties::write_time_node_property_list;
