//! Presentation color-map values and bounded XML parsing.

mod codec;
mod model;

pub use codec::{parse_master, parse_override};
pub use model::{Map, Override, Role, Slot};
