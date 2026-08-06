//! Shared markup-compatibility preprocessing for OOXML parts.

pub mod alternative;

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{
    active_offsets, process_markup_compatibility, process_ooxml, process_part, process_part_arc,
    process_str,
};
pub use model::{Capabilities, Error, Limits, NAMESPACE, Name, OffsetLimits, Output, Report};
