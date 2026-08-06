//! Table-template facade combining semantic, XML codec, and validation layers.

mod codec;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_parts};
pub use semantic::{Axis, Region, Style, Template};
