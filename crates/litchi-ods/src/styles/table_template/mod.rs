//! Table-template facade combining semantic, XML codec, and validation layers.

mod codec;
mod semantic;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_parts};
pub use semantic::{Axis, Region, Style, Template};
pub(crate) use transaction::styles_xml_for_templates;
pub use transaction::{Commit, Edit, Patch, Snapshot};
