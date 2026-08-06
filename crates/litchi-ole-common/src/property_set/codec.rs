//! MS-OLEPS parsing, serialization, and transactional CFB integration.

mod binary;
mod editor;
mod package;
mod semantic;
mod support;
#[cfg(test)]
mod tests;

pub use editor::Editor;
pub use package::PropertySetReader;

use super::model::Section;
use litchi_cfb::OleError;

pub(crate) fn validate_section(section: &Section, version: u16) -> Result<(), OleError> {
    semantic::validate_section(section, version)
}

// Keep the existing private property_set::codec test seam available to the
// parent module without exposing binary implementation details to consumers.
#[cfg(test)]
pub(super) use binary::{filetime_to_date, filetime_to_duration, parse_typed_property};
