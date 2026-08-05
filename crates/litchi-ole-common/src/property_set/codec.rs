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

// Keep the existing private property_set::codec test seam available to the
// parent module without exposing binary implementation details to consumers.
#[cfg(test)]
pub(super) use binary::{filetime_to_date, filetime_to_duration, parse_typed_property};
