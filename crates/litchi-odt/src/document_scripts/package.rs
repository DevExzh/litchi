//! Package-facing boundary for document script content.

use litchi_core::Result;

use super::{Scripts, codec::parse_scripts};

/// Read the optional document-level script inventory from package content XML.
///
/// Package mutation and resource CRUD remain owned by the existing package host;
/// this boundary keeps document XML decoding local to the semantic owner.
pub(crate) fn read(content: &str) -> Result<Option<Scripts>> {
    parse_scripts(content)
}
