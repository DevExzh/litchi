//! Chart package authoring.

use super::{Definition, serialize_content};
use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

/// Detached typed builder for a standalone chart package.
#[derive(Clone, Debug)]
pub struct Builder {
    definition: Definition,
}

impl Builder {
    #[must_use]
    pub fn new() -> Self {
        Self {
            definition: Definition::new("chart:line"),
        }
    }

    /// Supply the typed chart definition to publish.
    #[must_use]
    pub fn with_definition(mut self, definition: Definition) -> Self {
        self.definition = definition;
        self
    }

    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.definition
    }

    pub fn definition_mut(&mut self) -> &mut Definition {
        &mut self.definition
    }

    /// Serialize the definition into a validated chart package.
    ///
    /// # Errors
    ///
    /// Returns an error if the definition fails serialization or validation,
    /// or if the package cannot be written.
    pub fn build(self) -> Result<Vec<u8>> {
        let content_xml = serialize_content(&self.definition)?;
        crate::codec::validate(&content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}
