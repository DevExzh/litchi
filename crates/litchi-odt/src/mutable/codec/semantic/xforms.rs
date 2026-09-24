//! Lossless XForms model declaration operations.

use super::super::super::model::MutableDocument;
use crate::xforms::{self, Model};
use litchi_core::{Position, Result};

impl MutableDocument {
    /// Inspect inert XForms model declarations in `office:forms`.
    pub fn xforms_models(&self) -> Result<Vec<Model>> {
        self.with_content_xml(xforms::parse_models)
    }

    /// Insert a model at the end of the existing `office:forms` container.
    pub fn insert_xforms_model(&mut self, model: &Model) -> Result<()> {
        let updated = self.with_content_xml(|xml| xforms::insert_model(xml, model))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one model in `office:forms`.
    pub fn replace_xforms_model(&mut self, position: Position, model: &Model) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| xforms::replace_model(xml, position.get(), model))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove one model from `office:forms`.
    pub fn remove_xforms_model(&mut self, position: Position) -> Result<()> {
        let updated = self.with_content_xml(|xml| xforms::remove_model(xml, position.get()))?;
        self.content_xml = Some(updated);
        Ok(())
    }
}
