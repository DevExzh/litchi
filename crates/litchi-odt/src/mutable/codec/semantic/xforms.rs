//! Lossless XForms model declaration operations.

use super::super::super::model::MutableDocument;
use crate::xforms::{self, Model};
use litchi_core::{Position, Result};

impl MutableDocument {
    /// Inspect inert XForms model declarations in `office:forms`.
    pub(crate) fn xforms_models(&self) -> Result<Vec<Model>> {
        self.with_content_xml(|xml| xforms::parse_models_with_xml_limits(xml, self.xml_limits))
    }

    /// Insert a model at the end of the existing `office:forms` container.
    pub(crate) fn insert_xforms_model(&mut self, model: &Model) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            xforms::insert_model_with_xml_limits(xml, model, self.xml_limits)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace one model in `office:forms`.
    pub(crate) fn replace_xforms_model(&mut self, position: Position, model: &Model) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            xforms::replace_model_with_xml_limits(xml, position.get(), model, self.xml_limits)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove one model from `office:forms`.
    pub(crate) fn remove_xforms_model(&mut self, position: Position) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            xforms::remove_model_with_xml_limits(xml, position.get(), self.xml_limits)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }
}
