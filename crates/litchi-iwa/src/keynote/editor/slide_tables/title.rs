//! Typed title visibility and outline settings for Keynote slide tables.

use super::*;

/// Lossless optional title settings stored by a Keynote table model.
pub type KeynoteTableTitleSettings = crate::numbers::NumbersTableTitleSettings;

impl KeynoteEditor {
    /// Read a slide table's lossless title visibility and outline settings.
    pub fn slide_table_title_settings(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<KeynoteTableTitleSettings> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_title_settings_in_package(self.package(), model_object_id)
    }

    /// Replace a slide table's title visibility and outline settings transactionally.
    pub fn set_slide_table_title_settings(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        settings: KeynoteTableTitleSettings,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self.slide_table_title_settings(slide_index, model_object_id)? == settings {
            return Ok(());
        }
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_title_settings_in_package(
            &mut staged,
            model_object_id,
            settings,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified.slide_table_title_settings(slide_index, model_object_id)? != settings {
            return Err(Error::InvalidFormat(
                "Keynote table title settings failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}
