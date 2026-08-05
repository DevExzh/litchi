//! Typed title visibility and outline settings for Pages body tables.

use super::*;

/// Lossless optional title settings stored by a Pages table model.
pub type PagesTableTitleSettings = crate::numbers::Settings;

impl PagesEditor {
    /// Read a body table's lossless title visibility and outline settings.
    pub fn table_title_settings(&self, model_object_id: u64) -> Result<PagesTableTitleSettings> {
        self.require_body_table(model_object_id)?;
        crate::numbers::editor::table_title_settings_in_package(self.package(), model_object_id)
    }

    /// Replace a body table's title visibility and outline settings transactionally.
    pub fn set_table_title_settings(
        &mut self,
        model_object_id: u64,
        settings: PagesTableTitleSettings,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_title_settings(model_object_id)? == settings {
            return Ok(());
        }
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_title_settings_in_package(
            &mut staged,
            model_object_id,
            settings,
        )?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_title_settings(model_object_id)? != settings {
            return Err(Error::InvalidFormat(
                "Pages table title settings failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}
