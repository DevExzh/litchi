//! Typed hidden-row and hidden-column CRUD for Pages body tables.

use super::*;

/// One zero-based row or column position in a Pages body table.
pub type PagesTableAxisIndex = crate::table_hidden_axes::TableAxisIndex;
/// Canonical, duplicate-free user-hidden axes of a Pages body table.
pub type PagesTableHiddenAxes = crate::table_hidden_axes::TableHiddenAxes;

impl PagesEditor {
    /// Read the canonical user-hidden rows and columns of a body table.
    pub fn table_hidden_axes(&self, model_object_id: u64) -> Result<PagesTableHiddenAxes> {
        self.require_body_table(model_object_id)?;
        crate::table_hidden_axes::table_hidden_axes(self.package(), model_object_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_table_hidden_axes(
        &mut self,
        model_object_id: u64,
        hidden: &PagesTableHiddenAxes,
    ) -> Result<()> {
        self.require_body_table(model_object_id)?;
        if self.table_hidden_axes(model_object_id)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package().clone();
        crate::table_hidden_axes::set_table_hidden_axes(&mut staged, model_object_id, hidden)?;
        let verified = Self::from_package(staged)?;
        verified.require_body_table(model_object_id)?;
        if verified.table_hidden_axes(model_object_id)? != *hidden {
            return Err(Error::InvalidFormat(
                "Pages table hidden axes failed round-trip validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesDocumentBuilder;

    #[test]
    fn scratch_body_table_roundtrips_hidden_axes_transactionally() {
        let mut editor = PagesDocumentBuilder::new()
            .body_table("Hidden", 4, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].model_object_id;
        let hidden = PagesTableHiddenAxes::new([
            PagesTableAxisIndex::row(2),
            PagesTableAxisIndex::column(1),
        ])
        .unwrap();

        editor.set_table_hidden_axes(table_id, &hidden).unwrap();
        assert_eq!(editor.table_hidden_axes(table_id).unwrap(), hidden);
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_hidden_axes(table_id).unwrap(), hidden);

        let before = editor.to_bytes().unwrap();
        let invalid = PagesTableHiddenAxes::new([PagesTableAxisIndex::row(4)]).unwrap();
        assert!(editor.set_table_hidden_axes(table_id, &invalid).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);

        editor
            .set_table_hidden_axes(table_id, &PagesTableHiddenAxes::empty())
            .unwrap();
        assert!(editor.table_hidden_axes(table_id).unwrap().is_empty());
    }
}
