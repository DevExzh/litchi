//! Typed hidden-row and hidden-column CRUD for Keynote slide tables.

use super::*;

/// One zero-based row or column position in a Keynote slide table.
pub type KeynoteTableAxisIndex = crate::table_hidden_axes::TableAxisIndex;
/// Canonical, duplicate-free user-hidden axes of a Keynote slide table.
pub type KeynoteTableHiddenAxes = crate::table_hidden_axes::TableHiddenAxes;

impl KeynoteEditor {
    /// Read the canonical user-hidden rows and columns of a slide table.
    pub fn slide_table_hidden_axes(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<KeynoteTableHiddenAxes> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::table_hidden_axes::table_hidden_axes(self.package(), model_object_id)
    }

    /// Replace all user-hidden rows and columns transactionally.
    pub fn set_slide_table_hidden_axes(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        hidden: &KeynoteTableHiddenAxes,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        if self.slide_table_hidden_axes(slide_index, model_object_id)? == *hidden {
            return Ok(());
        }
        let mut staged = self.package().clone();
        crate::table_hidden_axes::set_table_hidden_axes(&mut staged, model_object_id, hidden)?;
        let verified = Self::from_package(staged)?;
        require_table_model(&verified, slide_index, model_object_id)?;
        if verified.slide_table_hidden_axes(slide_index, model_object_id)? != *hidden {
            return Err(Error::InvalidFormat(
                "Keynote table hidden axes failed round-trip validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn scratch_slide_table_roundtrips_hidden_axes_transactionally() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let table = editor
            .add_slide_table(
                0,
                "Hidden",
                4,
                3,
                DrawablePoint { x: 320.0, y: 360.0 },
                DrawableSize {
                    width: 1_280.0,
                    height: 480.0,
                },
            )
            .unwrap();
        let hidden = KeynoteTableHiddenAxes::new([
            KeynoteTableAxisIndex::row(2),
            KeynoteTableAxisIndex::column(1),
        ])
        .unwrap();

        editor
            .set_slide_table_hidden_axes(0, table.model_object_id, &hidden)
            .unwrap();
        assert_eq!(
            editor
                .slide_table_hidden_axes(0, table.model_object_id)
                .unwrap(),
            hidden
        );
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_table_hidden_axes(0, table.model_object_id)
                .unwrap(),
            hidden
        );

        let before = editor.to_bytes().unwrap();
        let invalid = KeynoteTableHiddenAxes::new([KeynoteTableAxisIndex::column(3)]).unwrap();
        assert!(
            editor
                .set_slide_table_hidden_axes(0, table.model_object_id, &invalid)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);

        editor
            .set_slide_table_hidden_axes(0, table.model_object_id, &KeynoteTableHiddenAxes::empty())
            .unwrap();
        assert!(
            editor
                .slide_table_hidden_axes(0, table.model_object_id)
                .unwrap()
                .is_empty()
        );
    }
}
