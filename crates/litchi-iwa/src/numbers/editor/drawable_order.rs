//! Native back-to-front ordering for Numbers sheet drawables.

use super::*;
use crate::drawable_order::{DrawableLayerMove, move_drawable_layer, validate_unique_drawables};

impl NumbersEditor {
    /// List one sheet's drawable identifiers from the back-most to the front-most layer.
    pub fn sheet_drawable_order(&self, sheet_id: u64) -> Result<Vec<u64>> {
        let (_, _, sheet) = numbers_sheet(&self.package, sheet_id)?;
        let ordered = sheet
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        validate_unique_drawables(&ordered, "Numbers sheet drawable order")?;
        Ok(ordered)
    }

    /// Set the exact back-to-front order of every drawable on one sheet.
    ///
    /// The supplied slice must be a permutation of [`Self::sheet_drawable_order`].
    /// Existing reference payloads and unrelated native wire fields are retained.
    pub fn set_sheet_drawable_order(
        &mut self,
        sheet_id: u64,
        ordered_drawable_ids: &[u64],
    ) -> Result<()> {
        let current = self.sheet_drawable_order(sheet_id)?;
        if current == ordered_drawable_ids {
            return Ok(());
        }
        let (archive_name, _, _) = numbers_sheet(&self.package, sheet_id)?;
        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            let (message_index, sheet) = decode_sheet(object)?;
            let previous = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            replace_sheet_drawable_references(
                object,
                message_index,
                &previous,
                ordered_drawable_ids,
            )
        })?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_drawable_order(sheet_id)? != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Numbers sheet drawable-order update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Move one sheet drawable using the native Arrange layer semantics.
    ///
    /// Returns `false` when the requested move is already at the requested
    /// boundary; otherwise the changed order is committed transactionally.
    pub fn move_sheet_drawable(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        movement: DrawableLayerMove,
    ) -> Result<bool> {
        let current = self.sheet_drawable_order(sheet_id)?;
        let Some(ordered) = move_drawable_layer(&current, drawable_object_id, movement)? else {
            return Ok(false);
        };
        self.set_sheet_drawable_order(sheet_id, &ordered)?;
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::DrawableLayerMove;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 180.0 };
    const OVERLAP_SIZE: DrawableSize = DrawableSize {
        width: 320.0,
        height: 240.0,
    };

    #[test]
    fn scratch_spreadsheet_supports_drawable_order_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Layered")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let first = append_rectangle(&mut editor, sheet_id, "First");
        let second = append_rectangle(&mut editor, sheet_id, "Second");
        let third = append_rectangle(&mut editor, sheet_id, "Third");
        let original = editor.sheet_drawable_order(sheet_id).unwrap();
        assert!(original.contains(&first));
        assert!(original.contains(&second));
        assert!(original.contains(&third));

        assert!(
            editor
                .move_sheet_drawable(sheet_id, first, DrawableLayerMove::ToFront)
                .unwrap()
        );
        assert_eq!(
            editor.sheet_drawable_order(sheet_id).unwrap().last(),
            Some(&first)
        );
        assert!(
            editor
                .move_sheet_drawable(sheet_id, first, DrawableLayerMove::ToBack)
                .unwrap()
        );
        assert_eq!(
            editor.sheet_drawable_order(sheet_id).unwrap().first(),
            Some(&first)
        );

        let mut reversed = original.clone();
        reversed.reverse();
        editor
            .set_sheet_drawable_order(sheet_id, &reversed)
            .unwrap();
        assert_eq!(editor.sheet_drawable_order(sheet_id).unwrap(), reversed);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheet_drawable_order(sheet_id).unwrap(), reversed);

        let bytes = editor.to_bytes().unwrap();
        let mut duplicate = reversed.clone();
        duplicate[1] = duplicate[0];
        assert!(
            editor
                .set_sheet_drawable_order(sheet_id, &duplicate)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), bytes);

        editor
            .set_sheet_drawable_order(sheet_id, &original)
            .unwrap();
        assert!(
            !editor
                .move_sheet_drawable(sheet_id, original[0], DrawableLayerMove::ToBack)
                .unwrap()
        );
    }

    fn append_rectangle(editor: &mut NumbersEditor, sheet_id: u64, text: &str) -> u64 {
        editor
            .add_sheet_rectangle(sheet_id, text, OVERLAP_POSITION, OVERLAP_SIZE)
            .unwrap()
            .drawable_object_id
    }
}
