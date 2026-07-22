//! Native back-to-front ordering for Keynote slide drawables.

use prost::Message;

use super::text_box_create::text_box_context;
use super::*;
use crate::drawable_order::{
    DrawableLayerMove, move_drawable_layer, reorder_reference_field, validate_unique_drawables,
};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;

impl KeynoteEditor {
    /// List one slide's drawable identifiers from the back-most to the front-most layer.
    pub fn slide_drawable_order(&self, slide_index: usize) -> Result<Vec<u64>> {
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_context(&graph, slide_index)?;
        let ordered = context
            .slide
            .drawables_z_order
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        validate_unique_drawables(&ordered, "Keynote slide drawable order")?;
        Ok(ordered)
    }

    /// Set the exact back-to-front order of every drawable on one slide.
    ///
    /// The supplied slice must be a permutation of [`Self::slide_drawable_order`].
    /// Existing reference payloads and unrelated native wire fields are retained.
    pub fn set_slide_drawable_order(
        &mut self,
        slide_index: usize,
        ordered_drawable_ids: &[u64],
    ) -> Result<()> {
        let current = self.slide_drawable_order(slide_index)?;
        if current == ordered_drawable_ids {
            return Ok(());
        }
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_context(&graph, slide_index)?;
        let archive_name = graph.archive_name(context.slide_id)?.to_owned();
        let mut staged = self.package().clone();
        replace_slide_drawable_order(
            &mut staged,
            &archive_name,
            context.slide_id,
            ordered_drawable_ids,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_drawable_order(slide_index)? != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Keynote slide drawable-order update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Move one slide drawable using the native Arrange layer semantics.
    ///
    /// Returns `false` when the requested move is already at the requested
    /// boundary; otherwise the changed order is committed transactionally.
    pub fn move_slide_drawable(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        movement: DrawableLayerMove,
    ) -> Result<bool> {
        let current = self.slide_drawable_order(slide_index)?;
        let Some(ordered) = move_drawable_layer(&current, drawable_object_id, movement)? else {
            return Ok(false);
        };
        self.set_slide_drawable_order(slide_index, &ordered)?;
        Ok(true)
    }
}

fn replace_slide_drawable_order(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    ordered_drawable_ids: &[u64],
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must have exactly one SlideArchive payload"
            )));
        };
        let message_index = *message_index;
        let original = object.messages[message_index].data.as_slice();
        let previous = kn::SlideArchive::decode(original)?;
        let current = previous
            .drawables_z_order
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        let data = reorder_reference_field(
            original,
            SLIDE_DRAWABLES_Z_ORDER_FIELD,
            &current,
            ordered_drawable_ids,
        )?;
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        if verified.owned_drawables != previous.owned_drawables
            || verified
                .drawables_z_order
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>()
                != ordered_drawable_ids
        {
            return Err(Error::InvalidFormat(
                "Keynote slide drawable-order wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::DrawableLayerMove;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 704.0, y: 284.0 };
    const OVERLAP_SIZE: DrawableSize = DrawableSize {
        width: 512.0,
        height: 512.0,
    };

    #[test]
    fn scratch_presentation_supports_drawable_order_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Layered")
            .subtitle("Arrange API")
            .build()
            .unwrap();
        let first = append_rectangle(&mut editor, "First");
        let second = append_rectangle(&mut editor, "Second");
        let third = append_rectangle(&mut editor, "Third");
        let original = editor.slide_drawable_order(0).unwrap();
        assert!(original.contains(&first));
        assert!(original.contains(&second));
        assert!(original.contains(&third));

        assert!(
            editor
                .move_slide_drawable(0, first, DrawableLayerMove::ToFront)
                .unwrap()
        );
        assert_eq!(editor.slide_drawable_order(0).unwrap().last(), Some(&first));
        assert!(
            editor
                .move_slide_drawable(0, first, DrawableLayerMove::ToBack)
                .unwrap()
        );
        assert_eq!(
            editor.slide_drawable_order(0).unwrap().first(),
            Some(&first)
        );

        let mut reversed = original.clone();
        reversed.reverse();
        editor.set_slide_drawable_order(0, &reversed).unwrap();
        assert_eq!(editor.slide_drawable_order(0).unwrap(), reversed);
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.slide_drawable_order(0).unwrap(), reversed);

        let bytes = editor.to_bytes().unwrap();
        let mut duplicate = reversed.clone();
        duplicate[1] = duplicate[0];
        assert!(editor.set_slide_drawable_order(0, &duplicate).is_err());
        assert_eq!(editor.to_bytes().unwrap(), bytes);

        editor.set_slide_drawable_order(0, &original).unwrap();
        assert!(
            !editor
                .move_slide_drawable(0, original[0], DrawableLayerMove::ToBack)
                .unwrap()
        );
    }

    fn append_rectangle(editor: &mut KeynoteEditor, text: &str) -> u64 {
        editor
            .add_slide_rectangle(0, text, OVERLAP_POSITION, OVERLAP_SIZE)
            .unwrap()
            .drawable_object_id
    }
}
