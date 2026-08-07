//! Native back-to-front ordering for Pages document drawables.

use litchi_iwa_common::comment::DrawableId;
use prost::Message;

use super::*;
use crate::drawable_order::{
    DrawableLayerMove, move_drawable_layer, reorder_reference_field, validate_unique_drawables,
};

const DRAWABLES_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
const DRAWABLES_Z_ORDER_REFERENCES_FIELD: u32 = 1;

impl PagesEditor {
    /// List document drawable identifiers from the back-most to the front-most layer.
    pub fn body_drawable_order(&self) -> Result<Vec<DrawableId>> {
        pages_drawable_order(self.package())
    }

    /// Set the exact back-to-front order of every document drawable.
    ///
    /// The supplied slice must be a permutation of [`Self::body_drawable_order`].
    /// Reference payloads and unrelated native wire fields are retained verbatim.
    pub fn set_body_drawable_order(&mut self, ordered_drawable_ids: &[DrawableId]) -> Result<()> {
        let current = self.body_drawable_order()?;
        if current == ordered_drawable_ids {
            return Ok(());
        }
        let mut staged = self.package().clone();
        replace_pages_drawable_order(&mut staged, ordered_drawable_ids)?;
        let verified = Self::from_package(staged)?;
        if verified.body_drawable_order()? != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Pages drawable-order update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Move one drawable using the native Arrange layer semantics.
    ///
    /// Returns `false` when the requested move is already at the requested
    /// boundary; otherwise the changed order is committed transactionally.
    pub fn move_body_drawable(
        &mut self,
        drawable_object_id: DrawableId,
        movement: DrawableLayerMove,
    ) -> Result<bool> {
        let current = self.body_drawable_order()?;
        let Some(ordered) = move_drawable_layer(&current, drawable_object_id, movement)? else {
            return Ok(false);
        };
        self.set_body_drawable_order(&ordered)?;
        Ok(true)
    }
}

fn pages_drawable_order(package: &IWorkPackage) -> Result<Vec<DrawableId>> {
    let document = root_document(package)?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let archive_name = find_object_archive(package, z_order_id.identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(z_order_id.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages drawable z-order object {} is missing",
            z_order_id.identifier
        ))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == DRAWABLES_Z_ORDER_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages drawable z-order object {} must have exactly one payload",
            z_order_id.identifier
        )));
    };
    let z_order = tp::DrawablesZOrderArchive::decode(message.data.as_slice())?;
    let ordered = z_order
        .drawables
        .iter()
        .map(|reference| DrawableId::from_raw(reference.identifier).map_err(crate::Error::from))
        .collect::<Result<Vec<_>>>()?;
    validate_unique_drawables(&ordered, "Pages drawable order")?;
    Ok(ordered)
}

fn replace_pages_drawable_order(
    package: &mut IWorkPackage,
    ordered_drawable_ids: &[DrawableId],
) -> Result<()> {
    let document = root_document(package)?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let archive_name = find_object_archive(package, z_order_id.identifier)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(z_order_id.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable z-order object {} is missing",
                z_order_id.identifier
            ))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == DRAWABLES_Z_ORDER_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages drawable z-order object {} must have exactly one payload",
                z_order_id.identifier
            )));
        };
        let message_index = *message_index;
        let original = object.messages[message_index].data.as_slice();
        let previous = tp::DrawablesZOrderArchive::decode(original)?;
        let current = previous
            .drawables
            .iter()
            .map(|reference| DrawableId::from_raw(reference.identifier).map_err(crate::Error::from))
            .collect::<Result<Vec<_>>>()?;
        let current_raw = current
            .iter()
            .map(|identifier| identifier.get())
            .collect::<Vec<_>>();
        let requested_raw = ordered_drawable_ids
            .iter()
            .map(|identifier| identifier.get())
            .collect::<Vec<_>>();
        let data = reorder_reference_field(
            original,
            DRAWABLES_Z_ORDER_REFERENCES_FIELD,
            &current_raw,
            &requested_raw,
        )?;
        let verified = tp::DrawablesZOrderArchive::decode(data.as_slice())?;
        let verified = verified
            .drawables
            .iter()
            .map(|reference| DrawableId::from_raw(reference.identifier).map_err(crate::Error::from))
            .collect::<Result<Vec<_>>>()?;
        if verified != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Pages drawable-order wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: DRAWABLES_Z_ORDER_MESSAGE_TYPE,
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
    use crate::shapes::{DrawablePoint, DrawableSize};

    const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const OVERLAP_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 180.0,
    };

    #[test]
    fn scratch_document_supports_drawable_order_crud() {
        let mut editor = PagesEditor::create_with_text("Layered").unwrap();
        let first = append_rectangle(&mut editor, "First");
        let second = append_rectangle(&mut editor, "Second");
        let third = append_rectangle(&mut editor, "Third");
        let original = editor.body_drawable_order().unwrap();
        assert!(original.contains(&first));
        assert!(original.contains(&second));
        assert!(original.contains(&third));

        assert!(
            editor
                .move_body_drawable(first, DrawableLayerMove::ToFront)
                .unwrap()
        );
        assert_eq!(editor.body_drawable_order().unwrap().last(), Some(&first));
        assert!(
            editor
                .move_body_drawable(first, DrawableLayerMove::ToBack)
                .unwrap()
        );
        assert_eq!(editor.body_drawable_order().unwrap().first(), Some(&first));

        let mut reversed = original.clone();
        reversed.reverse();
        editor.set_body_drawable_order(&reversed).unwrap();
        assert_eq!(editor.body_drawable_order().unwrap(), reversed);
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_drawable_order().unwrap(), reversed);

        let bytes = editor.to_bytes().unwrap();
        let mut duplicate = reversed.clone();
        duplicate[1] = duplicate[0];
        assert!(editor.set_body_drawable_order(&duplicate).is_err());
        assert_eq!(editor.to_bytes().unwrap(), bytes);

        editor.set_body_drawable_order(&original).unwrap();
        assert!(
            !editor
                .move_body_drawable(original[0], DrawableLayerMove::ToBack)
                .unwrap()
        );
    }

    fn append_rectangle(editor: &mut PagesEditor, text: &str) -> DrawableId {
        let anchor = editor.body_text().unwrap().encode_utf16().count();
        DrawableId::from_raw(
            editor
                .add_body_rectangle(anchor, text, OVERLAP_POSITION, OVERLAP_SIZE)
                .unwrap()
                .drawable_object_id,
        )
        .unwrap()
    }
}
