//! Invalidation of stale Keynote slide-node thumbnail caches.

use super::*;

const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const DATABASE_THUMBNAIL_FIELD: u32 = 3;
const DATABASE_THUMBNAILS_FIELD: u32 = 9;
const THUMBNAIL_SIZES_FIELD: u32 = 10;
const THUMBNAILS_DIRTY_FIELD: u32 = 14;
const THUMBNAILS_FIELD: u32 = 16;
const THUMBNAIL_DIGESTS_FIELD: u32 = 25;

/// Remove rendered previews whose pixels no longer represent a changed slide.
#[allow(deprecated)]
pub(super) fn invalidate(
    package: &mut IWorkPackage,
    archive_name: &str,
    node_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SLIDE_NODE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide node {node_id} must contain exactly one node payload"
            )));
        };
        let original = object.messages[*index].data.as_slice();
        let original_node = kn::SlideNodeArchive::decode(original)?;
        let removed_object_references = original_node
            .database_thumbnail
            .iter()
            .chain(&original_node.database_thumbnails)
            .map(|reference| reference.identifier)
            .collect::<HashSet<_>>();

        let mut data = original.to_vec();
        for field in [
            DATABASE_THUMBNAILS_FIELD,
            THUMBNAIL_SIZES_FIELD,
            THUMBNAILS_FIELD,
            THUMBNAIL_DIGESTS_FIELD,
        ] {
            data = rewrite_repeated_length_delimited_fields(&data, field, &[])?;
        }
        data = patch_length_delimited_field(
            &data,
            DATABASE_THUMBNAIL_FIELD,
            original_node.database_thumbnail.is_some(),
            None,
        )?;
        data = patch_varint_field(
            &data,
            THUMBNAILS_DIRTY_FIELD,
            original_node.thumbnails_are_dirty.is_some(),
            Some(1),
        )?;

        let mut expected = original_node;
        expected.database_thumbnail = None;
        expected.database_thumbnails.clear();
        expected.thumbnail_sizes.clear();
        expected.thumbnails.clear();
        expected
            .digests_for_datas_needing_download_for_thumbnail
            .clear();
        expected.thumbnails_are_dirty = Some(true);
        if kn::SlideNodeArchive::decode(data.as_slice())? != expected {
            return Err(Error::InvalidFormat(
                "Keynote slide preview invalidation failed validation".to_owned(),
            ));
        }
        object.replace_message(
            *index,
            RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        info.object_references
            .retain(|identifier| !removed_object_references.contains(identifier));
        info.data_references.clear();
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| !removed_object_references.contains(identifier));
            field.data_references.clear();
        }
        Ok(())
    })
}
