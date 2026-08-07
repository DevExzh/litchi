//! Transactional slide-component deletion and orphaned-media reclamation.

use super::*;
use crate::data_reference_registry::component_data_identifiers;
use crate::media::MediaAssetId;

impl KeynoteEditor {
    /// Remove a slide and its slide-tree node.
    ///
    /// A dedicated `Index/Slide-<id>.iwa` component is removed in one operation,
    /// including media whose final package reference belonged to that component.
    /// Otherwise only the slide object is removed so unrelated colocated objects
    /// are preserved. Removing the final slide is rejected.
    pub fn remove_slide(&mut self, slide_index: usize) -> Result<KeynoteSlideInfo> {
        let slides = self.slides()?;
        if slides.len() <= 1 {
            return Err(Error::ParseError(
                "Cannot remove the final Keynote slide".to_owned(),
            ));
        }
        let removed = slides.get(slide_index).cloned().ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;

        let graph = ObjectGraph::read(self.package())?;
        let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let show_archive = graph.archive_name(show_id)?.to_owned();
        let node_archive = graph.archive_name(removed.node_id)?.to_owned();
        let slide_archive = graph.archive_name(removed.slide_id)?.to_owned();
        let removed_slide_object_ids = self
            .package()
            .archive(&slide_archive)?
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        let mut staged = self.package().clone();

        remove_slide_node_from_show(&mut staged, &show_archive, show_id, removed.node_id)?;
        remove_object(&mut staged, &node_archive, removed.node_id)?;
        if let Some(document_component) = component_identifier_for_entry(&staged, &node_archive)? {
            remove_component_object_uuids(&mut staged, document_component, &[removed.node_id])?;
        }

        let dedicated_slide_archive = format!("Index/Slide-{}.iwa", removed.slide_id);
        if slide_archive == dedicated_slide_archive {
            staged = remove_dedicated_slide_component(staged, &slide_archive)?;
        } else {
            remove_object(&mut staged, &slide_archive, removed.slide_id)?;
        }

        let mut removed_object_ids = removed_slide_object_ids;
        removed_object_ids.push(removed.node_id);
        release_package_identifier_suffix(&mut staged, &removed_object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.slides()?.len() + 1 != slides.len() {
            return Err(Error::InvalidFormat(
                "Keynote slide deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(removed)
    }
}

fn remove_slide_node_from_show(
    package: &mut IWorkPackage,
    archive_name: &str,
    show_id: u64,
    node_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(show_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| kn::ShowArchive::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote show object {show_id} has no ShowArchive payload"
                ))
            })?;
        let show = kn::ShowArchive::decode(object.messages[message_index].data.as_slice())?;
        let desired = show
            .slide_tree
            .slides
            .iter()
            .filter(|reference| reference.identifier != node_id)
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        if desired.len() + 1 != show.slide_tree.slides.len() {
            return Err(Error::InvalidFormat(format!(
                "Keynote show does not contain slide node {node_id} exactly once"
            )));
        }
        let message_type = object.messages[message_index].type_;
        let data = rewrite_show_slide_references(
            object.messages[message_index].data.as_slice(),
            &show.slide_tree.slides,
            &desired,
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .retain(|&identifier| identifier != node_id);
        for field in &mut object.archive_info.message_infos[message_index].field_infos {
            field
                .object_references
                .retain(|&identifier| identifier != node_id);
        }
        Ok(())
    })
}

fn remove_dedicated_slide_component(
    mut package: IWorkPackage,
    archive_name: &str,
) -> Result<IWorkPackage> {
    let component = component_identifier_for_entry(&package, archive_name)?;
    let data_identifiers = component
        .map(|identifier| component_data_identifiers(&package, identifier))
        .transpose()?
        .unwrap_or_default();
    package.remove_entry(archive_name).ok_or_else(|| {
        Error::InvalidFormat(format!("Keynote slide component {archive_name} is missing"))
    })?;
    let Some(component) = component else {
        return Ok(package);
    };
    remove_component_registration(&mut package, component)?;

    let mut media = IWorkMediaEditor::from_package(package)?;
    for identifier in data_identifiers {
        let identifier = MediaAssetId::try_from(identifier)?;
        if media
            .asset(identifier)
            .is_some_and(|asset| !asset.is_referenced())
        {
            media.remove_unreferenced(identifier)?;
        }
    }
    Ok(media.into_package())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::media::MediaAssetId;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use litchi_keynote::slide::image::Options as ImageOptions;

    const IMAGE_POSITION: DrawablePoint = DrawablePoint { x: 80.0, y: 90.0 };
    const IMAGE_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 180.0,
    };

    fn package_shape(package: &IWorkPackage) -> (Vec<String>, Vec<u64>) {
        let entries = package.entry_names().map(str::to_owned).collect::<Vec<_>>();
        let mut identifiers = package
            .iwa_entry_names()
            .map(str::to_owned)
            .flat_map(|name| {
                package
                    .archive(&name)
                    .unwrap()
                    .objects
                    .iter()
                    .filter_map(|object| object.archive_info.identifier)
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        identifiers.sort_unstable();
        (entries, identifiers)
    }

    #[test]
    fn deleting_created_slide_reclaims_its_component_and_media() {
        let mut editor = KeynoteEditor::create().unwrap();
        let baseline = package_shape(editor.package());
        let layout = editor.default_slide_layout().unwrap();
        editor.add_slide(layout).unwrap();
        editor
            .add_slide_image(
                1,
                "litchi_logo.png",
                include_bytes!("../../../../../media/litchi_logo.png"),
                ImageOptions::new(IMAGE_POSITION, IMAGE_SIZE).unwrap(),
            )
            .unwrap();
        assert_eq!(editor.media_assets().unwrap().len(), 1);

        editor.remove_slide(1).unwrap();

        assert_eq!(editor.slides().unwrap().len(), 1);
        assert!(editor.media_assets().unwrap().is_empty());
        assert_eq!(package_shape(editor.package()), baseline);
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(package_shape(reopened.package()), baseline);
    }

    #[test]
    fn deleting_slide_preserves_media_shared_with_a_survivor() {
        let mut editor = KeynoteEditor::create().unwrap();
        let layout = editor.default_slide_layout().unwrap();
        editor.add_slide(layout).unwrap();
        let created = editor
            .add_slide_image(
                1,
                "litchi_logo.png",
                include_bytes!("../../../../../media/litchi_logo.png"),
                ImageOptions::new(IMAGE_POSITION, IMAGE_SIZE).unwrap(),
            )
            .unwrap();
        editor.duplicate_slide(1).unwrap();

        editor.remove_slide(2).unwrap();

        let assets = editor.media_assets().unwrap();
        assert_eq!(assets.len(), 1);
        let image_data_identifier =
            MediaAssetId::try_from(created.image_data_identifier).expect("valid image media ID");
        assert_eq!(assets[0].data_identifier, image_data_identifier);
        assert_eq!(
            editor.extract_media(image_data_identifier.get()).unwrap(),
            include_bytes!("../../../../../media/litchi_logo.png")
        );
        assert_eq!(editor.slide_images(1).unwrap(), [created]);
    }

    #[test]
    fn rejected_slide_deletions_are_transactional() {
        let mut editor = KeynoteEditor::create().unwrap();
        let before = editor.to_bytes().unwrap();

        assert!(editor.remove_slide(0).is_err());
        assert!(editor.remove_slide(1).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
