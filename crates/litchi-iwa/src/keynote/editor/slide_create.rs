//! Fresh slide creation from presentation theme layouts.

use super::*;

mod component;
mod graph;
pub(super) mod layout;
mod wire;

use component::register_created_slide;
use graph::{find_note_source, take_identifier, template_clone_object_ids};
use layout::{LayoutCatalog, default_layout_node_id, read_layout_graph, resolve_layout};
use wire::{clear_user_guides, insert_slide_node, materialize_slide_object, prepare_slide_number};

impl KeynoteEditor {
    /// List the slide layouts stored by the presentation theme.
    pub fn slide_layouts(&self) -> Result<Vec<KeynoteSlideLayoutInfo>> {
        let graph = ObjectGraph::read(self.package())?;
        let layout_graph = read_layout_graph(&graph)?;
        LayoutCatalog::read(&graph, &layout_graph.theme).map(LayoutCatalog::into_infos)
    }

    /// Return the theme's default slide layout.
    pub fn default_slide_layout(&self) -> Result<KeynoteSlideLayoutId> {
        let graph = ObjectGraph::read(self.package())?;
        let layout_graph = read_layout_graph(&graph)?;
        let identifier = default_layout_node_id(&layout_graph.theme)?;
        if !layout_graph
            .theme
            .templates
            .iter()
            .any(|reference| reference.identifier == identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote default layout node {identifier} is not in the theme layout list"
            )));
        }
        Ok(KeynoteSlideLayoutId(identifier))
    }

    /// Append a fresh, empty slide using a theme layout.
    pub fn add_slide(&mut self, layout: KeynoteSlideLayoutId) -> Result<KeynoteSlideInfo> {
        let index = self.slides()?.len();
        self.insert_slide(index, layout)
    }

    /// Insert a fresh, empty slide using a theme layout.
    pub fn insert_slide(
        &mut self,
        index: usize,
        layout: KeynoteSlideLayoutId,
    ) -> Result<KeynoteSlideInfo> {
        let slides = self.slides()?;
        if index > slides.len() {
            return Err(Error::ParseError(format!(
                "Keynote slide insertion index {index} is out of range for {} slides",
                slides.len()
            )));
        }
        let graph = ObjectGraph::read(self.package())?;
        let layout_graph = read_layout_graph(&graph)?;
        if !layout_graph
            .theme
            .templates
            .iter()
            .any(|reference| reference.identifier == layout.0)
        {
            return Err(Error::ParseError(format!(
                "Keynote theme has no slide layout {}",
                layout.0
            )));
        }
        let resolved = resolve_layout(&graph, layout.0)?;
        let template_archive = self.package().archive(&resolved.archive_name)?;
        let template_ids =
            template_clone_object_ids(&template_archive, resolved.slide_id, &resolved.slide)?;
        let note_source = find_note_source(self, &graph, &slides)?;
        let note_archive = self.package().archive(&note_source.archive_name)?;

        let mut next_identifier = next_object_identifier(self.package())?;
        let new_node_id = take_identifier(&mut next_identifier)?;
        let mut remap = HashMap::with_capacity(template_ids.len() + note_source.object_ids.len());
        for identifier in template_ids.iter().chain(&note_source.object_ids) {
            if remap
                .insert(*identifier, take_identifier(&mut next_identifier)?)
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide creation graph repeats object {identifier}"
                )));
            }
        }
        let attachment_id = take_identifier(&mut next_identifier)?;
        let new_slide_id = remap[&resolved.slide_id];
        let new_note_id = remap[&note_source.note_id];
        let new_note_storage_id = remap[&note_source.storage_id];

        let mut new_archive = Archive {
            objects: template_ids
                .iter()
                .map(|identifier| {
                    let source = template_archive.object(*identifier).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Keynote layout object {identifier} disappeared during creation"
                        ))
                    })?;
                    clone_slide_object(source, &remap)
                })
                .chain(note_source.object_ids.iter().map(|identifier| {
                    let source = note_archive.object(*identifier).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Keynote note object {identifier} disappeared during creation"
                        ))
                    })?;
                    clone_slide_object(source, &remap)
                }))
                .collect::<Result<Vec<_>>>()?,
        };
        materialize_slide_object(
            new_archive.object_mut(new_slide_id).ok_or_else(|| {
                Error::InvalidFormat("Created Keynote slide object is missing".to_owned())
            })?,
            resolved.slide_id,
            new_note_id,
        )?;
        if let Some(reference) = &resolved.slide.user_defined_guide_storage {
            clear_user_guides(&mut new_archive, remap[&reference.identifier])?;
        }
        if let Some(attachment) =
            prepare_slide_number(&mut new_archive, new_slide_id, attachment_id)?
        {
            new_archive.insert_object(attachment)?;
        }

        let mut staged = self.package().clone();
        let new_archive_name = format!("Index/Slide-{new_slide_id}.iwa");
        if staged.contains_entry(&new_archive_name) {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide component {new_archive_name} already exists"
            )));
        }
        staged.replace_archive(&new_archive_name, &new_archive)?;
        let node_archive_name = graph.archive_name(resolved.node_id)?.to_owned();
        let source_node_archive = self.package().archive(&node_archive_name)?;
        let source_node = source_node_archive
            .object(resolved.node_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote layout node {} is missing",
                    resolved.node_id
                ))
            })?;
        let new_node = clone_slide_node(source_node, new_node_id, resolved.slide_id, new_slide_id)?;
        staged.update_archive(&node_archive_name, |archive| {
            archive.insert_object(new_node)?;
            Ok(())
        })?;
        insert_slide_node(
            &mut staged,
            &layout_graph.show_archive_name,
            layout_graph.show_id,
            index,
            new_node_id,
            slides.len(),
        )?;
        register_created_slide(
            &mut staged,
            &resolved,
            &new_archive_name,
            &node_archive_name,
            new_node_id,
            new_slide_id,
            &remap,
            &note_source,
        )?;
        set_package_last_object_identifier(&mut staged, next_identifier - 1)?;

        let preview = KeynoteEditor::from_package(staged.clone())?;
        let text_storage_ids = preview
            .slide_text_storages(index)?
            .into_iter()
            .map(|storage| storage.storage.object_id)
            .collect::<Vec<_>>();
        let mut text = IWorkTextEditor::from_package(staged);
        for storage_id in text_storage_ids {
            text.set_text(storage_id, "")?;
        }
        text.set_text(new_note_storage_id, "")?;
        let staged = text.into_package();
        let verified = KeynoteEditor::from_package(staged)?;
        let created = verified.slides()?.get(index).cloned().ok_or_else(|| {
            Error::InvalidFormat(
                "Created Keynote slide is missing from its insertion point".to_owned(),
            )
        })?;
        if created.slide_id != new_slide_id || created.node_id != new_node_id {
            return Err(Error::InvalidFormat(
                "Keynote slide creation produced the wrong graph identity".to_owned(),
            ));
        }
        if created.notes.as_deref() != Some("") {
            return Err(Error::InvalidFormat(
                "Keynote slide creation did not initialize empty speaker notes".to_owned(),
            ));
        }
        if created
            .title
            .as_deref()
            .is_some_and(|text| !text.is_empty())
            || created.body.as_deref().is_some_and(|text| !text.is_empty())
            || verified
                .slide_text_storages(index)?
                .iter()
                .any(|storage| !storage.storage.text.is_empty())
        {
            return Err(Error::InvalidFormat(
                "Keynote slide creation did not initialize empty layout text".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }
}
