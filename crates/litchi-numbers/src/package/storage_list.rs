//! Paragraph lists of one `TSWP.StorageArchive` in a structural working set
//! (record 0370b).
//!
//! W1-D's list engine ([`litchi_iwa_text_wire::list`], record 0374) plans
//! each change from the storage and its list styles; this module publishes
//! the plan as the host text editor's list families do: list-style
//! boundaries are rewritten with their reference sequence, presets and
//! paragraph variations are stored beside their stylesheet and registered
//! there, and the metadata registry gains or loses the matching UUIDs and
//! cross-component references. Any storage of the package can be published
//! this way; the Numbers cell-list family ([`super::table_cell_list`]) runs
//! it on a cell's rich-text storage.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the publisher keeps each step beside the operation that runs it"
)]

use std::collections::BTreeMap;

use litchi_iwa_common::WireLimits;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Change, Plan, PropertyKind, Styles};

use super::structure::{
    self, Fault, Working,
    metadata::{Registry, Transition},
};
use super::table_topology::wire::{repeated_bytes, single_bytes, single_reference};

/// `TSWP.StorageArchive`.
pub(super) const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;

/// Why a storage-list publication failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Failure {
    /// A structural fault of the working set.
    Fault(Fault),
    /// The list engine's error.
    List(list::Error),
}

impl From<Fault> for Failure {
    fn from(fault: Fault) -> Self {
        Self::Fault(fault)
    }
}

impl From<list::Error> for Failure {
    fn from(error: list::Error) -> Self {
        Self::List(error)
    }
}

type Result<T> = std::result::Result<T, Failure>;

/// The working set's list styles, as the engine reads them.
pub(super) struct ListStyles<'w, 'p>(pub(super) &'w Working<'p>);

impl Styles for ListStyles<'_, '_> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == list::MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

/// One located storage: its component, message index and payload.
pub(super) struct Storage {
    pub(super) identifier: u64,
    pub(super) entry: String,
    pub(super) index: usize,
    pub(super) payload: Vec<u8>,
}

/// The storage `identifier`.
pub(super) fn storage(working: &Working<'_>, identifier: u64) -> Result<Storage> {
    let entry = working.locate(identifier)?.ok_or(Fault::Invalid)?;
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?;
    let payload = structure::copy_bytes(&object.messages[index].data)?;
    Ok(Storage {
        identifier,
        entry,
        index,
        payload,
    })
}

/// The publisher of one transaction: the working set, the source registry
/// and the registry transition it accumulates.
pub(super) struct Publisher<'a, 'p> {
    pub(super) working: &'a mut Working<'p>,
    pub(super) registry: &'a Registry,
    pub(super) transition: &'a mut Transition,
    pub(super) limits: WireLimits,
}

impl Publisher<'_, '_> {
    fn component(&self, entry: &str) -> Option<u64> {
        self.registry
            .component_for_entry(entry)
            .ok()
            .map(|component| component.identifier)
    }

    /// The list-style boundaries' styles, in order.
    fn boundary_styles(&self, payload: &[u8]) -> Result<Vec<u64>> {
        Ok(list::boundaries(payload, self.limits)?
            .into_iter()
            .map(|entry| entry.1)
            .collect())
    }

    /// The host's `replace_reference_sequence`: replace the storage payload
    /// and swap the old boundary styles for the new ones, in place, in the
    /// message's references and in every field info naming one of them.
    fn replace_with_sequence(
        &mut self,
        storage: &Storage,
        payload: Vec<u8>,
        old: &[u64],
        new: &[u64],
    ) -> Result<()> {
        fn sequence(references: &mut Vec<u64>, old: &[u64], new: &[u64]) -> Result<()> {
            let positions = references
                .iter()
                .filter(|reference| old.contains(reference))
                .count();
            if positions != old.len() {
                return Err(Fault::Invalid.into());
            }
            let insertion = references
                .iter()
                .position(|reference| old.contains(reference))
                .unwrap_or(references.len());
            references.retain(|reference| !old.contains(reference));
            structure::reserve(references, new.len())?;
            references.splice(insertion..insertion, new.iter().copied());
            Ok(())
        }
        super::table_cell_text_style::replace_message(
            self.working,
            storage.identifier,
            storage.index,
            payload,
        )?;
        let object = self.working.object_mut(storage.identifier)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(storage.index)
            .ok_or(Fault::Invalid)?;
        sequence(&mut info.object_references, old, new)?;
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|reference| old.contains(reference))
            {
                sequence(&mut field.object_references, old, new)?;
            }
        }
        Ok(())
    }

    /// Replace the storage payload, keeping its references.
    fn replace(&mut self, storage: &Storage, payload: Vec<u8>) -> Result<()> {
        super::table_cell_text_style::replace_message(
            self.working,
            storage.identifier,
            storage.index,
            payload,
        )?;
        Ok(())
    }

    /// The stylesheet message of `stylesheet`: its index and payload.
    fn stylesheet(&self, stylesheet: u64) -> Result<(usize, Vec<u8>)> {
        let object = self.working.object(stylesheet)?;
        let index = structure::unique_message(object, &[STYLESHEET_MESSAGE_TYPE])?;
        Ok((index, structure::copy_bytes(&object.messages[index].data)?))
    }

    /// Replace the stylesheet payload and append `style` to its references.
    fn replace_stylesheet(&mut self, stylesheet: u64, data: Vec<u8>, style: u64) -> Result<()> {
        let (index, _) = self.stylesheet(stylesheet)?;
        super::table_cell_text_style::replace_message(self.working, stylesheet, index, data)?;
        let info = self
            .working
            .object_mut(stylesheet)?
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Fault::Invalid)?;
        if info.object_references.contains(&style) {
            return Err(Fault::Invalid.into());
        }
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(style);
        Ok(())
    }

    /// A new list-style object: one message at the engine's version.
    fn new_style(
        &self,
        identifier: u64,
        payload: Vec<u8>,
        references: Vec<u64>,
    ) -> Result<litchi_iwa_core::ArchiveObject> {
        let mut object = structure::new_object(
            identifier,
            None,
            list::MESSAGE_TYPE,
            payload,
            references,
            self.working.archive_limits(),
        )?;
        object.archive_info.message_infos[0].versions = list::MESSAGE_VERSION.to_vec().into();
        Ok(object)
    }

    /// The host's `register_style_reference`: the storage's component
    /// references `style` in the style's component.
    fn register_reference(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        style: u64,
    ) -> Result<()> {
        if let (Some(owner), Some(target)) =
            (self.component(owner_entry), self.component(style_entry))
        {
            self.transition
                .reference(self.registry, owner, target, style)?;
        }
        Ok(())
    }

    /// The host's `register_private_style`: a UUID in the style's component
    /// and the storage component's reference.
    fn register_private(&mut self, owner_entry: &str, style_entry: &str, style: u64) -> Result<()> {
        let Some(target) = self.component(style_entry) else {
            return Ok(());
        };
        let seed = structure::metadata::identity_seed(self.working, style)?;
        self.transition
            .register_uuid(self.registry, target, style, seed)?;
        self.register_reference(owner_entry, style_entry, style)
    }

    /// The host's `unregister_owner_reference_if_unused`: the storage
    /// component's reference to `style` goes once none of its objects
    /// references the style.
    fn unregister_if_unused(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        style: u64,
    ) -> Result<()> {
        let (Some(owner), Some(target)) =
            (self.component(owner_entry), self.component(style_entry))
        else {
            return Ok(());
        };
        if owner == target {
            return Ok(());
        }
        let archive = self.working.archive(owner_entry)?;
        if archive.objects.iter().any(|object| {
            object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&style))
        }) {
            return Ok(());
        }
        self.transition
            .remove_reference(self.registry, owner, target, style)?;
        Ok(())
    }

    /// The stylesheet's preset style for `preset`, stored and registered
    /// when it has none (the host's `preset_style_id` with creation).
    fn preset_object(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        stylesheet: u64,
        preset: ParagraphList,
    ) -> Result<u64> {
        let existing = {
            let archive = self.working.archive(style_entry)?;
            let candidates = archive.objects.iter().filter_map(|object| {
                let identifier = object.archive_info.identifier?;
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == list::MESSAGE_TYPE)
                    .map(|message| (identifier, message.data.as_slice()))
            });
            list::find_preset(candidates, stylesheet, preset, self.limits)?
        };
        if let Some(identifier) = existing {
            self.register_reference(owner_entry, style_entry, identifier)?;
            return Ok(identifier);
        }
        let created = *self.working.allocate(1)?.first().ok_or(Fault::Invalid)?;
        let object = self.new_style(
            created,
            list::preset_payload(stylesheet, preset, self.limits)?,
            Vec::new(),
        )?;
        // The host's `insert_private_style`: appended to the stylesheet's
        // styles (field 1).
        let (_, sheet) = self.stylesheet(stylesheet)?;
        let mut reference = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut reference, 1, created)
            .map_err(|error| structure::map_wire(&error))?;
        let data =
            litchi_iwa_common::wire::append_repeated_length_delimited_field(&sheet, 1, &reference)
                .map_err(|error| structure::map_wire(&error))?;
        self.replace_stylesheet(stylesheet, data, created)?;
        self.working.insert(style_entry, object)?;
        self.register_private(owner_entry, style_entry, created)?;
        Ok(created)
    }

    /// Give every paragraph of `storage` the one preset `preset` (the host
    /// text editor's `set_paragraph_list`).
    pub(super) fn set_list(&mut self, storage: u64, preset: ParagraphList) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        let style = list::uniform_style(&located.payload, self.limits)?;
        if list::preset(&ListStyles(self.working), style, self.limits).ok() == Some(preset) {
            return Ok(false);
        }
        let style_entry = self.working.locate(style)?.ok_or(Fault::Invalid)?;
        let stylesheet = list::stylesheet(&ListStyles(self.working), style, self.limits)?;
        let target = self.preset_object(&located.entry, &style_entry, stylesheet, preset)?;
        let located = self::storage(self.working, storage)?;
        let payload = list::replace_boundaries(&located.payload, &[(0, target)], self.limits)?;
        self.replace_with_sequence(&located, payload, &[style], &[target])?;
        self.unregister_if_unused(&located.entry, &style_entry, style)?;
        Ok(true)
    }

    /// Give the paragraphs of `storage` the presets of `placements` (the
    /// host text editor's `set_paragraph_lists`).
    pub(super) fn set_lists(
        &mut self,
        storage: u64,
        placements: &[ParagraphListPlacement],
    ) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        let normalized = list::normalize(&located.payload, placements, self.limits)?;
        if list::placements(&located.payload, &ListStyles(self.working), self.limits)? == normalized
        {
            return Ok(false);
        }
        let boundaries = list::boundaries(&located.payload, self.limits)?;
        let first = boundaries.first().ok_or(Fault::Invalid)?.1;
        let style_entry = self.working.locate(first)?.ok_or(Fault::Invalid)?;
        let stylesheet = list::stylesheet(&ListStyles(self.working), first, self.limits)?;
        let slot = |preset: ParagraphList| match preset {
            ParagraphList::None => 0,
            ParagraphList::Bullet => 1,
            ParagraphList::Numbered => 2,
        };
        let mut slots: [Option<u64>; 3] = [None; 3];
        for placement in &normalized {
            if slots[slot(placement.list)].is_none() {
                slots[slot(placement.list)] = Some(self.preset_object(
                    &located.entry,
                    &style_entry,
                    stylesheet,
                    placement.list,
                )?);
            }
        }
        let mut new = Vec::new();
        structure::reserve(&mut new, normalized.len())?;
        for placement in &normalized {
            new.push((
                placement.paragraph.utf16_index(),
                slots[slot(placement.list)].ok_or(Fault::Invalid)?,
            ));
        }
        let located = self::storage(self.working, storage)?;
        let old: Vec<u64> = boundaries.iter().map(|entry| entry.1).collect();
        let payload = list::replace_boundaries(&located.payload, &new, self.limits)?;
        let styles: Vec<u64> = new.iter().map(|entry| entry.1).collect();
        self.replace_with_sequence(&located, payload, &old, &styles)?;
        let mut unique = old;
        unique.sort_unstable();
        unique.dedup();
        for style in unique {
            self.unregister_if_unused(&located.entry, &style_entry, style)?;
        }
        Ok(true)
    }

    /// Set the list level of the paragraph at `paragraph`.
    pub(super) fn set_level(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        let Some(payload) = list::with_level(&located.payload, paragraph, level, self.limits)?
        else {
            return Ok(false);
        };
        self.replace(&located, payload)?;
        Ok(true)
    }

    /// Set the numbering of the paragraph at `paragraph`.
    pub(super) fn set_numbering(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        let Some(payload) = list::with_numbering(
            &located.payload,
            &ListStyles(self.working),
            paragraph,
            numbering,
            self.limits,
        )?
        else {
            return Ok(false);
        };
        self.replace(&located, payload)?;
        Ok(true)
    }

    /// Set or reset one list property of the paragraph at `paragraph`; a
    /// reset that leaves the paragraph's variation overriding nothing
    /// collapses it onto its parent (the host's reset clean-up).
    pub(super) fn property(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        change: &Change,
    ) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        let exclusive = Exclusive::load(self.working)?;
        let plan = list::plan(
            &located.payload,
            &ListStyles(self.working),
            paragraph,
            change,
            |style| exclusive.is_exclusive(style),
            self.limits,
        )?;
        let changed = self.apply(storage, paragraph, plan)?;
        if let Change::Reset(kind) = change
            && changed
        {
            self.collapse(storage, paragraph, *kind)?;
        }
        Ok(changed)
    }

    fn collapse(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<()> {
        let located = self::storage(self.working, storage)?;
        let exclusive = Exclusive::load(self.working)?;
        let plan = list::plan_collapse(
            &located.payload,
            &ListStyles(self.working),
            paragraph,
            kind,
            |style| exclusive.is_exclusive(style),
            self.limits,
        )?;
        self.apply(storage, paragraph, plan)?;
        Ok(())
    }

    fn apply(&mut self, storage: u64, paragraph: TextPosition, plan: Plan) -> Result<bool> {
        let located = self::storage(self.working, storage)?;
        match plan {
            Plan::Unchanged => return Ok(false),
            Plan::Rewrite { style, payload } => {
                let object = self.working.object(style)?;
                let index = structure::unique_message(object, &[list::MESSAGE_TYPE])?;
                super::table_cell_text_style::replace_message(self.working, style, index, payload)?;
            },
            Plan::Derive {
                parent,
                stylesheet,
                payload,
            } => {
                let style_entry = self.working.locate(parent)?.ok_or(Fault::Invalid)?;
                let created = *self.working.allocate(1)?.first().ok_or(Fault::Invalid)?;
                let object = self.new_style(created, payload, vec![parent])?;
                // The host's `insert_style_variation`.
                let (_, sheet) = self.stylesheet(stylesheet)?;
                let data = litchi_iwa_text_wire::paragraph_style::stylesheet_with_variation(
                    &sheet,
                    parent,
                    created,
                    self.limits,
                )
                .map_err(|_error| Fault::Invalid)?;
                self.replace_stylesheet(stylesheet, data, created)?;
                self.working.insert(&style_entry, object)?;
                self.register_private(&located.entry, &style_entry, created)?;
                let located = self::storage(self.working, storage)?;
                let old = self.boundary_styles(&located.payload)?;
                let new =
                    list::boundaries_with_style(&located.payload, paragraph, created, self.limits)?;
                let payload = list::replace_boundaries(&located.payload, &new, self.limits)?;
                let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
                self.replace_with_sequence(&located, payload, &old, &new)?;
            },
            Plan::Collapse {
                style,
                parent,
                stylesheet,
            } => {
                let style_entry = self.working.locate(style)?.ok_or(Fault::Invalid)?;
                let old = self.boundary_styles(&located.payload)?;
                let new =
                    list::boundaries_with_style(&located.payload, paragraph, parent, self.limits)?;
                let payload = list::replace_boundaries(&located.payload, &new, self.limits)?;
                let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
                self.replace_with_sequence(&located, payload, &old, &new)?;
                // The host's `remove_style_variation`.
                let (index, sheet) = self.stylesheet(stylesheet)?;
                let data = litchi_iwa_text_wire::paragraph_style::stylesheet_without_variation(
                    &sheet,
                    parent,
                    style,
                    self.limits,
                )
                .map_err(|_error| Fault::Invalid)?;
                super::table_cell_text_style::replace_message(
                    self.working,
                    stylesheet,
                    index,
                    data,
                )?;
                let info = self
                    .working
                    .object_mut(stylesheet)?
                    .archive_info
                    .message_infos
                    .get_mut(index)
                    .ok_or(Fault::Invalid)?;
                info.object_references
                    .retain(|reference| *reference != style);
                for field in &mut info.field_infos {
                    field
                        .object_references
                        .retain(|reference| *reference != style);
                }
                self.working.remove(style)?;
                // The host's `unregister_private_style`, with the parent as
                // the storage component's replacement reference.
                self.transition.remove_objects(self.registry, &[style])?;
                self.register_reference(&located.entry, &style_entry, parent)?;
                self.transition.release_suffix(&[style])?;
            },
            _ => return Err(Fault::Invalid.into()),
        }
        Ok(true)
    }
}

/// Which list styles one storage boundary alone references and nobody
/// names as parent (the host's `style_is_exclusive`).
struct Exclusive {
    references: BTreeMap<u64, usize>,
    parents: Vec<u64>,
}

impl Exclusive {
    fn load(working: &Working<'_>) -> Result<Self> {
        let mut references = BTreeMap::new();
        let mut parents = Vec::new();
        for name in working.component_names() {
            let archive = working.archive(&name)?;
            for object in &archive.objects {
                for message in &object.messages {
                    match message.type_ {
                        STORAGE_MESSAGE_TYPE => {
                            // `table_list_style` entries' objects.
                            let Ok(Some(table)) = single_bytes(&message.data, 7) else {
                                continue;
                            };
                            for entry in repeated_bytes(table, 1).unwrap_or_default() {
                                if let Ok(Some(style)) = single_reference(entry, 2) {
                                    *references.entry(style).or_default() += 1;
                                }
                            }
                        },
                        list::MESSAGE_TYPE => {
                            // `super.parent`.
                            if let Ok(Some(base)) = single_bytes(&message.data, 1)
                                && let Ok(Some(parent)) = single_reference(base, 3)
                            {
                                parents.push(parent);
                            }
                        },
                        _ => {},
                    }
                }
            }
        }
        Ok(Self {
            references,
            parents,
        })
    }

    fn is_exclusive(&self, style: u64) -> bool {
        self.references.get(&style) == Some(&1) && !self.parents.contains(&style)
    }
}
