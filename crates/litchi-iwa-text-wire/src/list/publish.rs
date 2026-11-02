//! Publication of paragraph-list plans on one `TSWP.StorageArchive`
//! (record 0413; the publisher of record 0370b, lifted out of
//! `litchi-numbers`).
//!
//! The planners of this module's parent read and plan; [`Publisher`]
//! publishes their plans as the host text editor's list families do:
//! list-style boundaries are rewritten with their reference sequence,
//! presets and paragraph variations are stored beside their stylesheet and
//! registered there, and the metadata registry gains or loses the matching
//! UUIDs and cross-component references.
//!
//! The package side (objects, archive headers, identifiers and the registry
//! transition) belongs to the leaf format, which implements [`Workspace`]
//! over its structural working set.

use std::collections::BTreeMap;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    append_repeated_length_delimited_field, append_varint_field, parse_wire_view,
};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;

use super::{
    Change, Error as ListError, MESSAGE_TYPE, MESSAGE_VERSION, Plan, PropertyKind, Styles,
    boundaries, boundaries_with_style, find_preset, normalize, placements, plan, plan_collapse,
    preset, preset_payload, replace_boundaries, stylesheet as list_stylesheet, uniform_style,
    with_level, with_numbering,
};

/// `TSWP.StorageArchive`.
pub const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSS.StylesheetArchive`.
pub const STYLESHEET_MESSAGE_TYPE: u32 = 401;

/// The package side of a list publication: a leaf's structural working set,
/// its metadata registry and the registry transition it accumulates.
///
/// Objects are addressed by identifier and, where a message index is
/// needed, by the index of their unique message of a type.
pub trait Workspace {
    /// The leaf's structural error.
    type Error;

    /// The archive member that stores object `identifier`.
    ///
    /// # Errors
    ///
    /// The leaf's error when the object is missing.
    fn entry(&self, identifier: u64) -> Result<String, Self::Error>;

    /// The index and payload of the unique message of `message_type` of
    /// object `identifier`.
    ///
    /// # Errors
    ///
    /// The leaf's error when the object or a unique message is missing.
    fn message(&self, identifier: u64, message_type: u32) -> Result<(usize, &[u8]), Self::Error>;

    /// Replace the payload of message `index` of object `identifier`,
    /// keeping its header.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn replace(
        &mut self,
        identifier: u64,
        index: usize,
        payload: Vec<u8>,
    ) -> Result<(), Self::Error>;

    /// In the header of message `index` of object `identifier` (its
    /// references and every field info naming one of `old`), replace the
    /// references `old` by `new`, in place.
    ///
    /// # Errors
    ///
    /// The leaf's error when `old` is not exactly present.
    fn replace_reference_sequence(
        &mut self,
        identifier: u64,
        index: usize,
        old: &[u64],
        new: &[u64],
    ) -> Result<(), Self::Error>;

    /// Append `reference` to the header references of message `index`.
    ///
    /// # Errors
    ///
    /// The leaf's error when the reference is already present.
    fn append_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> Result<(), Self::Error>;

    /// Remove `reference` from the header of message `index` (its
    /// references and every field info).
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn remove_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> Result<(), Self::Error>;

    /// A fresh object identifier.
    ///
    /// # Errors
    ///
    /// The leaf's error when identifiers are exhausted.
    fn allocate(&mut self) -> Result<u64, Self::Error>;

    /// Store a new one-message object in member `entry`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn insert(
        &mut self,
        entry: &str,
        identifier: u64,
        message_type: u32,
        version: &[u32],
        payload: Vec<u8>,
        references: Vec<u64>,
    ) -> Result<(), Self::Error>;

    /// Remove object `identifier`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn remove(&mut self, identifier: u64) -> Result<(), Self::Error>;

    /// Every object of member `entry` with its message of `message_type`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn entry_messages(
        &self,
        entry: &str,
        message_type: u32,
    ) -> Result<Vec<(u64, &[u8])>, Self::Error>;

    /// Whether any message header of member `entry` references `reference`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn entry_references(&self, entry: &str, reference: u64) -> Result<bool, Self::Error>;

    /// The payload of every message of `message_type` in the package.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn all_messages(&self, message_type: u32) -> Result<Vec<&[u8]>, Self::Error>;

    /// The registry component of member `entry`, when it has one.
    fn component(&self, entry: &str) -> Option<u64>;

    /// Register component `owner`'s reference to `object` in `target`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn add_external_reference(
        &mut self,
        owner: u64,
        target: u64,
        object: u64,
    ) -> Result<(), Self::Error>;

    /// Remove component `owner`'s reference to `object` in `target`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn remove_external_reference(
        &mut self,
        owner: u64,
        target: u64,
        object: u64,
    ) -> Result<(), Self::Error>;

    /// Register a UUID for the new object `object` in component
    /// `component`.
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn register_uuid(&mut self, component: u64, object: u64) -> Result<(), Self::Error>;

    /// Unregister removed `object` (UUID, references, watermark suffix).
    ///
    /// # Errors
    ///
    /// The leaf's error.
    fn unregister(&mut self, object: u64) -> Result<(), Self::Error>;

    /// The wire error that a malformed payload or header maps to.
    fn invalid() -> Self::Error;
}

/// Why a list publication failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Failure<E> {
    /// The leaf workspace failed.
    Workspace(E),
    /// The list engine failed.
    List(ListError),
}

impl<E> From<ListError> for Failure<E> {
    fn from(error: ListError) -> Self {
        Self::List(error)
    }
}

type Outcome<T, E> = Result<T, Failure<E>>;

/// The workspace's list styles, as the planners read them.
pub struct WorkspaceStyles<'w, W>(pub &'w W);

impl<W: Workspace> Styles for WorkspaceStyles<'_, W> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        self.0
            .message(identifier, MESSAGE_TYPE)
            .ok()
            .map(|(_, payload)| payload)
    }
}

/// One located storage: its member, message index and payload.
struct Storage {
    identifier: u64,
    entry: String,
    index: usize,
    payload: Vec<u8>,
}

/// The publisher of one list transaction over a leaf [`Workspace`].
pub struct Publisher<'w, W: Workspace> {
    workspace: &'w mut W,
    limits: WireLimits,
}

fn copy<E>(bytes: &[u8]) -> Outcome<Vec<u8>, E> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| ListError::Allocation {
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

impl<'w, W: Workspace> Publisher<'w, W> {
    /// A publisher over `workspace`, decoding payloads under `limits`.
    pub fn new(workspace: &'w mut W, limits: WireLimits) -> Self {
        Self { workspace, limits }
    }

    fn ws<T>(result: Result<T, W::Error>) -> Outcome<T, W::Error> {
        result.map_err(Failure::Workspace)
    }

    fn storage(&self, identifier: u64) -> Outcome<Storage, W::Error> {
        let entry = Self::ws(self.workspace.entry(identifier))?;
        let (index, payload) = Self::ws(self.workspace.message(identifier, STORAGE_MESSAGE_TYPE))?;
        let payload = copy(payload)?;
        Ok(Storage {
            identifier,
            entry,
            index,
            payload,
        })
    }

    /// The text payload of storage `identifier`.
    ///
    /// # Errors
    ///
    /// The workspace's error.
    pub fn storage_payload(&self, identifier: u64) -> Outcome<Vec<u8>, W::Error> {
        Ok(self.storage(identifier)?.payload)
    }

    fn boundary_styles(&self, payload: &[u8]) -> Outcome<Vec<u64>, W::Error> {
        Ok(boundaries(payload, self.limits)?
            .into_iter()
            .map(|entry| entry.1)
            .collect())
    }

    fn replace_with_sequence(
        &mut self,
        storage: &Storage,
        payload: Vec<u8>,
        old: &[u64],
        new: &[u64],
    ) -> Outcome<(), W::Error> {
        Self::ws(
            self.workspace
                .replace(storage.identifier, storage.index, payload),
        )?;
        Self::ws(self.workspace.replace_reference_sequence(
            storage.identifier,
            storage.index,
            old,
            new,
        ))
    }

    fn stylesheet(&self, stylesheet: u64) -> Outcome<(usize, Vec<u8>), W::Error> {
        let (index, payload) =
            Self::ws(self.workspace.message(stylesheet, STYLESHEET_MESSAGE_TYPE))?;
        Ok((index, copy(payload)?))
    }

    fn replace_stylesheet(
        &mut self,
        stylesheet: u64,
        data: Vec<u8>,
        style: u64,
    ) -> Outcome<(), W::Error> {
        let (index, _) = self.stylesheet(stylesheet)?;
        Self::ws(self.workspace.replace(stylesheet, index, data))?;
        Self::ws(self.workspace.append_reference(stylesheet, index, style))
    }

    /// The host's `register_style_reference`.
    fn register_reference(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        style: u64,
    ) -> Outcome<(), W::Error> {
        if let (Some(owner), Some(target)) = (
            self.workspace.component(owner_entry),
            self.workspace.component(style_entry),
        ) {
            Self::ws(self.workspace.add_external_reference(owner, target, style))?;
        }
        Ok(())
    }

    /// The host's `register_private_style`.
    fn register_private(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        style: u64,
    ) -> Outcome<(), W::Error> {
        let Some(target) = self.workspace.component(style_entry) else {
            return Ok(());
        };
        Self::ws(self.workspace.register_uuid(target, style))?;
        self.register_reference(owner_entry, style_entry, style)
    }

    /// The host's `unregister_owner_reference_if_unused`.
    fn unregister_if_unused(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        style: u64,
    ) -> Outcome<(), W::Error> {
        let (Some(owner), Some(target)) = (
            self.workspace.component(owner_entry),
            self.workspace.component(style_entry),
        ) else {
            return Ok(());
        };
        if owner == target || Self::ws(self.workspace.entry_references(owner_entry, style))? {
            return Ok(());
        }
        Self::ws(
            self.workspace
                .remove_external_reference(owner, target, style),
        )
    }

    /// The stylesheet's preset style for `preset`, stored and registered
    /// when it has none (the host's `preset_style_id` with creation).
    fn preset_object(
        &mut self,
        owner_entry: &str,
        style_entry: &str,
        stylesheet: u64,
        preset: ParagraphList,
    ) -> Outcome<u64, W::Error> {
        let existing = {
            let candidates = Self::ws(self.workspace.entry_messages(style_entry, MESSAGE_TYPE))?;
            find_preset(candidates, stylesheet, preset, self.limits)?
        };
        if let Some(identifier) = existing {
            self.register_reference(owner_entry, style_entry, identifier)?;
            return Ok(identifier);
        }
        let created = Self::ws(self.workspace.allocate())?;
        let payload = preset_payload(stylesheet, preset, self.limits)?;
        // The host's `insert_private_style`: appended to the stylesheet's
        // styles (field 1).
        let (_, sheet) = self.stylesheet(stylesheet)?;
        let mut reference = Vec::new();
        append_varint_field(&mut reference, 1, created)
            .map_err(|_| Failure::Workspace(W::invalid()))?;
        let data = append_repeated_length_delimited_field(&sheet, 1, &reference)
            .map_err(|_| Failure::Workspace(W::invalid()))?;
        self.replace_stylesheet(stylesheet, data, created)?;
        Self::ws(self.workspace.insert(
            style_entry,
            created,
            MESSAGE_TYPE,
            &MESSAGE_VERSION,
            payload,
            Vec::new(),
        ))?;
        self.register_private(owner_entry, style_entry, created)?;
        Ok(created)
    }

    /// Give every paragraph of `storage` the one preset `preset` (the host
    /// text editor's `set_paragraph_list`).
    ///
    /// # Errors
    ///
    /// The list engine's and the workspace's errors.
    pub fn set_list(
        &mut self,
        storage: u64,
        preset_list: ParagraphList,
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        let style = uniform_style(&located.payload, self.limits)?;
        if preset(&WorkspaceStyles(&*self.workspace), style, self.limits).ok() == Some(preset_list)
        {
            return Ok(false);
        }
        let style_entry = Self::ws(self.workspace.entry(style))?;
        let stylesheet = list_stylesheet(&WorkspaceStyles(&*self.workspace), style, self.limits)?;
        let target = self.preset_object(&located.entry, &style_entry, stylesheet, preset_list)?;
        let located = self.storage(storage)?;
        let payload = replace_boundaries(&located.payload, &[(0, target)], self.limits)?;
        self.replace_with_sequence(&located, payload, &[style], &[target])?;
        self.unregister_if_unused(&located.entry, &style_entry, style)?;
        Ok(true)
    }

    /// Give the paragraphs of `storage` the presets of `placements` (the
    /// host text editor's `set_paragraph_lists`).
    ///
    /// # Errors
    ///
    /// The list engine's and the workspace's errors.
    pub fn set_lists(
        &mut self,
        storage: u64,
        requested: &[ParagraphListPlacement],
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        let normalized = normalize(&located.payload, requested, self.limits)?;
        if placements(
            &located.payload,
            &WorkspaceStyles(&*self.workspace),
            self.limits,
        )? == normalized
        {
            return Ok(false);
        }
        let current = boundaries(&located.payload, self.limits)?;
        let first = current
            .first()
            .ok_or_else(|| Failure::Workspace(W::invalid()))?
            .1;
        let style_entry = Self::ws(self.workspace.entry(first))?;
        let stylesheet = list_stylesheet(&WorkspaceStyles(&*self.workspace), first, self.limits)?;
        let slot = |list: ParagraphList| match list {
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
        new.try_reserve_exact(normalized.len())
            .map_err(|_| ListError::Allocation {
                amount: normalized.len(),
            })?;
        for placement in &normalized {
            new.push((
                placement.paragraph.utf16_index(),
                slots[slot(placement.list)].ok_or_else(|| Failure::Workspace(W::invalid()))?,
            ));
        }
        let located = self.storage(storage)?;
        let old: Vec<u64> = current.iter().map(|entry| entry.1).collect();
        let payload = replace_boundaries(&located.payload, &new, self.limits)?;
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
    ///
    /// # Errors
    ///
    /// The list engine's and the workspace's errors.
    pub fn set_level(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        let Some(payload) = with_level(&located.payload, paragraph, level, self.limits)? else {
            return Ok(false);
        };
        Self::ws(
            self.workspace
                .replace(located.identifier, located.index, payload),
        )?;
        Ok(true)
    }

    /// Set the numbering of the paragraph at `paragraph`.
    ///
    /// # Errors
    ///
    /// The list engine's and the workspace's errors.
    pub fn set_numbering(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        let Some(payload) = with_numbering(
            &located.payload,
            &WorkspaceStyles(&*self.workspace),
            paragraph,
            numbering,
            self.limits,
        )?
        else {
            return Ok(false);
        };
        Self::ws(
            self.workspace
                .replace(located.identifier, located.index, payload),
        )?;
        Ok(true)
    }

    /// Set or reset one list property of the paragraph at `paragraph`; a
    /// reset that leaves the paragraph's variation overriding nothing
    /// collapses it onto its parent (the host's reset clean-up).
    ///
    /// # Errors
    ///
    /// The list engine's and the workspace's errors.
    pub fn property(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        change: &Change,
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        let exclusive = Exclusive::load(&*self.workspace)?;
        let planned = plan(
            &located.payload,
            &WorkspaceStyles(&*self.workspace),
            paragraph,
            change,
            |style| exclusive.is_exclusive(style),
            self.limits,
        )?;
        let changed = self.apply(storage, paragraph, planned)?;
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
    ) -> Outcome<(), W::Error> {
        let located = self.storage(storage)?;
        let exclusive = Exclusive::load(&*self.workspace)?;
        let planned = plan_collapse(
            &located.payload,
            &WorkspaceStyles(&*self.workspace),
            paragraph,
            kind,
            |style| exclusive.is_exclusive(style),
            self.limits,
        )?;
        self.apply(storage, paragraph, planned)?;
        Ok(())
    }

    fn apply(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        planned: Plan,
    ) -> Outcome<bool, W::Error> {
        let located = self.storage(storage)?;
        match planned {
            Plan::Unchanged => return Ok(false),
            Plan::Rewrite { style, payload } => {
                let (index, _) = Self::ws(self.workspace.message(style, MESSAGE_TYPE))?;
                Self::ws(self.workspace.replace(style, index, payload))?;
            },
            Plan::Derive {
                parent,
                stylesheet,
                payload,
            } => {
                let style_entry = Self::ws(self.workspace.entry(parent))?;
                let created = Self::ws(self.workspace.allocate())?;
                // The host's `insert_style_variation`.
                let (_, sheet) = self.stylesheet(stylesheet)?;
                let data = crate::paragraph_style::stylesheet_with_variation(
                    &sheet,
                    parent,
                    created,
                    self.limits,
                )
                .map_err(|_| Failure::Workspace(W::invalid()))?;
                self.replace_stylesheet(stylesheet, data, created)?;
                Self::ws(self.workspace.insert(
                    &style_entry,
                    created,
                    MESSAGE_TYPE,
                    &MESSAGE_VERSION,
                    payload,
                    vec![parent],
                ))?;
                self.register_private(&located.entry, &style_entry, created)?;
                let located = self.storage(storage)?;
                let old = self.boundary_styles(&located.payload)?;
                let new = boundaries_with_style(&located.payload, paragraph, created, self.limits)?;
                let payload = replace_boundaries(&located.payload, &new, self.limits)?;
                let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
                self.replace_with_sequence(&located, payload, &old, &new)?;
            },
            Plan::Collapse {
                style,
                parent,
                stylesheet,
            } => {
                let style_entry = Self::ws(self.workspace.entry(style))?;
                let old = self.boundary_styles(&located.payload)?;
                let new = boundaries_with_style(&located.payload, paragraph, parent, self.limits)?;
                let payload = replace_boundaries(&located.payload, &new, self.limits)?;
                let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
                self.replace_with_sequence(&located, payload, &old, &new)?;
                // The host's `remove_style_variation`.
                let (index, sheet) = self.stylesheet(stylesheet)?;
                let data = crate::paragraph_style::stylesheet_without_variation(
                    &sheet,
                    parent,
                    style,
                    self.limits,
                )
                .map_err(|_| Failure::Workspace(W::invalid()))?;
                Self::ws(self.workspace.replace(stylesheet, index, data))?;
                Self::ws(self.workspace.remove_reference(stylesheet, index, style))?;
                Self::ws(self.workspace.remove(style))?;
                // The host's `unregister_private_style`, with the parent as
                // the storage component's replacement reference.
                Self::ws(self.workspace.unregister(style))?;
                self.register_reference(&located.entry, &style_entry, parent)?;
            },
        }
        Ok(true)
    }
}

/// The payload of the unique length-delimited field `number`, if present.
fn single_bytes(data: &[u8], number: u32) -> Option<&[u8]> {
    let view = parse_wire_view(data, WireLimits::DEFAULT).ok()?;
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || found.replace(field.payload()).is_some() {
            return None;
        }
    }
    found
}

/// The reference identifier (`TSP.Reference.identifier`, field 1) of the
/// unique field `number`.
fn single_reference(data: &[u8], number: u32) -> Option<u64> {
    let reference = single_bytes(data, number)?;
    let view = parse_wire_view(reference, WireLimits::DEFAULT).ok()?;
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == 1) {
        if field.wire_type() != 0 || found.is_some() {
            return None;
        }
        let (value, length) =
            litchi_iwa_common::varint::decode_varint_from_bytes(field.payload()).ok()?;
        if length != field.payload().len() {
            return None;
        }
        found = Some(value);
    }
    found
}

/// Which list styles one storage boundary alone references and nobody
/// names as parent (the host's `style_is_exclusive`).
struct Exclusive {
    references: BTreeMap<u64, usize>,
    parents: Vec<u64>,
}

impl Exclusive {
    fn load<W: Workspace>(workspace: &W) -> Outcome<Self, W::Error> {
        let mut references = BTreeMap::new();
        let mut parents = Vec::new();
        for payload in workspace
            .all_messages(STORAGE_MESSAGE_TYPE)
            .map_err(Failure::Workspace)?
        {
            // `table_list_style` entries' objects.
            let Some(table) = single_bytes(payload, 7) else {
                continue;
            };
            let Ok(view) = parse_wire_view(table, WireLimits::DEFAULT) else {
                continue;
            };
            for entry in view.fields().filter(|field| field.number() == 1) {
                if entry.wire_type() != 2 {
                    continue;
                }
                if let Some(style) = single_reference(entry.payload(), 2) {
                    *references.entry(style).or_default() += 1;
                }
            }
        }
        for payload in workspace
            .all_messages(MESSAGE_TYPE)
            .map_err(Failure::Workspace)?
        {
            // `super.parent`.
            if let Some(base) = single_bytes(payload, 1)
                && let Some(parent) = single_reference(base, 3)
            {
                parents.push(parent);
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
