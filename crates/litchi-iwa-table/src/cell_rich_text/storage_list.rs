//! Paragraph lists of one `TSWP.StorageArchive` in a structural working set
//! (records 0370b, 0412).
//!
//! The publication itself is the shared text engine's
//! ([`litchi_iwa_text_wire::list::publish`], record 0413): list-style
//! boundaries are rewritten with their reference sequence, presets and
//! paragraph variations are stored beside their stylesheet and registered
//! there, and the metadata registry gains or loses the matching UUIDs and
//! cross-component references. This module is its package side over the
//! table engines' [`Working`](crate::structure::Working) set: [`Workspace`](crate::cell_rich_text::storage_list::Workspace) maps the publisher's object,
//! header, identifier and registry requests onto the working set and its
//! registry transition. The cell-list engine ([`list`](crate::cell_rich_text::list)) runs it on a
//! cell's rich-text storage, and the Numbers sheet drawables on a text box's
//! storage.

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::publish;
use litchi_iwa_text_wire::list::{self, Change, Styles};

use crate::structure::metadata::{self, Registry, Transition};
use crate::structure::{self, Fault, FaultResult, TableLeaf, Working};

/// `TSWP.StorageArchive`.
pub const STORAGE_MESSAGE_TYPE: u32 = publish::STORAGE_MESSAGE_TYPE;

/// Why a storage-list publication failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Failure {
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

impl From<publish::Failure<Fault>> for Failure {
    fn from(failure: publish::Failure<Fault>) -> Self {
        match failure {
            publish::Failure::Workspace(fault) => Self::Fault(fault),
            publish::Failure::List(error) => Self::List(error),
        }
    }
}

type Result<T> = std::result::Result<T, Failure>;

/// The working set's list styles, as the engine reads them.
pub struct ListStyles<'w, 'p, P>(pub &'w Working<'p, P>);

impl<P: TableLeaf> Styles for ListStyles<'_, '_, P> {
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
pub struct Storage {
    /// The storage object.
    pub identifier: u64,
    /// Its component.
    pub entry: String,
    /// The storage message's index in the object.
    pub index: usize,
    /// The storage payload.
    pub payload: Vec<u8>,
}

/// The storage `identifier`.
///
/// # Errors
///
/// A [`Failure`] for a missing or ambiguous storage.
pub fn storage<P: TableLeaf>(working: &Working<'_, P>, identifier: u64) -> Result<Storage> {
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

/// The package side of the shared list publisher: the working set, the
/// source registry and the registry transition it accumulates.
pub struct Workspace<'a, 'p, P> {
    working: &'a mut Working<'p, P>,
    registry: &'a Registry,
    transition: &'a mut Transition,
}

impl<P: TableLeaf> Workspace<'_, '_, P> {
    fn info(&mut self, identifier: u64, index: usize) -> FaultResult<&mut MessageInfo> {
        self.working
            .object_mut(identifier)?
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Fault::Invalid)
    }
}

/// The host's `replace_reference_sequence` over one reference list.
fn sequence(references: &mut Vec<u64>, old: &[u64], new: &[u64]) -> FaultResult<()> {
    let positions = references
        .iter()
        .filter(|reference| old.contains(reference))
        .count();
    if positions != old.len() {
        return Err(Fault::Invalid);
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

impl<P: TableLeaf> publish::Workspace for Workspace<'_, '_, P> {
    type Error = Fault;

    fn entry(&self, identifier: u64) -> FaultResult<String> {
        self.working.locate(identifier)?.ok_or(Fault::Invalid)
    }

    fn message(&self, identifier: u64, message_type: u32) -> FaultResult<(usize, &[u8])> {
        let object = self.working.object(identifier)?;
        let index = structure::unique_message(object, &[message_type])?;
        Ok((index, object.messages[index].data.as_slice()))
    }

    fn replace(&mut self, identifier: u64, index: usize, payload: Vec<u8>) -> FaultResult<()> {
        let limits = self.working.archive_limits();
        let object = self.working.object_mut(identifier)?;
        let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: message_type,
                    data: payload,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    }

    fn replace_reference_sequence(
        &mut self,
        identifier: u64,
        index: usize,
        old: &[u64],
        new: &[u64],
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
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

    fn append_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
        if info.object_references.contains(&reference) {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(reference);
        Ok(())
    }

    fn remove_reference(
        &mut self,
        identifier: u64,
        index: usize,
        reference: u64,
    ) -> FaultResult<()> {
        let info = self.info(identifier, index)?;
        info.object_references.retain(|item| *item != reference);
        for field in &mut info.field_infos {
            field.object_references.retain(|item| *item != reference);
        }
        Ok(())
    }

    fn allocate(&mut self) -> FaultResult<u64> {
        self.working
            .allocate(1)?
            .first()
            .copied()
            .ok_or(Fault::Invalid)
    }

    fn insert(
        &mut self,
        entry: &str,
        identifier: u64,
        message_type: u32,
        version: &[u32],
        payload: Vec<u8>,
        references: Vec<u64>,
    ) -> FaultResult<()> {
        let mut object: ArchiveObject = structure::new_object(
            identifier,
            None,
            message_type,
            payload,
            references,
            self.working.archive_limits(),
        )?;
        object.archive_info.message_infos[0].versions = version.to_vec().into();
        self.working.insert(entry, object)
    }

    fn remove(&mut self, identifier: u64) -> FaultResult<()> {
        self.working.remove(identifier).map(drop)
    }

    fn entry_messages(&self, entry: &str, message_type: u32) -> FaultResult<Vec<(u64, &[u8])>> {
        let archive = self.working.archive(entry)?;
        let mut output = Vec::new();
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == message_type)
            {
                structure::reserve(&mut output, 1)?;
                output.push((identifier, message.data.as_slice()));
            }
        }
        Ok(output)
    }

    fn entry_references(&self, entry: &str, reference: u64) -> FaultResult<bool> {
        let archive = self.working.archive(entry)?;
        Ok(archive.objects.iter().any(|object| {
            object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&reference))
        }))
    }

    fn all_messages(&self, message_type: u32) -> FaultResult<Vec<&[u8]>> {
        let mut output = Vec::new();
        for name in self.working.component_names() {
            let archive = self.working.archive(&name)?;
            for object in &archive.objects {
                for message in &object.messages {
                    if message.type_ == message_type {
                        structure::reserve(&mut output, 1)?;
                        output.push(message.data.as_slice());
                    }
                }
            }
        }
        Ok(output)
    }

    fn component(&self, entry: &str) -> Option<u64> {
        self.registry
            .component_for_entry(entry)
            .ok()
            .map(|component| component.identifier)
    }

    fn add_external_reference(&mut self, owner: u64, target: u64, object: u64) -> FaultResult<()> {
        self.transition
            .reference(self.registry, owner, target, object)
    }

    fn remove_external_reference(
        &mut self,
        owner: u64,
        target: u64,
        object: u64,
    ) -> FaultResult<()> {
        self.transition
            .remove_reference(self.registry, owner, target, object)
    }

    fn register_uuid(&mut self, component: u64, object: u64) -> FaultResult<()> {
        let seed = metadata::identity_seed(self.working, object)?;
        self.transition
            .register_uuid(self.registry, component, object, seed)
    }

    fn unregister(&mut self, object: u64) -> FaultResult<()> {
        self.transition.remove_objects(self.registry, &[object])?;
        self.transition.release_suffix(&[object])
    }

    fn invalid() -> Fault {
        Fault::Invalid
    }
}

/// The publisher of one transaction: the working set, the source registry
/// and the registry transition it accumulates. Each operation runs the
/// shared publisher over a [`Workspace`] of these.
pub struct Publisher<'a, 'p, P> {
    /// The working set.
    pub working: &'a mut Working<'p, P>,
    /// The source package metadata registry.
    pub registry: &'a Registry,
    /// The registry transition the publication accumulates.
    pub transition: &'a mut Transition,
    /// The shared text engines' wire profile.
    pub limits: WireLimits,
}

impl<P: TableLeaf> Publisher<'_, '_, P> {
    fn run<T>(
        &mut self,
        operation: impl FnOnce(
            &mut publish::Publisher<'_, Workspace<'_, '_, P>>,
        ) -> std::result::Result<T, publish::Failure<Fault>>,
    ) -> Result<T> {
        let mut workspace = Workspace {
            working: self.working,
            registry: self.registry,
            transition: self.transition,
        };
        let mut publisher = publish::Publisher::new(&mut workspace, self.limits);
        Ok(operation(&mut publisher)?)
    }

    /// Give every paragraph of `storage` the one preset `preset` (the host
    /// text editor's `set_paragraph_list`).
    ///
    /// # Errors
    ///
    /// A [`Failure`] of the working set or the list engine.
    pub fn set_list(&mut self, storage: u64, preset: ParagraphList) -> Result<bool> {
        self.run(|publisher| publisher.set_list(storage, preset))
    }

    /// Give the paragraphs of `storage` the presets of `placements` (the
    /// host text editor's `set_paragraph_lists`).
    ///
    /// # Errors
    ///
    /// A [`Failure`] of the working set or the list engine.
    pub fn set_lists(
        &mut self,
        storage: u64,
        placements: &[ParagraphListPlacement],
    ) -> Result<bool> {
        self.run(|publisher| publisher.set_lists(storage, placements))
    }

    /// Set the list level of the paragraph at `paragraph`.
    ///
    /// # Errors
    ///
    /// A [`Failure`] of the working set or the list engine.
    pub fn set_level(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<bool> {
        self.run(|publisher| publisher.set_level(storage, paragraph, level))
    }

    /// Set the numbering of the paragraph at `paragraph`.
    ///
    /// # Errors
    ///
    /// A [`Failure`] of the working set or the list engine.
    pub fn set_numbering(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<bool> {
        self.run(|publisher| publisher.set_numbering(storage, paragraph, numbering))
    }

    /// Set or reset one list property of the paragraph at `paragraph`; a
    /// reset that leaves the paragraph's variation overriding nothing
    /// collapses it onto its parent (the host's reset clean-up).
    ///
    /// # Errors
    ///
    /// A [`Failure`] of the working set or the list engine.
    pub fn property(
        &mut self,
        storage: u64,
        paragraph: TextPosition,
        change: &Change,
    ) -> Result<bool> {
        self.run(|publisher| publisher.property(storage, paragraph, change))
    }
}
