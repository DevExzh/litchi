//! The structural working set as the package side of the shared
//! storage-list publisher (`litchi_iwa_text_wire::list::publish`, record
//! 0413). Pages and Numbers publish their text-box and cell lists through
//! it.

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text_wire::list::publish::Workspace;
use litchi_iwa_text_wire::list::{MESSAGE_TYPE, Styles};

use super::metadata::{Registry, Transition};
use super::{Fault, FaultResult, TableLeaf, Working};

/// The list styles of a working set, as the read planners see them.
pub struct ReadStyles<'w, 'p, P: TableLeaf>(pub &'w Working<'p, P>);

impl<P: TableLeaf> Styles for ReadStyles<'_, '_, P> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        let object = self.0.object(identifier).ok()?;
        object
            .messages
            .iter()
            .find(|message| message.type_ == MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

/// The working-set side of the shared list publisher: the structural working
/// set, the source registry and the registry transition it accumulates.
pub struct ListWorkspace<'a, 'p, P: TableLeaf> {
    /// The working set.
    pub working: &'a mut Working<'p, P>,
    /// The source registry.
    pub registry: &'a Registry,
    /// The registry transition the publication accumulates.
    pub transition: &'a mut Transition,
}

impl<P: TableLeaf> ListWorkspace<'_, '_, P> {
    fn info(
        &mut self,
        identifier: u64,
        index: usize,
    ) -> FaultResult<&mut litchi_iwa_core::MessageInfo> {
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
    super::reserve(references, new.len())?;
    references.splice(insertion..insertion, new.iter().copied());
    Ok(())
}

impl<P: TableLeaf> Workspace for ListWorkspace<'_, '_, P> {
    type Error = Fault;

    fn entry(&self, identifier: u64) -> FaultResult<String> {
        self.working.locate(identifier)?.ok_or(Fault::Invalid)
    }

    fn message(&self, identifier: u64, message_type: u32) -> FaultResult<(usize, &[u8])> {
        let object = self.working.object(identifier)?;
        let index = super::unique_message(object, &[message_type])?;
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
            .map_err(|error| super::map_core(&error))
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
        super::reserve(&mut info.object_references, 1)?;
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
        let mut object: ArchiveObject = super::new_object(
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
        Ok(archive
            .objects
            .iter()
            .filter_map(|object| {
                let identifier = object.archive_info.identifier?;
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == message_type)
                    .map(|message| (identifier, message.data.as_slice()))
            })
            .collect())
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
                        super::reserve(&mut output, 1)?;
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
        let seed = super::metadata::identity_seed(self.working, object)?;
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

/// The `TSWP.StorageArchive` payload of storage `identifier`.
///
/// # Errors
///
/// [`Fault`] when the object or its unique storage message is missing.
pub fn storage_payload<P: TableLeaf>(
    working: &Working<'_, P>,
    identifier: u64,
) -> FaultResult<Vec<u8>> {
    let object = working.object(identifier)?;
    let index = super::unique_message(
        object,
        &[litchi_iwa_text_wire::list::publish::STORAGE_MESSAGE_TYPE],
    )?;
    super::copy_bytes(&object.messages[index].data)
}
