//! The comment-storage graph the text annotation and drawable comment
//! families share (record 0378): the generated annotation author, comment
//! storage creation and removal, and their package-metadata registrations,
//! as the host editor's `comments` module keeps them.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::comment::{AuthorId, Uuid};
use litchi_iwa_core::{ArchiveObject, FieldInfo, FieldPath, RawMessage, UnknownFieldRule};
use litchi_iwa_text::date_time::Instant;
use litchi_iwa_text_wire::annotation::{self as engine, COMMENT_STORAGE_MESSAGE_TYPE};

use super::structure::{
    self, Fault, FaultResult, Working,
    metadata::{self, Registry, Transition},
    references,
};

/// `TSK.AnnotationAuthorArchive`.
const AUTHOR_MESSAGE_TYPE: u32 = 212;
/// `TSK.AnnotationAuthorStorageArchive`.
const AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
/// The generated author's name and public identifier (the host editor's).
const AUTHOR_NAME: &str = "litchi-iwa";
const AUTHOR_PUBLIC_ID: &str = "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6";
const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;

/// The fault an annotation-engine failure maps to.
pub(super) const fn engine_fault(error: &engine::Error) -> Fault {
    match error {
        engine::Error::Allocation { amount } => Fault::Allocation { amount: *amount },
        _ => Fault::Invalid,
    }
}

/// The decoded comment storage of `identifier`.
pub(super) fn comment_storage(
    working: &Working<'_>,
    identifier: u64,
) -> FaultResult<(usize, engine::CommentStorage)> {
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, &[COMMENT_STORAGE_MESSAGE_TYPE])?;
    let storage = engine::comment_storage(&object.messages[index].data, WireLimits::DEFAULT)
        .map_err(|error| engine_fault(&error))?;
    Ok((index, storage))
}

/// The generated author's `TSK.AnnotationAuthorArchive` payload (the host
/// editor's `generated_author`, in its encoder's field order).
fn generated_author(public: bool) -> Vec<u8> {
    let mut color = Vec::new();
    let float = |out: &mut Vec<u8>, number: u8, value: f32| {
        out.push(number << 3 | 5);
        out.extend_from_slice(&value.to_le_bytes());
    };
    color.extend_from_slice(&[0x08, 0x01]);
    float(&mut color, 3, 0.368_627_46);
    float(&mut color, 4, 0.568_627_5);
    float(&mut color, 5, 0.937_254_9);
    float(&mut color, 6, 1.0);
    color.extend_from_slice(&[0x60, 0x01]);
    let mut output = Vec::new();
    let bytes = |out: &mut Vec<u8>, number: u8, payload: &[u8]| {
        out.push(number << 3 | 2);
        let mut length = payload.len();
        loop {
            #[allow(clippy::cast_possible_truncation, reason = "seven-bit groups")]
            let byte = (length & 0x7f) as u8;
            length >>= 7;
            if length == 0 {
                out.push(byte);
                break;
            }
            out.push(byte | 0x80);
        }
        out.extend_from_slice(payload);
    };
    bytes(&mut output, 1, AUTHOR_NAME.as_bytes());
    bytes(&mut output, 2, &color);
    if public {
        bytes(&mut output, 3, AUTHOR_PUBLIC_ID.as_bytes());
    }
    output.extend_from_slice(&[0x20, 0x00]);
    if public {
        bytes(&mut output, 5, AUTHOR_PUBLIC_ID.as_bytes());
    }
    output
}

/// The document's author storage, if any.
fn author_storage(working: &Working<'_>) -> FaultResult<Option<u64>> {
    let mut found = None;
    for object in working.source().state.components.iter_objects() {
        if object
            .messages
            .iter()
            .any(|message| message.type_ == AUTHOR_STORAGE_MESSAGE_TYPE)
            && found
                .replace(object.archive_info.identifier.ok_or(Fault::Invalid)?)
                .is_some()
        {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

fn author_payload(working: &Working<'_>, author: u64) -> FaultResult<Vec<u8>> {
    let object = working.object(author)?;
    let index = structure::unique_message(object, &[AUTHOR_MESSAGE_TYPE])?;
    Ok(object.messages[index].data.clone())
}

fn is_generated(working: &Working<'_>, author: u64) -> FaultResult<bool> {
    let payload = author_payload(working, author)?;
    Ok(payload == generated_author(true) || payload == generated_author(false))
}

/// Replace message `index` of `identifier` canonically and adjust its
/// aggregate references.
pub(super) fn replace(
    working: &mut Working<'_>,
    identifier: u64,
    index: usize,
    data: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(identifier)?;
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    let mut staged = object.clone();
    staged
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    let info = staged
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Fault::Invalid)?;
    if let Some(new) = added {
        if info.object_references.contains(&new) {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut info.object_references, 1)?;
        info.object_references.push(new);
    }
    if let Some(old) = removed {
        if info
            .object_references
            .iter()
            .filter(|reference| **reference == old)
            .count()
            != 1
        {
            return Err(Fault::Invalid);
        }
        info.object_references.retain(|reference| *reference != old);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != old);
        }
    }
    staged
        .validate(limits)
        .map_err(|error| structure::map_core(&error))?;
    *working.object_mut(identifier)? = staged;
    Ok(())
}

/// The current instant, for a new comment storage's creation date.
pub(super) fn now() -> FaultResult<Instant> {
    let unix = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_err(|_error| Fault::Invalid)?
        .as_secs_f64();
    Instant::from_reference_date_seconds(unix - APPLE_EPOCH_UNIX_OFFSET_SECONDS)
        .map_err(|_error| Fault::Invalid)
}

/// One transaction's comment graph: the working set, its registry and the
/// metadata transition being built.
pub(super) struct Graph<'w, 'p> {
    pub(super) working: &'w mut Working<'p>,
    pub(super) registry: Registry,
    pub(super) transition: Transition,
}

impl Graph<'_, '_> {
    /// Every comment-storage payload of the working set, including the
    /// objects this transaction inserted.
    fn comment_storages(&self) -> FaultResult<Vec<&[u8]>> {
        let mut found = Vec::new();
        for name in self.working.component_names() {
            for object in &self.working.archive(&name)?.objects {
                for message in &object.messages {
                    if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
                        structure::reserve(&mut found, 1)?;
                        found.push(message.data.as_slice());
                    }
                }
            }
        }
        Ok(found)
    }

    pub(super) fn component(&self, member: &str) -> FaultResult<u64> {
        Ok(self.registry.component_for_entry(member)?.identifier)
    }

    /// The registry component of the member that stores `object`.
    pub(super) fn component_of(&self, object: u64) -> FaultResult<u64> {
        let member = self.working.locate(object)?.ok_or(Fault::Invalid)?;
        self.component(&member)
    }

    pub(super) fn register_uuid(&mut self, member: &str, object: u64) -> FaultResult<()> {
        let component = self.component(member)?;
        let seed = metadata::identity_seed(self.working, object)?;
        self.transition
            .register_uuid(&self.registry, component, object, seed)
    }

    /// The host's `ensure_annotation_author` (public generated author): the
    /// author, and whether it was created.
    pub(super) fn ensure_author(&mut self) -> FaultResult<Option<(u64, bool)>> {
        let Some(storage) = author_storage(self.working)? else {
            return Ok(None);
        };
        let object = self.working.object(storage)?;
        let index = structure::unique_message(object, &[AUTHOR_STORAGE_MESSAGE_TYPE])?;
        let payload = object.messages[index].data.clone();
        for record in references::records(&payload, &[1])? {
            let author = references::reference_identifier(record)?;
            if author_payload(self.working, author)? == generated_author(true) {
                return Ok(Some((author, false)));
            }
        }
        let author = *self.working.allocate(1)?.first().ok_or(Fault::Invalid)?;
        let mut data = payload;
        data.extend_from_slice(&[0x0a]);
        let reference = references::reference_payload(author)?;
        data.push(u8::try_from(reference.len()).map_err(|_error| Fault::Invalid)?);
        data.extend_from_slice(&reference);
        replace(self.working, storage, index, data, None, None)?;
        let limits = self.working.archive_limits();
        let mut created = ArchiveObject::new(
            author,
            vec![RawMessage {
                type_: AUTHOR_MESSAGE_TYPE,
                data: generated_author(true),
            }],
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
        for field in [4_u32, 3] {
            let mut info = FieldInfo::new(
                FieldPath::new(vec![field], limits).map_err(|_| Fault::Invalid)?,
                limits,
            )
            .map_err(|_| Fault::Invalid)?;
            info.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserve);
            created.archive_info.message_infos[0].field_infos.push(info);
        }
        let member = self.working.locate(storage)?.ok_or(Fault::Invalid)?;
        self.working.insert(&member, created)?;
        self.register_uuid(&member, author)?;
        Ok(Some((author, true)))
    }

    /// The member of the document's author storage.
    pub(super) fn author_member(&self) -> FaultResult<Option<String>> {
        match author_storage(self.working)? {
            Some(storage) => self.working.locate(storage),
            None => Ok(None),
        }
    }

    /// Register that `member`'s component references `author`.
    pub(super) fn add_author_reference(
        &mut self,
        member: &str,
        author: Option<u64>,
    ) -> FaultResult<()> {
        let Some(author) = author else {
            return Ok(());
        };
        let source = self.component(member)?;
        let target = self.component_of(author)?;
        self.transition
            .reference(&self.registry, source, target, author)
    }

    /// Remove object `identifier` with its registrations (the host's
    /// `remove_registered_object` and `remove_object`).
    pub(super) fn forget(&mut self, identifier: u64) -> FaultResult<()> {
        self.transition
            .remove_objects(&self.registry, &[identifier])?;
        self.transition.release_suffix(&[identifier])?;
        self.working.remove(identifier)?;
        Ok(())
    }

    /// The host's `remove_generated_annotation_author_if_unused`: whether
    /// the author was removed.
    pub(super) fn remove_author_if_unused(&mut self, author: u64) -> FaultResult<bool> {
        if !is_generated(self.working, author)? {
            return Ok(false);
        }
        for payload in self.comment_storages()? {
            if engine::comment_storage(payload, WireLimits::DEFAULT)
                .map_err(|error| engine_fault(&error))?
                .metadata()
                .author()
                .map(AuthorId::get)
                == Some(author)
            {
                return Ok(false);
            }
        }
        let storage = author_storage(self.working)?.ok_or(Fault::Invalid)?;
        let object = self.working.object(storage)?;
        let index = structure::unique_message(object, &[AUTHOR_STORAGE_MESSAGE_TYPE])?;
        let data = litchi_iwa_common::wire::remove_repeated_length_delimited_field_where(
            &object.messages[index].data,
            1,
            |payload| {
                Ok::<bool, litchi_iwa_common::Error>(
                    references::reference_identifier(payload).is_ok_and(|id| id == author),
                )
            },
        )
        .map_err(|error| structure::map_wire(&error))?;
        replace(self.working, storage, index, data, None, None)?;
        self.forget(author)?;
        Ok(true)
    }

    /// Insert a new comment storage `identifier` into `member`.
    pub(super) fn insert_comment_storage(
        &mut self,
        member: &str,
        identifier: u64,
        body: &str,
        author: Option<u64>,
    ) -> FaultResult<()> {
        let uuid = self.fresh_uuid(identifier)?;
        let payload = engine::encode_comment_storage(
            body,
            now()?,
            author
                .map(AuthorId::from_raw)
                .transpose()
                .map_err(|_| Fault::Invalid)?,
            uuid,
            WireLimits::DEFAULT,
        )
        .map_err(|error| engine_fault(&error))?;
        let limits = self.working.archive_limits();
        let mut object = ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data: payload,
            }],
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
        if let Some(author) = author {
            object.archive_info.message_infos[0]
                .object_references
                .push(author);
        }
        self.working.insert(member, object)?;
        self.register_uuid(member, identifier)
    }

    /// A deterministic comment-storage UUID no comment storage uses yet.
    pub(super) fn fresh_uuid(&self, seed: u64) -> FaultResult<Uuid> {
        let mut used = Vec::new();
        for payload in self.comment_storages()? {
            if let Ok(decoded) = engine::comment_storage(payload, WireLimits::DEFAULT) {
                structure::reserve(&mut used, 1)?;
                used.push(decoded.metadata().storage_uuid());
            }
        }
        let mut lower = metadata::mix(seed ^ 0x6c69_7463_6869_0378);
        loop {
            let uuid =
                Uuid::from_parts(lower, 0x6c69_7463_6869_7577).map_err(|_error| Fault::Invalid)?;
            if !used.contains(&uuid) {
                return Ok(uuid);
            }
            lower = lower.wrapping_add(1);
        }
    }
}
