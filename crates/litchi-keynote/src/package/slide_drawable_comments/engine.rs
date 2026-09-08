//! Source preserving mutation of Keynote drawable comment threads.
//!
//! This module is deliberately a transaction engine rather than a facade.  It
//! receives a checked semantic selection and one already bounded operation,
//! changes only the component archives required by that operation, and
//! returns a candidate ZIP.  The coordinator is responsible for reopening the
//! candidate and publishing a commit.  Native identifiers therefore remain
//! entirely inside this private module and its graph/metadata siblings.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_arguments,
    clippy::too_many_lines,
    reason = "The private transaction keeps admission, graph ownership, and candidate staging together."
)]

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::mem::size_of;
use std::time::{SystemTime, UNIX_EPOCH};

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::wire::{NestedFieldEdit, NestedFieldReplacement};
use litchi_iwa_common::{WireLimits, wire};
use litchi_iwa_core::archive::{
    ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, FieldObjectReferenceTransition, ObjectReferenceTransition,
};
use litchi_iwa_core::{Archive, ArchiveLimits, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::comment_storage_codec::{
    self, CommentStorageReplyRewrite, DateSnapshot, DecodeOptions, ReferenceRecord,
    RewriteExecutionRequirements, UuidSnapshot,
};

use super::super::{Package, PhysicalSource};
use super::{Budget, Error, Operation};
use super::{graph, metadata};
use graph::Selection;

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const COMMENT_TEXT_FIELD: u32 = 1;
const COMMENT_UUID_FIELD: u32 = 5;
const COMMENT_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;
const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;
const COMMENT_RECURSION_LIMIT: u32 = 64;
const MAX_RESERVED_IDENTIFIERS: usize = 3;

/// The only mutable state owned by the engine.
///
/// A component is copied into this map at most once.  Unselected components
/// stay borrowed from the immutable source, so a no-op or a small comment
/// edit never materializes the complete package in memory.
pub(super) struct WorkingSet<'source> {
    source: &'source Package,
    archives: BTreeMap<String, Archive>,
    reserved_identifiers: Vec<u64>,
}

impl<'source> WorkingSet<'source> {
    pub(super) fn new(source: &'source Package) -> Self {
        Self {
            source,
            archives: BTreeMap::new(),
            reserved_identifiers: Vec::new(),
        }
    }

    pub(super) const fn source(&self) -> &'source Package {
        self.source
    }

    pub(super) fn get(&self, name: &str) -> Option<&Archive> {
        self.archives.get(name)
    }

    pub(super) fn reserve_identifier(
        &mut self,
        identifier: u64,
        budget: &mut Budget,
    ) -> Result<bool, Error> {
        if identifier == 0 || self.reserved_identifiers.contains(&identifier) {
            return Ok(false);
        }
        if self.reserved_identifiers.len() >= MAX_RESERVED_IDENTIFIERS {
            return Err(Error::InvalidSource);
        }
        if self.reserved_identifiers.is_empty() {
            let amount = MAX_RESERVED_IDENTIFIERS
                .checked_mul(size_of::<u64>())
                .ok_or(Error::InvalidSource)?;
            budget.charge_allocations(amount)?;
            self.reserved_identifiers
                .try_reserve_exact(MAX_RESERVED_IDENTIFIERS)
                .map_err(|_| Error::Allocation { amount })?;
        }
        self.reserved_identifiers.push(identifier);
        Ok(true)
    }

    pub(super) fn load(&mut self, name: &str, budget: &mut Budget) -> Result<&mut Archive, Error> {
        if !self.archives.contains_key(name) {
            // Establish the complete admission charge before allocating or
            // cloning any source state.  `encoded` accounts for payload and
            // nested message buffers; the object header charge covers the
            // per-object clone allocation and the name charge covers the map
            // key copied into the candidate.
            let (object_count, clone_bytes) = {
                let component = self
                    .source
                    .state
                    .source
                    .components()
                    .get(name)
                    .ok_or(Error::InvalidSource)?;
                let archive = component.archive();
                let encoded = archive
                    .encoded_len_with_limits(
                        self.source
                            .limits()
                            .effective_archive_limits()
                            .map_err(|_| Error::InvalidSource)?,
                    )
                    .map_err(|_| Error::InvalidSource)?;
                (
                    archive.objects.len(),
                    archive_clone_allocation_bound(archive, encoded)?,
                )
            };
            budget.charge_allocation_plan(clone_bytes, 2)?;
            let cloned = {
                let component = self
                    .source
                    .state
                    .source
                    .components()
                    .get(name)
                    .ok_or(Error::InvalidSource)?;
                let archive = component.archive();
                let mut cloned = Archive::new();
                let object_bytes = object_count
                    .checked_mul(size_of::<ArchiveObject>())
                    .ok_or(Error::InvalidSource)?;
                cloned
                    .objects
                    .try_reserve_exact(object_count)
                    .map_err(|_| Error::Allocation {
                        amount: object_bytes,
                    })?;
                cloned.objects.extend(archive.objects.iter().cloned());
                cloned
            };
            let owned_name = copy_component_name(name, budget)?;
            self.archives.insert(owned_name, cloned);
        }
        self.archives.get_mut(name).ok_or(Error::InvalidSource)
    }

    fn changed_names(&self) -> impl Iterator<Item = &str> {
        self.archives.keys().map(String::as_str)
    }

    /// Visit each effective component once, with changed archives taking
    /// precedence over the immutable source archive of the same name.
    pub(super) fn for_each_archive(
        &self,
        mut visit: impl FnMut(&str, &Archive) -> Result<(), Error>,
    ) -> Result<(), Error> {
        for component in self.source.state.source.components().iter() {
            let name = component.name();
            if let Some(archive) = self.archives.get(name) {
                visit(name, archive)?;
            } else {
                visit(name, component.archive())?;
            }
        }
        for (name, archive) in &self.archives {
            if self.source.state.source.components().get(name).is_none() {
                visit(name, archive)?;
            }
        }
        Ok(())
    }

    /// Find an object in the effective candidate view without copying its
    /// payload.  This is used by metadata planning and graph cleanup.
    pub(super) fn find_object(&self, identifier: u64) -> Option<(&str, &ArchiveObject)> {
        if let Some((source_name, _)) = self.source.object_with_component(identifier) {
            if let Some(archive) = self.archives.get(source_name) {
                return archive
                    .object(identifier)
                    .map(|object| (source_name, object));
            }
            return self.source.object_with_component(identifier);
        }
        for (name, archive) in &self.archives {
            if let Some(object) = archive.object(identifier) {
                return Some((name.as_str(), object));
            }
        }
        None
    }
}

/// Bound the owned state created by `Archive::clone`.  The encoded length
/// covers the source payload/header bytes, while the structural terms account
/// for the independent Rust collections that `Clone` creates for every
/// object, message, metadata record, field, reference, and preserved header.
/// This is intentionally conservative because the core clone implementation
/// owns those allocations internally and cannot debit our transaction ledger.
fn archive_clone_allocation_bound(archive: &Archive, encoded_bytes: usize) -> Result<usize, Error> {
    let mut bytes = encoded_bytes
        .checked_add(size_of::<Archive>())
        .and_then(|bytes| bytes.checked_add(size_of::<(String, Archive)>()))
        .ok_or(Error::InvalidSource)?;
    bytes = bytes
        .checked_add(
            archive
                .objects
                .len()
                .checked_mul(size_of::<ArchiveObject>())
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    for object in &archive.objects {
        bytes = bytes
            .checked_add(
                object
                    .messages
                    .len()
                    .checked_mul(size_of::<RawMessage>())
                    .ok_or(Error::InvalidSource)?,
            )
            .and_then(|bytes| {
                bytes.checked_add(
                    object
                        .archive_info
                        .message_infos
                        .len()
                        .checked_mul(size_of::<litchi_iwa_core::MessageInfo>())?,
                )
            })
            .ok_or(Error::InvalidSource)?;
        bytes = bytes
            .checked_add(
                usize::try_from(object.header_length)
                    .map_err(|_| Error::InvalidSource)?
                    .checked_mul(2)
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        for message in &object.messages {
            bytes = bytes
                .checked_add(message.data.len())
                .ok_or(Error::InvalidSource)?;
        }
        for info in &object.archive_info.message_infos {
            bytes = bytes
                .checked_add(
                    info.versions
                        .len()
                        .checked_mul(size_of::<u32>())
                        .and_then(|value| {
                            value.checked_add(
                                info.diff_merge_version
                                    .len()
                                    .checked_mul(size_of::<u32>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                info.diff_read_version.len().checked_mul(size_of::<u32>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                info.object_references.len().checked_mul(size_of::<u64>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                info.data_references.len().checked_mul(size_of::<u64>())?,
                            )
                        })
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?;
            if let Some(path) = &info.diff_field_path {
                bytes = bytes
                    .checked_add(
                        path.path
                            .len()
                            .checked_mul(size_of::<u32>())
                            .ok_or(Error::InvalidSource)?,
                    )
                    .ok_or(Error::InvalidSource)?;
            }
            for path in &info.fields_to_remove {
                bytes = bytes
                    .checked_add(size_of::<litchi_iwa_core::FieldPath>())
                    .and_then(|value| {
                        value.checked_add(path.path.len().checked_mul(size_of::<u32>())?)
                    })
                    .ok_or(Error::InvalidSource)?;
            }
            bytes = bytes
                .checked_add(
                    info.field_infos
                        .len()
                        .checked_mul(size_of::<litchi_iwa_core::FieldInfo>())
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?;
            for field in &info.field_infos {
                bytes = bytes
                    .checked_add(
                        field
                            .path
                            .path
                            .len()
                            .checked_mul(size_of::<u32>())
                            .and_then(|value| {
                                value.checked_add(
                                    field
                                        .object_references
                                        .len()
                                        .checked_mul(size_of::<u64>())?,
                                )
                            })
                            .and_then(|value| {
                                value.checked_add(
                                    field.data_references.len().checked_mul(size_of::<u64>())?,
                                )
                            })
                            .and_then(|value| {
                                value.checked_add(
                                    field
                                        .known_field_version
                                        .len()
                                        .checked_mul(size_of::<u32>())?,
                                )
                            })
                            .and_then(|value| {
                                value.checked_add(
                                    field
                                        .known_field_feature_identifier
                                        .as_ref()
                                        .map_or(0, String::len),
                                )
                            })
                            .ok_or(Error::InvalidSource)?,
                    )
                    .ok_or(Error::InvalidSource)?;
            }
        }
    }
    Ok(bytes)
}

/// Admit the source-sized arenas used by the core header-preserving mutation
/// primitives.  Their internal buffers are intentionally private to the core
/// crate, so this mirrors the same source-header/metadata bound used by the
/// lifecycle owner and narrows the limits passed to the core call.
fn core_header_limits(
    object: &ArchiveObject,
    additional_references: usize,
    limits: ArchiveLimits,
    budget: &mut Budget,
) -> Result<ArchiveLimits, Error> {
    let source_bytes = usize::try_from(object.header_length)
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    let mut variable_fields = additional_references
        .checked_add(2)
        .ok_or(Error::InvalidSource)?;
    let mut containers = object
        .archive_info
        .message_infos
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    for info in &object.archive_info.message_infos {
        variable_fields = variable_fields
            .checked_add(info.object_references.len())
            .and_then(|value| value.checked_add(2))
            .ok_or(Error::InvalidSource)?;
        containers = containers
            .checked_add(info.field_infos.len())
            .ok_or(Error::InvalidSource)?;
        for field in &info.field_infos {
            variable_fields = variable_fields
                .checked_add(field.object_references.len())
                .ok_or(Error::InvalidSource)?;
        }
    }
    let growth = variable_fields
        .checked_add(containers)
        .and_then(|value| value.checked_mul(10))
        .ok_or(Error::InvalidSource)?;
    let header_bytes = source_bytes
        .checked_add(growth)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_header_bytes())
        .max(1);
    let header_memory = header_bytes
        .checked_mul(256)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_header_memory_bytes())
        .max(1);
    let narrowed = limits
        .with_header_bytes(header_bytes)
        .and_then(|value| value.with_header_fields(header_bytes.min(limits.max_header_fields())))
        .and_then(|value| value.with_metadata_items(header_bytes.min(limits.max_metadata_items())))
        .and_then(|value| value.with_header_memory_bytes(header_memory))
        .map_err(|_| Error::InvalidSource)?;
    let bytes = header_memory
        .checked_mul(4)
        .and_then(|value| value.checked_add(header_bytes.checked_mul(16)?))
        .ok_or(Error::InvalidSource)?;
    let events = containers
        .checked_mul(64)
        .and_then(|value| value.checked_add(32))
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocation_plan(bytes, events)?;
    budget.charge_wire_work(bytes)?;
    Ok(narrowed)
}

/// Bound the independent vectors and payload buffers created by one core
/// object clone.  Header scratch is admitted separately by
/// [`core_header_limits`].
fn object_clone_allocation_bound(
    object: &ArchiveObject,
    replacement_messages: &[RawMessage],
) -> Result<usize, Error> {
    if object.messages.len() != replacement_messages.len() {
        return Err(Error::InvalidSource);
    }
    let mut bytes = size_of::<ArchiveObject>()
        .checked_add(
            replacement_messages
                .len()
                .checked_mul(size_of::<RawMessage>())
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|value| {
            value.checked_add(
                object
                    .archive_info
                    .message_infos
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::MessageInfo>())?,
            )
        })
        .and_then(|value| {
            value.checked_add(usize::try_from(object.header_length).ok()?.checked_mul(2)?)
        })
        .ok_or(Error::InvalidSource)?;
    for message in replacement_messages {
        bytes = bytes
            .checked_add(message.data.len())
            .ok_or(Error::InvalidSource)?;
    }
    for info in &object.archive_info.message_infos {
        let mut info_bytes = info
            .versions
            .len()
            .checked_mul(size_of::<u32>())
            .and_then(|value| {
                value.checked_add(
                    info.diff_merge_version
                        .len()
                        .checked_mul(size_of::<u32>())?,
                )
            })
            .and_then(|value| {
                value.checked_add(info.diff_read_version.len().checked_mul(size_of::<u32>())?)
            })
            .and_then(|value| {
                value.checked_add(info.object_references.len().checked_mul(size_of::<u64>())?)
            })
            .and_then(|value| {
                value.checked_add(info.data_references.len().checked_mul(size_of::<u64>())?)
            })
            .ok_or(Error::InvalidSource)?;
        if let Some(path) = &info.diff_field_path {
            info_bytes = info_bytes
                .checked_add(
                    path.path
                        .len()
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?;
        }
        info_bytes = info_bytes
            .checked_add(
                info.fields_to_remove
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::FieldPath>())
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        for path in &info.fields_to_remove {
            info_bytes = info_bytes
                .checked_add(
                    path.path
                        .len()
                        .checked_mul(size_of::<u32>())
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?;
        }
        info_bytes = info_bytes
            .checked_add(
                info.field_infos
                    .len()
                    .checked_mul(size_of::<litchi_iwa_core::FieldInfo>())
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        for field in &info.field_infos {
            info_bytes = info_bytes
                .checked_add(
                    field
                        .path
                        .path
                        .len()
                        .checked_mul(size_of::<u32>())
                        .and_then(|value| {
                            value.checked_add(
                                field
                                    .object_references
                                    .len()
                                    .checked_mul(size_of::<u64>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                field.data_references.len().checked_mul(size_of::<u64>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                field
                                    .known_field_version
                                    .len()
                                    .checked_mul(size_of::<u32>())?,
                            )
                        })
                        .and_then(|value| {
                            value.checked_add(
                                field
                                    .known_field_feature_identifier
                                    .as_ref()
                                    .map_or(0, String::len),
                            )
                        })
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?;
        }
        bytes = bytes.checked_add(info_bytes).ok_or(Error::InvalidSource)?;
    }
    Ok(bytes)
}

fn reserve_vec<T>(
    output: &mut Vec<T>,
    additional: usize,
    budget: &mut Budget,
) -> Result<(), Error> {
    if additional == 0 {
        return Ok(());
    }
    let amount = additional
        .checked_mul(size_of::<T>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(amount)?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| Error::Allocation { amount })
}

fn copy_component_name(name: &str, budget: &mut Budget) -> Result<String, Error> {
    budget.charge_allocations(name.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(name.len())
        .map_err(|_| Error::Allocation { amount: name.len() })?;
    output.push_str(name);
    Ok(output)
}

fn copy_text_value(value: &str, budget: &mut Budget) -> Result<Box<str>, Error> {
    budget.charge_allocations(value.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| Error::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

/// Private evidence emitted with the staged candidate.
///
/// The coordinator turns this into public diagnostics after candidate
/// readback.  Keeping the evidence private means no facade can accidentally
/// make object identifiers part of the supported API.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct EngineDiagnostics {
    pub(super) changed: bool,
    pub(super) created_object_count: usize,
    pub(super) removed_object_count: usize,
    pub(super) touched_components: usize,
    pub(super) generated_author: bool,
}

impl EngineDiagnostics {
    fn noop() -> Self {
        Self {
            changed: false,
            created_object_count: 0,
            removed_object_count: 0,
            touched_components: 0,
            generated_author: false,
        }
    }
}

/// Candidate bytes plus private staging evidence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct EngineOutput {
    pub(super) bytes: Vec<u8>,
    pub(super) diagnostics: EngineDiagnostics,
}

/// Execute one complete direct-drawable comment operation.
///
/// Graph validation happens before any archive is copied.  All subsequent
/// allocations use the caller's operation-wide budget.  Every changed
/// storage node is copy-on-written except a globally single-owned root text
/// replacement, which can safely update that root in place.
pub(super) fn execute(
    source: &Package,
    selection: &Selection,
    operation: &Operation,
    budget: &mut Budget,
) -> Result<EngineOutput, Error> {
    if !source_has_exact_catalog(source) {
        return Err(Error::UnsupportedSource);
    }

    // The graph adapter performs the complete wrapper-route, rooted-closure,
    // metadata-edge, and global-owner proof.  The engine intentionally keeps
    // no second ownership implementation; it only uses the selected root ID
    // and the payloads needed for the source-preserving rewrites below.
    let thread = graph::read_comment_graph(source, selection, budget)?;

    let mut working = WorkingSet::new(source);
    let mut diagnostics = EngineDiagnostics::noop();
    let root = selection.comment_identifier;

    match operation {
        Operation::Set { text } => {
            if let Some(root_identifier) = root {
                let facts = read_storage(source, root_identifier, budget)?;
                let current = facts.text.as_deref().ok_or(Error::InvalidSource)?;
                if current == text.as_ref() {
                    return exact_noop(source, budget);
                }
                if selection.direct_users == 1 {
                    rewrite_root_in_place(
                        source,
                        &mut working,
                        selection,
                        root_identifier,
                        text,
                        budget,
                    )?;
                    advance_selected_save_token(source, &mut working, selection, budget)?;
                } else {
                    let new_root = metadata::allocate_identifier(source, &mut working, budget)?;
                    let uuid =
                        metadata::fresh_storage_uuid_in_working_set(source, &working, budget)?;
                    clone_root_with_text_and_uuid(
                        source,
                        &mut working,
                        selection,
                        root_identifier,
                        new_root,
                        text,
                        uuid,
                        budget,
                    )?;
                    replace_drawable_reference(
                        &mut working,
                        selection,
                        Some(root_identifier),
                        Some(new_root),
                        budget,
                    )?;
                    metadata::reserve_last_identifier(&mut working, new_root, budget)?;
                    cleanup_old_graph(
                        source,
                        &mut working,
                        selection,
                        thread.as_ref(),
                        root_identifier,
                        budget,
                    )?;
                    diagnostics.created_object_count = 1;
                }
                diagnostics.changed = true;
            } else {
                create_root(
                    source,
                    &mut working,
                    selection,
                    text,
                    budget,
                    &mut diagnostics,
                )?;
                advance_selected_save_token(source, &mut working, selection, budget)?;
                diagnostics.changed = true;
            }
        },
        Operation::Clear => {
            let Some(root_identifier) = root else {
                return exact_noop(source, budget);
            };
            // `read_comment_graph` above is intentionally before this edge
            // removal so malformed graphs fail atomically.
            replace_drawable_reference(
                &mut working,
                selection,
                Some(root_identifier),
                None,
                budget,
            )?;
            cleanup_old_graph(
                source,
                &mut working,
                selection,
                thread.as_ref(),
                root_identifier,
                budget,
            )?;
            diagnostics.changed = true;
        },
        Operation::AddReply { text } => {
            let root_identifier = root.ok_or(Error::InvalidSource)?;
            let _facts = read_storage(source, root_identifier, budget)?;
            let new_root = metadata::allocate_identifier(source, &mut working, budget)?;
            clone_storage_object(source, &mut working, root_identifier, new_root, budget)?;
            replace_drawable_reference(
                &mut working,
                selection,
                Some(root_identifier),
                Some(new_root),
                budget,
            )?;

            let author = metadata::ensure_generated_author(source, &mut working, budget)?;
            let author_identifier = author.identifier();
            let reply_identifier = metadata::allocate_identifier(source, &mut working, budget)?;
            let reply_uuid = metadata::fresh_storage_uuid_in_working_set(source, &working, budget)?;
            let reply_data = canonical_leaf(text, author_identifier, reply_uuid, budget)?;
            insert_storage_object(
                source,
                &mut working,
                selection.component_name.as_ref(),
                reply_identifier,
                reply_data,
                author_identifier,
                budget,
            )?;
            rewrite_root_replies(
                source,
                &mut working,
                new_root,
                CommentStorageReplyRewrite::append(reply_identifier),
                budget,
            )?;
            if let Some(author_identifier) = author_identifier {
                metadata::rewrite_for_comment_edges(
                    source,
                    &mut working,
                    metadata::EdgeEdit::add_author(
                        selection.component_name.as_ref(),
                        reply_identifier,
                        author_identifier,
                        budget,
                    )?,
                    budget,
                )?;
            }
            metadata::reserve_last_identifier(&mut working, reply_identifier, budget)?;
            cleanup_old_graph(
                source,
                &mut working,
                selection,
                thread.as_ref(),
                root_identifier,
                budget,
            )?;
            diagnostics.changed = true;
            diagnostics.created_object_count = 2;
            diagnostics.generated_author = author.created();
        },
        Operation::SetReply { selector, text } => {
            let root_identifier = root.ok_or(Error::InvalidSource)?;
            let facts = read_storage(source, root_identifier, budget)?;
            let reply_identifier = facts
                .reply_ids
                .get(selector.as_index())
                .copied()
                .ok_or(Error::InvalidSource)?;
            let reply = read_storage(source, reply_identifier, budget)?;
            if reply.text.as_deref() == Some(text.as_ref()) {
                return exact_noop(source, budget);
            }
            let new_root = metadata::allocate_identifier(source, &mut working, budget)?;
            clone_storage_object(source, &mut working, root_identifier, new_root, budget)?;
            replace_drawable_reference(
                &mut working,
                selection,
                Some(root_identifier),
                Some(new_root),
                budget,
            )?;
            let new_reply = metadata::allocate_identifier(source, &mut working, budget)?;
            clone_storage_object(source, &mut working, reply_identifier, new_reply, budget)?;
            rewrite_storage_text(&mut working, source, new_reply, text, budget)?;
            rewrite_root_replies(
                source,
                &mut working,
                new_root,
                CommentStorageReplyRewrite::replace(
                    selector.as_index(),
                    reply_identifier,
                    new_reply,
                ),
                budget,
            )?;
            metadata::reserve_last_identifier(&mut working, new_reply, budget)?;
            cleanup_old_graph(
                source,
                &mut working,
                selection,
                thread.as_ref(),
                root_identifier,
                budget,
            )?;
            diagnostics.changed = true;
            diagnostics.created_object_count = 2;
        },
        Operation::RemoveReply { selector } => {
            let root_identifier = root.ok_or(Error::InvalidSource)?;
            let facts = read_storage(source, root_identifier, budget)?;
            let reply_identifier = facts
                .reply_ids
                .get(selector.as_index())
                .copied()
                .ok_or(Error::InvalidSource)?;
            let new_root = metadata::allocate_identifier(source, &mut working, budget)?;
            clone_storage_object(source, &mut working, root_identifier, new_root, budget)?;
            replace_drawable_reference(
                &mut working,
                selection,
                Some(root_identifier),
                Some(new_root),
                budget,
            )?;
            rewrite_root_replies(
                source,
                &mut working,
                new_root,
                CommentStorageReplyRewrite::remove(selector.as_index(), reply_identifier),
                budget,
            )?;
            metadata::reserve_last_identifier(&mut working, new_root, budget)?;
            cleanup_old_graph(
                source,
                &mut working,
                selection,
                thread.as_ref(),
                root_identifier,
                budget,
            )?;
            diagnostics.changed = true;
            diagnostics.created_object_count = 1;
        },
    }

    if !diagnostics.changed {
        return exact_noop(source, budget);
    }
    let bytes = serialize_candidate(source, &working, budget)?;
    diagnostics.touched_components = working.archives.len();
    Ok(EngineOutput { bytes, diagnostics })
}

fn source_has_exact_catalog(source: &Package) -> bool {
    matches!(&source.state.source, PhysicalSource::Package(_))
}

fn exact_noop(source: &Package, budget: &mut Budget) -> Result<EngineOutput, Error> {
    budget.charge_output(source.source_bytes().len())?;
    budget.charge_allocations(source.source_bytes().len())?;
    Ok(EngineOutput {
        bytes: source.source_bytes().to_vec(),
        diagnostics: EngineDiagnostics::noop(),
    })
}

fn create_root(
    source: &Package,
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    text: &str,
    budget: &mut Budget,
    diagnostics: &mut EngineDiagnostics,
) -> Result<(), Error> {
    let author = metadata::ensure_generated_author(source, working, budget)?;
    let author_identifier = author.identifier();
    let identifier = metadata::allocate_identifier(source, working, budget)?;
    let uuid = metadata::fresh_storage_uuid_in_working_set(source, working, budget)?;
    let data = canonical_leaf(text, author_identifier, uuid, budget)?;
    insert_storage_object(
        source,
        working,
        selection.component_name.as_ref(),
        identifier,
        data,
        author_identifier,
        budget,
    )?;
    replace_drawable_reference(working, selection, None, Some(identifier), budget)?;
    if let Some(author_identifier) = author_identifier {
        metadata::rewrite_for_comment_edges(
            source,
            working,
            metadata::EdgeEdit::add_author(
                selection.component_name.as_ref(),
                identifier,
                author_identifier,
                budget,
            )?,
            budget,
        )?;
    }
    metadata::reserve_last_identifier(working, identifier, budget)?;
    diagnostics.created_object_count = 1;
    diagnostics.generated_author = author.created();
    Ok(())
}

fn advance_selected_save_token(
    source: &Package,
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    budget: &mut Budget,
) -> Result<(), Error> {
    budget.charge_allocations(selection.component_name.len())?;
    let component_name: Box<str> = selection.component_name.as_ref().into();
    metadata::advance_save_tokens(
        source,
        working,
        std::slice::from_ref(&component_name),
        budget,
    )
}

fn rewrite_root_in_place(
    source: &Package,
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    root_identifier: u64,
    text: &str,
    budget: &mut Budget,
) -> Result<(), Error> {
    let component = object_component(source, root_identifier)?;
    let archive = working.load(component, budget)?;
    let (index, payload) = storage_message(archive, root_identifier)?;
    let facts = decode_payload(root_identifier, payload, budget)?;
    let wire_limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let data = patch_text(payload, facts.text.is_some(), text, wire_limits, budget)?;
    replace_message(archive, root_identifier, index, data, source, budget)?;
    // The selection is checked before this engine call. Keep it in the
    // signature to make accidental cross-drawable use impossible.
    let _ = selection;
    Ok(())
}

fn clone_root_with_text_and_uuid(
    source: &Package,
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    old_identifier: u64,
    new_identifier: u64,
    text: &str,
    uuid: UuidSnapshot,
    budget: &mut Budget,
) -> Result<(), Error> {
    let facts = clone_storage_object_with_uuid(
        source,
        working,
        old_identifier,
        new_identifier,
        uuid,
        budget,
    )?;
    let component = copy_component_name(object_component(source, old_identifier)?, budget)?;
    let archive = working.load(component.as_str(), budget)?;
    let (index, payload) = storage_message(archive, new_identifier)?;
    let wire_limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let data = patch_text(payload, facts.text.is_some(), text, wire_limits, budget)?;
    replace_message(archive, new_identifier, index, data, source, budget)?;
    let _ = selection;
    Ok(())
}

fn clone_storage_object(
    source: &Package,
    working: &mut WorkingSet<'_>,
    old_identifier: u64,
    new_identifier: u64,
    budget: &mut Budget,
) -> Result<(), Error> {
    if new_identifier == 0 || old_identifier == 0 || old_identifier == new_identifier {
        return Err(Error::InvalidSource);
    }
    let component = object_component(source, old_identifier)?;
    let source_object = source.object(old_identifier).ok_or(Error::InvalidSource)?;
    if source_object.messages.len() != 1
        || source_object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
    {
        return Err(Error::InvalidSource);
    }
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let core_limits = core_header_limits(source_object, 0, archive_limits, budget)?;
    let clone_bytes = object_clone_allocation_bound(source_object, &source_object.messages)?;
    budget.charge_allocation_plan(clone_bytes, 1)?;
    // Source headers are preserved by the core object clone operation. This
    // also retains unknown ArchiveInfo fields and all object-reference order.
    let clone = source_object
        .clone_with_identity_remap_with_limits(
            new_identifier,
            &[],
            &source_object.messages,
            core_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let archive = working.load(component, budget)?;
    budget.charge_allocations(size_of::<ArchiveObject>())?;
    archive
        .insert_object_with_limits(clone, archive_limits)
        .map_err(|_| Error::InvalidSource)
}

fn insert_storage_object(
    source: &Package,
    working: &mut WorkingSet<'_>,
    component: &str,
    identifier: u64,
    data: Vec<u8>,
    author_identifier: Option<u64>,
    budget: &mut Budget,
) -> Result<(), Error> {
    if identifier == 0 || data.is_empty() {
        return Err(Error::InvalidSource);
    }
    budget.charge_allocations(data.len())?;
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let constructor_bytes = size_of::<ArchiveObject>()
        .checked_add(size_of::<RawMessage>())
        .and_then(|bytes| bytes.checked_add(size_of::<litchi_iwa_core::MessageInfo>()))
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocation_plan(constructor_bytes, 3)?;
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data,
        }],
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    if let Some(author_identifier) = author_identifier {
        let info = object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or(Error::InvalidSource)?;
        reserve_vec(&mut info.object_references, 1, budget)?;
        info.object_references.push(author_identifier);
    }
    let archive = working.load(component, budget)?;
    budget.charge_allocations(size_of::<ArchiveObject>())?;
    archive
        .insert_object_with_limits(object, archive_limits)
        .map_err(|_| Error::InvalidSource)
}

fn clone_storage_object_with_uuid(
    source: &Package,
    working: &mut WorkingSet<'_>,
    old_identifier: u64,
    new_identifier: u64,
    uuid: UuidSnapshot,
    budget: &mut Budget,
) -> Result<StorageFacts, Error> {
    clone_storage_object(source, working, old_identifier, new_identifier, budget)?;
    let component = copy_component_name(object_component(source, old_identifier)?, budget)?;
    let archive = working.load(component.as_str(), budget)?;
    let (index, payload) = storage_message(archive, new_identifier)?;
    let facts = decode_payload(new_identifier, payload, budget)?;
    let wire_limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let data = patch_uuid(payload, facts.uuid.is_some(), uuid, wire_limits, budget)?;
    replace_message(archive, new_identifier, index, data, source, budget)?;
    Ok(facts)
}

fn rewrite_root_replies(
    source: &Package,
    working: &mut WorkingSet<'_>,
    root_identifier: u64,
    operation: CommentStorageReplyRewrite,
    budget: &mut Budget,
) -> Result<(), Error> {
    let component_name = working
        .find_object(root_identifier)
        .map(|(name, _)| name)
        .ok_or(Error::InvalidSource)?;
    let component = copy_component_name(component_name, budget)?;
    let archive = working.load(component.as_str(), budget)?;
    let (index, payload) = storage_message(archive, root_identifier)?;
    let facts = decode_payload(root_identifier, payload, budget)?;
    let info = archive
        .object(root_identifier)
        .and_then(|object| object.archive_info.message_infos.get(index))
        .ok_or(Error::InvalidSource)?;
    let aggregate_before = copy_reference_ids(&info.object_references, budget)?;
    let (removed, added) = reply_reference_mutation(&facts.reply_ids, operation)?;
    let aggregate_after = transition_reply_references(&aggregate_before, removed, added, budget)?;

    // A native CommentStorageArchive may attribute replies through one
    // aggregate FieldInfo or through a sequence of nested paths. Update every
    // selected field that actually owns the removed reply; untouched and
    // unknown metadata stays source-authoritative. New replies are appended to
    // an existing reply-owning field when such a field exists. If the source
    // has aggregate-only metadata, the aggregate transition below is the
    // complete native representation and no synthetic FieldInfo is invented.
    let mut field_transitions = Vec::new();
    let field_count = info.field_infos.len();
    let field_transition_bytes = field_count
        .checked_mul(size_of::<OwnedFieldReferenceTransition>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(field_transition_bytes)?;
    field_transitions
        .try_reserve_exact(field_count)
        .map_err(|_| Error::Allocation {
            amount: field_transition_bytes,
        })?;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let owns_removed =
            removed.is_some_and(|identifier| field.object_references.contains(&identifier));
        let owns_reply = field
            .object_references
            .iter()
            .any(|identifier| facts.reply_ids.contains(identifier));
        if !(owns_removed || (removed.is_none() && added.is_some() && owns_reply)) {
            continue;
        }
        let before = copy_reference_ids(&field.object_references, budget)?;
        let after = transition_field_references(&before, removed, added, budget)?;
        if before == after {
            continue;
        }
        let path = copy_field_path(field.path.as_slice(), budget)?;
        field_transitions.push(OwnedFieldReferenceTransition {
            field_info_index,
            path,
            before,
            after,
        });
    }

    let options = comment_options(payload.len())?;
    let prepared =
        comment_storage_codec::prepare_comment_storage_reply_rewrite(payload, operation, options)
            .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    charge_requirements(requirements, budget)?;
    let rewritten = prepared
        .execute(requirements.exact())
        .map_err(|_| Error::InvalidSource)?
        .into_bytes();
    let transition_bytes = field_transitions
        .len()
        .checked_mul(size_of::<FieldObjectReferenceTransition<'static>>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(transition_bytes)?;
    let mut transitions = Vec::new();
    transitions
        .try_reserve_exact(field_transitions.len())
        .map_err(|_| Error::Allocation {
            amount: transition_bytes,
        })?;
    transitions.extend(
        field_transitions
            .iter()
            .map(|field| FieldObjectReferenceTransition {
                field_info_index: field.field_info_index,
                expected_path: field.path.as_slice(),
                before: field.before.as_slice(),
                after: field.after.as_slice(),
            }),
    );
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let core_limits = {
        let object = archive
            .object(root_identifier)
            .ok_or(Error::InvalidSource)?;
        core_header_limits(object, usize::from(added.is_some()), archive_limits, budget)?
    };
    archive
        .object_mut(root_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data: rewritten,
            },
            ObjectReferenceTransition {
                aggregate_before: aggregate_before.as_slice(),
                aggregate_after: aggregate_after.as_slice(),
                fields: transitions.as_slice(),
            },
            core_limits,
        )
        .map(|_| ())
        .map_err(|_| Error::InvalidSource)
}

#[derive(Debug)]
struct OwnedFieldReferenceTransition {
    field_info_index: usize,
    path: Vec<u32>,
    before: Vec<u64>,
    after: Vec<u64>,
}

fn copy_reference_ids(ids: &[u64], budget: &mut Budget) -> Result<Vec<u64>, Error> {
    let bytes = ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut copy = Vec::new();
    copy.try_reserve_exact(ids.len())
        .map_err(|_| Error::Allocation { amount: bytes })?;
    copy.extend_from_slice(ids);
    Ok(copy)
}

fn copy_field_path(path: &[u32], budget: &mut Budget) -> Result<Vec<u32>, Error> {
    let bytes = path
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut copy = Vec::new();
    copy.try_reserve_exact(path.len())
        .map_err(|_| Error::Allocation { amount: bytes })?;
    copy.extend_from_slice(path);
    Ok(copy)
}

fn reply_reference_mutation(
    reply_ids: &[u64],
    operation: CommentStorageReplyRewrite,
) -> Result<(Option<u64>, Option<u64>), Error> {
    match operation {
        CommentStorageReplyRewrite::Append { identifier } => {
            if identifier == 0 || reply_ids.contains(&identifier) {
                return Err(Error::InvalidSource);
            }
            Ok((None, Some(identifier)))
        },
        CommentStorageReplyRewrite::Replace {
            ordinal,
            expected_identifier,
            replacement_identifier,
        } => {
            if replacement_identifier == 0
                || reply_ids.get(ordinal).copied() != Some(expected_identifier)
                || expected_identifier == replacement_identifier
                || reply_ids.contains(&replacement_identifier)
            {
                return Err(Error::InvalidSource);
            }
            Ok((Some(expected_identifier), Some(replacement_identifier)))
        },
        CommentStorageReplyRewrite::Remove {
            ordinal,
            expected_identifier,
        } => {
            if expected_identifier == 0
                || reply_ids.get(ordinal).copied() != Some(expected_identifier)
            {
                return Err(Error::InvalidSource);
            }
            Ok((Some(expected_identifier), None))
        },
    }
}

fn transition_reply_references(
    before: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
    budget: &mut Budget,
) -> Result<Vec<u64>, Error> {
    let capacity = before
        .len()
        .checked_add(usize::from(added.is_some()))
        .ok_or(Error::InvalidSource)?;
    let bytes = capacity
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: bytes })?;
    after.extend_from_slice(before);
    if let Some(identifier) = removed {
        let before_len = after.len();
        after.retain(|candidate| *candidate != identifier);
        if after.len() == before_len {
            return Err(Error::InvalidSource);
        }
    }
    if let Some(identifier) = added {
        if after.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        after.push(identifier);
    }
    Ok(after)
}

fn transition_field_references(
    before: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
    budget: &mut Budget,
) -> Result<Vec<u64>, Error> {
    let capacity = before
        .len()
        .checked_add(usize::from(added.is_some()))
        .ok_or(Error::InvalidSource)?;
    let bytes = capacity
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: bytes })?;
    after.extend_from_slice(before);
    if let Some(identifier) = removed {
        let before_len = after.len();
        after.retain(|candidate| *candidate != identifier);
        if after.len() == before_len {
            return Err(Error::InvalidSource);
        }
    }
    if let Some(identifier) = added {
        if after.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        after.push(identifier);
    }
    Ok(after)
}

fn rewrite_storage_text(
    working: &mut WorkingSet<'_>,
    source: &Package,
    identifier: u64,
    text: &str,
    budget: &mut Budget,
) -> Result<(), Error> {
    let component_name = working
        .find_object(identifier)
        .map(|(name, _)| name)
        .ok_or(Error::InvalidSource)?;
    let component = copy_component_name(component_name, budget)?;
    let archive = working.load(component.as_str(), budget)?;
    let (index, payload) = storage_message(archive, identifier)?;
    let facts = decode_payload(identifier, payload, budget)?;
    let wire_limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let data = patch_text(payload, facts.text.is_some(), text, wire_limits, budget)?;
    replace_message(archive, identifier, index, data, source, budget)
}

fn replace_drawable_reference(
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    old: Option<u64>,
    new: Option<u64>,
    budget: &mut Budget,
) -> Result<(), Error> {
    let archive_limits = working
        .source()
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let wire_limits = working
        .source()
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let archive = working.load(selection.component_name.as_ref(), budget)?;
    let object = archive
        .object_mut(selection.drawable_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(selection.message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != selection.message_type {
        return Err(Error::InvalidSource);
    }
    let current = decode_drawable_reference(
        message.data.as_slice(),
        selection.comment_wire_path,
        wire_limits,
        budget,
    )?;
    if current != old {
        return Err(Error::PatchConflict);
    }
    if new == Some(0) {
        return Err(Error::InvalidSource);
    }
    let mut identifier_path = [0u32; 5];
    let identifier_path_len = selection
        .comment_wire_path
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    if identifier_path_len > identifier_path.len() {
        return Err(Error::InvalidSource);
    }
    identifier_path[..selection.comment_wire_path.len()]
        .copy_from_slice(selection.comment_wire_path);
    identifier_path[selection.comment_wire_path.len()] = COMMENT_REFERENCE_IDENTIFIER_FIELD;
    let first_envelope_missing = if old.is_none() && selection.comment_wire_path.len() > 1 {
        budget.charge_allocations(message.data.len())?;
        budget.charge_wire_work(message.data.len().max(1))?;
        let fields = wire::parse_wire_fields_with_limits(message.data.as_slice(), wire_limits)
            .map_err(|_| Error::InvalidSource)?;
        budget.charge_wire_fields(fields.len())?;
        !fields
            .iter()
            .any(|field| field.number() == selection.comment_wire_path[0])
    } else {
        false
    };

    let data = match (old, new) {
        // Preserve the complete source reference envelope, including unknown
        // fields, while changing only its canonical identifier leaf.
        (Some(_), Some(identifier)) => {
            let edit = NestedFieldEdit::new(
                &identifier_path[..identifier_path_len],
                true,
                NestedFieldReplacement::Varint(Some(identifier)),
            );
            patch_nested_fields(message.data.as_slice(), &[edit], wire_limits, 16, budget)?
        },
        // A new attachment has no source envelope to preserve.  The canonical
        // reference payload intentionally contains only its identifier.
        (None, Some(identifier)) => {
            let reference_path = [COMMENT_REFERENCE_IDENTIFIER_FIELD];
            let reference_edit = NestedFieldEdit::new(
                &reference_path,
                false,
                NestedFieldReplacement::Varint(Some(identifier)),
            );
            let reference = patch_nested_fields(&[], &[reference_edit], wire_limits, 16, budget)?;
            if first_envelope_missing {
                let data = create_missing_route_reference(
                    message.data.as_slice(),
                    selection.comment_wire_path,
                    reference.as_slice(),
                    wire_limits,
                    budget,
                )?;
                update_drawable_reference_metadata(
                    &mut object.archive_info.message_infos[selection.message_index],
                    old,
                    new,
                    budget,
                )?;
                let core_limits = core_header_limits(
                    &*object,
                    usize::from(old.is_none() && new.is_some()),
                    archive_limits,
                    budget,
                )?;
                object
                    .replace_message_preserving_header_with_limits(
                        selection.message_index,
                        RawMessage {
                            type_: selection.message_type,
                            data,
                        },
                        core_limits,
                    )
                    .map_err(|_| Error::InvalidSource)?;
                return Ok(());
            }
            let route_edit = NestedFieldEdit::new(
                selection.comment_wire_path,
                false,
                NestedFieldReplacement::LengthDelimited(Some(reference.as_slice())),
            );
            patch_nested_fields(
                message.data.as_slice(),
                &[route_edit],
                wire_limits,
                reference.len().saturating_add(16),
                budget,
            )?
        },
        // Clearing removes the complete reference envelope, as native hosts
        // do, rather than retaining an empty reference message.
        (Some(_), None) => {
            let edit = NestedFieldEdit::new(
                selection.comment_wire_path,
                true,
                NestedFieldReplacement::LengthDelimited(None),
            );
            patch_nested_fields(message.data.as_slice(), &[edit], wire_limits, 0, budget)?
        },
        (None, None) => return Err(Error::InvalidSource),
    };

    update_drawable_reference_metadata(
        &mut object.archive_info.message_infos[selection.message_index],
        old,
        new,
        budget,
    )?;
    let core_limits = core_header_limits(
        &*object,
        usize::from(old.is_none() && new.is_some()),
        archive_limits,
        budget,
    )?;
    object
        .replace_message_preserving_header_with_limits(
            selection.message_index,
            RawMessage {
                type_: selection.message_type,
                data,
            },
            core_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

fn create_missing_route_reference(
    source_data: &[u8],
    route: &[u32],
    reference: &[u8],
    wire_limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let (&root_field, descendants) = route.split_first().ok_or(Error::InvalidSource)?;
    if descendants.is_empty() {
        return Err(Error::InvalidSource);
    }
    budget.charge_allocations(reference.len())?;
    let mut nested = reference.to_vec();
    for &field_number in descendants.iter().rev() {
        let path = [field_number];
        let edit = NestedFieldEdit::new(
            &path,
            false,
            NestedFieldReplacement::LengthDelimited(Some(nested.as_slice())),
        );
        nested = patch_nested_fields(
            &[],
            &[edit],
            wire_limits,
            nested.len().saturating_add(16),
            budget,
        )?;
    }
    let path = [root_field];
    let edit = NestedFieldEdit::new(
        &path,
        false,
        NestedFieldReplacement::LengthDelimited(Some(nested.as_slice())),
    );
    patch_nested_fields(
        source_data,
        &[edit],
        wire_limits,
        nested.len().saturating_add(16),
        budget,
    )
}

fn patch_nested_fields(
    data: &[u8],
    edits: &[NestedFieldEdit<'_>],
    limits: WireLimits,
    replacement_headroom: usize,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    // The common wire helper performs a bounded planning pass but cannot see
    // the operation-wide lifecycle ledger.  Admit an upper bound before it
    // reserves its trie, plans, and output; the exact output is checked after
    // execution so a future wire implementation cannot exceed the admission.
    let output_bound = data
        .len()
        .checked_add(replacement_headroom)
        .ok_or(Error::InvalidSource)?;
    let allocation_bound = data
        .len()
        .checked_add(output_bound)
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocation_plan(allocation_bound, 2)?;
    budget.charge_wire_work(
        data.len()
            .checked_add(output_bound)
            .ok_or(Error::InvalidSource)?,
    )?;
    let output = wire::patch_nested_fields_batched_with_limits(data, edits, limits)
        .map_err(|_| Error::InvalidSource)?;
    if output.len() > output_bound {
        return Err(Error::InvalidSource);
    }
    Ok(output)
}

fn update_drawable_reference_metadata(
    info: &mut litchi_iwa_core::MessageInfo,
    old: Option<u64>,
    new: Option<u64>,
    budget: &mut Budget,
) -> Result<(), Error> {
    let old_matches = old.map_or(0, |identifier| {
        (if info.object_references.contains(&identifier) {
            1
        } else {
            0
        }) + info
            .field_infos
            .iter()
            .map(|field| {
                field
                    .object_references
                    .iter()
                    .filter(|candidate| **candidate == identifier)
                    .count()
            })
            .sum::<usize>()
    });
    if old.is_some() && old_matches != 1 {
        return Err(Error::InvalidSource);
    }
    if old.is_none() {
        if let Some(identifier) = new
            && (info.object_references.contains(&identifier)
                || info
                    .field_infos
                    .iter()
                    .any(|field| field.object_references.contains(&identifier)))
        {
            return Err(Error::InvalidSource);
        }
        if let Some(identifier) = new {
            reserve_vec(&mut info.object_references, 1, budget)?;
            info.object_references.push(identifier);
        }
        return Ok(());
    }

    let old = old.ok_or(Error::InvalidSource)?;
    // Match the native editor's metadata placement: a replacement is emitted
    // in the message aggregate, while stale field-level witnesses are
    // removed. This preserves unknown field bytes without inventing a field
    // path for a newly attached reference.
    info.object_references.retain(|candidate| *candidate != old);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|candidate| *candidate != old);
    }
    if let Some(new) = new
        && !info.object_references.contains(&new)
    {
        reserve_vec(&mut info.object_references, 1, budget)?;
        info.object_references.push(new);
    }
    Ok(())
}

struct RetainedReferenceVisitor<'a> {
    closure: &'a HashSet<u64>,
    retained: &'a mut HashSet<u64>,
    source_in_closure: bool,
}

impl ArchiveReferenceVisitor for RetainedReferenceVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if !self.source_in_closure
            && occurrence.kind == ArchiveReferenceKind::Object
            && self.closure.contains(&occurrence.referenced_identifier)
        {
            self.retained.insert(occurrence.referenced_identifier);
        }
        Ok(())
    }
}

fn cleanup_old_graph(
    source: &Package,
    working: &mut WorkingSet<'_>,
    selection: &Selection,
    thread: Option<&graph::Thread>,
    old_root: u64,
    budget: &mut Budget,
) -> Result<(), Error> {
    let Some(thread) = thread else {
        return Err(Error::InvalidSource);
    };
    if thread.root_identifier != old_root {
        return Err(Error::PatchConflict);
    }

    // A candidate's old graph is removable as one closure only when no
    // object outside that closure owns any of its nodes.  Internal reply
    // edges are ignored while computing the initial witness, then retained
    // nodes are closed over their descendants so a shared root cannot strand
    // or accidentally delete one of its replies.
    let node_count = thread.nodes.len();
    let set_bytes = node_count
        .checked_mul(size_of::<u64>())
        .and_then(|bytes| bytes.checked_mul(4))
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(set_bytes)?;
    let mut closure = HashSet::new();
    closure
        .try_reserve(node_count)
        .map_err(|_| Error::Allocation { amount: set_bytes })?;
    for node in &thread.nodes {
        closure.insert(node.identifier);
    }
    let mut retained = HashSet::new();
    retained
        .try_reserve(node_count)
        .map_err(|_| Error::Allocation { amount: set_bytes })?;
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    working.for_each_archive(|_name, archive| {
        budget.charge_entries(archive.objects.len())?;
        for object in &archive.objects {
            let object_identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if object_identifier == 0 {
                return Err(Error::InvalidSource);
            }
            if object.archive_info.message_infos.len() != object.messages.len() {
                return Err(Error::InvalidSource);
            }
            let mut fields = 0usize;
            let mut work = 0usize;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if info.type_ != message.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(Error::InvalidSource);
                }
                fields = fields
                    .checked_add(1usize.saturating_add(info.field_infos.len()))
                    .ok_or(Error::InvalidSource)?;
                work = work
                    .checked_add(message.data.len().max(1))
                    .ok_or(Error::InvalidSource)?;
            }
            let source_in_closure = closure.contains(&object_identifier);
            let mut visitor = RetainedReferenceVisitor {
                closure: &closure,
                retained: &mut retained,
                source_in_closure,
            };
            let references = object
                .inspect_references_with_policy_and_limits(
                    &mut visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            budget.charge_wire_fields(fields)?;
            budget.charge_references(references)?;
            budget.charge_wire_work(work.saturating_add(fields).saturating_add(references))?;
            if source_in_closure {
                continue;
            }
            for message in &object.messages {
                if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
                    continue;
                }
                let facts = decode_payload(object_identifier, message.data.as_slice(), budget)?;
                for identifier in facts.reply_ids {
                    if closure.contains(&identifier) {
                        retained.insert(identifier);
                    }
                }
            }
        }
        Ok(())
    })?;

    // A shared root remains attached through another drawable and therefore
    // cannot be reclaimed by this edit.  This explicit witness also protects
    // against future metadata codecs that do not project a drawable edge in
    // ArchiveInfo.
    if selection.direct_users > 1 {
        retained.insert(old_root);
    }

    // Any retained root or reply keeps its complete reachable reply suffix.
    // Index the validated closure once and use a queue so a long reply chain
    // is traversed in O(nodes + edges), with every auxiliary allocation
    // admitted before its container is reserved.
    let index_bytes = node_count
        .checked_mul(size_of::<(u64, usize)>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(index_bytes)?;
    let mut node_indices = HashMap::new();
    node_indices
        .try_reserve(node_count)
        .map_err(|_| Error::Allocation {
            amount: index_bytes,
        })?;
    for (index, node) in thread.nodes.iter().enumerate() {
        if node_indices.insert(node.identifier, index).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let queue_bytes = node_count
        .checked_mul(size_of::<usize>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(queue_bytes)?;
    let mut queue = VecDeque::new();
    queue
        .try_reserve_exact(node_count)
        .map_err(|_| Error::Allocation {
            amount: queue_bytes,
        })?;
    for (index, node) in thread.nodes.iter().enumerate() {
        if retained.contains(&node.identifier) {
            queue.push_back(index);
        }
    }
    while let Some(index) = queue.pop_front() {
        let node = thread.nodes.get(index).ok_or(Error::InvalidSource)?;
        for identifier in &node.reply_identifiers {
            if !closure.contains(identifier) || !retained.insert(*identifier) {
                continue;
            }
            let child_index = *node_indices.get(identifier).ok_or(Error::InvalidSource)?;
            queue.push_back(child_index);
        }
    }

    let removed_count = thread
        .nodes
        .iter()
        .filter(|node| !retained.contains(&node.identifier))
        .count();
    let removed_bytes = removed_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    let removed_set_bytes = removed_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(
        removed_bytes
            .checked_add(removed_set_bytes)
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(removed_count)
        .map_err(|_| Error::Allocation {
            amount: removed_bytes,
        })?;
    let mut removed_set = HashSet::new();
    removed_set
        .try_reserve(removed_count)
        .map_err(|_| Error::Allocation {
            amount: removed_set_bytes,
        })?;
    for node in &thread.nodes {
        if !retained.contains(&node.identifier) {
            removed.push(node.identifier);
            removed_set.insert(node.identifier);
        }
    }
    if removed.is_empty() {
        return advance_selected_save_token(source, working, selection, budget);
    }

    // Remove an existing optional PackageMetadata object-to-UUID witness before
    // deleting its archive object. Native sources may carry only the payload
    // UUID; the identity adapter treats that mapping as an exact, already
    // satisfied no-op. The metadata rewrite must happen first because the
    // archive object is the source witness for the component/object pair.
    for node in &thread.nodes {
        if !removed_set.contains(&node.identifier) {
            continue;
        }
        if let Some((lower, upper)) = node.storage_uuid {
            metadata::remove_storage_identity(
                source,
                working,
                node.component_name.as_ref(),
                node.identifier,
                UuidSnapshot::from_parts(lower, upper),
                budget,
            )?;
        }
    }
    for node in &thread.nodes {
        if !removed_set.contains(&node.identifier) {
            continue;
        }
        let archive = working.load(node.component_name.as_ref(), budget)?;
        archive
            .remove_object_checked_with_limits(
                node.identifier,
                source
                    .limits()
                    .effective_archive_limits()
                    .map_err(|_| Error::InvalidSource)?,
            )
            .map_err(|_| Error::InvalidSource)?;
    }
    let author_set_bytes = removed_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(author_set_bytes)?;
    let mut author_identifiers = HashSet::new();
    author_identifiers
        .try_reserve(removed_count)
        .map_err(|_| Error::Allocation {
            amount: author_set_bytes,
        })?;
    for node in &thread.nodes {
        if removed_set.contains(&node.identifier) {
            if let Some(identifier) = node.author_identifier {
                author_identifiers.insert(identifier);
            }
        }
    }
    let author_count = author_identifiers.len();
    let suffix_author_bytes = author_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(suffix_author_bytes)?;
    let mut suffix_removed = removed;
    suffix_removed
        .try_reserve_exact(author_count)
        .map_err(|_| Error::Allocation {
            amount: suffix_author_bytes,
        })?;
    for identifier in author_identifiers {
        if metadata::cleanup_generated_author_for_identifier(source, working, identifier, budget)? {
            suffix_removed.push(identifier);
        }
    }
    metadata::release_identifier_suffix(working, &suffix_removed, budget)?;
    budget.charge_allocations(selection.component_name.len())?;
    let component_name: Box<str> = selection.component_name.as_ref().into();
    metadata::advance_save_tokens(
        source,
        working,
        std::slice::from_ref(&component_name),
        budget,
    )
}

fn replace_message(
    archive: &mut Archive,
    identifier: u64,
    index: usize,
    data: Vec<u8>,
    source: &Package,
    budget: &mut Budget,
) -> Result<(), Error> {
    budget.charge_output(data.len())?;
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let core_limits = {
        let object = archive.object(identifier).ok_or(Error::InvalidSource)?;
        core_header_limits(object, 0, archive_limits, budget)?
    };
    let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
    let message_type = object
        .messages
        .get(index)
        .ok_or(Error::InvalidSource)?
        .type_;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            core_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

fn storage_message(archive: &Archive, identifier: u64) -> Result<(usize, &[u8]), Error> {
    let object = archive.object(identifier).ok_or(Error::InvalidSource)?;
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            continue;
        }
        if selected.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn object_component(source: &Package, identifier: u64) -> Result<&str, Error> {
    source
        .object_with_component(identifier)
        .map(|(component, _)| component)
        .ok_or(Error::InvalidSource)
}

#[derive(Debug)]
struct StorageFacts {
    text: Option<Box<str>>,
    reply_ids: Vec<u64>,
    uuid: Option<UuidSnapshot>,
}

fn read_storage(
    source: &Package,
    identifier: u64,
    budget: &mut Budget,
) -> Result<StorageFacts, Error> {
    let (_, object) = source
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)?;
    let (index, payload) = storage_message_in_object(object)?;
    let _ = index;
    decode_payload(identifier, payload, budget)
}

fn storage_message_in_object(object: &ArchiveObject) -> Result<(usize, &[u8]), Error> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            continue;
        }
        if selected.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn decode_payload(
    _identifier: u64,
    payload: &[u8],
    budget: &mut Budget,
) -> Result<StorageFacts, Error> {
    let options = comment_options_with_limits(payload.len())?;
    // The first borrowed pass gives the exact reply cardinality without
    // collecting references. Admit that backing allocation before the second
    // pass fills it; the visitor never grows beyond the proven cardinality.
    let (_, count_report) =
        comment_storage_codec::decode_comment_storage_archive_with_report(payload, options)
            .map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(count_report.fields())?;
    budget.charge_wire_work(count_report.work_bytes())?;
    budget.charge_nesting(count_report.max_depth() as usize)?;
    budget.charge_references(count_report.references())?;
    let maximum = count_report.replies();
    let reply_bytes = maximum
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(reply_bytes)?;
    let mut visitor = ReplyCollector {
        identifiers: Vec::new(),
        maximum,
        overflow: false,
    };
    visitor
        .identifiers
        .try_reserve_exact(maximum)
        .map_err(|_| Error::Allocation {
            amount: reply_bytes,
        })?;
    let (snapshot, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references())?;
    if visitor.overflow || visitor.identifiers.len() != maximum || report.replies() != maximum {
        return Err(Error::InvalidSource);
    }
    let text = snapshot
        .text()
        .map(|value| copy_text_value(value, budget))
        .transpose()?;
    if let Some(reference) = snapshot.author() {
        if reference.identifier() == 0
            || reference.deprecated_type().is_some()
            || reference.deprecated_is_external().is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(StorageFacts {
        text,
        reply_ids: visitor.identifiers,
        uuid: snapshot.storage_uuid(),
    })
}

#[derive(Debug)]
struct ReplyCollector {
    identifiers: Vec<u64>,
    maximum: usize,
    overflow: bool,
}

impl comment_storage_codec::CommentStorageVisitor for ReplyCollector {
    fn visit_reply(
        &mut self,
        reply: ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.identifiers.len() >= self.maximum {
            self.overflow = true;
        } else {
            self.identifiers.push(reply.identifier());
        }
        Ok(())
    }
}

fn patch_text(
    source: &[u8],
    present: bool,
    text: &str,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let path = [COMMENT_TEXT_FIELD];
    let edit = NestedFieldEdit::new(
        &path,
        present,
        NestedFieldReplacement::LengthDelimited(Some(text.as_bytes())),
    );
    patch_nested_fields(
        source,
        &[edit],
        limits,
        text.len().saturating_add(16),
        budget,
    )
}

fn patch_uuid(
    source: &[u8],
    present: bool,
    uuid: UuidSnapshot,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    if present {
        let lower_path = [COMMENT_UUID_FIELD, UUID_LOWER_FIELD];
        let upper_path = [COMMENT_UUID_FIELD, UUID_UPPER_FIELD];
        let edits = [
            NestedFieldEdit::new(
                &lower_path,
                true,
                NestedFieldReplacement::Varint(Some(uuid.lower())),
            ),
            NestedFieldEdit::new(
                &upper_path,
                true,
                NestedFieldReplacement::Varint(Some(uuid.upper())),
            ),
        ];
        return patch_nested_fields(source, &edits, limits, 32, budget);
    }

    // Some native roots have no UUID leaf.  Build the canonical UUID payload
    // and append it as one outer length-delimited field rather than rejecting
    // an otherwise valid source or fabricating a metadata identity witness.
    let lower_path = [UUID_LOWER_FIELD];
    let upper_path = [UUID_UPPER_FIELD];
    let uuid_edits = [
        NestedFieldEdit::new(
            &lower_path,
            false,
            NestedFieldReplacement::Varint(Some(uuid.lower())),
        ),
        NestedFieldEdit::new(
            &upper_path,
            false,
            NestedFieldReplacement::Varint(Some(uuid.upper())),
        ),
    ];
    let uuid_payload = patch_nested_fields(&[], &uuid_edits, limits, 32, budget)?;
    let outer_path = [COMMENT_UUID_FIELD];
    let outer_edit = NestedFieldEdit::new(
        &outer_path,
        false,
        NestedFieldReplacement::LengthDelimited(Some(uuid_payload.as_slice())),
    );
    patch_nested_fields(
        source,
        &[outer_edit],
        limits,
        uuid_payload.len().saturating_add(16),
        budget,
    )
}

fn canonical_leaf(
    text: &str,
    author_identifier: Option<u64>,
    uuid: UuidSnapshot,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let date = current_apple_date()?;
    let options = comment_options_with_limits(text.len().saturating_add(128))?;
    let write = match author_identifier {
        Some(identifier) => {
            comment_storage_codec::CommentStorageLeafWrite::new(text, date, identifier, uuid)
        },
        None => comment_storage_codec::CommentStorageLeafWrite::without_author(text, date, uuid),
    };
    let prepared = comment_storage_codec::prepare_comment_storage_leaf_write(write, options)
        .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    charge_requirements(requirements, budget)?;
    prepared
        .execute(requirements.exact())
        .map_err(|_| Error::InvalidSource)
        .map(comment_storage_codec::RewriteOutput::into_bytes)
}

fn current_apple_date() -> Result<DateSnapshot, Error> {
    let seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|_| Error::InvalidSource)?
        .as_secs_f64()
        - APPLE_EPOCH_UNIX_OFFSET_SECONDS;
    if !seconds.is_finite() || seconds < 0.0 {
        return Err(Error::InvalidSource);
    }
    Ok(DateSnapshot::from_bits(seconds.to_bits()))
}

fn comment_options(source_length: usize) -> Result<DecodeOptions, Error> {
    comment_options_with_limits(source_length)
}

fn comment_options_with_limits(source_length: usize) -> Result<DecodeOptions, Error> {
    let bytes = source_length.max(1);
    let fields = bytes.checked_mul(8).ok_or(Error::InvalidSource)?.max(1);
    let work = bytes.checked_mul(64).ok_or(Error::InvalidSource)?.max(1);
    let references = bytes.max(1);
    Ok(DecodeOptions::new(
        bytes,
        fields,
        work,
        COMMENT_RECURSION_LIMIT,
        references,
        bytes,
    ))
}

fn charge_requirements(
    requirements: RewriteExecutionRequirements,
    budget: &mut Budget,
) -> Result<(), Error> {
    budget.charge_input(requirements.input_bytes)?;
    budget.charge_output(requirements.output_bytes)?;
    budget.charge_wire_fields(requirements.fields)?;
    budget.charge_wire_work(requirements.work_bytes)?;
    budget.charge_wire_work(requirements.reference_bytes)?;
    budget.charge_nesting(requirements.max_depth as usize)?;
    budget.charge_references(requirements.references)?;
    budget.charge_allocation_plan(
        requirements
            .scratch_bytes
            .checked_add(requirements.retained_bytes)
            .ok_or(Error::InvalidSource)?,
        requirements.allocations,
    )?;
    Ok(())
}

fn serialize_candidate(
    source: &Package,
    working: &WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let source_catalog = match &source.state.source {
        PhysicalSource::Package(source) => source,
        PhysicalSource::Semantic(_) => return Err(Error::UnsupportedSource),
    };
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let changed_count = working.archives.len();
    let edit_storage_bytes = changed_count
        .checked_mul(size_of::<Vec<u8>>())
        .and_then(|bytes| {
            changed_count
                .checked_mul(size_of::<EntryEdit<'_>>())
                .and_then(|edit_bytes| bytes.checked_add(edit_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(edit_storage_bytes)?;
    let mut encoded_archives = Vec::new();
    encoded_archives
        .try_reserve_exact(changed_count)
        .map_err(|_| Error::Allocation {
            amount: edit_storage_bytes,
        })?;
    let snappy_limits = source
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource)?;
    for name in working.changed_names() {
        let archive = working.get(name).ok_or(Error::InvalidSource)?;
        let encoded_length = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource)?;
        let compressed_upper_bound = SnappyStream::maximum_compressed_len(encoded_length)
            .map_err(|_| Error::InvalidSource)?;
        if compressed_upper_bound > snappy_limits.max_compressed_stream() {
            return Err(Error::InvalidSource);
        }
        budget.charge_allocations(
            encoded_length
                .checked_add(compressed_upper_bound)
                .ok_or(Error::InvalidSource)?,
        )?;
        let decompressed = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource)?;
        if decompressed.len() != encoded_length {
            return Err(Error::InvalidSource);
        }
        let bytes = SnappyStream::compress(&decompressed).map_err(|_| Error::InvalidSource)?;
        if bytes.len() > snappy_limits.max_compressed_stream() {
            return Err(Error::InvalidSource);
        }
        encoded_archives.push(bytes);
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(changed_count)
        .map_err(|_| Error::Allocation {
            amount: edit_storage_bytes,
        })?;
    for (name, bytes) in working.changed_names().zip(&encoded_archives) {
        edits.push(EntryEdit::new(name, bytes));
    }
    let prepared = source_catalog
        .package()
        .prepare_reassembly(&edits, source_catalog.limits())
        .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.charge_output(requirements.output_bytes())?;
    budget.charge_allocation_plan(
        requirements
            .scratch_bytes()
            .checked_add(requirements.retained_bytes())
            .ok_or(Error::InvalidSource)?,
        requirements.allocations(),
    )?;
    prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::InvalidSource)
}

fn decode_drawable_reference(
    source: &[u8],
    path: &[u32],
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Option<u64>, Error> {
    let mut current = source;
    for (depth, field_number) in path.iter().copied().enumerate() {
        if current.len() > limits.max_input_bytes() {
            return Err(Error::InvalidSource);
        }
        budget.charge_allocations(current.len())?;
        budget.charge_wire_work(current.len().max(1))?;
        budget.charge_nesting(depth.saturating_add(1))?;
        let fields = wire::parse_wire_fields_with_limits(current, limits)
            .map_err(|_| Error::InvalidSource)?;
        budget.charge_wire_fields(fields.len())?;
        let mut selected = None;
        for field in fields {
            if field.number() != field_number {
                continue;
            }
            if selected.is_some() || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            selected = Some(field.payload(current).map_err(|_| Error::InvalidSource)?);
        }
        let Some(selected) = selected else {
            return Ok(None);
        };
        if depth + 1 == path.len() {
            budget.charge_allocations(selected.len())?;
            budget.charge_wire_work(selected.len().max(1))?;
            let reference = comment_storage_codec::decode_reference(
                selected,
                comment_options_with_limits(selected.len())?,
            )
            .map_err(|_| Error::InvalidSource)?;
            return Ok(Some(reference.identifier()));
        }
        current = selected;
    }
    Err(Error::InvalidSource)
}
