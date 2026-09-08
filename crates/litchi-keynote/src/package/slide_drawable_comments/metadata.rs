//! Source-bound metadata planning for drawable-comment transactions.
//!
//! Comment payloads live in archive objects, while the identities that make a
//! cloned thread valid live in several independent native projections.  This
//! module keeps the two concerns separate from the transaction engine: it
//! discovers the authoritative author storage, validates the exact current
//! metadata witnesses, and returns owned plans that the engine can apply to a
//! copy-on-write [`super::engine::WorkingSet`].  No native identifier or
//! generated protobuf value crosses the public package API.
//!
//! The engine owns candidate publication.  Every function here is therefore
//! source-or-plan based: a failed witness never mutates an archive and a
//! required component is reported by its discovered source name rather than
//! by an assumed `Index/AnnotationAuthorStorage.iwa` path.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The adapter keeps source discovery, identity planning, and strict witnesses together."
)]

use std::collections::HashSet;
use std::mem::size_of;

use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor,
};
use litchi_iwa_protos::{
    annotation_author_codec, comment_storage_codec, package_metadata_codec as identity_codec,
};

use super::super::Package;
use super::engine::WorkingSet;
use super::{Budget, Error};

mod identities;
pub(super) use identities::{
    register_storage_identity_if_registered, reject_data_reference_owners_for_removed_objects,
    remove_storage_identity,
};

const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;

const DOCUMENT_SUPER_FIELD: u32 = 3;
const TSA_DOCUMENT_SUPER_FIELD: u32 = 1;
const TSK_ANNOTATION_AUTHOR_STORAGE_FIELD: u32 = 7;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const GENERATED_AUTHOR_NAME: &str = "litchi-iwa";
const GENERATED_AUTHOR_PUBLIC_ID: &str = "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6";

/// Native UUID facts retained by an owned metadata plan.
pub(super) type StorageUuid = comment_storage_codec::UuidSnapshot;

/// An exact current component selector owned by a pending transaction.
///
/// The identity codec consumes borrowed selectors.  Keeping this owned form
/// in a plan prevents a transaction from borrowing a temporary source census
/// or accidentally routing an edge by basename alone.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct ComponentRef {
    identifier: u64,
    locator: Box<str>,
}

impl ComponentRef {
    #[must_use]
    pub(super) fn new(identifier: u64, locator: &str) -> Self {
        Self {
            identifier,
            locator: locator.into(),
        }
    }

    #[must_use]
    pub(super) const fn identifier(&self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub(super) fn selector(&self) -> identity_codec::ComponentSelector<'_> {
        identity_codec::ComponentSelector::new(self.identifier, &self.locator)
    }
}

/// One exact component-external-reference transition.
///
/// A comment author edge is strong in native packages (`is_weak` omitted or
/// false).  The value is retained in both additions and removals so the
/// neutral metadata codec can verify the selected source record before
/// preserving every unselected/versioned/unknown record byte-for-byte.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum EdgeEdit {
    /// Resolve the source/target component identities from the authoritative
    /// PackageMetadata census using a native component name and author ID.
    /// This is the compact request used by the comment engine when it has not
    /// yet borrowed the metadata selectors.
    AddAuthor {
        component: Box<str>,
        author_identifier: u64,
    },
}

/// One source-component ownership transition for a copied comment storage
/// object.  The metadata edge is optional in native packages; a missing old
/// edge therefore remains missing unless the source census proves that the
/// package already carried the dependency witness.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CommentStorageEdgeTransition<'a> {
    pub(super) source_component: &'a str,
    pub(super) old_target_component: Option<&'a str>,
    pub(super) old_identifier: Option<u64>,
    pub(super) new_target_component: Option<&'a str>,
    pub(super) new_identifier: Option<u64>,
}

impl EdgeEdit {
    pub(super) fn add_author(
        component: &str,
        _object_identifier: u64,
        author_identifier: u64,
        budget: &mut Budget,
    ) -> Result<Self, Error> {
        budget.charge_allocations(component.len())?;
        Ok(Self::AddAuthor {
            component: component.into(),
            author_identifier,
        })
    }
}

/// The source location of the single authoritative author-storage object.
///
/// The component name is discovered through the root Document reference.  A
/// valid package may co-locate this object with Document or another native
/// component, so callers must never replace it with a guessed filename.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct AuthorStorageLocation {
    pub(super) component: Box<str>,
    pub(super) object_identifier: u64,
    pub(super) message_index: usize,
}

/// A source author binding used by comment clone/remove plans.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct AuthorBinding {
    pub(super) author_identifier: Option<u64>,
    pub(super) author_component: Option<Box<str>>,
    pub(super) author_message_index: Option<usize>,
    pub(super) author_storage_identifier: Option<u64>,
    pub(super) author_storage_component: Option<Box<str>>,
    pub(super) author_storage_message_index: Option<usize>,
    pub(super) author_storage_ordinal: Option<usize>,
    pub(super) generated: bool,
}

impl AuthorBinding {
    /// Return the native identifier when an author was found or allocated.
    ///
    /// The `Option` shape lets the engine represent a package with no rooted
    /// or compatibility author-storage witness without manufacturing an ID.
    /// Root creation may preserve that authorless state; operations that
    /// require an existing author reject `None` before using the identifier.
    #[must_use]
    pub(super) const fn identifier(&self) -> Option<u64> {
        self.author_identifier
    }

    #[must_use]
    pub(super) const fn created(&self) -> bool {
        self.generated
    }
}

/// The component/object pair that owns one selected message type.
#[derive(Debug, Clone, Copy)]
struct MessageLocation<'a> {
    message_index: usize,
    payload: &'a [u8],
}

/// Resolve annotation-author storage from the rooted Document ownership edge.
///
/// A valid field-7 edge is authoritative and unrelated type-213 objects are
/// ignored.  Only an absent optional field-7 edge enters the narrow legacy
/// fallback, which accepts one globally unique type-213 message or preserves
/// the authorless no-registry state when none exists.  A present malformed,
/// ambiguous, or dangling rooted edge never falls back to that scan.
///
/// This is intentionally stricter than scanning for a basename.  The root
/// Document reference is the native ownership witness; a second unreferenced
/// type-213 object must not become an accidental author registry.
pub(super) fn annotation_author_storage_location(
    package: &Package,
    budget: &mut Budget,
) -> Result<Option<AuthorStorageLocation>, Error> {
    let root = package.root_document_payload().map_err(|_| Error::Read)?;
    let limits = package
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let tsa_payload = unique_length_field(root, DOCUMENT_SUPER_FIELD, limits, budget, 1)?
        .ok_or(Error::InvalidSource)?;
    let tsk_payload =
        unique_length_field(tsa_payload, TSA_DOCUMENT_SUPER_FIELD, limits, budget, 2)?
            .ok_or(Error::InvalidSource)?;
    let reference_payload = unique_length_field(
        tsk_payload,
        TSK_ANNOTATION_AUTHOR_STORAGE_FIELD,
        limits,
        budget,
        3,
    )?;
    let Some(reference_payload) = reference_payload else {
        return global_annotation_author_storage_location(package, budget);
    };
    let storage_identifier = strict_reference_identifier(reference_payload, limits, budget, 3)?;
    let Some((component, object)) = package.object_with_component(storage_identifier) else {
        return Err(Error::InvalidSource);
    };
    let location = unique_message_location(
        component,
        object,
        storage_identifier,
        ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
    )?;
    budget.charge_allocations(component.len())?;
    Ok(Some(AuthorStorageLocation {
        component: component.into(),
        object_identifier: storage_identifier,
        message_index: location.message_index,
    }))
}

/// Find the compatibility registry used by the legacy comment host.
///
/// The rooted Document field is the authoritative route whenever it exists.
/// This scan is therefore reached only when the optional field-7 registry
/// reference is absent.  In that narrow case the old host accepted exactly
/// one type-213 message anywhere in the package, so retain that source-order
/// behavior while keeping the candidate source-bound and bounded.
fn global_annotation_author_storage_location(
    package: &Package,
    budget: &mut Budget,
) -> Result<Option<AuthorStorageLocation>, Error> {
    let mut selected: Option<(&str, u64, usize)> = None;
    for component in package.state.source.components().iter() {
        let archive = component.archive();
        budget.charge_entries(archive.objects.len())?;
        let mut message_count = 0usize;
        for object in &archive.objects {
            message_count = message_count
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?;
        }
        budget.charge_wire_work(message_count.max(1))?;

        budget.charge_entries(archive.objects.len())?;
        budget.charge_wire_work(message_count.max(1))?;
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE {
                    continue;
                }
                if selected.is_some() {
                    return Err(Error::InvalidSource);
                }
                selected = Some((component.name(), identifier, message_index));
            }
        }
    }

    let Some((component_name, identifier, message_index)) = selected else {
        return Ok(None);
    };
    let (actual_component, object) = package
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)?;
    if actual_component != component_name {
        return Err(Error::InvalidSource);
    }
    let location = unique_message_location(
        actual_component,
        object,
        identifier,
        ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
    )?;
    if location.message_index != message_index {
        return Err(Error::InvalidSource);
    }
    budget.charge_allocations(component_name.len())?;
    Ok(Some(AuthorStorageLocation {
        component: component_name.into(),
        object_identifier: identifier,
        message_index,
    }))
}

/// Resolve and validate the selected storage object using the source root
/// witness.  The returned payload remains borrowed from the immutable package
/// and is never copied by discovery.
pub(super) fn annotation_author_storage_payload<'a>(
    package: &'a Package,
    budget: &mut Budget,
) -> Result<Option<(AuthorStorageLocation, &'a [u8])>, Error> {
    let Some(location) = annotation_author_storage_location(package, budget)? else {
        return Ok(None);
    };
    let (_component, object) = package
        .object_with_component(location.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(location.message_index)
        .ok_or(Error::InvalidSource)?;
    Ok(Some((location, message.data.as_slice())))
}

fn unique_length_field<'a>(
    payload: &'a [u8],
    number: u32,
    limits: WireLimits,
    budget: &mut Budget,
    nesting: usize,
) -> Result<Option<&'a [u8]>, Error> {
    let view = parse_wire(payload, limits, budget, nesting)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == number) {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if field.wire_type() != 2 || selected.replace(field.payload()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(selected)
}

fn strict_reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut Budget,
    nesting: usize,
) -> Result<u64, Error> {
    let view = parse_wire(payload, limits, budget, nesting)?;
    let mut identifier = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if field.number() != REFERENCE_IDENTIFIER_FIELD
            || field.wire_type() != 0
            || identifier.is_some()
        {
            // Deprecated reference metadata and unknown extensions are not
            // safe ownership witnesses.  They stay untouched in unselected
            // messages, but a selected root edge is fail-closed.
            return Err(Error::InvalidSource);
        }
        let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
            .map_err(|_| Error::InvalidSource)?;
        if width != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != width
            || value == 0
        {
            return Err(Error::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(Error::InvalidSource)
}

fn parse_wire<'a>(
    payload: &'a [u8],
    limits: WireLimits,
    budget: &mut Budget,
    nesting: usize,
) -> Result<WireView<'a>, Error> {
    if payload.len() > limits.max_input_bytes() {
        return Err(Error::InvalidSource);
    }
    budget.charge_wire_work(payload.len().max(1))?;
    budget.charge_nesting(nesting.max(1))?;
    let view = WireView::parse_with_limits(payload, limits).map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(view.len())?;
    Ok(view)
}

fn unique_message_location<'a>(
    _component: &str,
    object: &'a ArchiveObject,
    identifier: u64,
    message_type: u32,
) -> Result<MessageLocation<'a>, Error> {
    if object.archive_info.identifier != Some(identifier)
        || object.archive_info.message_infos.len() != object.messages.len()
    {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (message_index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ != message_type {
            continue;
        }
        if selected
            .replace(MessageLocation {
                message_index,
                payload: message.data.as_slice(),
            })
            .is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

/// Ensure that the native generated author is registered in the authoritative
/// storage list and staged in the same component as that storage object.
///
/// Existing source authors are validated strictly and reused in source order.
/// A missing generated author is appended with the neutral Buffa codec, then
/// its object is inserted into the candidate `WorkingSet`; the immutable
/// package is never changed.  The returned binding is intentionally tiny so
/// the engine can put the native identifier back into a private comment
/// payload without exposing it through the public API.
pub(super) fn ensure_generated_author(
    source: &Package,
    working: &mut WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<AuthorBinding, Error> {
    let Some((location, storage_payload)) = annotation_author_storage_payload(source, budget)?
    else {
        return Ok(AuthorBinding {
            author_identifier: None,
            author_component: None,
            author_message_index: None,
            author_storage_identifier: None,
            author_storage_component: None,
            author_storage_message_index: None,
            author_storage_ordinal: None,
            generated: false,
        });
    };
    let storage_options = annotation_author_codec::DecodeOptions::for_source(storage_payload);
    let (storage, storage_report) =
        annotation_author_codec::decode_annotation_author_storage_with_report(
            storage_payload,
            storage_options,
        )
        .map_err(|_| Error::InvalidSource)?;
    charge_author_decode_report(budget, storage_report)?;

    let author_ref_count = storage.author_ref_count();
    let author_ids_bytes = author_ref_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(author_ids_bytes)?;
    let mut author_ids = Vec::new();
    author_ids
        .try_reserve_exact(author_ref_count)
        .map_err(|_| Error::Allocation {
            amount: author_ids_bytes,
        })?;
    budget.charge_allocations(author_ids_bytes)?;
    let mut seen = HashSet::new();
    seen.try_reserve(author_ref_count)
        .map_err(|_| Error::Allocation {
            amount: author_ids_bytes,
        })?;
    for (ordinal, reference) in storage.author_refs().enumerate() {
        let identifier = reference.identifier();
        if identifier == 0
            || reference.deprecated_type().is_some()
            || reference.deprecated_is_external().is_some()
            || !seen.insert(identifier)
        {
            return Err(Error::InvalidSource);
        }
        author_ids.push(identifier);
        let Some((component, object)) = working.find_object(identifier) else {
            return Err(Error::InvalidSource);
        };
        let author = unique_message_location(
            component,
            object,
            identifier,
            ANNOTATION_AUTHOR_MESSAGE_TYPE,
        )?;
        let options = annotation_author_codec::DecodeOptions::for_source(author.payload);
        let (snapshot, report) =
            annotation_author_codec::decode_annotation_author_with_report(author.payload, options)
                .map_err(|_| Error::InvalidSource)?;
        charge_author_decode_report(budget, report)?;
        if is_generated_author(&snapshot, true) {
            budget.charge_allocations(component.len())?;
            return Ok(AuthorBinding {
                author_identifier: Some(identifier),
                author_component: Some(component.into()),
                author_message_index: Some(author.message_index),
                author_storage_identifier: Some(location.object_identifier),
                author_storage_component: Some(location.component),
                author_storage_message_index: Some(location.message_index),
                author_storage_ordinal: Some(ordinal),
                generated: false,
            });
        }
    }

    let identifier = allocate_identifier(source, working, budget)?;
    let public_ids = [GENERATED_AUTHOR_PUBLIC_ID];
    let color = annotation_author_codec::AuthorColorWrite::new(
        1,
        Some(0.368_627_46),
        Some(0.568_627_5),
        Some(0.937_254_9),
        Some(1.0),
        None,
        None,
        None,
        None,
        None,
        Some(1),
    );
    let author_write = annotation_author_codec::AnnotationAuthorWrite::new(
        Some(GENERATED_AUTHOR_NAME),
        Some(color),
        Some(GENERATED_AUTHOR_PUBLIC_ID),
        Some(false),
        &public_ids,
    );
    let (author_bytes, author_report) =
        annotation_author_codec::encode_annotation_author_with_report(
            &author_write,
            annotation_author_codec::EncodeOptions::for_author(&author_write),
        )
        .map_err(|_| Error::InvalidSource)?;
    charge_author_encode_report(budget, author_report)?;
    let author_payload = author_bytes.into_boxed_slice();

    // Build the storage candidate before borrowing the archive mutably.  The
    // prepared rewrite verifies both the selected source ordinal and every
    // untouched field, including unknown fields and source ordering.
    let storage_options = annotation_author_codec::DecodeOptions::for_source(storage_payload);
    let prepared = annotation_author_codec::prepare_annotation_author_storage_rewrite(
        storage_payload,
        annotation_author_codec::AnnotationAuthorStorageRewrite::append(identifier),
        storage_options,
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.requirements();
    charge_author_rewrite_requirements(budget, requirements)?;
    let storage_candidate = prepared
        .execute(requirements.exact())
        .map_err(|_| Error::InvalidSource)?
        .0;

    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let message_vec_bytes = size_of::<litchi_iwa_core::RawMessage>();
    let message_info_vec_bytes = size_of::<litchi_iwa_core::MessageInfo>();
    let version_vec_bytes = 3usize
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource)?;
    let nested_allocation_bytes = message_vec_bytes
        .checked_add(message_info_vec_bytes)
        .and_then(|bytes| bytes.checked_add(version_vec_bytes))
        .ok_or(Error::InvalidSource)?;
    // `encode_annotation_author_with_report` already charges the payload
    // buffer. Transfer that buffer into RawMessage without copying it, then
    // precharge the three heap vectors built by ArchiveObject::new_with_limits
    // (messages, MessageInfo records, and their conventional version list).
    budget.charge_allocation_plan(nested_allocation_bytes, 3)?;
    let object = ArchiveObject::new_with_limits(
        identifier,
        vec![litchi_iwa_core::RawMessage {
            type_: ANNOTATION_AUTHOR_MESSAGE_TYPE,
            data: author_payload.into_vec(),
        }],
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource)?;

    let storage_archive = working.load(&location.component, budget)?;
    let storage_object = storage_archive
        .object_mut(location.object_identifier)
        .ok_or(Error::InvalidSource)?;
    replace_author_storage_message(
        storage_object,
        location.message_index,
        storage_candidate,
        &author_ids,
        None,
        Some(identifier),
        archive_limits,
        budget,
    )?;
    storage_archive
        .insert_object_with_limits(object, archive_limits)
        .map_err(|_| Error::InvalidSource)?;

    Ok(AuthorBinding {
        author_identifier: Some(identifier),
        author_component: Some({
            budget.charge_allocations(location.component.len())?;
            location.component.clone()
        }),
        author_message_index: Some(0),
        author_storage_identifier: Some(location.object_identifier),
        author_storage_component: Some(location.component),
        author_storage_message_index: Some(location.message_index),
        author_storage_ordinal: Some(storage.author_ref_count()),
        generated: true,
    })
}

fn is_generated_author(
    author: &annotation_author_codec::AnnotationAuthorSnapshot<'_>,
    with_public_id: bool,
) -> bool {
    let Some(color) = author.color() else {
        return false;
    };
    let mut public_ids = author.public_ids();
    let public_ids_match = if with_public_id {
        public_ids.next() == Some(GENERATED_AUTHOR_PUBLIC_ID) && public_ids.next().is_none()
    } else {
        public_ids.next().is_none()
    };
    author.name() == Some(GENERATED_AUTHOR_NAME)
        && author.public_id() == with_public_id.then_some(GENERATED_AUTHOR_PUBLIC_ID)
        && author.is_public_author() == Some(false)
        && public_ids_match
        && color.model() == 1
        && color.red().map(f32::to_bits) == Some(0.368_627_46_f32.to_bits())
        && color.green().map(f32::to_bits) == Some(0.568_627_5_f32.to_bits())
        && color.blue().map(f32::to_bits) == Some(0.937_254_9_f32.to_bits())
        && color.alpha().map(f32::to_bits) == Some(1.0_f32.to_bits())
        && color.cyan().is_none()
        && color.magenta().is_none()
        && color.yellow().is_none()
        && color.black().is_none()
        && color.white().is_none()
        && color.rgbspace() == Some(1)
}

fn charge_author_decode_report(
    budget: &mut Budget,
    report: annotation_author_codec::DecodeReport,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references())?;
    budget.charge_allocation_plan(report.text_bytes(), report.allocations())?;
    Ok(())
}

fn charge_author_rewrite_requirements(
    budget: &mut Budget,
    requirements: annotation_author_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    budget.charge_input(requirements.input_bytes())?;
    budget.charge_output(requirements.output_bytes())?;
    budget.charge_wire_fields(requirements.fields())?;
    budget.charge_wire_work(requirements.work_bytes())?;
    budget.charge_nesting(requirements.max_depth() as usize)?;
    budget.charge_references(requirements.references())?;
    let scratch_and_retained = requirements
        .scratch_bytes()
        .checked_add(requirements.retained_bytes())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocation_plan(scratch_and_retained, requirements.allocations())?;
    Ok(())
}

fn charge_author_encode_report(
    budget: &mut Budget,
    report: annotation_author_codec::EncodeReport,
) -> Result<(), Error> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_references(report.references())?;
    budget.charge_allocation_plan(report.output_bytes(), report.allocations())?;
    Ok(())
}

fn comment_decode_options(
    payload: &[u8],
    limits: WireLimits,
) -> Result<comment_storage_codec::DecodeOptions, Error> {
    if payload.len() > limits.max_input_bytes() {
        return Err(Error::InvalidSource);
    }
    let bytes = payload.len().max(1);
    let fields = bytes
        .checked_mul(8)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_fields())
        .max(1);
    let work = bytes
        .checked_mul(64)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_rewrite_work())
        .max(1);
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let references = bytes.min(limits.max_fields()).max(1);
    let text = bytes.min(limits.max_input_bytes()).max(1);
    Ok(comment_storage_codec::DecodeOptions::new(
        bytes, fields, work, recursion, references, text,
    ))
}

/// Allocate the next object identifier from the effective candidate view.
///
/// Archive object IDs and the PackageMetadata watermark are both considered;
/// either one can be ahead of the other in a valid source.  The reservation is
/// recorded in `WorkingSet` before the caller inserts the object, so a second
/// allocation in the same transaction cannot reuse the first ID.
pub(super) fn allocate_identifier(
    source: &Package,
    working: &mut WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<u64, Error> {
    let mut maximum = 0_u64;
    working.for_each_archive(|component, archive| {
        budget.charge_entries(archive.objects.len())?;
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            maximum = maximum.max(identifier);
        }
        if component == METADATA_COMPONENT {
            let payload = unique_message_payload(archive, PACKAGE_METADATA_MESSAGE_TYPE)?;
            let options = metadata_options(source, payload.len(), budget)?;
            let mut visitor = MaximumIdentifierVisitor { maximum };
            let inspection = identity_codec::inspect_package_metadata_with_visitor(
                payload,
                options,
                &mut visitor,
            )
            .map_err(|_| Error::InvalidSource)?;
            maximum = maximum.max(inspection.last_object_identifier());
            maximum = maximum.max(visitor.maximum);
        }
        Ok(())
    })?;

    let mut candidate = maximum.checked_add(1).ok_or(Error::InvalidSource)?;
    loop {
        if working.reserve_identifier(candidate, budget)? {
            return Ok(candidate);
        }
        candidate = candidate.checked_add(1).ok_or(Error::InvalidSource)?;
    }
}

/// UUID allocation variant that includes already staged archives.
pub(super) fn fresh_storage_uuid_in_working_set(
    source: &Package,
    working: &WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<StorageUuid, Error> {
    let limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut comment_message_count = 0usize;
    let mut scanned_object_count = 0usize;
    let mut scanned_message_count = 0usize;
    working.for_each_archive(|_component, archive| {
        scanned_object_count = scanned_object_count
            .checked_add(archive.objects.len())
            .ok_or(Error::InvalidSource)?;
        for object in &archive.objects {
            scanned_message_count = scanned_message_count
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?;
            comment_message_count = comment_message_count
                .checked_add(
                    object
                        .messages
                        .iter()
                        .filter(|message| message.type_ == 3_056)
                        .count(),
                )
                .ok_or(Error::InvalidSource)?;
        }
        Ok(())
    })?;
    budget.charge_entries(scanned_object_count)?;
    budget.charge_wire_work(scanned_message_count.max(1))?;
    let existing_bytes = comment_message_count
        .checked_mul(size_of::<(u64, u64)>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(existing_bytes)?;
    let mut existing = HashSet::new();
    existing
        .try_reserve(comment_message_count)
        .map_err(|_| Error::Allocation {
            amount: existing_bytes,
        })?;
    working.for_each_archive(|_component, archive| {
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != 3_056 {
                    continue;
                }
                let options = comment_decode_options(&message.data, limits)?;
                let (snapshot, report) =
                    comment_storage_codec::decode_comment_storage_archive_with_report(
                        &message.data,
                        options,
                    )
                    .map_err(|_| Error::InvalidSource)?;
                budget.charge_wire_fields(report.fields())?;
                budget.charge_wire_work(report.work_bytes())?;
                budget.charge_nesting(report.max_depth() as usize)?;
                budget.charge_references(report.references())?;
                budget.charge_allocations(message.data.len())?;
                if let Some(uuid) = snapshot.storage_uuid() {
                    if uuid.lower() == 0 && uuid.upper() == 0 {
                        return Err(Error::InvalidSource);
                    }
                    existing.insert((uuid.lower(), uuid.upper()));
                }
            }
        }
        Ok(())
    })?;
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0_u8; 8];
        let mut upper = [0_u8; 8];
        lower.copy_from_slice(&bytes[..8]);
        upper.copy_from_slice(&bytes[8..]);
        let uuid = StorageUuid::from_parts(u64::from_le_bytes(lower), u64::from_le_bytes(upper));
        if uuid.lower() != 0 && uuid.upper() != 0 && existing.insert((uuid.lower(), uuid.upper())) {
            return Ok(uuid);
        }
    }
}

struct MaximumIdentifierVisitor {
    maximum: u64,
}

impl MaximumIdentifierVisitor {
    fn observe(&mut self, identifier: u64) {
        self.maximum = self.maximum.max(identifier);
    }
}

impl identity_codec::PackageMetadataVisitor for MaximumIdentifierVisitor {
    fn visit_component(
        &mut self,
        component: identity_codec::ComponentDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.observe(component.identifier());
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        object: identity_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.observe(object.object_identifier());
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if let Some(identifier) = reference.object_identifier() {
            self.observe(identifier);
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: identity_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.observe(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: identity_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), identity_codec::RewriteError> {
        self.observe(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), identity_codec::RewriteError> {
        self.observe(object_identifier);
        Ok(())
    }
}

fn unique_message_payload(archive: &Archive, message_type: u32) -> Result<&[u8], Error> {
    let mut selected = None;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != message_type {
                continue;
            }
            if selected.replace(message.data.as_slice()).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn metadata_options(
    source: &Package,
    payload_length: usize,
    budget: &mut Budget,
) -> Result<identity_codec::RewriteOptions, Error> {
    let semantic = source.semantic_limits();
    let wire = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let output = payload_length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let fields = payload_length
        .checked_mul(16)
        .and_then(|value| value.checked_add(64))
        .ok_or(Error::InvalidSource)?;
    let work = payload_length
        .checked_mul(128)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let recursion = u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_work(work)?;
    Ok(identity_codec::RewriteOptions::new(
        payload_length.max(1),
        output,
        fields,
        work,
        recursion,
        semantic.max_objects(),
        semantic.max_references(),
        semantic.max_objects(),
    ))
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ComponentWitness {
    identifier: u64,
    locator: Box<str>,
    current: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ExternalWitness {
    source_identifier: u64,
    target_identifier: u64,
    object_identifier: Option<u64>,
    is_weak: Option<bool>,
    versioned: bool,
    unknown_fields: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ObjectUuidWitness {
    component_identifier: u64,
    object_identifier: u64,
    uuid: identity_codec::UuidBits,
    current: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DataReferenceOwnerWitness {
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    count: u32,
    current: bool,
    unknown_fields: bool,
}

#[derive(Debug, Default)]
struct MetadataCensus {
    components: Vec<ComponentWitness>,
    external_references: Vec<ExternalWitness>,
    object_uuids: Vec<ObjectUuidWitness>,
    data_reference_owners: Vec<DataReferenceOwnerWitness>,
    last_identifier: u64,
}

#[derive(Debug, Default)]
struct MetadataCensusCounts {
    components: usize,
    external_references: usize,
    object_uuids: usize,
    data_reference_owners: usize,
    locator_bytes: usize,
}

impl MetadataCensusCounts {
    fn add(value: &mut usize, amount: usize) -> Result<(), identity_codec::RewriteError> {
        *value = value
            .checked_add(amount)
            .ok_or_else(|| identity_codec::RewriteError::allocation(usize::MAX))?;
        Ok(())
    }
}

impl identity_codec::PackageMetadataVisitor for MetadataCensusCounts {
    fn visit_component(
        &mut self,
        component: identity_codec::ComponentDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        Self::add(&mut self.components, 1)?;
        Self::add(&mut self.locator_bytes, component.effective_locator().len())
    }

    fn visit_external_reference(
        &mut self,
        _reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        Self::add(&mut self.external_references, 1)
    }

    fn visit_object_uuid(
        &mut self,
        _binding: identity_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        Self::add(&mut self.object_uuids, 1)
    }

    fn visit_data_reference_owner(
        &mut self,
        _owner: identity_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        Self::add(&mut self.data_reference_owners, 1)
    }
}

impl MetadataCensus {
    fn inspect(
        source: &[u8],
        options: identity_codec::RewriteOptions,
        budget: &mut Budget,
    ) -> Result<Self, Error> {
        let mut counts = MetadataCensusCounts::default();
        let count_inspection =
            identity_codec::inspect_package_metadata_with_visitor(source, options, &mut counts)
                .map_err(|_| Error::InvalidSource)?;
        budget.charge_wire_fields(count_inspection.report().fields())?;
        budget.charge_wire_work(count_inspection.report().work_bytes())?;
        budget.charge_nesting(count_inspection.report().max_depth() as usize)?;
        budget.charge_references(count_inspection.report().references_scanned())?;

        let component_bytes = counts
            .components
            .checked_mul(size_of::<ComponentWitness>())
            .and_then(|bytes| bytes.checked_add(counts.locator_bytes))
            .ok_or(Error::InvalidSource)?;
        let external_bytes = counts
            .external_references
            .checked_mul(size_of::<ExternalWitness>())
            .ok_or(Error::InvalidSource)?;
        let uuid_bytes = counts
            .object_uuids
            .checked_mul(size_of::<ObjectUuidWitness>())
            .ok_or(Error::InvalidSource)?;
        let data_owner_bytes = counts
            .data_reference_owners
            .checked_mul(size_of::<DataReferenceOwnerWitness>())
            .ok_or(Error::InvalidSource)?;
        let allocation_bytes = component_bytes
            .checked_add(external_bytes)
            .and_then(|bytes| bytes.checked_add(uuid_bytes))
            .and_then(|bytes| bytes.checked_add(data_owner_bytes))
            .ok_or(Error::InvalidSource)?;
        let allocation_events = counts
            .components
            .checked_add(4)
            .ok_or(Error::InvalidSource)?;
        budget.charge_allocation_plan(allocation_bytes, allocation_events)?;

        let mut visitor = Self::default();
        visitor
            .components
            .try_reserve_exact(counts.components)
            .map_err(|_| Error::Allocation {
                amount: component_bytes,
            })?;
        visitor
            .external_references
            .try_reserve_exact(counts.external_references)
            .map_err(|_| Error::Allocation {
                amount: external_bytes,
            })?;
        visitor
            .object_uuids
            .try_reserve_exact(counts.object_uuids)
            .map_err(|_| Error::Allocation { amount: uuid_bytes })?;
        visitor
            .data_reference_owners
            .try_reserve_exact(counts.data_reference_owners)
            .map_err(|_| Error::Allocation {
                amount: data_owner_bytes,
            })?;
        let inspection =
            identity_codec::inspect_package_metadata_with_visitor(source, options, &mut visitor)
                .map_err(|_| Error::InvalidSource)?;
        visitor.last_identifier = inspection.last_object_identifier();
        budget.charge_wire_fields(inspection.report().fields())?;
        budget.charge_wire_work(inspection.report().work_bytes())?;
        budget.charge_nesting(inspection.report().max_depth() as usize)?;
        budget.charge_references(inspection.report().references_scanned())?;
        Ok(visitor)
    }

    fn current_component(&self, name: &str, budget: &mut Budget) -> Result<ComponentRef, Error> {
        let locator = metadata_locator(name).ok_or(Error::InvalidSource)?;
        budget.charge_wire_work(self.components.len().max(1))?;
        let mut selected = None;
        for component in &self.components {
            if !component.current || component.locator.as_ref() != locator {
                continue;
            }
            let allocation = size_of::<ComponentRef>()
                .checked_add(component.locator.len())
                .ok_or(Error::InvalidSource)?;
            budget.charge_allocations(allocation)?;
            if selected
                .replace(ComponentRef::new(component.identifier, &component.locator))
                .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
        selected.ok_or(Error::InvalidSource)
    }

    fn current_component_identifier(&self, name: &str, budget: &mut Budget) -> Result<u64, Error> {
        let locator = metadata_locator(name).ok_or(Error::InvalidSource)?;
        budget.charge_wire_work(self.components.len().max(1))?;
        let mut selected = None;
        for component in &self.components {
            if !component.current || component.locator.as_ref() != locator {
                continue;
            }
            if selected.replace(component.identifier).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        selected.ok_or(Error::InvalidSource)
    }

    fn component_by_identifier(
        &self,
        identifier: u64,
        budget: &mut Budget,
    ) -> Result<ComponentRef, Error> {
        budget.charge_wire_work(self.components.len().max(1))?;
        let mut selected = None;
        for component in &self.components {
            if !component.current || component.identifier != identifier {
                continue;
            }
            let allocation = size_of::<ComponentRef>()
                .checked_add(component.locator.len())
                .ok_or(Error::InvalidSource)?;
            budget.charge_allocations(allocation)?;
            if selected
                .replace(ComponentRef::new(component.identifier, &component.locator))
                .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
        selected.ok_or(Error::InvalidSource)
    }

    fn edge_state(
        &self,
        source: &ComponentRef,
        target: &ComponentRef,
        object_identifier: u64,
        budget: &mut Budget,
    ) -> Result<EdgeState, Error> {
        let mut state = EdgeState::Missing;
        for edge in &self.external_references {
            budget.charge_wire_work(1)?;
            if edge.source_identifier != source.identifier
                || edge.target_identifier != target.identifier
                || edge.object_identifier != Some(object_identifier)
            {
                continue;
            }
            state = match (state, edge.versioned, edge.unknown_fields, edge.is_weak) {
                (EdgeState::Missing, false, false, is_weak) => EdgeState::Current(is_weak),
                (EdgeState::Missing, _, _, _) => EdgeState::Hostile,
                _ => EdgeState::Hostile,
            };
        }
        Ok(state)
    }

    fn object_uuid_state(
        &self,
        component: &ComponentRef,
        object_identifier: u64,
        uuid: identity_codec::UuidBits,
        budget: &mut Budget,
    ) -> Result<ObjectUuidState, Error> {
        let mut state = ObjectUuidState::Missing;
        for binding in &self.object_uuids {
            budget.charge_wire_work(1)?;
            if !binding.current {
                continue;
            }
            if binding.object_identifier == object_identifier {
                if binding.component_identifier != component.identifier || binding.uuid != uuid {
                    return Ok(ObjectUuidState::Hostile);
                }
                state = ObjectUuidState::Current;
            }
            if binding.uuid == uuid
                && (binding.component_identifier != component.identifier
                    || binding.object_identifier != object_identifier)
            {
                return Ok(ObjectUuidState::Hostile);
            }
        }
        Ok(state)
    }

    fn object_uuid_for_object(
        &self,
        component: &ComponentRef,
        object_identifier: u64,
        budget: &mut Budget,
    ) -> Result<Option<identity_codec::UuidBits>, Error> {
        let mut selected = None;
        for binding in &self.object_uuids {
            budget.charge_wire_work(1)?;
            if !binding.current {
                continue;
            }
            if binding.object_identifier == object_identifier
                && binding.component_identifier != component.identifier
            {
                return Err(Error::InvalidSource);
            }
            if binding.object_identifier == object_identifier
                && binding.component_identifier == component.identifier
            {
                if selected.replace(binding.uuid).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
        if let Some(uuid) = selected {
            for binding in &self.object_uuids {
                budget.charge_wire_work(1)?;
                if !binding.current || binding.uuid != uuid {
                    continue;
                }
                if binding.object_identifier != object_identifier
                    || binding.component_identifier != component.identifier
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
        Ok(selected)
    }

    fn component_witness_by_identifier(
        &self,
        identifier: u64,
        budget: &mut Budget,
    ) -> Result<&ComponentWitness, Error> {
        budget.charge_wire_work(self.components.len().max(1))?;
        let mut selected = None;
        for component in &self.components {
            if !component.current || component.identifier != identifier {
                continue;
            }
            if selected.replace(component).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        selected.ok_or(Error::InvalidSource)
    }
}

impl identity_codec::PackageMetadataVisitor for MetadataCensus {
    fn visit_component(
        &mut self,
        component: identity_codec::ComponentDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.components.push(ComponentWitness {
            identifier: component.identifier(),
            locator: component.effective_locator().into(),
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.external_references.push(ExternalWitness {
            source_identifier: reference.source().identifier(),
            target_identifier: reference.target_component_identifier(),
            object_identifier: reference.object_identifier(),
            is_weak: reference.is_weak(),
            versioned: reference.is_versioned() || !reference.source().is_current(),
            unknown_fields: reference.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: identity_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.object_uuids.push(ObjectUuidWitness {
            component_identifier: binding.component().identifier(),
            object_identifier: binding.object_identifier(),
            uuid: binding.uuid(),
            current: binding.component().is_current(),
        });
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: identity_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        self.data_reference_owners.push(DataReferenceOwnerWitness {
            component_identifier: owner.component().identifier(),
            data_identifier: owner.data_identifier(),
            object_identifier: owner.object_identifier(),
            count: owner.count(),
            current: owner.component().is_current(),
            unknown_fields: owner.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        _object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), identity_codec::RewriteError> {
        Ok(())
    }
}

#[derive(Debug, Default)]
struct ArchiveObjectReferenceCounts {
    object_references: usize,
}

impl ArchiveReferenceVisitor for ArchiveObjectReferenceCounts {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object {
            self.object_references = self.object_references.checked_add(1).ok_or(
                litchi_iwa_core::Error::InvalidArchive {
                    offset: 0,
                    reason: "reference count overflow",
                },
            )?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ArchiveObjectReferenceWitness {
    component_identifier: u64,
    referenced_identifier: u64,
}

struct ArchiveObjectReferenceCollector<'a> {
    component_identifier: u64,
    witnesses: &'a mut Vec<ArchiveObjectReferenceWitness>,
}

impl ArchiveReferenceVisitor for ArchiveObjectReferenceCollector<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object {
            self.witnesses.push(ArchiveObjectReferenceWitness {
                component_identifier: self.component_identifier,
                referenced_identifier: occurrence.referenced_identifier,
            });
        }
        Ok(())
    }
}

fn effective_object_reference_census(
    source: &Package,
    working: &WorkingSet<'_>,
    metadata: &MetadataCensus,
    budget: &mut Budget,
) -> Result<Vec<ArchiveObjectReferenceWitness>, Error> {
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut reference_count = 0usize;
    working.for_each_archive(|name, archive| {
        if name == METADATA_COMPONENT {
            return Ok(());
        }
        metadata.current_component_identifier(name, budget)?;
        budget.charge_entries(archive.objects.len())?;
        for object in &archive.objects {
            let mut counts = ArchiveObjectReferenceCounts::default();
            let references = object
                .inspect_references_with_policy_and_limits(
                    &mut counts,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            budget.charge_references(references)?;
            budget.charge_wire_work(references.max(1))?;
            reference_count = reference_count
                .checked_add(counts.object_references)
                .ok_or(Error::InvalidSource)?;
        }
        Ok(())
    })?;

    let witness_bytes = reference_count
        .checked_mul(size_of::<ArchiveObjectReferenceWitness>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(witness_bytes)?;
    let mut witnesses = Vec::new();
    witnesses
        .try_reserve_exact(reference_count)
        .map_err(|_| Error::Allocation {
            amount: witness_bytes,
        })?;
    working.for_each_archive(|name, archive| {
        if name == METADATA_COMPONENT {
            return Ok(());
        }
        let component_identifier = metadata.current_component_identifier(name, budget)?;
        budget.charge_entries(archive.objects.len())?;
        for object in &archive.objects {
            let mut collector = ArchiveObjectReferenceCollector {
                component_identifier,
                witnesses: &mut witnesses,
            };
            let references = object
                .inspect_references_with_policy_and_limits(
                    &mut collector,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            budget.charge_references(references)?;
            budget.charge_wire_work(references.max(1))?;
        }
        Ok(())
    })?;
    if witnesses.len() != reference_count {
        return Err(Error::InvalidSource);
    }
    Ok(witnesses)
}

fn has_surviving_object_reference(
    witnesses: &[ArchiveObjectReferenceWitness],
    component_identifier: u64,
    object_identifier: u64,
    budget: &mut Budget,
) -> Result<bool, Error> {
    budget.charge_wire_work(witnesses.len().max(1))?;
    Ok(witnesses.iter().any(|witness| {
        witness.component_identifier == component_identifier
            && witness.referenced_identifier == object_identifier
    }))
}

fn metadata_payload(
    source: &Package,
    working: &WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<Vec<u8>, Error> {
    let mut payload = None;
    working.for_each_archive(|name, archive| {
        if name != METADATA_COMPONENT {
            return Ok(());
        }
        if payload.is_some() {
            return Err(Error::InvalidSource);
        }
        let source_payload = unique_message_payload(archive, PACKAGE_METADATA_MESSAGE_TYPE)?;
        budget.charge_allocations(source_payload.len())?;
        payload = Some(source_payload.to_vec());
        Ok(())
    })?;
    let payload = payload.ok_or(Error::InvalidSource)?;
    let _ = source;
    Ok(payload)
}

fn unique_message_location_in_archive(
    archive: &Archive,
    message_type: u32,
) -> Result<(u64, usize), Error> {
    let mut selected = None;
    for object in &archive.objects {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != message_type {
                continue;
            }
            if selected.replace((identifier, message_index)).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn charge_identity_report(
    budget: &mut Budget,
    report: identity_codec::RewriteReport,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references_scanned())?;
    budget.charge_output(report.output_bytes())?;
    budget.charge_allocations(report.retained_bytes())?;
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum EdgeState {
    Missing,
    Current(Option<bool>),
    Hostile,
}

#[derive(Debug)]
struct ExternalEdgePlan {
    source: ComponentRef,
    target: ComponentRef,
    object_identifier: u64,
    is_weak: Option<bool>,
}

fn current_external_edge_state(
    census: &MetadataCensus,
    source: &ComponentRef,
    target: &ComponentRef,
    object_identifier: u64,
    budget: &mut Budget,
) -> Result<EdgeState, Error> {
    let mut current = None;
    for edge in &census.external_references {
        budget.charge_wire_work(1)?;
        if edge.source_identifier != source.identifier()
            || edge.target_identifier != target.identifier()
            || edge.object_identifier != Some(object_identifier)
        {
            continue;
        }
        if edge.versioned {
            // Historical records are intentionally preserved by the neutral
            // removal codec and do not make a unique current witness
            // ambiguous.
            continue;
        }
        if edge.unknown_fields {
            return Err(Error::InvalidSource);
        }
        if current.replace(edge.is_weak).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(current.map_or(EdgeState::Missing, EdgeState::Current))
}

struct SourceEdgeStateProbe {
    source_identifier: u64,
    target_identifier: u64,
    object_identifier: u64,
    state: EdgeState,
}

impl identity_codec::PackageMetadataVisitor for SourceEdgeStateProbe {
    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if reference.source().identifier() != self.source_identifier
            || reference.target_component_identifier() != self.target_identifier
            || reference.object_identifier() != Some(self.object_identifier)
            || reference.is_versioned()
            || !reference.source().is_current()
        {
            return Ok(());
        }
        if reference.has_unknown_fields() {
            self.state = EdgeState::Hostile;
        } else if matches!(self.state, EdgeState::Missing) {
            self.state = EdgeState::Current(reference.is_weak());
        } else {
            self.state = EdgeState::Hostile;
        }
        Ok(())
    }
}

fn source_external_edge_state(
    source: &Package,
    source_component: &ComponentRef,
    target_component: &ComponentRef,
    object_identifier: u64,
    budget: &mut Budget,
) -> Result<EdgeState, Error> {
    let source_archive = source
        .state
        .source
        .components()
        .get(METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    let payload = unique_message_payload(source_archive.archive(), PACKAGE_METADATA_MESSAGE_TYPE)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let mut probe = SourceEdgeStateProbe {
        source_identifier: source_component.identifier(),
        target_identifier: target_component.identifier(),
        object_identifier,
        state: EdgeState::Missing,
    };
    let inspection =
        identity_codec::inspect_package_metadata_with_visitor(payload, options, &mut probe)
            .map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(inspection.report().fields())?;
    budget.charge_wire_work(inspection.report().work_bytes())?;
    budget.charge_nesting(inspection.report().max_depth() as usize)?;
    budget.charge_references(inspection.report().references_scanned())?;
    Ok(probe.state)
}

fn push_external_edge_plan(
    plans: &mut Vec<ExternalEdgePlan>,
    source: &ComponentRef,
    target: &ComponentRef,
    object_identifier: u64,
    is_weak: Option<bool>,
    budget: &mut Budget,
) -> Result<(), Error> {
    for existing in plans.iter() {
        budget.charge_wire_work(1)?;
        if existing.source.identifier() != source.identifier()
            || existing.target.identifier() != target.identifier()
            || existing.object_identifier != object_identifier
        {
            continue;
        }
        if existing.is_weak != is_weak {
            return Err(Error::InvalidSource);
        }
        return Ok(());
    }
    plans.push(ExternalEdgePlan {
        source: clone_component_ref(source, budget)?,
        target: clone_component_ref(target, budget)?,
        object_identifier,
        is_weak,
    });
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ObjectUuidState {
    Missing,
    Current,
    Hostile,
}

fn metadata_locator(component: &str) -> Option<&str> {
    component
        .strip_prefix("Index/")
        .and_then(|component| component.strip_suffix(".iwa"))
}

fn copy_boxed_name(value: &str, budget: &mut Budget) -> Result<Box<str>, Error> {
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

fn clone_component_ref(
    component: &ComponentRef,
    budget: &mut Budget,
) -> Result<ComponentRef, Error> {
    budget.charge_allocations(size_of::<ComponentRef>())?;
    Ok(ComponentRef {
        identifier: component.identifier,
        locator: copy_boxed_name(&component.locator, budget)?,
    })
}

fn package_component_name_for_ref<'source>(
    source: &'source Package,
    component: &ComponentRef,
    budget: &mut Budget,
) -> Result<&'source str, Error> {
    let mut selected = None;
    for candidate in source.state.source.components().iter() {
        budget.charge_wire_work(1)?;
        if metadata_locator(candidate.name()) != Some(component.locator.as_ref()) {
            continue;
        }
        if selected.replace(candidate.name()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

/// Rewrite one or more exact author dependency edges in the staged metadata
/// component.  The operation uses the neutral PackageMetadata codec, so
/// versioned records, unknown selected fields, duplicate edges, and locator
/// ambiguity fail before any candidate is installed in the `WorkingSet`.
pub(super) fn rewrite_for_comment_edges(
    source: &Package,
    working: &mut WorkingSet<'_>,
    edit: EdgeEdit,
    budget: &mut Budget,
) -> Result<(), Error> {
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let (source_component, target_component, object_identifier, source_name) = match edit {
        EdgeEdit::AddAuthor {
            component,
            author_identifier,
        } => {
            let source_component = census.current_component(&component, budget)?;
            let Some((target_name, _)) = working.find_object(author_identifier) else {
                return Err(Error::InvalidSource);
            };
            let target_component = census.current_component(target_name, budget)?;
            (
                source_component,
                target_component,
                author_identifier,
                component,
            )
        },
    };

    // Same-component objects are already owned by the component and native
    // Keynote does not add a self-edge to PackageMetadata.
    if source_component.identifier() == target_component.identifier() {
        return Ok(());
    }
    let state = census.edge_state(
        &source_component,
        &target_component,
        object_identifier,
        budget,
    )?;
    match state {
        // Replaying an already committed strong edge is an idempotent
        // operation.  A weak or structurally hostile witness cannot be
        // silently upgraded because that would discard source metadata.
        EdgeState::Current(Some(true)) | EdgeState::Hostile => {
            return Err(Error::InvalidSource);
        },
        EdgeState::Current(None | Some(false)) => return Ok(()),
        EdgeState::Missing => {},
    }
    if object_identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let additions = [identity_codec::ExternalReferenceAddition::new(
        source_component.selector(),
        target_component.selector(),
        object_identifier,
        None,
    )];
    let new_last = census
        .last_identifier
        .checked_add(1)
        .ok_or(Error::InvalidSource)?
        .max(object_identifier);
    let batch = identity_codec::Batch::new(census.last_identifier, new_last, &[], &additions);
    let output = identity_codec::rewrite_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    working.mark_save_token_component(&source_name, budget)?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)
}

/// Reconcile PackageMetadata edges after comment-storage copy-on-write.
///
/// The archive ownership census is intentionally performed once for the whole
/// transition set.  This matters for recursive reply deletion: a removed
/// reply can be the source of a foreign metadata edge even when no explicit
/// drawable transition names that component.  Current records are removed
/// only when no effective ArchiveInfo object-reference witness remains in the
/// same source component.  Historical records and optional native edges stay
/// byte-preserved, while a selected current record with unknown fields or
/// duplicate ownership fails closed.
pub(super) fn reconcile_comment_storage_edges(
    source: &Package,
    working: &mut WorkingSet<'_>,
    transitions: &[CommentStorageEdgeTransition<'_>],
    removed_storage_ids: &[u64],
    budget: &mut Budget,
) -> Result<(), Error> {
    if transitions.is_empty() && removed_storage_ids.is_empty() {
        return Ok(());
    }
    for (index, identifier) in removed_storage_ids.iter().copied().enumerate() {
        budget.charge_wire_work(index.saturating_add(1))?;
        if identifier == 0 || removed_storage_ids[..index].contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }

    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let metadata = MetadataCensus::inspect(&payload, options, budget)?;
    let witnesses = effective_object_reference_census(source, working, &metadata, budget)?;

    let transition_capacity = transitions
        .len()
        .checked_mul(2)
        .ok_or(Error::InvalidSource)?;
    let edge_capacity = metadata
        .external_references
        .len()
        .checked_add(transition_capacity)
        .ok_or(Error::InvalidSource)?;
    let edge_bytes = edge_capacity
        .checked_mul(size_of::<ExternalEdgePlan>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(edge_bytes.checked_mul(2).ok_or(Error::InvalidSource)?)?;
    let mut removals = Vec::new();
    removals
        .try_reserve_exact(edge_capacity)
        .map_err(|_| Error::Allocation { amount: edge_bytes })?;
    let mut additions = Vec::new();
    additions
        .try_reserve_exact(edge_capacity)
        .map_err(|_| Error::Allocation { amount: edge_bytes })?;

    for transition in transitions {
        if transition.source_component.is_empty() {
            return Err(Error::InvalidSource);
        }
        let source_component = metadata.current_component(transition.source_component, budget)?;
        let source_identifier = source_component.identifier();
        let old = match (transition.old_target_component, transition.old_identifier) {
            (None, None) => None,
            (Some(target), Some(identifier)) if identifier != 0 => {
                Some((metadata.current_component(target, budget)?, identifier))
            },
            _ => return Err(Error::InvalidSource),
        };
        let new = match (transition.new_target_component, transition.new_identifier) {
            (None, None) => None,
            (Some(target), Some(identifier)) if identifier != 0 => {
                Some((metadata.current_component(target, budget)?, identifier))
            },
            _ => return Err(Error::InvalidSource),
        };

        let candidate_old_state_result =
            old.as_ref()
                .map_or(Ok(EdgeState::Missing), |(target, identifier)| {
                    current_external_edge_state(
                        &metadata,
                        &source_component,
                        target,
                        *identifier,
                        budget,
                    )
                });
        let candidate_old_state = candidate_old_state_result?;
        let source_old_state =
            old.as_ref()
                .map_or(Ok(EdgeState::Missing), |(target, identifier)| {
                    if source_component.identifier() == target.identifier() {
                        Ok(EdgeState::Missing)
                    } else {
                        source_external_edge_state(
                            source,
                            &source_component,
                            target,
                            *identifier,
                            budget,
                        )
                    }
                })?;
        if matches!(source_old_state, EdgeState::Hostile) {
            return Err(Error::InvalidSource);
        }
        let old_state_for_addition = match (candidate_old_state, source_old_state) {
            (EdgeState::Missing, EdgeState::Current(is_weak)) => EdgeState::Current(is_weak),
            (candidate, _) => candidate,
        };
        let old_witness = old.as_ref().map_or(Ok(false), |(_, identifier)| {
            has_surviving_object_reference(&witnesses, source_identifier, *identifier, budget)
        })?;
        if matches!(candidate_old_state, EdgeState::Hostile) && (new.is_some() || !old_witness) {
            return Err(Error::InvalidSource);
        }

        if let Some((old_target, old_identifier)) = old.as_ref()
            && !old_witness
            && let EdgeState::Current(is_weak) = candidate_old_state
        {
            push_external_edge_plan(
                &mut removals,
                &source_component,
                old_target,
                *old_identifier,
                is_weak,
                budget,
            )?;
        }

        let Some((new_target, new_identifier)) = new.as_ref() else {
            continue;
        };
        if new_target.identifier() == source_identifier {
            continue;
        }
        let new_witness =
            has_surviving_object_reference(&witnesses, source_identifier, *new_identifier, budget)?;
        if !new_witness {
            continue;
        }
        let EdgeState::Current(is_weak) = old_state_for_addition else {
            continue;
        };
        let new_state = current_external_edge_state(
            &metadata,
            &source_component,
            new_target,
            *new_identifier,
            budget,
        )?;
        if matches!(new_state, EdgeState::Hostile) {
            return Err(Error::InvalidSource);
        }
        if matches!(new_state, EdgeState::Missing) {
            push_external_edge_plan(
                &mut additions,
                &source_component,
                new_target,
                *new_identifier,
                is_weak,
                budget,
            )?;
        }
    }

    for edge in &metadata.external_references {
        let Some(object_identifier) = edge.object_identifier else {
            continue;
        };
        budget.charge_wire_work(removed_storage_ids.len().max(1))?;
        if !removed_storage_ids.contains(&object_identifier) || edge.versioned {
            continue;
        }
        let source_component = metadata.component_by_identifier(edge.source_identifier, budget)?;
        if has_surviving_object_reference(
            &witnesses,
            source_component.identifier(),
            object_identifier,
            budget,
        )? {
            continue;
        }
        let target_component = metadata.component_by_identifier(edge.target_identifier, budget)?;
        let state = current_external_edge_state(
            &metadata,
            &source_component,
            &target_component,
            object_identifier,
            budget,
        )?;
        match state {
            EdgeState::Current(is_weak) => {
                push_external_edge_plan(
                    &mut removals,
                    &source_component,
                    &target_component,
                    object_identifier,
                    is_weak,
                    budget,
                )?;
            },
            EdgeState::Missing => {},
            EdgeState::Hostile => return Err(Error::InvalidSource),
        }
    }

    if removals.is_empty() && additions.is_empty() {
        return Ok(());
    }
    let conflict_work = removals
        .len()
        .checked_mul(additions.len())
        .ok_or(Error::InvalidSource)?;
    budget.charge_wire_work(conflict_work.max(1))?;
    for removal in &removals {
        if additions.iter().any(|addition| {
            addition.source.identifier() == removal.source.identifier()
                && addition.target.identifier() == removal.target.identifier()
                && addition.object_identifier == removal.object_identifier
        }) {
            return Err(Error::InvalidSource);
        }
    }

    let request_bytes = removals
        .len()
        .checked_mul(size_of::<identity_codec::ExternalReferenceRemoval<'_>>())
        .and_then(|bytes| {
            bytes.checked_add(
                additions
                    .len()
                    .checked_mul(size_of::<identity_codec::ExternalReferenceAddition<'_>>())?,
            )
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(request_bytes)?;
    let mut removal_requests = Vec::new();
    removal_requests
        .try_reserve_exact(removals.len())
        .map_err(|_| Error::Allocation {
            amount: request_bytes,
        })?;
    for removal in &removals {
        removal_requests.push(identity_codec::ExternalReferenceRemoval::new(
            removal.source.selector(),
            removal.target.selector(),
            removal.object_identifier,
            removal.is_weak,
        ));
    }
    let mut addition_requests = Vec::new();
    addition_requests
        .try_reserve_exact(additions.len())
        .map_err(|_| Error::Allocation {
            amount: request_bytes,
        })?;
    for addition in &additions {
        addition_requests.push(identity_codec::ExternalReferenceAddition::new(
            addition.source.selector(),
            addition.target.selector(),
            addition.object_identifier,
            addition.is_weak,
        ));
    }

    let mut candidate = payload;
    if !removal_requests.is_empty() {
        let batch = identity_codec::RemovalBatch::new(
            metadata.last_identifier,
            &[],
            &removal_requests,
            &[],
        );
        let output = identity_codec::remove_package_metadata(&candidate, batch, options)
            .map_err(|_| Error::InvalidSource)?;
        charge_identity_report(budget, output.report())?;
        candidate = output.into_bytes();
    }
    if !addition_requests.is_empty() {
        let mut new_last = metadata
            .last_identifier
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        for addition in &additions {
            new_last = new_last.max(addition.object_identifier);
        }
        let batch =
            identity_codec::Batch::new(metadata.last_identifier, new_last, &[], &addition_requests);
        let output = identity_codec::rewrite_package_metadata(&candidate, batch, options)
            .map_err(|_| Error::InvalidSource)?;
        charge_identity_report(budget, output.report())?;
        candidate = output.into_bytes();
    }

    for edge in removals.iter().chain(additions.iter()) {
        let name = package_component_name_for_ref(source, &edge.source, budget)?;
        working.mark_save_token_component(name, budget)?;
    }
    replace_metadata_payload(working, candidate, source, budget)
}

/// Advance the package watermark to cover a newly staged object identifier.
/// Equal targets are intentionally a no-op so callers can reserve a sequence
/// of IDs and publish one monotonic watermark after each object insertion.
pub(super) fn reserve_last_identifier(
    working: &mut WorkingSet<'_>,
    identifier: u64,
    budget: &mut Budget,
) -> Result<(), Error> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let source = working.source();
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    if identifier <= census.last_identifier {
        return Ok(());
    }
    let batch = identity_codec::Batch::new(census.last_identifier, identifier, &[], &[]);
    let output = identity_codec::rewrite_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(())
}

/// Advance the root and selected current-component save tokens in one
/// source-bound rewrite.  The caller supplies native component names only as
/// private engine facts; selectors are resolved against current metadata
/// identifiers and effective locators before the codec runs.
pub(super) fn advance_save_tokens(
    source: &Package,
    working: &mut WorkingSet<'_>,
    component_names: &[Box<str>],
    budget: &mut Budget,
) -> Result<(), Error> {
    if component_names.is_empty() {
        return Ok(());
    }
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let component_bytes = component_names
        .len()
        .checked_mul(size_of::<ComponentRef>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(component_bytes)?;
    let mut components = Vec::new();
    components
        .try_reserve_exact(component_names.len())
        .map_err(|_| Error::Allocation {
            amount: component_bytes,
        })?;
    for name in component_names {
        let component = census.current_component(name, budget)?;
        budget.charge_wire_work(components.len().max(1))?;
        if components
            .iter()
            .any(|selected: &ComponentRef| selected == &component)
        {
            return Err(Error::InvalidSource);
        }
        components.push(component);
    }
    let selector_bytes = components
        .len()
        .checked_mul(size_of::<identity_codec::ComponentSelector<'_>>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(selector_bytes)?;
    let mut selectors = Vec::new();
    selectors
        .try_reserve_exact(components.len())
        .map_err(|_| Error::Allocation {
            amount: selector_bytes,
        })?;
    selectors.extend(components.iter().map(ComponentRef::selector));
    let batch = identity_codec::SaveTokenBatch::new(&selectors);
    let output = identity_codec::rewrite_package_metadata_save_tokens(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(())
}

/// Lower the metadata watermark only when the current watermark object is
/// physically removed and no surviving object remains at or above it. Any
/// unrelated non-tail removal leaves the watermark unchanged.
pub(super) fn release_identifier_suffix(
    working: &mut WorkingSet<'_>,
    removed_identifiers: &[u64],
    budget: &mut Budget,
) -> Result<(), Error> {
    if removed_identifiers.is_empty() {
        return Ok(());
    }
    let removed_bytes = removed_identifiers
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(removed_bytes)?;
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(removed_identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: removed_bytes,
        })?;
    removed.extend_from_slice(removed_identifiers);
    removed.sort_unstable();
    if removed.windows(2).any(|window| window[0] >= window[1]) || removed.contains(&0) {
        return Err(Error::InvalidSource);
    }

    let source = working.source();
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let last = census.last_identifier;
    if removed.binary_search(&last).is_err() {
        return Ok(());
    }

    // The working archives already exclude reclaimed objects. Bind the
    // removed watermark to the immutable source, then inspect every
    // effective survivor. A staged clone or a pre-existing object above the
    // current watermark is valid: it means this edit cannot lower the scalar
    // watermark, so leave the metadata payload unchanged after validating the
    // removal set.
    if source.object_with_component(last).is_none() {
        return Err(Error::InvalidSource);
    }
    let mut maximum_remaining = 0_u64;
    working.for_each_archive(|_name, archive| {
        budget.charge_entries(archive.objects.len())?;
        let work = archive
            .objects
            .len()
            .checked_mul(usize::BITS as usize)
            .ok_or(Error::InvalidSource)?;
        budget.charge_wire_work(work.max(1))?;
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if identifier == 0 {
                return Err(Error::InvalidSource);
            }
            if removed.binary_search(&identifier).is_ok() {
                return Err(Error::InvalidSource);
            }
            maximum_remaining = maximum_remaining.max(identifier);
        }
        Ok(())
    })?;
    if maximum_remaining == 0 {
        return Err(Error::InvalidSource);
    }
    if maximum_remaining >= last {
        return Ok(());
    }
    let batch = identity_codec::RemovalBatch::new(last, &[], &[], &[])
        .with_new_last_object_identifier(maximum_remaining);
    // This transition only releases the scalar watermark. The engine advances
    // the affected component save tokens after graph cleanup is complete.
    let output = identity_codec::remove_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(())
}

fn replace_metadata_payload(
    working: &mut WorkingSet<'_>,
    payload: Vec<u8>,
    source: &Package,
    budget: &mut Budget,
) -> Result<(), Error> {
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let archive = working.load(METADATA_COMPONENT, budget)?;
    let (identifier, message_index) =
        unique_message_location_in_archive(archive, PACKAGE_METADATA_MESSAGE_TYPE)?;
    archive
        .object_mut(identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            litchi_iwa_core::RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

#[derive(Debug)]
struct OwnedAuthorFieldReferenceTransition {
    field_info_index: usize,
    path: Vec<u32>,
    before: Vec<u64>,
    after: Vec<u64>,
}

fn copy_author_reference_ids(ids: &[u64], budget: &mut Budget) -> Result<Vec<u64>, Error> {
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

fn copy_author_field_path(path: &[u32], budget: &mut Budget) -> Result<Vec<u32>, Error> {
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

fn transition_author_references(
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
        if identifier == 0 || after.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        after.push(identifier);
    }
    Ok(after)
}

fn prepare_author_storage_header_transition(
    object: &ArchiveObject,
    message_index: usize,
    author_ids: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
    budget: &mut Budget,
) -> Result<(Vec<u64>, Vec<u64>, Vec<OwnedAuthorFieldReferenceTransition>), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let aggregate_before = copy_author_reference_ids(&info.object_references, budget)?;
    let aggregate_after = transition_author_references(&aggregate_before, removed, added, budget)?;
    let field_capacity_bytes = info
        .field_infos
        .len()
        .checked_mul(size_of::<OwnedAuthorFieldReferenceTransition>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(field_capacity_bytes)?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(info.field_infos.len())
        .map_err(|_| Error::Allocation {
            amount: field_capacity_bytes,
        })?;
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        budget.charge_references(field.object_references.len())?;
        budget.charge_wire_work(
            field
                .object_references
                .len()
                .saturating_mul(author_ids.len().max(1)),
        )?;
        let selected = removed
            .is_some_and(|identifier| field.object_references.contains(&identifier))
            || (removed.is_none()
                && added.is_some()
                && field
                    .object_references
                    .iter()
                    .any(|identifier| author_ids.contains(identifier)));
        if !selected {
            continue;
        }
        let before = copy_author_reference_ids(&field.object_references, budget)?;
        let after = transition_author_references(&before, removed, added, budget)?;
        if before == after {
            continue;
        }
        let path = copy_author_field_path(field.path.as_slice(), budget)?;
        fields.push(OwnedAuthorFieldReferenceTransition {
            field_info_index,
            path,
            before,
            after,
        });
    }
    Ok((aggregate_before, aggregate_after, fields))
}

fn replace_author_storage_message(
    object: &mut ArchiveObject,
    message_index: usize,
    payload: Vec<u8>,
    author_ids: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
    limits: litchi_iwa_core::Limits,
    budget: &mut Budget,
) -> Result<(), Error> {
    let (aggregate_before, aggregate_after, owned_fields) =
        prepare_author_storage_header_transition(
            object,
            message_index,
            author_ids,
            removed,
            added,
            budget,
        )?;
    let transition_bytes = owned_fields
        .len()
        .checked_mul(size_of::<
            litchi_iwa_core::archive::FieldObjectReferenceTransition<'_>,
        >())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(transition_bytes)?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(owned_fields.len())
        .map_err(|_| Error::Allocation {
            amount: transition_bytes,
        })?;
    for field in &owned_fields {
        fields.push(litchi_iwa_core::archive::FieldObjectReferenceTransition {
            field_info_index: field.field_info_index,
            expected_path: field.path.as_slice(),
            before: field.before.as_slice(),
            after: field.after.as_slice(),
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            litchi_iwa_core::RawMessage {
                type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                data: payload,
            },
            litchi_iwa_core::archive::ObjectReferenceTransition {
                aggregate_before: aggregate_before.as_slice(),
                aggregate_after: aggregate_after.as_slice(),
                fields: fields.as_slice(),
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

/// Remove every exact current external edge that owns a generated author.
///
/// PackageMetadata stores the edge on the component containing the comment
/// storage, while the author object may live in a different component.  A
/// single `RemoveAuthor` edit therefore cannot clean up a package-wide
/// author: it would leave another slide component pointing at the object
/// that is about to be deleted.  The census below selects all current strong
/// edges to the exact author object, rejects versioned/unknown/weak witnesses,
/// and publishes one atomic metadata rewrite with save-token transitions for
/// every changed source component.
fn remove_author_external_edges(
    source: &Package,
    working: &mut WorkingSet<'_>,
    author_component_name: &str,
    author_identifier: u64,
    budget: &mut Budget,
) -> Result<(), Error> {
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let target = census.current_component(author_component_name, budget)?;

    let mut matching_count = 0usize;
    let mut locator_bytes = 0usize;
    for edge in &census.external_references {
        if edge.target_identifier != target.identifier()
            || edge.object_identifier != Some(author_identifier)
        {
            continue;
        }
        if edge.versioned || edge.unknown_fields || edge.is_weak == Some(true) {
            return Err(Error::InvalidSource);
        }
        let component = census.component_by_identifier(edge.source_identifier, budget)?;
        matching_count = matching_count.checked_add(1).ok_or(Error::InvalidSource)?;
        locator_bytes = locator_bytes
            .checked_add(component.locator.len())
            .ok_or(Error::InvalidSource)?;
    }
    if matching_count == 0 {
        return Ok(());
    }

    let match_bytes = matching_count
        .checked_mul(size_of::<(ComponentRef, Option<bool>)>())
        .and_then(|bytes| bytes.checked_add(locator_bytes))
        .and_then(|bytes| {
            bytes.checked_add(
                matching_count
                    .checked_mul(size_of::<identity_codec::ExternalReferenceRemoval<'_>>())?,
            )
        })
        .and_then(|bytes| {
            bytes.checked_add(
                matching_count.checked_mul(size_of::<identity_codec::ComponentSelector<'_>>())?,
            )
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(match_bytes)?;

    let mut matches = Vec::new();
    matches
        .try_reserve_exact(matching_count)
        .map_err(|_| Error::Allocation {
            amount: match_bytes,
        })?;
    for edge in &census.external_references {
        if edge.target_identifier != target.identifier()
            || edge.object_identifier != Some(author_identifier)
        {
            continue;
        }
        let component = census.component_by_identifier(edge.source_identifier, budget)?;
        if matches
            .iter()
            .any(|(selected, _): &(ComponentRef, Option<bool>)| {
                selected.identifier() == component.identifier()
            })
        {
            return Err(Error::InvalidSource);
        }
        matches.push((component, edge.is_weak));
    }

    let mut removals = Vec::new();
    removals
        .try_reserve_exact(matches.len())
        .map_err(|_| Error::Allocation {
            amount: match_bytes,
        })?;
    for (component, expected_is_weak) in &matches {
        removals.push(identity_codec::ExternalReferenceRemoval::new(
            component.selector(),
            target.selector(),
            author_identifier,
            *expected_is_weak,
        ));
    }

    let batch = identity_codec::RemovalBatch::new(census.last_identifier, &[], &removals, &[]);
    let output = identity_codec::remove_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    for (component, _) in &matches {
        let name = package_component_name_for_ref(source, component, budget)?;
        working.mark_save_token_component(name, budget)?;
    }
    Ok(())
}

fn effective_author_storage_payload(
    source: &Package,
    working: &WorkingSet<'_>,
    budget: &mut Budget,
) -> Result<Option<(AuthorStorageLocation, Vec<u8>)>, Error> {
    let Some(location) = annotation_author_storage_location(source, budget)? else {
        return Ok(None);
    };
    let (component, object) = working
        .find_object(location.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = unique_message_location(
        component,
        object,
        location.object_identifier,
        ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
    )?;
    budget.charge_allocations(message.payload.len())?;
    Ok(Some((location, message.payload.to_vec())))
}

struct AuthorReferenceVisitor {
    author_identifier: u64,
    storage_identifier: u64,
    storage_message_index: usize,
    membership_proven: bool,
    found: bool,
}

impl ArchiveReferenceVisitor for AuthorReferenceVisitor {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if self.membership_proven
            && occurrence.object_identifier == self.storage_identifier
            && occurrence.message_index == self.storage_message_index
            && occurrence.kind == ArchiveReferenceKind::Object
            && occurrence.referenced_identifier == self.author_identifier
        {
            return Ok(());
        }
        if occurrence.kind == ArchiveReferenceKind::Object
            && occurrence.referenced_identifier == self.author_identifier
        {
            self.found = true;
        }
        Ok(())
    }
}

/// Remove a generated author only after a package-wide comment census proves
/// that no surviving comment-storage object still references it.  The source
/// author registry and its component are resolved from the root Document each
/// time, so co-located or staged author storage remains supported.
pub(super) fn cleanup_generated_author_for_identifier(
    source: &Package,
    working: &mut WorkingSet<'_>,
    author_identifier: u64,
    budget: &mut Budget,
) -> Result<bool, Error> {
    if author_identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let Some((location, storage_payload)) =
        effective_author_storage_payload(source, working, budget)?
    else {
        return Ok(false);
    };
    let storage_options = annotation_author_codec::DecodeOptions::for_source(&storage_payload);
    let (storage, storage_report) =
        annotation_author_codec::decode_annotation_author_storage_with_report(
            &storage_payload,
            storage_options,
        )
        .map_err(|_| Error::InvalidSource)?;
    charge_author_decode_report(budget, storage_report)?;

    let author_ref_count = storage.author_ref_count();
    let author_ids_bytes = author_ref_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(author_ids_bytes)?;
    let mut author_ids = Vec::new();
    author_ids
        .try_reserve_exact(author_ref_count)
        .map_err(|_| Error::Allocation {
            amount: author_ids_bytes,
        })?;
    budget.charge_allocations(author_ids_bytes)?;
    let mut seen = HashSet::new();
    seen.try_reserve(author_ref_count)
        .map_err(|_| Error::Allocation {
            amount: author_ids_bytes,
        })?;
    let mut selected_ordinal = None;
    let mut author_component: Option<Box<str>> = None;
    for (ordinal, reference) in storage.author_refs().enumerate() {
        let identifier = reference.identifier();
        if identifier == 0
            || reference.deprecated_type().is_some()
            || reference.deprecated_is_external().is_some()
            || !seen.insert(identifier)
        {
            return Err(Error::InvalidSource);
        }
        author_ids.push(identifier);
        let Some((component, object)) = working.find_object(identifier) else {
            return Err(Error::InvalidSource);
        };
        let author = unique_message_location(
            component,
            object,
            identifier,
            ANNOTATION_AUTHOR_MESSAGE_TYPE,
        )?;
        let options = annotation_author_codec::DecodeOptions::for_source(author.payload);
        let (snapshot, report) =
            annotation_author_codec::decode_annotation_author_with_report(author.payload, options)
                .map_err(|_| Error::InvalidSource)?;
        charge_author_decode_report(budget, report)?;
        if identifier == author_identifier {
            if !is_generated_author(&snapshot, true) {
                return Ok(false);
            }
            selected_ordinal = Some(ordinal);
            budget.charge_allocations(component.len())?;
            author_component = Some(component.into());
        }
    }
    let Some(selected_ordinal) = selected_ordinal else {
        return Ok(false);
    };
    let author_component = author_component.ok_or(Error::InvalidSource)?;

    // The selected storage object's aggregate ArchiveInfo references are an
    // ownership projection of its repeated author list.  Exempt exactly that
    // proven membership edge from the global census; any extra header or
    // nested-field occurrence still keeps the author alive.
    let (_, storage_object) = working
        .find_object(location.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let storage_info = storage_object
        .archive_info
        .message_infos
        .get(location.message_index)
        .ok_or(Error::InvalidSource)?;
    let membership_proven = storage_info.object_references.contains(&author_identifier);
    if membership_proven {
        // ArchiveInfo aggregate references are an optional set projection of
        // the selected payload fields; native writers may order that
        // projection differently from the repeated author list.  Prove the
        // selected membership while rejecting duplicate or unrelated header
        // references.
        let mut aggregate = copy_author_reference_ids(&storage_info.object_references, budget)?;
        aggregate.sort_unstable();
        let aggregate_matches = !aggregate.windows(2).any(|window| window[0] == window[1])
            && aggregate.binary_search(&author_identifier).is_ok()
            && aggregate
                .iter()
                .all(|identifier| author_ids.contains(identifier));
        if !aggregate_matches {
            return Err(Error::InvalidSource);
        }
    }
    if membership_proven
        && storage_info.field_infos.iter().any(|field| {
            field
                .object_references
                .iter()
                .any(|identifier| author_ids.contains(identifier))
                && field
                    .object_references
                    .iter()
                    .any(|identifier| !author_ids.contains(identifier))
        })
    {
        return Err(Error::InvalidSource);
    }

    let wire_limits = source
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut used = false;
    working.for_each_archive(|_name, archive| {
        for object in &archive.objects {
            let mut references = AuthorReferenceVisitor {
                author_identifier,
                storage_identifier: location.object_identifier,
                storage_message_index: location.message_index,
                membership_proven,
                found: false,
            };
            let reference_count = object
                .inspect_references_with_policy_and_limits(
                    &mut references,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            budget.charge_references(reference_count)?;
            budget.charge_wire_work(reference_count.max(1))?;
            if references.found {
                used = true;
            }
            for message in &object.messages {
                if message.type_ != 3_056 {
                    continue;
                }
                let options = comment_decode_options(&message.data, wire_limits)?;
                let (snapshot, report) =
                    comment_storage_codec::decode_comment_storage_archive_with_report(
                        &message.data,
                        options,
                    )
                    .map_err(|_| Error::InvalidSource)?;
                budget.charge_wire_fields(report.fields())?;
                budget.charge_wire_work(report.work_bytes())?;
                budget.charge_nesting(report.max_depth() as usize)?;
                budget.charge_references(report.references())?;
                budget.charge_allocations(message.data.len())?;
                if snapshot
                    .author()
                    .is_some_and(|reference| reference.identifier() == author_identifier)
                {
                    used = true;
                }
            }
        }
        Ok(())
    })?;
    if used {
        return Ok(false);
    }

    // Keep this strict: a failed edge witness must abort the candidate before
    // the author object is removed, otherwise PackageMetadata would retain a
    // dangling external reference.
    remove_author_external_edges(
        source,
        working,
        &author_component,
        author_identifier,
        budget,
    )?;

    let archive_limits = source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let storage_archive = working.load(&location.component, budget)?;
    let storage_object = storage_archive
        .object_mut(location.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let current_storage_message = {
        let message = storage_object
            .messages
            .get(location.message_index)
            .ok_or(Error::InvalidSource)?;
        budget.charge_allocations(message.data.len())?;
        message.data.clone()
    };
    let options = annotation_author_codec::DecodeOptions::for_source(&current_storage_message);
    let prepared = annotation_author_codec::prepare_annotation_author_storage_rewrite(
        &current_storage_message,
        annotation_author_codec::AnnotationAuthorStorageRewrite::remove(
            selected_ordinal,
            author_identifier,
        ),
        options,
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.requirements();
    charge_author_rewrite_requirements(budget, requirements)?;
    let candidate = prepared
        .execute(requirements.exact())
        .map_err(|_| Error::InvalidSource)?
        .0;
    if membership_proven {
        replace_author_storage_message(
            storage_object,
            location.message_index,
            candidate,
            &author_ids,
            Some(author_identifier),
            None,
            archive_limits,
            budget,
        )?;
    } else {
        storage_object
            .replace_message_preserving_header_with_limits(
                location.message_index,
                litchi_iwa_core::RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: candidate,
                },
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
    }

    let archive = working.load(&author_component, budget)?;
    let object = archive
        .object(author_identifier)
        .ok_or(Error::InvalidSource)?;
    let author = unique_message_location(
        &author_component,
        object,
        author_identifier,
        ANNOTATION_AUTHOR_MESSAGE_TYPE,
    )?;
    let options = annotation_author_codec::DecodeOptions::for_source(author.payload);
    let (snapshot, report) =
        annotation_author_codec::decode_annotation_author_with_report(author.payload, options)
            .map_err(|_| Error::InvalidSource)?;
    charge_author_decode_report(budget, report)?;
    if !is_generated_author(&snapshot, true) {
        return Err(Error::InvalidSource);
    }
    archive.remove_object(author_identifier);
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::{
        Budget, ComponentRef, EdgeState, Error, ExternalWitness, MetadataCensus, Package,
        current_external_edge_state, metadata_locator,
    };

    const TEST_SOURCE: &[u8] = include_bytes!(
        "../../../../../test-data/iwork/keynote/drawable-comments-source-native.key"
    );

    fn edge_census(external_references: Vec<ExternalWitness>) -> MetadataCensus {
        MetadataCensus {
            external_references,
            ..MetadataCensus::default()
        }
    }

    fn test_budget() -> Budget {
        let package = Package::from_bytes(TEST_SOURCE).expect("native drawable-comment package");
        Budget::for_package(&package).expect("drawable-comment test budget")
    }

    #[test]
    fn metadata_locator_matches_native_component_entry_names() {
        assert_eq!(metadata_locator("Index/Document.iwa"), Some("Document"));
        assert_eq!(metadata_locator("Index/Metadata.iwa"), Some("Metadata"));
        assert_eq!(metadata_locator("Document.iwa"), None);
        assert_eq!(metadata_locator("Index/Document"), None);
    }

    #[test]
    fn current_external_edge_state_rejects_unknown_current_records() {
        let source = ComponentRef::new(11, "Slide");
        let target = ComponentRef::new(22, "AnnotationAuthorStorage");
        let census = edge_census(vec![ExternalWitness {
            source_identifier: source.identifier(),
            target_identifier: target.identifier(),
            object_identifier: Some(33),
            is_weak: None,
            versioned: false,
            unknown_fields: true,
        }]);
        let mut budget = test_budget();

        assert!(matches!(
            current_external_edge_state(&census, &source, &target, 33, &mut budget),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn current_external_edge_state_rejects_duplicate_current_records() {
        let source = ComponentRef::new(11, "Slide");
        let target = ComponentRef::new(22, "AnnotationAuthorStorage");
        let witness = ExternalWitness {
            source_identifier: source.identifier(),
            target_identifier: target.identifier(),
            object_identifier: Some(33),
            is_weak: None,
            versioned: false,
            unknown_fields: false,
        };
        let census = edge_census(vec![witness, witness]);
        let mut budget = test_budget();

        assert!(matches!(
            current_external_edge_state(&census, &source, &target, 33, &mut budget),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn current_external_edge_state_preserves_historical_records() {
        let source = ComponentRef::new(11, "Slide");
        let target = ComponentRef::new(22, "AnnotationAuthorStorage");
        let historical = ExternalWitness {
            source_identifier: source.identifier(),
            target_identifier: target.identifier(),
            object_identifier: Some(33),
            is_weak: None,
            versioned: true,
            unknown_fields: true,
        };

        let mut historical_only_budget = test_budget();
        assert_eq!(
            current_external_edge_state(
                &edge_census(vec![historical]),
                &source,
                &target,
                33,
                &mut historical_only_budget,
            )
            .expect("historical records are not current witnesses"),
            EdgeState::Missing
        );

        let current = ExternalWitness {
            versioned: false,
            unknown_fields: false,
            is_weak: Some(false),
            ..historical
        };
        let mut mixed_budget = test_budget();
        assert_eq!(
            current_external_edge_state(
                &edge_census(vec![historical, current]),
                &source,
                &target,
                33,
                &mut mixed_budget,
            )
            .expect("historical records must not make a valid current edge hostile"),
            EdgeState::Current(Some(false))
        );
    }
}
