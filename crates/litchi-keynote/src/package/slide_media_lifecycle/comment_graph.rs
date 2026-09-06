//! Strict, source-bound planning for a drawable comment graph.
//!
//! The media lifecycle owner cannot clone or remove a drawable while leaving
//! an opaque comment/reply edge behind.  This adapter therefore resolves the
//! complete rooted `TSD.CommentStorageArchive` graph before a transaction is
//! allowed to touch the drawable.  Storage objects and replies belong to the
//! selected component; annotation authors are preserved dependencies and are
//! deliberately excluded from the storage clone/removal closure.
//!
//! This module supplies the private plan consumed by the lifecycle clone
//! transaction.  It does not rewrite a drawable, clone a storage object,
//! reclaim an author, or expose a public comment API; the transaction owner
//! remains responsible for applying the validated closure atomically.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The private planner keeps source admission, strict decoding, and closure facts together."
)]

use std::mem::size_of;

use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{ArchiveObject, Limits as ArchiveLimits, MessageInfo};
use litchi_iwa_protos::comment_storage_codec;

use super::budget::LifecycleBudget;
use super::{Package, SlideMediaLifecycleError, SuperUuid};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;

const COMMENT_TEXT_FIELD: u32 = 1;
const COMMENT_DATE_FIELD: u32 = 2;
const COMMENT_AUTHOR_FIELD: u32 = 3;
const COMMENT_REPLIES_FIELD: u32 = 4;
const COMMENT_UUID_FIELD: u32 = 5;
const DATE_SECONDS_FIELD: u32 = 1;
const WIRE_VIEW_SPAN_BYTES_PER_INPUT_BYTE: usize = 64;

const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;

/// One source storage identity retained for the future clone/removal seam.
///
/// `uuid` stays optional because the native archive declares `storage_uuid`
/// optional.  When present, the value is source-authoritative and may be
/// shared by native copy-on-write roots; it is not a package-wide uniqueness
/// witness.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CommentStorageIdentity {
    pub(super) identifier: u64,
    pub(super) uuid: Option<SuperUuid>,
}

/// One comment storage to annotation-author edge.
///
/// The author component is retained so a later transaction can preserve the
/// exact source dependency without accidentally treating the author as a
/// clone root.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct CommentAuthorDependency {
    pub(super) storage_identifier: u64,
    pub(super) author_identifier: u64,
    pub(super) component_name: Box<str>,
}

/// Complete rooted comment/reply closure for one selected drawable.
///
/// Every storage identifier is sorted and belongs to `component_name`.
/// `author_ids` is sorted and unique, while `author_dependencies` retains the
/// per-storage edge and the exact source component of each dependency.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct CommentGraphPlan {
    pub(super) root_storage_identifier: u64,
    pub(super) component_name: Box<str>,
    pub(super) storage_ids: Vec<u64>,
    pub(super) storage_identities: Vec<CommentStorageIdentity>,
    pub(super) author_ids: Vec<u64>,
    pub(super) author_dependencies: Vec<CommentAuthorDependency>,
}

impl CommentGraphPlan {
    /// Return the selected root's source UUID, when the root encoded one.
    #[must_use]
    pub(super) fn root_storage_uuid(&self) -> Option<SuperUuid> {
        self.storage_identities
            .iter()
            .find(|identity| identity.identifier == self.root_storage_identifier)
            .and_then(|identity| identity.uuid)
    }
}

/// Plan the complete same-component comment/reply closure rooted at one
/// selected drawable comment storage object.
///
/// The caller supplies the component already proven to own the selected
/// drawable.  Every storage lookup is repeated against the exact immutable
/// [`Package`] object and must resolve to that component.  Author objects are
/// looked up in the same package, validated as type-212 objects, and returned
/// as dependencies only; they are never added to `storage_ids`.
pub(super) fn plan_comment_graph(
    package: &Package,
    component_name: &str,
    root_storage_identifier: u64,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<CommentGraphPlan, SlideMediaLifecycleError> {
    if component_name.is_empty() || root_storage_identifier == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let (actual_component, _) = package
        .object_with_component(root_storage_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if actual_component != component_name {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let component_name = copy_boxed(component_name, budget)?;
    let mut pending = Vec::new();
    reserve_vec(&mut pending, 1, budget)?;
    pending.push(root_storage_identifier);

    let mut storage_ids = Vec::new();
    insert_sorted_unique(&mut storage_ids, root_storage_identifier, budget)?;
    budget.charge_entries(1)?;

    let mut storage_identities = Vec::new();
    let mut author_ids = Vec::new();
    let mut author_dependencies = Vec::new();

    while let Some(storage_identifier) = pending.pop() {
        let facts = read_storage_facts(
            package,
            component_name.as_ref(),
            storage_identifier,
            limits,
            budget,
        )?;

        push_storage_identity(
            &mut storage_identities,
            CommentStorageIdentity {
                identifier: storage_identifier,
                uuid: facts.uuid,
            },
            budget,
        )?;

        if let Some(author_identifier) = facts.author_identifier {
            let (author_component, author) = package
                .object_with_component(author_identifier)
                .ok_or(SlideMediaLifecycleError::InvalidSource)?;
            if author_ids.binary_search(&author_identifier).is_err() {
                validate_author_object(
                    author_identifier,
                    author,
                    package
                        .limits()
                        .effective_archive_limits()
                        .map_err(|_| SlideMediaLifecycleError::InvalidSource)?,
                    budget,
                )?;
                insert_sorted_unique(&mut author_ids, author_identifier, budget)?;
                budget.charge_entries(1)?;
            }
            push_vec(
                &mut author_dependencies,
                CommentAuthorDependency {
                    storage_identifier,
                    author_identifier,
                    component_name: copy_boxed(author_component, budget)?,
                },
                budget,
            )?;
        }

        for reply_identifier in facts.reply_identifiers {
            // `storage_ids` contains pending nodes as soon as they are
            // discovered.  Seeing one again therefore catches both a direct
            // duplicate and any recursive cycle, including a reply to root.
            if reply_identifier == 0 || storage_ids.binary_search(&reply_identifier).is_ok() {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            insert_sorted_unique(&mut storage_ids, reply_identifier, budget)?;
            budget.charge_entries(1)?;
            push_vec(&mut pending, reply_identifier, budget)?;
        }
    }

    storage_identities.sort_unstable_by_key(|identity| identity.identifier);
    if storage_identities
        .windows(2)
        .any(|window| window[0].identifier >= window[1].identifier)
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    author_dependencies.sort_unstable_by(|left, right| {
        (left.storage_identifier, left.author_identifier)
            .cmp(&(right.storage_identifier, right.author_identifier))
    });
    if author_dependencies.windows(2).any(|window| {
        (window[0].storage_identifier, window[0].author_identifier)
            >= (window[1].storage_identifier, window[1].author_identifier)
    }) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    if storage_identities.len() != storage_ids.len() {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    Ok(CommentGraphPlan {
        root_storage_identifier,
        component_name,
        storage_ids,
        storage_identities,
        author_ids,
        author_dependencies,
    })
}

#[derive(Debug)]
pub(super) struct StorageFacts {
    pub(super) author_identifier: Option<u64>,
    pub(super) reply_identifiers: Vec<u64>,
    pub(super) uuid: Option<SuperUuid>,
}

fn read_storage_facts(
    package: &Package,
    component_name: &str,
    storage_identifier: u64,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<StorageFacts, SlideMediaLifecycleError> {
    let (actual_component, object) = package
        .object_with_component(storage_identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if actual_component != component_name
        || object.archive_info.identifier != Some(storage_identifier)
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let (info, payload) = exact_comment_storage_payload(
        object,
        storage_identifier,
        package
            .limits()
            .effective_archive_limits()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?,
        budget,
    )?;
    decode_storage_payload(storage_identifier, info, payload, limits, budget, true)
}

fn exact_comment_storage_payload<'source>(
    object: &'source ArchiveObject,
    storage_identifier: u64,
    limits: ArchiveLimits,
    budget: &mut LifecycleBudget,
) -> Result<(&'source MessageInfo, &'source [u8]), SlideMediaLifecycleError> {
    if object.archive_info.identifier != Some(storage_identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let message = object
        .messages
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || info.type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || usize::try_from(info.length).ok() != Some(message.data.len())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    budget.charge_wire_fields(
        1usize
            .checked_add(info.field_infos.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?,
    )?;
    budget.charge_wire_work(message.data.len().max(1))?;
    validate_archive_reference_shape(info)?;
    let header_limits = super::graph::reserve_core_header_inspection_work(object, limits, budget)?;
    super::graph::strict_reference_census(object, header_limits, budget)?;
    Ok((info, message.data.as_slice()))
}

/// Validate one comment-storage payload against its ArchiveInfo references
/// without requiring opaque root extensions to be understood.  The selected
/// clone path remains strict; removal scans use this neutral semantic census
/// for untouched storage objects and preserve their source bytes.
pub(super) fn validate_comment_storage_payload_relationship(
    storage_identifier: u64,
    info: &MessageInfo,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<StorageFacts, SlideMediaLifecycleError> {
    if storage_identifier == 0
        || info.type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || usize::try_from(info.length).ok() != Some(payload.len())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    validate_archive_reference_shape(info)?;
    decode_storage_payload(storage_identifier, info, payload, limits, budget, false)
}

fn decode_storage_payload(
    storage_identifier: u64,
    info: &MessageInfo,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    reject_unknown_root_fields: bool,
) -> Result<StorageFacts, SlideMediaLifecycleError> {
    // The neutral codec owns a bounded lazy projection and may allocate its
    // own temporary Buffa state.  Reserve an operation-wide upper bound before
    // entering it so this external allocation remains charged to the shared
    // lifecycle ledger as well.
    budget.charge_allocation_plan(payload.len().max(1), 1)?;
    let options = comment_decode_options(payload, limits)?;
    let mut visitor = ReplyCollector::new(budget);
    let decoded = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        payload,
        options,
        &mut visitor,
    );
    let (reply_identifiers, callback_error) = visitor.finish();
    let (snapshot, report) = decoded.map_err(map_comment_decode_error)?;

    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references())?;
    if report.replies() != reply_identifiers.len() {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    if let Some(error) = callback_error {
        return Err(error);
    }

    validate_comment_payload_wire(payload, limits, budget, reject_unknown_root_fields)?;

    let author_identifier = snapshot.author().map(|reference| {
        if reference.deprecated_type().is_some() || reference.deprecated_is_external().is_some() {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        nonzero_identifier(reference.identifier())
    });
    let author_identifier = match author_identifier {
        Some(result) => Some(result?),
        None => None,
    };

    let uuid = snapshot.storage_uuid().map(|uuid| {
        if uuid.lower() == 0 && uuid.upper() == 0 {
            Err(SlideMediaLifecycleError::InvalidSource)
        } else {
            Ok(SuperUuid {
                lower: uuid.lower(),
                upper: uuid.upper(),
            })
        }
    });
    let uuid = match uuid {
        Some(result) => Some(result?),
        None => None,
    };

    let mut reply_identifiers = reply_identifiers;
    reply_identifiers.sort_unstable();
    if reply_identifiers
        .windows(2)
        .any(|window| window[0] == window[1])
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    if reply_identifiers.contains(&0) {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let facts = StorageFacts {
        author_identifier,
        reply_identifiers,
        uuid,
    };
    validate_storage_metadata(storage_identifier, info, &facts, budget)?;
    Ok(facts)
}

struct ReplyCollector<'budget> {
    identifiers: Vec<u64>,
    budget: &'budget mut LifecycleBudget,
    failure: Option<SlideMediaLifecycleError>,
}

impl<'budget> ReplyCollector<'budget> {
    fn new(budget: &'budget mut LifecycleBudget) -> Self {
        Self {
            identifiers: Vec::new(),
            budget,
            failure: None,
        }
    }

    fn finish(self) -> (Vec<u64>, Option<SlideMediaLifecycleError>) {
        (self.identifiers, self.failure)
    }
}

impl comment_storage_codec::CommentStorageVisitor for ReplyCollector<'_> {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.failure.is_some() {
            return Ok(());
        }
        let identifier = reply.identifier();
        if identifier == 0
            || reply.reference().deprecated_type().is_some()
            || reply.reference().deprecated_is_external().is_some()
        {
            self.failure = Some(SlideMediaLifecycleError::InvalidSource);
            return Ok(());
        }
        if self.identifiers.len() == self.identifiers.capacity() {
            if let Err(error) = reserve_vec(&mut self.identifiers, 1, self.budget) {
                self.failure = Some(error);
                return Ok(());
            }
        }
        self.identifiers.push(identifier);
        Ok(())
    }
}

fn validate_author_object(
    author_identifier: u64,
    object: &ArchiveObject,
    limits: ArchiveLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if object.archive_info.identifier != Some(author_identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    let message = object
        .messages
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .first()
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    if message.type_ != ANNOTATION_AUTHOR_MESSAGE_TYPE
        || info.type_ != ANNOTATION_AUTHOR_MESSAGE_TYPE
        || usize::try_from(info.length).ok() != Some(message.data.len())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    validate_archive_reference_shape(info)?;
    let header_limits = super::graph::reserve_core_header_inspection_work(object, limits, budget)?;
    super::graph::strict_reference_census(object, header_limits, budget)?;
    budget.charge_entries(1)?;
    budget.charge_wire_fields(
        1usize
            .checked_add(info.field_infos.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)?,
    )?;
    budget.charge_wire_work(message.data.len().max(1))?;
    Ok(())
}

fn validate_archive_reference_shape(info: &MessageInfo) -> Result<(), SlideMediaLifecycleError> {
    if !info.data_references.is_empty()
        || info
            .field_infos
            .iter()
            .any(|field| !field.data_references.is_empty())
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    if info.object_references.contains(&0)
        || info
            .field_infos
            .iter()
            .flat_map(|field| &field.object_references)
            .any(|reference| *reference == 0)
    {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }
    Ok(())
}

fn validate_storage_metadata(
    storage_identifier: u64,
    info: &MessageInfo,
    facts: &StorageFacts,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if storage_identifier == 0 {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    // The aggregate header must match the decoded references exactly, while
    // field-level records may partition that same multiset.  Known nested
    // ownership fields were checked by the wire pass; opaque root extensions
    // remain source bytes and do not enter this reference census.
    let expected_len = usize::from(facts.author_identifier.is_some())
        .checked_add(facts.reply_identifiers.len())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    let mut expected = Vec::new();
    reserve_vec(&mut expected, expected_len, budget)?;
    if let Some(author_identifier) = facts.author_identifier {
        push_vec(&mut expected, author_identifier, budget)?;
    }
    for &reply_identifier in &facts.reply_identifiers {
        push_vec(&mut expected, reply_identifier, budget)?;
    }
    expected.sort_unstable();

    let mut aggregate = Vec::new();
    reserve_vec(&mut aggregate, info.object_references.len(), budget)?;
    for &reference in &info.object_references {
        if reference == 0 {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        push_vec(&mut aggregate, reference, budget)?;
    }
    aggregate.sort_unstable();
    if aggregate != expected {
        return Err(SlideMediaLifecycleError::InvalidSource);
    }

    let field_reference_count = info.field_infos.iter().try_fold(0usize, |count, field| {
        count
            .checked_add(field.object_references.len())
            .ok_or(SlideMediaLifecycleError::InvalidSource)
    })?;
    let metadata_reference_count = info
        .object_references
        .len()
        .checked_add(field_reference_count)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_references(metadata_reference_count)?;
    let mut field_references = Vec::new();
    reserve_vec(&mut field_references, field_reference_count, budget)?;
    for field in &info.field_infos {
        for &reference in &field.object_references {
            if reference == 0 {
                return Err(SlideMediaLifecycleError::InvalidSource);
            }
            push_vec(&mut field_references, reference, budget)?;
        }
    }
    field_references.sort_unstable();

    let mut aggregate_index = 0usize;
    for reference in field_references {
        while aggregate_index < aggregate.len() && aggregate[aggregate_index] < reference {
            aggregate_index += 1;
        }
        if aggregate_index == aggregate.len() || aggregate[aggregate_index] != reference {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
        aggregate_index += 1;
    }
    Ok(())
}

fn validate_comment_payload_wire(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut LifecycleBudget,
    reject_unknown_root_fields: bool,
) -> Result<(), SlideMediaLifecycleError> {
    let root = parse_wire_view(payload, limits, 1, budget)?;
    for field in root.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        match field.number() {
            COMMENT_TEXT_FIELD => require_wire_type(field.wire_type(), 2)?,
            COMMENT_DATE_FIELD => {
                require_wire_type(field.wire_type(), 2)?;
                validate_nested_fields(
                    field.payload(),
                    limits,
                    2,
                    budget,
                    |number, wire_type| number == DATE_SECONDS_FIELD && wire_type == 1,
                    true,
                )?;
            },
            COMMENT_AUTHOR_FIELD | COMMENT_REPLIES_FIELD => {
                require_wire_type(field.wire_type(), 2)?;
                validate_nested_fields(
                    field.payload(),
                    limits,
                    2,
                    budget,
                    |number, wire_type| {
                        (number == REFERENCE_IDENTIFIER_FIELD
                            || number == REFERENCE_DEPRECATED_TYPE_FIELD
                            || number == REFERENCE_EXTERNAL_FIELD)
                            && wire_type == 0
                    },
                    true,
                )?;
            },
            COMMENT_UUID_FIELD => {
                require_wire_type(field.wire_type(), 2)?;
                validate_nested_fields(
                    field.payload(),
                    limits,
                    2,
                    budget,
                    |number, wire_type| {
                        (number == UUID_LOWER_FIELD || number == UUID_UPPER_FIELD) && wire_type == 0
                    },
                    true,
                )?;
            },
            // Selected comment storage is admitted to a clone transaction only
            // when every payload field has a known ownership shape.  The
            // neutral codec still preserves these bytes during the later
            // rewrite; unselected existing comments are never routed here.
            _ if reject_unknown_root_fields => {
                return Err(SlideMediaLifecycleError::InvalidSource);
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_nested_fields<F>(
    payload: &[u8],
    limits: WireLimits,
    nesting: usize,
    budget: &mut LifecycleBudget,
    is_known: F,
    reject_unknown: bool,
) -> Result<(), SlideMediaLifecycleError>
where
    F: Fn(u32, u8) -> bool,
{
    let view = parse_wire_view(payload, limits, nesting, budget)?;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
        if reject_unknown && !is_known(field.number(), field.wire_type()) {
            return Err(SlideMediaLifecycleError::InvalidSource);
        }
    }
    Ok(())
}

fn parse_wire_view<'source>(
    payload: &'source [u8],
    limits: WireLimits,
    nesting: usize,
    budget: &mut LifecycleBudget,
) -> Result<WireView<'source>, SlideMediaLifecycleError> {
    budget.charge_wire_work(payload.len().max(1))?;
    budget.charge_nesting(nesting.max(1))?;
    let span_allocation = payload
        .len()
        .checked_mul(WIRE_VIEW_SPAN_BYTES_PER_INPUT_BYTE)
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocations(span_allocation)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.charge_wire_fields(view.len())?;
    Ok(view)
}

fn require_wire_type(actual: u8, expected: u8) -> Result<(), SlideMediaLifecycleError> {
    (actual == expected)
        .then_some(())
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn nonzero_identifier(identifier: u64) -> Result<u64, SlideMediaLifecycleError> {
    (identifier != 0)
        .then_some(identifier)
        .ok_or(SlideMediaLifecycleError::InvalidSource)
}

fn comment_decode_options(
    payload: &[u8],
    limits: WireLimits,
) -> Result<comment_storage_codec::DecodeOptions, SlideMediaLifecycleError> {
    if payload.len() > limits.max_input_bytes() {
        return Err(SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::InputBytes,
            observed: payload.len() as u64,
            maximum: limits.max_input_bytes() as u64,
        });
    }
    let source = payload.len().max(1);
    let fields = source.min(limits.max_fields()).max(1);
    let work = source
        .checked_mul(32)
        .unwrap_or(usize::MAX)
        .min(limits.max_rewrite_work())
        .max(1);
    let nesting =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMediaLifecycleError::InvalidSource)?;
    let references = source.min(limits.max_fields()).max(1);
    let text = source.min(limits.max_input_bytes()).max(1);
    Ok(comment_storage_codec::DecodeOptions::new(
        source, fields, work, nesting, references, text,
    ))
}

fn map_comment_decode_error(error: comment_storage_codec::DecodeError) -> SlideMediaLifecycleError {
    use comment_storage_codec::DecodeLimit;
    let Some(limit) = error.resource_limit() else {
        return SlideMediaLifecycleError::InvalidSource;
    };
    match limit {
        DecodeLimit::Bytes { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::InputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::OutputBytes { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::OutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::References { observed, maximum }
        | DecodeLimit::Replies { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::References,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::ReferenceBytes { observed, maximum }
        | DecodeLimit::Work { observed, maximum }
        | DecodeLimit::Scratch { observed, maximum }
        | DecodeLimit::Retained { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Text { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Fields { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Nesting { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        DecodeLimit::Allocations { observed, maximum } => SlideMediaLifecycleError::LimitExceeded {
            kind: super::SlideMediaLifecycleLimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideMediaLifecycleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => {
            let kind = match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    super::SlideMediaLifecycleLimitKind::InputBytes
                },
                litchi_iwa_common::LimitKind::Fields => {
                    super::SlideMediaLifecycleLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::OutputBytes => {
                    super::SlideMediaLifecycleLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    super::SlideMediaLifecycleLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork => {
                    super::SlideMediaLifecycleLimitKind::WireWork
                },
                _ => return SlideMediaLifecycleError::InvalidSource,
            };
            SlideMediaLifecycleError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideMediaLifecycleError::Allocation { amount }
        },
        _ => SlideMediaLifecycleError::InvalidSource,
    }
}

fn copy_boxed(
    value: &str,
    budget: &mut LifecycleBudget,
) -> Result<Box<str>, SlideMediaLifecycleError> {
    budget.charge_allocations(value.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| SlideMediaLifecycleError::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

fn reserve_vec<T>(
    output: &mut Vec<T>,
    additional: usize,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if additional == 0 {
        return Ok(());
    }
    let amount = additional
        .checked_mul(size_of::<T>())
        .ok_or(SlideMediaLifecycleError::InvalidSource)?;
    budget.charge_allocations(amount)?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| SlideMediaLifecycleError::Allocation { amount })
}

fn push_vec<T>(
    output: &mut Vec<T>,
    value: T,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    if output.len() == output.capacity() {
        reserve_vec(output, 1, budget)?;
    }
    output.push(value);
    Ok(())
}

fn insert_sorted_unique(
    output: &mut Vec<u64>,
    value: u64,
    budget: &mut LifecycleBudget,
) -> Result<bool, SlideMediaLifecycleError> {
    match output.binary_search(&value) {
        Ok(_) => Ok(false),
        Err(index) => {
            budget.charge_wire_work(output.len().max(1))?;
            if output.len() == output.capacity() {
                reserve_vec(output, 1, budget)?;
            }
            output.insert(index, value);
            Ok(true)
        },
    }
}

fn push_storage_identity(
    output: &mut Vec<CommentStorageIdentity>,
    identity: CommentStorageIdentity,
    budget: &mut LifecycleBudget,
) -> Result<(), SlideMediaLifecycleError> {
    push_vec(output, identity, budget)
}

#[cfg(test)]
mod tests;
