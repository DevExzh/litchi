//! Bounded, selector-first graph reads for Keynote drawable comments.
//!
//! This adapter is deliberately independent of generated drawable messages.
//! It recognizes only the native wrapper routes that the legacy editor knew,
//! follows their comment reference as a lazy wire projection, and leaves every
//! other payload opaque.  The returned selection owns identities and source
//! positions; payload bytes are read from the immutable [`Package`] only by
//! the transaction engine.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The graph keeps route admission, source-order selection, and closure validation together."
)]

use std::collections::HashSet;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{
    ArchiveLimits, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, MessageInfo,
};
use litchi_iwa_protos::{annotation_author_codec, comment_storage_codec};

use super::super::Package;
use super::{Budget, DrawableKind, DrawableSelector, DrawableSummary, Error};
use crate::SlideSelector;
use crate::package::slide_media_lifecycle::comment_graph::{
    CommentGraphPlan, plan_comment_graph_cross_component,
};
use crate::slide::comment::{Comment, CommentAuthor, CommentTimestamp, Reply};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;

const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;

/// An owned, source-bound drawable selection used by the comment engine.
///
/// No wire payload is retained here.  This is important for copy-on-write:
/// the engine can load exactly the selected component and still use this
/// value after any borrowed wire views have expired.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Selection {
    pub(super) slide_position: Position,
    pub(super) drawable_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) component_name: Arc<str>,
    pub(super) drawable_identifier: u64,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) comment_identifier: Option<u64>,
    pub(super) comment_wire_path: &'static [u32],
    pub(super) kind: DrawableKind,
    pub(super) direct_users: usize,
}

/// A private, source-bound comment thread.
///
/// `Thread` intentionally carries storage identities and source message
/// positions for the transaction owner alongside an ID-free root snapshot;
/// generated protobuf values never cross this module boundary.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Thread {
    pub(super) root_identifier: u64,
    pub(super) snapshot: ThreadSnapshot,
    pub(super) nodes: Box<[Node]>,
    pub(super) direct_users: usize,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub(super) struct ThreadSnapshot {
    pub(super) comment: Option<Arc<Comment>>,
    pub(super) replies: Box<[Reply]>,
}

/// One validated comment-storage node in a rooted closure.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Node {
    pub(super) identifier: u64,
    pub(super) component_name: Arc<str>,
    pub(super) message_index: usize,
    pub(super) text: Option<Box<str>>,
    pub(super) creation_date_seconds: Option<f64>,
    pub(super) author_identifier: Option<u64>,
    pub(super) author: Option<CommentAuthor>,
    pub(super) storage_uuid: Option<(u64, u64)>,
    pub(super) reply_identifiers: Box<[u64]>,
}

impl Thread {
    pub(super) fn snapshot(self) -> ThreadSnapshot {
        self.snapshot
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Route {
    kind: DrawableKind,
    path: &'static [u32],
    first_envelope_optional: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Projection {
    comment_identifier: Option<u64>,
    comment_wire_path: &'static [u32],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ResolvedDrawable {
    position: Position,
    identifier: u64,
    message_index: usize,
    message_type: u32,
    kind: DrawableKind,
    comment_identifier: Option<u64>,
    comment_wire_path: &'static [u32],
}

/// One bounded physical-header witness for a selected comment root.
///
/// The drawable route is only one of the native ownership witnesses.  An
/// unsupported payload wrapper can still carry an incoming edge in its
/// ArchiveInfo, so ownership admission must inspect the complete core header
/// while leaving that payload opaque.
struct RootReferenceVisitor {
    root_identifier: u64,
    found: bool,
}

impl ArchiveReferenceVisitor for RootReferenceVisitor {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object
            && occurrence.referenced_identifier == self.root_identifier
        {
            self.found = true;
        }
        Ok(())
    }
}

#[derive(Debug)]
struct SlideContext {
    slide_position: Position,
    slide_identifier: u64,
    drawables: Vec<(Position, u64)>,
}

/// Select one drawable by a semantic slide selector and source-order position.
pub(super) fn select_drawable(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    drawable_selector: DrawableSelector,
    budget: &mut Budget,
) -> Result<Selection, Error> {
    let limits = package.semantic_wire_limits().map_err(|_| Error::Read)?;
    let context = resolve_slide_context(package, slide_selector, limits, budget)?;
    let drawable_position = drawable_selector.as_position();
    if context.drawables.get(drawable_position.get()).is_none() {
        return Err(Error::DrawablePositionNotFound {
            position: drawable_position,
        });
    }
    let resolved = resolve_drawables(package, &context, limits, budget)?
        .into_iter()
        .find(|drawable| drawable.position == drawable_position)
        .ok_or(Error::UnsupportedDrawable)?;

    let direct_users = match resolved.comment_identifier {
        Some(identifier) => global_direct_users(package, identifier, limits, budget)?,
        None => 0,
    };
    let (drawable_component_name, _) = package
        .object_with_component(resolved.identifier)
        .ok_or(Error::InvalidSource)?;
    let component_name = copy_arc(drawable_component_name, budget)?;
    Ok(Selection {
        slide_position: context.slide_position,
        drawable_position,
        slide_identifier: context.slide_identifier,
        component_name,
        drawable_identifier: resolved.identifier,
        message_index: resolved.message_index,
        message_type: resolved.message_type,
        comment_identifier: resolved.comment_identifier,
        comment_wire_path: resolved.comment_wire_path,
        kind: resolved.kind,
        direct_users,
    })
}

/// List known drawables in their original slide-owned source order.
///
/// Unknown native drawable kinds are skipped without renumbering the source
/// positions.  A later selector still addresses the original source ordinal,
/// so an unknown entry cannot cause an edit to drift onto a different object.
pub(super) fn inventory_drawables(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    budget: &mut Budget,
) -> Result<Box<[DrawableSummary]>, Error> {
    let limits = package.semantic_wire_limits().map_err(|_| Error::Read)?;
    let context = resolve_slide_context(package, slide_selector, limits, budget)?;
    let resolved = resolve_drawables(package, &context, limits, budget)?;
    let mut output = Vec::new();
    reserve_vec(&mut output, resolved.len(), budget)?;
    for drawable in resolved {
        let reply_count = drawable
            .comment_identifier
            .map(|identifier| read_root_reply_count(package, identifier, limits, budget))
            .transpose()?
            .unwrap_or(0);
        output.push(DrawableSummary::new(
            DrawableSelector::index(drawable.position.get()),
            drawable.kind,
            drawable.comment_identifier.is_some(),
            reply_count,
        ));
    }
    Ok(output.into_boxed_slice())
}

/// Read and validate the complete rooted comment/reply closure for a selection.
pub(super) fn read_comment_graph(
    package: &Package,
    selection: &Selection,
    budget: &mut Budget,
) -> Result<Option<Thread>, Error> {
    let Some(root_identifier) = selection.comment_identifier else {
        return Ok(None);
    };
    if root_identifier == 0 || selection.direct_users == 0 {
        return Err(Error::InvalidSource);
    }
    let limits = package.semantic_wire_limits().map_err(|_| Error::Read)?;
    let component_name = selection.component_name.as_ref();
    let plan = plan_comment_graph_cross_component(
        package,
        component_name,
        root_identifier,
        limits,
        budget,
    )
    .map_err(map_lifecycle_error)?;

    let mut nodes = Vec::new();
    reserve_vec(&mut nodes, plan.storage_ids.len(), budget)?;
    for &identifier in &plan.storage_ids {
        let node = read_storage_node(package, &plan, identifier, limits, budget)?;
        nodes.push(node);
    }
    let lookup_work = binary_search_work(nodes.len())?;
    budget.charge_wire_work(lookup_work)?;
    let root_index = nodes
        .binary_search_by_key(&root_identifier, |node| node.identifier)
        .map_err(|_| Error::InvalidSource)?;
    // The decoded text and author values already own their bounded storage.
    // Move them into the public snapshot instead of cloning them a second
    // time.  The node keeps its identity and graph edges for the mutation
    // engine; only the semantic payload that has crossed the API boundary is
    // consumed here.
    let (root_text, root_timestamp, root_author, reply_count) = {
        let root = &mut nodes[root_index];
        (
            root.text.take().unwrap_or_default(),
            root.creation_date_seconds
                .take()
                .map(comment_timestamp)
                .transpose()?,
            root.author.take(),
            root.reply_identifiers.len(),
        )
    };
    let comment_allocation = size_of::<Comment>()
        .checked_add(
            size_of::<usize>()
                .checked_mul(2)
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(comment_allocation)?;
    let comment = Arc::new(Comment::with_metadata(
        root_text,
        root_timestamp,
        root_author,
    ));
    let mut replies = Vec::new();
    reserve_vec(&mut replies, reply_count, budget)?;
    for reply_ordinal in 0..reply_count {
        let reply_identifier = nodes[root_index].reply_identifiers[reply_ordinal];
        budget.charge_wire_work(lookup_work)?;
        let reply_index = nodes
            .binary_search_by_key(&reply_identifier, |node| node.identifier)
            .map_err(|_| Error::InvalidSource)?;
        let reply = &mut nodes[reply_index];
        let text = reply.text.take().unwrap_or_default();
        let timestamp = reply
            .creation_date_seconds
            .take()
            .map(comment_timestamp)
            .transpose()?;
        let author = reply.author.take();
        replies.push(Reply::with_metadata(text, timestamp, author));
    }
    Ok(Some(Thread {
        root_identifier,
        snapshot: ThreadSnapshot {
            comment: Some(comment),
            replies: replies.into_boxed_slice(),
        },
        nodes: nodes.into_boxed_slice(),
        direct_users: selection.direct_users,
    }))
}

fn resolve_slide_context(
    package: &Package,
    selector: SlideSelector<'_>,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<SlideContext, Error> {
    let slide_position = resolve_slide_position(package, selector, budget)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(|_| Error::Read)?
        .ok_or(Error::InvalidSource)?;
    let (_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(Error::InvalidSource)?;
    validate_object_shape(slide, record.slide_identifier)?;
    let slide_payload = unique_typed_payload(slide, SLIDE_MESSAGE_TYPE)?;
    let drawable_ids =
        references_in_field(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    let mut drawables = Vec::new();
    reserve_vec(&mut drawables, drawable_ids.len(), budget)?;
    let mut seen = HashSet::new();
    let set_bytes = drawable_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(set_bytes)?;
    seen.try_reserve(drawable_ids.len())
        .map_err(|_| Error::Allocation { amount: set_bytes })?;
    for (index, identifier) in drawable_ids.into_iter().enumerate() {
        if !seen.insert(identifier) {
            return Err(Error::InvalidSource);
        }
        let position = Position::new(index);
        drawables.push((position, identifier));
    }
    Ok(SlideContext {
        slide_position,
        slide_identifier: record.slide_identifier,
        drawables,
    })
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut Budget,
) -> Result<Position, Error> {
    budget.charge_wire_work(1)?;
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(|_| Error::Read)?
            .map(|_| position)
            .ok_or(Error::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(Error::EmptySlideName);
            }
            let selected = package
                .show()
                .map_err(|_| Error::Read)?
                .select_slide(SlideSelector::name(name))
                .map_err(|error| match error {
                    crate::SlideSelectorError::EmptySlideName => Error::EmptySlideName,
                    crate::SlideSelectorError::DuplicateSlideName { .. } => {
                        Error::AmbiguousSelector
                    },
                })?;
            selected
                .map(|slide| Position::new(slide.index()))
                .ok_or(Error::SlideNameNotFound)
        },
    }
}

fn resolve_drawables(
    package: &Package,
    context: &SlideContext,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<ResolvedDrawable>, Error> {
    let mut output = Vec::new();
    reserve_vec(&mut output, context.drawables.len(), budget)?;
    for &(position, identifier) in &context.drawables {
        let Some((_component_name, object)) = package.object_with_component(identifier) else {
            return Err(Error::InvalidSource);
        };
        validate_object_shape(object, identifier)?;
        charge_message_scan_work(object.messages.len(), budget)?;
        let mut resolved = None;
        for (message_index, message) in object.messages.iter().enumerate() {
            let Some(route) = route_for(message.type_) else {
                continue;
            };
            let projection = decode_route(message.data.as_slice(), route, limits, budget)?;
            let Some(projection) = projection else {
                continue;
            };
            if resolved.is_some() {
                return Err(Error::InvalidSource);
            }
            validate_message_reference_metadata(
                object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(Error::InvalidSource)?,
                projection.comment_identifier,
            )?;
            if let Some(comment_identifier) = projection.comment_identifier {
                validate_comment_storage_target(package, comment_identifier)?;
            }
            resolved = Some(ResolvedDrawable {
                position,
                identifier,
                message_index,
                message_type: message.type_,
                kind: route.kind,
                comment_identifier: projection.comment_identifier,
                comment_wire_path: projection.comment_wire_path,
            });
        }
        if let Some(resolved) = resolved {
            output.push(resolved);
        }
    }
    Ok(output)
}

fn global_direct_users(
    package: &Package,
    root_identifier: u64,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<usize, Error> {
    validate_comment_storage_target(package, root_identifier)?;
    let archive_limits = package
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut users = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let Some(object_identifier) = object.archive_info.identifier else {
                return Err(Error::InvalidSource);
            };
            budget.charge_entries(1).map_err(map_lifecycle_error)?;
            validate_object_shape(object, object_identifier)?;
            let mut owns_root =
                census_core_header_root_reference(object, root_identifier, archive_limits, budget)?;
            charge_message_scan_work(object.messages.len(), budget)?;
            let mut recognized = false;
            for (message_index, message) in object.messages.iter().enumerate() {
                let Some(route) = route_for(message.type_) else {
                    continue;
                };
                let projection = decode_route(message.data.as_slice(), route, limits, budget)?;
                let Some(projection) = projection else {
                    continue;
                };
                if recognized {
                    return Err(Error::InvalidSource);
                }
                recognized = true;
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(Error::InvalidSource)?;
                validate_message_reference_metadata(info, projection.comment_identifier)?;
                let Some(identifier) = projection.comment_identifier else {
                    continue;
                };
                validate_comment_storage_target(package, identifier)?;
                if identifier == root_identifier {
                    owns_root = true;
                }
            }
            // A root's own header may contain a self-reference in malformed
            // or producer-specific metadata. It is not an incoming owner;
            // cycles are rejected by the rooted graph planner below.
            if owns_root && object_identifier != root_identifier {
                users = users.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
    }
    if users == 0 {
        return Err(Error::InvalidSource);
    }
    Ok(users)
}

/// Inspect one complete ArchiveInfo header under the deletion-grade policy.
///
/// This intentionally does not decode any application payload. Unknown
/// payload wrappers therefore remain preservable, while unknown core metadata
/// is rejected before a direct-owner count can authorize a mutation. The
/// expected reference count and header work are charged before the bounded
/// core inspection starts; the returned occurrence count is cross-checked so
/// the ledger cannot silently under-account a retained source projection.
fn census_core_header_root_reference(
    object: &ArchiveObject,
    root_identifier: u64,
    archive_limits: ArchiveLimits,
    budget: &mut Budget,
) -> Result<bool, Error> {
    let mut fields = 0usize;
    let mut references = 0usize;
    for info in &object.archive_info.message_infos {
        fields = fields
            .checked_add(1)
            .and_then(|count| count.checked_add(info.field_infos.len()))
            .ok_or(Error::InvalidSource)?;
        references = references
            .checked_add(info.object_references.len())
            .and_then(|count| count.checked_add(info.data_references.len()))
            .ok_or(Error::InvalidSource)?;
        for field in &info.field_infos {
            references = references
                .checked_add(field.object_references.len())
                .and_then(|count| count.checked_add(field.data_references.len()))
                .ok_or(Error::InvalidSource)?;
        }
    }

    let header_bytes = usize::try_from(object.header_length)
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    budget
        .charge_wire_fields(fields)
        .map_err(map_lifecycle_error)?;
    budget
        .charge_references(references)
        .map_err(map_lifecycle_error)?;
    budget
        .charge_wire_work(
            header_bytes
                .checked_add(fields)
                .and_then(|work| work.checked_add(references))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_lifecycle_error)?;

    let mut visitor = RootReferenceVisitor {
        root_identifier,
        found: false,
    };
    let observed = object
        .inspect_references_with_policy_and_limits(
            &mut visitor,
            ArchiveReferencePolicy::RejectUnknownMetadata,
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    if observed != references {
        return Err(Error::InvalidSource);
    }
    Ok(visitor.found)
}

fn read_storage_node(
    package: &Package,
    plan: &CommentGraphPlan,
    identifier: u64,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Node, Error> {
    let (component_name, object) = package
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)?;
    let expected_component = plan
        .storage_component(identifier)
        .ok_or(Error::InvalidSource)?;
    if component_name != expected_component || object.archive_info.identifier != Some(identifier) {
        return Err(Error::InvalidSource);
    }
    validate_object_shape(object, identifier)?;
    let mut matching = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE);
    let (message_index, message) = matching.next().ok_or(Error::InvalidSource)?;
    if matching.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info.type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || usize::try_from(info.length).ok() != Some(message.data.len())
    {
        return Err(Error::InvalidSource);
    }
    budget.charge_allocation_plan(message.data.len(), 1)?;
    let mut replies = ReplyCollector::new(budget);
    let decoded = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        message.data.as_slice(),
        comment_decode_options(message.data.as_slice(), limits)?,
        &mut replies,
    )
    .map_err(|_| Error::InvalidSource)?;
    let (reply_identifiers, callback_error) = replies.finish();
    if callback_error.is_some() || decoded.1.replies() != reply_identifiers.len() {
        return Err(Error::InvalidSource);
    }
    budget.charge_wire_fields(decoded.1.fields())?;
    budget.charge_wire_work(decoded.1.work_bytes())?;
    budget.charge_nesting(decoded.1.max_depth() as usize)?;
    budget.charge_references(decoded.1.references())?;
    let snapshot = decoded.0;
    let text = snapshot
        .text()
        .map(|value| copy_boxed(value, budget))
        .transpose()?;
    let creation_date_seconds = snapshot.creation_date().map(|date| date.seconds());
    let author_identifier = snapshot.author().map(|reference| {
        if reference.identifier() == 0
            || reference.deprecated_type().is_some()
            || reference.deprecated_is_external().is_some()
        {
            return Err(Error::InvalidSource);
        }
        Ok(reference.identifier())
    });
    let author_identifier = author_identifier.transpose()?;
    let author = author_identifier
        .map(|identifier| read_author(package, identifier, limits, budget))
        .transpose()?;
    let storage_uuid = snapshot.storage_uuid().map(|uuid| {
        if uuid.lower() == 0 && uuid.upper() == 0 {
            return Err(Error::InvalidSource);
        }
        Ok((uuid.lower(), uuid.upper()))
    });
    let storage_uuid = storage_uuid.transpose()?;
    let reply_ids = replies_to_box(reply_identifiers, budget)?;
    let mut seen_replies = HashSet::new();
    let set_bytes = reply_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(set_bytes)?;
    seen_replies
        .try_reserve(reply_ids.len())
        .map_err(|_| Error::Allocation { amount: set_bytes })?;
    for &reply_identifier in &reply_ids {
        if reply_identifier == 0 || !seen_replies.insert(reply_identifier) {
            return Err(Error::InvalidSource);
        }
    }
    // The planner has already checked the rooted closure and ownership.  The
    // source-order reply list remains intact for ordinal ReplySelector edits.
    let mut component = String::new();
    budget.charge_allocations(component_name.len())?;
    component
        .try_reserve_exact(component_name.len())
        .map_err(|_| Error::Allocation {
            amount: component_name.len(),
        })?;
    component.push_str(component_name);
    let component_name = Arc::<str>::from(component);
    Ok(Node {
        identifier,
        component_name,
        message_index,
        text,
        creation_date_seconds,
        author_identifier,
        author,
        storage_uuid,
        reply_identifiers: reply_ids,
    })
}

fn validate_comment_storage_target(package: &Package, identifier: u64) -> Result<(), Error> {
    let Some((_component, object)) = package.object_with_component(identifier) else {
        return Err(Error::InvalidSource);
    };
    if object.archive_info.identifier != Some(identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || object.archive_info.message_infos[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
        || usize::try_from(object.archive_info.message_infos[0].length).ok()
            != Some(object.messages[0].data.len())
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn read_root_reply_count(
    package: &Package,
    identifier: u64,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<usize, Error> {
    validate_comment_storage_target(package, identifier)?;
    let object = package.object(identifier).ok_or(Error::InvalidSource)?;
    let payload = object
        .messages
        .first()
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    budget.charge_allocation_plan(payload.len(), 1)?;
    let mut visitor = ReplyCollector::new(budget);
    let (_snapshot, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        payload,
        comment_decode_options(payload, limits)?,
        &mut visitor,
    )
    .map_err(|_| Error::InvalidSource)?;
    let (reply_identifiers, callback_error) = visitor.finish();
    if callback_error.is_some() || report.replies() != reply_identifiers.len() {
        return Err(Error::InvalidSource);
    }
    let mut seen = HashSet::new();
    let set_bytes = reply_identifiers
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(set_bytes)?;
    seen.try_reserve(reply_identifiers.len())
        .map_err(|_| Error::Allocation { amount: set_bytes })?;
    for &reply_identifier in &reply_identifiers {
        if !seen.insert(reply_identifier) {
            return Err(Error::InvalidSource);
        }
        validate_comment_storage_target(package, reply_identifier)?;
    }
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references())?;
    Ok(reply_identifiers.len())
}

fn read_author(
    package: &Package,
    identifier: u64,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<CommentAuthor, Error> {
    let Some((_component, object)) = package.object_with_component(identifier) else {
        return Err(Error::InvalidSource);
    };
    if object.archive_info.identifier != Some(identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || object.messages[0].type_ != 212
        || object.archive_info.message_infos[0].type_ != 212
        || usize::try_from(object.archive_info.message_infos[0].length).ok()
            != Some(object.messages[0].data.len())
    {
        return Err(Error::InvalidSource);
    }
    let payload = object.messages[0].data.as_slice();
    budget.charge_allocation_plan(payload.len(), 1)?;
    let (snapshot, report) = annotation_author_codec::decode_annotation_author_with_report(
        payload,
        author_decode_options(payload, limits)?,
    )
    .map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.references())?;
    let display_name = snapshot
        .name()
        .map(|value| copy_boxed(value, budget))
        .transpose()?;
    let public_id = snapshot
        .public_id()
        .map(|value| copy_boxed(value, budget))
        .transpose()?;
    Ok(CommentAuthor::new(display_name, public_id))
}

fn author_decode_options(
    payload: &[u8],
    limits: WireLimits,
) -> Result<annotation_author_codec::DecodeOptions, Error> {
    if payload.len() > limits.max_input_bytes() {
        return Err(Error::InvalidSource);
    }
    let bytes = payload.len().max(1);
    let fields = bytes.min(limits.max_fields()).max(1);
    let work = bytes
        .checked_mul(64)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_rewrite_work())
        .max(1);
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let references = bytes.min(limits.max_fields()).max(1);
    let text = bytes.min(limits.max_input_bytes()).max(1);
    let allocations = bytes.min(limits.max_fields()).max(1);
    Ok(annotation_author_codec::DecodeOptions::new(
        bytes,
        fields,
        work,
        recursion,
        references,
        text,
        allocations,
    ))
}

fn validate_object_shape(object: &ArchiveObject, identifier: u64) -> Result<(), Error> {
    if object.archive_info.identifier != Some(identifier)
        || object.archive_info.message_infos.len() != object.messages.len()
    {
        return Err(Error::InvalidSource);
    }
    for (message, info) in object
        .messages
        .iter()
        .zip(object.archive_info.message_infos.iter())
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn unique_typed_payload(object: &ArchiveObject, message_type: u32) -> Result<&[u8], Error> {
    let mut payload = None;
    for message in object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
    {
        if payload.replace(message.data.as_slice()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    payload.ok_or(Error::InvalidSource)
}

fn validate_message_reference_metadata(
    info: &MessageInfo,
    comment_identifier: Option<u64>,
) -> Result<(), Error> {
    let Some(comment_identifier) = comment_identifier else {
        return Ok(());
    };
    if comment_identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let aggregate = info
        .object_references
        .iter()
        .filter(|reference| **reference == comment_identifier)
        .count();
    let fields = info
        .field_infos
        .iter()
        .flat_map(|field| field.object_references.iter())
        .filter(|reference| **reference == comment_identifier)
        .count();
    if aggregate.checked_add(fields).ok_or(Error::InvalidSource)? != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn references_in_field(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<u64>, Error> {
    let view = parse_view(payload, limits, budget, 1)?;
    let candidate_count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut references = Vec::new();
    reserve_vec(&mut references, candidate_count, budget)?;
    let set_bytes = candidate_count
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    let mut seen = HashSet::new();
    if candidate_count != 0 {
        budget.charge_allocations(set_bytes)?;
        seen.try_reserve(candidate_count)
            .map_err(|_| Error::Allocation { amount: set_bytes })?;
    }
    for field in view.fields().filter(|field| field.number() == field_number) {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let identifier = strict_reference_identifier(field.payload(), limits, budget, 2)?;
        budget.charge_references(1)?;
        if !seen.insert(identifier) {
            return Err(Error::InvalidSource);
        }
        references.push(identifier);
    }
    Ok(references)
}

fn route_for(message_type: u32) -> Option<Route> {
    let (kind, path, first_envelope_optional) = match message_type {
        3_002 => (DrawableKind::Drawable, &[6][..], false),
        3_004 => (DrawableKind::Shape, &[1, 6][..], false),
        3_005 => (DrawableKind::Image, &[1, 6][..], false),
        3_006 => (DrawableKind::Mask, &[1, 6][..], false),
        3_007 => (DrawableKind::Movie, &[1, 6][..], false),
        3_008 => (DrawableKind::Group, &[1, 6][..], false),
        3_009 => (DrawableKind::ConnectionLine, &[1, 1, 6][..], false),
        5_021 => (DrawableKind::Chart, &[1, 6][..], true),
        6_000 => (DrawableKind::Table, &[1, 6][..], false),
        6_007 => (DrawableKind::WordProcessingTable, &[1, 1, 6][..], false),
        2_011 => (DrawableKind::Shape, &[1, 1, 6][..], false),
        2_014 => (DrawableKind::Shape, &[1, 1, 1, 6][..], false),
        7 | 12 => (DrawableKind::Placeholder, &[1, 1, 1, 6][..], false),
        _ => return None,
    };
    Some(Route {
        kind,
        path,
        first_envelope_optional,
    })
}

fn decode_route(
    payload: &[u8],
    route: Route,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Option<Projection>, Error> {
    let mut current = payload;
    for (depth, field_number) in route.path.iter().copied().enumerate() {
        let view = parse_view(current, limits, budget, depth.saturating_add(1))?;
        let mut selected = None;
        for field in view.fields().filter(|field| field.number() == field_number) {
            field
                .validate_canonical_framing()
                .map_err(|_| Error::InvalidSource)?;
            if selected.is_some() || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            selected = Some(field.payload());
        }
        let Some(selected) = selected else {
            if depth == 0 && route.first_envelope_optional {
                return Ok(None);
            }
            // The terminal comment field is optional.  A recognized drawable
            // without a comment remains in the inventory.
            if depth + 1 == route.path.len() {
                return Ok(Some(Projection {
                    comment_identifier: None,
                    comment_wire_path: route.path,
                }));
            }
            return Err(Error::InvalidSource);
        };
        if depth + 1 == route.path.len() {
            let comment_identifier =
                strict_reference_identifier(selected, limits, budget, depth + 2)?;
            return Ok(Some(Projection {
                comment_identifier: Some(comment_identifier),
                comment_wire_path: route.path,
            }));
        }
        current = selected;
    }
    Err(Error::InvalidSource)
}

fn strict_reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut Budget,
    nesting: usize,
) -> Result<u64, Error> {
    let view = parse_view(payload, limits, budget, nesting)?;
    let mut identifier = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        match field.number() {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                let (value, width) =
                    litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                        .map_err(|_| Error::InvalidSource)?;
                if width != field.payload().len()
                    || litchi_iwa_common::varint::encoded_len(value) != width
                    || value == 0
                {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD | REFERENCE_EXTERNAL_FIELD => {
                return Err(Error::InvalidSource);
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    identifier.ok_or(Error::InvalidSource)
}

fn parse_view<'source>(
    payload: &'source [u8],
    limits: WireLimits,
    budget: &mut Budget,
    nesting: usize,
) -> Result<WireView<'source>, Error> {
    if payload.len() > limits.max_input_bytes() {
        return Err(Error::InvalidSource);
    }
    budget.charge_wire_work(payload.len().max(1))?;
    budget.charge_nesting(nesting.max(1))?;
    budget.charge_allocations(payload.len())?;
    let view = WireView::parse_with_limits(payload, limits).map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(view.len())?;
    Ok(view)
}

fn comment_decode_options(
    payload: &[u8],
    limits: WireLimits,
) -> Result<comment_storage_codec::DecodeOptions, Error> {
    if payload.len() > limits.max_input_bytes() {
        return Err(Error::InvalidSource);
    }
    let source = payload.len().max(1);
    let fields = source.min(limits.max_fields()).max(1);
    let work = source
        .checked_mul(32)
        .ok_or(Error::InvalidSource)?
        .min(limits.max_rewrite_work())
        .max(1);
    let nesting = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let references = source.min(limits.max_fields()).max(1);
    let text = source.min(limits.max_input_bytes()).max(1);
    Ok(comment_storage_codec::DecodeOptions::new(
        source, fields, work, nesting, references, text,
    ))
}

struct ReplyCollector<'budget> {
    identifiers: Vec<u64>,
    budget: &'budget mut Budget,
    failure: Option<Error>,
}

impl<'budget> ReplyCollector<'budget> {
    fn new(budget: &'budget mut Budget) -> Self {
        Self {
            identifiers: Vec::new(),
            budget,
            failure: None,
        }
    }

    fn finish(self) -> (Vec<u64>, Option<Error>) {
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
            self.failure = Some(Error::InvalidSource);
            return Ok(());
        }
        if self.identifiers.len() == self.identifiers.capacity()
            && reserve_vec(&mut self.identifiers, 1, self.budget).is_err()
        {
            self.failure = Some(Error::Allocation { amount: 1 });
            return Ok(());
        }
        self.identifiers.push(identifier);
        Ok(())
    }
}

fn replies_to_box(values: Vec<u64>, budget: &mut Budget) -> Result<Box<[u64]>, Error> {
    let mut output = Vec::new();
    reserve_vec(&mut output, values.len(), budget)?;
    output.extend(values);
    Ok(output.into_boxed_slice())
}

fn copy_boxed(value: &str, budget: &mut Budget) -> Result<Box<str>, Error> {
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

fn copy_arc(value: &str, budget: &mut Budget) -> Result<Arc<str>, Error> {
    Ok(Arc::<str>::from(copy_boxed(value, budget)?))
}

fn comment_timestamp(seconds: f64) -> Result<CommentTimestamp, Error> {
    CommentTimestamp::new(seconds).ok_or(Error::InvalidSource)
}

fn reserve_vec<T>(
    output: &mut Vec<T>,
    additional: usize,
    budget: &mut Budget,
) -> Result<(), Error> {
    if additional == 0 {
        return Ok(());
    }
    let bytes = additional
        .checked_mul(size_of::<T>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| Error::Allocation { amount: bytes })
}

/// Charge the type-dispatch work performed before a route can be recognized.
///
/// Unknown and unsupported native messages are intentionally left opaque, but
/// their type tags still have to be inspected. Charge that scan before the
/// loop so a source containing only unsupported messages cannot bypass the
/// operation-wide work ceiling.
fn charge_message_scan_work(message_count: usize, budget: &mut Budget) -> Result<(), Error> {
    let bytes = message_count
        .checked_mul(size_of::<u32>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_wire_work(bytes)?;
    Ok(())
}

/// Return a bounded upper bound for one `binary_search_by_key` pass.
fn binary_search_work(length: usize) -> Result<usize, Error> {
    if length == 0 {
        return Ok(0);
    }
    let search_length = length.checked_add(1).ok_or(Error::InvalidSource)?;
    let leading_zeroes =
        usize::try_from(search_length.leading_zeros()).map_err(|_| Error::InvalidSource)?;
    let comparisons = usize::try_from(usize::BITS)
        .map_err(|_| Error::InvalidSource)?
        .checked_sub(leading_zeroes)
        .ok_or(Error::InvalidSource)?;
    comparisons
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)
}

fn map_lifecycle_error(
    error: crate::package::slide_media_lifecycle::SlideMediaLifecycleError,
) -> Error {
    error.into()
}
