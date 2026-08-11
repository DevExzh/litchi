//! Directional, allocation-free physical locality verification for cell edits.
//!
//! This module deliberately works on the archive-neutral physical catalog and
//! a compact borrowed mutation plan.  A writer can therefore prove that it
//! changed only the IWA messages it planned to change without materialising a
//! second catalogue, a member-name map, or an object-location set.

use litchi_iwa_archive::{ComponentCatalog, SourceCatalog, package::Entry};
use litchi_iwa_core::{Archive, ArchiveObject, MessageInfo, RawMessage};

use crate::table::cells::DirectionalReferenceTransition;

/// A compact position in the deterministic IWA component catalog.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub(in crate::package) struct ComponentLocation {
    /// Position in [`ComponentCatalog::iter`] order.
    pub(in crate::package) component_index: usize,
}

impl ComponentLocation {
    /// Construct an abstract component location.
    #[must_use]
    pub(in crate::package) const fn new(component_index: usize) -> Self {
        Self { component_index }
    }
}

/// A compact position of one object within an IWA component.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub(in crate::package) struct ObjectLocation {
    /// Owning IWA component.
    pub(in crate::package) component: ComponentLocation,
    /// Position in the component archive's object sequence.
    pub(in crate::package) object_index: usize,
}

impl ObjectLocation {
    /// Construct an abstract object location.
    #[must_use]
    pub(in crate::package) const fn new(component: ComponentLocation, object_index: usize) -> Self {
        Self {
            component,
            object_index,
        }
    }
}

/// A compact position of one payload message within an IWA object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub(in crate::package) struct MessageLocation {
    /// Owning object.
    pub(in crate::package) object: ObjectLocation,
    /// Position in the object's message sequence.
    pub(in crate::package) message_index: usize,
}

impl MessageLocation {
    /// Construct an abstract message location.
    #[must_use]
    pub(in crate::package) const fn new(object: ObjectLocation, message_index: usize) -> Self {
        Self {
            object,
            message_index,
        }
    }
}

/// The directional topology operation evidenced for one native message.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) enum DirectionalChange {
    /// One persistent object message changed in place (its object ordinal may
    /// nevertheless shift because another object was removed).
    Replace,
    /// Every message in one object exists only in the directional target.
    Append,
    /// Every message in one object exists only in the directional source.
    Delete,
}

/// One source/target message relation retained by an exact patch.
///
/// Coordinates are intentionally optional rather than represented by a
/// sentinel index.  This makes object append and deletion evidence unambiguous
/// and lets inverse application swap endpoints without a new allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct DirectionalMessage<'a> {
    /// Current directional source location.
    pub(in crate::package) source: Option<MessageLocation>,
    /// Current directional target location.
    pub(in crate::package) target: Option<MessageLocation>,
    /// Stable IWA object identifier shared by both directional endpoints.
    pub(in crate::package) object_identifier: u64,
    /// Stable message type at every present endpoint.
    pub(in crate::package) expected_type: u32,
    /// Authorized topology operation.
    pub(in crate::package) change: DirectionalChange,
    /// Exact target payload, borrowed from the already-reopened target.
    ///
    /// An apply operation derives this from the reopened candidate in its
    /// current direction; it is therefore not retained in the public patch.
    pub(in crate::package) target_payload: Option<&'a [u8]>,
    /// Exact object-reference transition, required for any reference edit.
    ///
    /// The facade owns the flattened reference table and direction-normalizes
    /// this borrowed view before it reaches locality. This avoids a per-proof
    /// conversion buffer while retaining exact aggregate and field metadata.
    pub(in crate::package) references: Option<DirectionalReferenceTransition<'a>>,
}

impl<'a> DirectionalMessage<'a> {
    /// Construct compact directional evidence.
    #[must_use]
    pub(in crate::package) const fn new(
        source: Option<MessageLocation>,
        target: Option<MessageLocation>,
        object_identifier: u64,
        expected_type: u32,
        change: DirectionalChange,
        target_payload: Option<&'a [u8]>,
    ) -> Self {
        Self {
            source,
            target,
            object_identifier,
            expected_type,
            change,
            target_payload,
            references: None,
        }
    }

    /// Authorize one complete, typed object-reference transition.
    #[must_use]
    pub(in crate::package) const fn with_reference_transition(
        mut self,
        references: DirectionalReferenceTransition<'a>,
    ) -> Self {
        self.references = Some(references);
        self
    }
}

/// Exact directional preview membership, expressed as a canonical-name mask.
///
/// A bit is set exactly when the name at that bit position exists in the
/// corresponding artifact.  Unlike a count, this rejects a partial preview
/// substitution with the same cardinality.
#[derive(Debug, Clone, Copy)]
pub(in crate::package) struct PreviewMask<'a> {
    /// Canonical preview names in fixed bit order.
    pub(in crate::package) names: &'a [&'a str],
    /// Exact source membership mask.
    pub(in crate::package) source_mask: u8,
    /// Exact target membership mask.
    pub(in crate::package) target_mask: u8,
}

/// Borrowed evidence for a directional append/delete-capable locality proof.
#[derive(Debug, Clone, Copy)]
pub(in crate::package) struct DirectionalPlan<'a> {
    /// Strictly component/source/target sorted message evidence.
    pub(in crate::package) messages: &'a [DirectionalMessage<'a>],
    /// Exact directional preview membership.
    pub(in crate::package) previews: PreviewMask<'a>,
}

/// Work ceiling for locality verification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct Limits {
    /// Maximum scalar inspections and byte comparisons.
    pub(in crate::package) max_work: u64,
}

impl Limits {
    /// An explicit unlimited profile for callers already bounded by package
    /// ingress limits.
    #[cfg(test)]
    #[must_use]
    pub(in crate::package) const fn unlimited() -> Self {
        Self { max_work: u64::MAX }
    }
}

/// Exact, monotonic accounting for a locality proof.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(in crate::package) struct Report {
    /// Physical ZIP entries inspected.
    pub(in crate::package) entries: u64,
    /// Selected IWA components inspected.
    pub(in crate::package) components: u64,
    /// IWA objects inspected inside selected components.
    pub(in crate::package) objects: u64,
    /// IWA messages inspected inside selected components.
    pub(in crate::package) messages: u64,
    /// Archive/message headers inspected.
    pub(in crate::package) headers: u64,
    /// Raw ZIP bytes compared.
    pub(in crate::package) bytes: u64,
    /// Root preview members observed in the directional source.
    pub(in crate::package) source_previews: u64,
    /// Root preview members observed in the directional candidate.
    pub(in crate::package) candidate_previews: u64,
    /// Aggregate work charged before each comparison or inspection.
    pub(in crate::package) work: u64,
}

impl Report {
    /// Convert this allocation-free proof's exact observations into the
    /// transaction-owned governed counters.
    ///
    /// Locality borrows both catalogs and its compact evidence slice. It owns
    /// no retained buffers, does not allocate scratch, and creates no output
    /// artifact, so every retained/scratch/allocation counter is exactly zero.
    /// Object checks, compared ZIP bytes, and charged scalar work map one-to-
    /// one to the three nonzero budget counters below.
    pub(in crate::package) fn usage(self) -> super::budget::Usage {
        super::budget::Usage {
            objects: self.objects,
            locality_bytes: self.bytes,
            transaction_work: self.work,
            ..super::budget::Usage::default()
        }
    }
}

/// Content-free locality verification failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) enum Error {
    /// The compact mutation plan has duplicate, contradictory, or impossible locations.
    InvalidPlan,
    /// The physical package or its parsed component layout is inconsistent.
    InvalidSource,
    /// A byte, member, object, message, or metadata value changed outside the plan.
    Verification,
    /// The proof exceeded its caller-provided bounded work ceiling.
    LimitExceeded {
        /// Work charged before the rejected operation.
        observed: u64,
        /// Configured maximum work.
        maximum: u64,
    },
}

/// Verify exact locality from directional replace/append/delete evidence.
///
/// The evidence is already normalized by the patch direction: inverse apply
/// swaps endpoints and maps append to delete before it reaches this function.
/// The caller obtains every `target_payload` from the reopened candidate, so
/// no authored payload bytes need be retained in the public patch.
pub(in crate::package) fn verify_directional(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    plan: DirectionalPlan<'_>,
    limits: Limits,
) -> Result<Report, Error> {
    let mut report = Report::default();
    validate_directional_plan(source, candidate, plan, limits, &mut report)?;
    verify_preview_mask(source, candidate, plan.previews, limits, &mut report)?;
    verify_directional_zip_members(source, candidate, plan, limits, &mut report)?;

    let mut start = 0usize;
    while start < plan.messages.len() {
        let component = directional_component(&plan.messages[start]).ok_or(Error::InvalidPlan)?;
        let end = directional_component_end(plan.messages, start, component);
        verify_directional_component(
            source,
            candidate,
            component,
            &plan.messages[start..end],
            limits,
            &mut report,
        )?;
        start = end;
    }
    Ok(report)
}

fn validate_directional_plan(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    plan: DirectionalPlan<'_>,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    validate_preview_mask(plan.previews)?;
    let mut previous_component = None;
    let mut previous_source = None;
    let mut previous_target = None;
    for message in plan.messages {
        charge(report, limits, 1)?;
        let component = directional_component(message).ok_or(Error::InvalidPlan)?;
        if previous_component.is_some_and(|previous| previous > component)
            || message.object_identifier == 0
            || message.expected_type == 0
            || !valid_directional_shape(*message)
        {
            return Err(Error::InvalidPlan);
        }
        if let Some(location) = message.source {
            validate_endpoint(
                source,
                location,
                message.object_identifier,
                message.expected_type,
            )?;
            if previous_source.is_some_and(|previous| previous >= location) {
                return Err(Error::InvalidPlan);
            }
            previous_source = Some(location);
        }
        if let Some(location) = message.target {
            validate_endpoint(
                candidate,
                location,
                message.object_identifier,
                message.expected_type,
            )?;
            if previous_target.is_some_and(|previous| previous >= location) {
                return Err(Error::InvalidPlan);
            }
            previous_target = Some(location);
        }
        previous_component = Some(component);
    }
    Ok(())
}

fn valid_directional_shape(message: DirectionalMessage<'_>) -> bool {
    match (
        message.change,
        message.source,
        message.target,
        message.target_payload,
    ) {
        (DirectionalChange::Replace, Some(source), Some(target), Some(_)) => {
            source.object.component == target.object.component
        },
        (DirectionalChange::Append, None, Some(_), Some(_))
        | (DirectionalChange::Delete, Some(_), None, None) => true,
        _ => false,
    }
}

fn validate_endpoint(
    catalog: &SourceCatalog,
    location: MessageLocation,
    identifier: u64,
    expected_type: u32,
) -> Result<(), Error> {
    let object = catalog
        .components()
        .get_index(location.object.component.component_index)
        .and_then(|component| {
            component
                .archive()
                .objects
                .get(location.object.object_index)
        })
        .filter(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(Error::InvalidPlan)?;
    object
        .messages
        .get(location.message_index)
        .filter(|message| message.type_ == expected_type)
        .map(|_message| ())
        .ok_or(Error::InvalidPlan)
}

fn validate_preview_mask(plan: PreviewMask<'_>) -> Result<(), Error> {
    if plan.names.len() > u8::BITS as usize {
        return Err(Error::InvalidPlan);
    }
    let valid_bits = if plan.names.len() == u8::BITS as usize {
        u8::MAX
    } else {
        (1u8 << plan.names.len()) - 1
    };
    if plan.source_mask & !valid_bits != 0 || plan.target_mask & !valid_bits != 0 {
        return Err(Error::InvalidPlan);
    }
    for (index, name) in plan.names.iter().enumerate() {
        if name.is_empty() || plan.names[..index].contains(name) {
            return Err(Error::InvalidPlan);
        }
    }
    Ok(())
}

fn verify_preview_mask(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    plan: PreviewMask<'_>,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    for (index, name) in plan.names.iter().enumerate() {
        let bit = 1u8 << index;
        let source_entry = unique_entry(source, name, limits, report)?;
        let candidate_entry = unique_entry(candidate, name, limits, report)?;
        if !preview_membership_matches(
            source_entry.is_some(),
            candidate_entry.is_some(),
            plan.source_mask & bit != 0,
            plan.target_mask & bit != 0,
        ) {
            return Err(Error::Verification);
        }
        if let (Some(left), Some(right)) = (source_entry, candidate_entry) {
            if !retained_member_preserved(left, right, limits, report)? {
                return Err(Error::Verification);
            }
        }
    }
    report.source_previews = u64::from(plan.source_mask.count_ones());
    report.candidate_previews = u64::from(plan.target_mask.count_ones());
    Ok(())
}

const fn preview_membership_matches(
    source_present: bool,
    candidate_present: bool,
    expected_source: bool,
    expected_candidate: bool,
) -> bool {
    source_present == expected_source && candidate_present == expected_candidate
}

fn unique_entry<'a>(
    catalog: &'a SourceCatalog,
    name: &str,
    limits: Limits,
    report: &mut Report,
) -> Result<Option<&'a Entry>, Error> {
    let mut result = None;
    for entry in catalog.package().iter() {
        charge(report, limits, 1)?;
        if entry.name() == name {
            if result.replace(entry).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(result)
}

fn verify_directional_zip_members(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    plan: DirectionalPlan<'_>,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    let mut source_entries = source
        .package()
        .iter()
        .filter(|entry| !is_preview(entry.name(), plan.previews.names));
    let mut candidate_entries = candidate
        .package()
        .iter()
        .filter(|entry| !is_preview(entry.name(), plan.previews.names));
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(left), Some(right)) => {
                charge(report, limits, 2)?;
                report.entries = report.entries.checked_add(1).ok_or(Error::InvalidSource)?;
                if left.name() != right.name() {
                    return Err(Error::Verification);
                }
                let selected =
                    directional_selected_component(source.components(), left.name(), plan.messages);
                let preserved = if selected {
                    selected_member_preserved(left, right, limits, report)?
                } else {
                    retained_member_preserved(left, right, limits, report)?
                };
                if !preserved {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    Ok(())
}

fn directional_selected_component(
    components: &ComponentCatalog,
    name: &str,
    messages: &[DirectionalMessage<'_>],
) -> bool {
    messages
        .binary_search_by(|message| {
            components
                .get_index(directional_component(message).unwrap_or(usize::MAX))
                .map_or(std::cmp::Ordering::Greater, |component| {
                    component.name().cmp(name)
                })
        })
        .is_ok()
}

fn directional_component(message: &DirectionalMessage<'_>) -> Option<usize> {
    let source = message
        .source
        .map(|location| location.object.component.component_index);
    let target = message
        .target
        .map(|location| location.object.component.component_index);
    match (source, target) {
        (Some(left), Some(right)) if left == right => Some(left),
        (Some(value), None) | (None, Some(value)) => Some(value),
        _ => None,
    }
}

fn directional_component_end(
    messages: &[DirectionalMessage<'_>],
    start: usize,
    component: usize,
) -> usize {
    let mut end = start;
    while messages
        .get(end)
        .is_some_and(|message| directional_component(message) == Some(component))
    {
        end += 1;
    }
    end
}

fn verify_directional_component(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    component_index: usize,
    evidence: &[DirectionalMessage<'_>],
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    let before = source
        .components()
        .get_index(component_index)
        .ok_or(Error::InvalidPlan)?;
    let after = candidate
        .components()
        .get(before.name())
        .ok_or(Error::Verification)?;
    charge(report, limits, 1)?;
    report.components = report
        .components
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;

    verify_directional_archives(before.archive(), after.archive(), evidence, limits, report)
}

fn verify_directional_archives(
    before: &Archive,
    after: &Archive,
    evidence: &[DirectionalMessage<'_>],
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    let mut source_object = 0usize;
    let mut target_object = 0usize;
    let mut source_cursor = 0usize;
    let mut target_cursor = 0usize;
    while source_object < before.objects.len() || target_object < after.objects.len() {
        let source_event = next_source(evidence, &mut source_cursor);
        let target_event = next_target(evidence, &mut target_cursor);
        let deleting = source_event.is_some_and(|message| {
            message.change == DirectionalChange::Delete
                && message
                    .source
                    .is_some_and(|location| location.object.object_index == source_object)
        });
        let appending = target_event.is_some_and(|message| {
            message.change == DirectionalChange::Append
                && message
                    .target
                    .is_some_and(|location| location.object.object_index == target_object)
        });
        if deleting {
            let object = before
                .objects
                .get(source_object)
                .ok_or(Error::Verification)?;
            verify_deleted_object(
                object,
                source_object,
                evidence,
                &mut source_cursor,
                limits,
                report,
            )?;
            source_object = source_object.checked_add(1).ok_or(Error::InvalidSource)?;
        } else if appending {
            let object = after
                .objects
                .get(target_object)
                .ok_or(Error::Verification)?;
            verify_appended_object(
                object,
                target_object,
                evidence,
                &mut target_cursor,
                limits,
                report,
            )?;
            target_object = target_object.checked_add(1).ok_or(Error::InvalidSource)?;
        } else {
            let left = before
                .objects
                .get(source_object)
                .ok_or(Error::Verification)?;
            let right = after
                .objects
                .get(target_object)
                .ok_or(Error::Verification)?;
            verify_persistent_object(
                left,
                right,
                source_object,
                target_object,
                evidence,
                &mut source_cursor,
                &mut target_cursor,
                limits,
                report,
            )?;
            source_object = source_object.checked_add(1).ok_or(Error::InvalidSource)?;
            target_object = target_object.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if next_source(evidence, &mut source_cursor).is_some()
        || next_target(evidence, &mut target_cursor).is_some()
    {
        return Err(Error::InvalidPlan);
    }
    Ok(())
}

fn verify_deleted_object(
    object: &ArchiveObject,
    object_index: usize,
    evidence: &[DirectionalMessage<'_>],
    cursor: &mut usize,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    charge(report, limits, 1)?;
    report.headers = report.headers.checked_add(1).ok_or(Error::InvalidSource)?;
    report.objects = report.objects.checked_add(1).ok_or(Error::InvalidSource)?;
    let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
    if !canonical_single_message_object(object) {
        return Err(Error::Verification);
    }
    for (message_index, message) in object.messages.iter().enumerate() {
        charge(report, limits, 1)?;
        report.messages = report.messages.checked_add(1).ok_or(Error::InvalidSource)?;
        let evidence_message = next_source(evidence, cursor).ok_or(Error::Verification)?;
        if evidence_message.change != DirectionalChange::Delete
            || evidence_message.object_identifier != identifier
            || evidence_message.expected_type != message.type_
            || evidence_message.target.is_some()
            || !evidence_message.source.is_some_and(|location| {
                location.object.object_index == object_index
                    && location.message_index == message_index
            })
            || object
                .archive_info
                .message_infos
                .get(message_index)
                .is_none_or(|info| info.type_ != message.type_)
        {
            return Err(Error::Verification);
        }
        *cursor = cursor.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    Ok(())
}

fn verify_appended_object(
    object: &ArchiveObject,
    object_index: usize,
    evidence: &[DirectionalMessage<'_>],
    cursor: &mut usize,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    charge(report, limits, 1)?;
    report.headers = report.headers.checked_add(1).ok_or(Error::InvalidSource)?;
    report.objects = report.objects.checked_add(1).ok_or(Error::InvalidSource)?;
    let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
    if !canonical_single_message_object(object) {
        return Err(Error::Verification);
    }
    for (message_index, message) in object.messages.iter().enumerate() {
        charge(report, limits, 1)?;
        report.messages = report.messages.checked_add(1).ok_or(Error::InvalidSource)?;
        let evidence_message = next_target(evidence, cursor).ok_or(Error::Verification)?;
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource)?;
        if evidence_message.change != DirectionalChange::Append
            || evidence_message.object_identifier != identifier
            || evidence_message.expected_type != message.type_
            || evidence_message.source.is_some()
            || !evidence_message.target.is_some_and(|location| {
                location.object.object_index == object_index
                    && location.message_index == message_index
            })
            || info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
            || !equal_bytes(
                message.data.as_slice(),
                evidence_message.target_payload.ok_or(Error::InvalidPlan)?,
                limits,
                report,
            )?
        {
            return Err(Error::Verification);
        }
        *cursor = cursor.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    Ok(())
}

fn canonical_single_message_object(object: &ArchiveObject) -> bool {
    let Some(message) = object.messages.first() else {
        return false;
    };
    let Some(info) = object.archive_info.message_infos.first() else {
        return false;
    };
    object.messages.len() == 1
        && object.archive_info.message_infos.len() == 1
        && object
            .archive_info
            .identifier
            .is_some_and(|identifier| identifier != 0)
        && object.archive_info.should_merge.is_none()
        && info.type_ == message.type_
        && usize::try_from(info.length).ok() == Some(message.data.len())
        && info.versions.as_slice() == [1, 0, 5]
        && info.field_infos.is_empty()
        && info.object_references.is_empty()
        && info.data_references.is_empty()
        && info.base_message_index.is_none()
        && info.diff_merge_version.is_empty()
        && info.diff_field_path.is_none()
        && info.fields_to_remove.is_empty()
        && info.diff_read_version.is_empty()
}

#[allow(
    clippy::too_many_arguments,
    reason = "two independent directional cursors are intentionally explicit"
)]
fn verify_persistent_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    source_object: usize,
    target_object: usize,
    evidence: &[DirectionalMessage<'_>],
    source_cursor: &mut usize,
    target_cursor: &mut usize,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    charge(report, limits, 1)?;
    report.objects = report.objects.checked_add(1).ok_or(Error::InvalidSource)?;
    if source.archive_info.identifier.is_none()
        || source.archive_info.identifier != candidate.archive_info.identifier
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
        || source.archive_info.should_merge != candidate.archive_info.should_merge
    {
        return Err(Error::Verification);
    }
    for (message_index, (left, right)) in
        source.messages.iter().zip(&candidate.messages).enumerate()
    {
        charge(report, limits, 1)?;
        report.messages = report.messages.checked_add(1).ok_or(Error::InvalidSource)?;
        let source_evidence = next_source(evidence, source_cursor).filter(|message| {
            message.source.is_some_and(|location| {
                location.object.object_index == source_object
                    && location.message_index == message_index
            })
        });
        if let Some(message) = source_evidence {
            let target_evidence =
                next_target(evidence, target_cursor).ok_or(Error::Verification)?;
            if message != target_evidence
                || message.change != DirectionalChange::Replace
                || message.object_identifier
                    != source.archive_info.identifier.ok_or(Error::InvalidSource)?
                || message.expected_type != left.type_
                || !message.target.is_some_and(|location| {
                    location.object.object_index == target_object
                        && location.message_index == message_index
                })
            {
                return Err(Error::Verification);
            }
            verify_directional_replacement(
                left,
                right,
                source.archive_info.message_infos.get(message_index),
                candidate.archive_info.message_infos.get(message_index),
                message.target_payload.ok_or(Error::InvalidPlan)?,
                message.references,
                limits,
                report,
            )?;
            *source_cursor = source_cursor.checked_add(1).ok_or(Error::InvalidSource)?;
            *target_cursor = target_cursor.checked_add(1).ok_or(Error::InvalidSource)?;
        } else if !raw_message_preserved(left, right, limits, report)?
            || source.archive_info.message_infos.get(message_index)
                != candidate.archive_info.message_infos.get(message_index)
        {
            return Err(Error::Verification);
        }
    }
    let source_left = next_source(evidence, source_cursor).is_some_and(|message| {
        message
            .source
            .is_some_and(|location| location.object.object_index == source_object)
    });
    let target_left = next_target(evidence, target_cursor).is_some_and(|message| {
        message
            .target
            .is_some_and(|location| location.object.object_index == target_object)
    });
    if source_left || target_left {
        return Err(Error::Verification);
    }
    Ok(())
}

fn next_source<'a>(
    evidence: &'a [DirectionalMessage<'a>],
    cursor: &mut usize,
) -> Option<DirectionalMessage<'a>> {
    while evidence
        .get(*cursor)
        .is_some_and(|message| message.source.is_none())
    {
        *cursor += 1;
    }
    evidence.get(*cursor).copied()
}

fn next_target<'a>(
    evidence: &'a [DirectionalMessage<'a>],
    cursor: &mut usize,
) -> Option<DirectionalMessage<'a>> {
    while evidence
        .get(*cursor)
        .is_some_and(|message| message.target.is_none())
    {
        *cursor += 1;
    }
    evidence.get(*cursor).copied()
}

fn verify_directional_replacement(
    source: &RawMessage,
    candidate: &RawMessage,
    source_info: Option<&MessageInfo>,
    candidate_info: Option<&MessageInfo>,
    expected_target: &[u8],
    references: Option<DirectionalReferenceTransition<'_>>,
    limits: Limits,
    report: &mut Report,
) -> Result<(), Error> {
    charge(report, limits, 1)?;
    report.headers = report.headers.checked_add(2).ok_or(Error::InvalidSource)?;
    let source_info = source_info.ok_or(Error::InvalidSource)?;
    let candidate_info = candidate_info.ok_or(Error::InvalidSource)?;
    if source.type_ != candidate.type_
        || source_info.type_ != source.type_
        || candidate_info.type_ != candidate.type_
        || usize::try_from(source_info.length).ok() != Some(source.data.len())
        || usize::try_from(candidate_info.length).ok() != Some(candidate.data.len())
        || !references.map_or_else(
            || message_info_preserved_except_length(source_info, candidate_info),
            |transition| {
                message_info_matches_reference_transition(source_info, candidate_info, transition)
            },
        )
        || !equal_bytes(candidate.data.as_slice(), expected_target, limits, report)?
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn raw_message_preserved(
    source: &RawMessage,
    candidate: &RawMessage,
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    Ok(source.type_ == candidate.type_
        && equal_bytes(
            source.data.as_slice(),
            candidate.data.as_slice(),
            limits,
            report,
        )?)
}

fn is_preview(name: &str, previews: &[&str]) -> bool {
    previews.contains(&name)
}

fn retained_member_preserved(
    source: &Entry,
    candidate: &Entry,
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    Ok(source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && equal_bytes(
            source.raw_record().local_record(),
            candidate.raw_record().local_record(),
            limits,
            report,
        )?
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
            limits,
            report,
        )?)
}

fn selected_member_preserved(
    source: &Entry,
    candidate: &Entry,
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    Ok(source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate, limits, report)?
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
            limits,
            report,
        )?)
}

fn central_record_preserved(
    source: &[u8],
    candidate: &[u8],
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    const OFFSET: std::ops::Range<usize> = 42..46;
    Ok(source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && equal_bytes(
            &source[..OFFSET.start],
            &candidate[..OFFSET.start],
            limits,
            report,
        )?
        && equal_bytes(
            &source[OFFSET.end..],
            &candidate[OFFSET.end..],
            limits,
            report,
        )?)
}

fn selected_local_record_preserved(
    source: &Entry,
    candidate: &Entry,
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let left = source.raw_record().local_record();
    let right = candidate.raw_record().local_record();
    let (Some(left_header), Some(right_header)) = (
        zip_local_header_length(left),
        zip_local_header_length(right),
    ) else {
        return Ok(false);
    };
    if left_header != right_header
        || !equal_bytes(
            &left[..CRC_AND_SIZES.start],
            &right[..CRC_AND_SIZES.start],
            limits,
            report,
        )?
        || !equal_bytes(
            &left[CRC_AND_SIZES.end..left_header],
            &right[CRC_AND_SIZES.end..right_header],
            limits,
            report,
        )?
    {
        return Ok(false);
    }
    let Some(left_end) = left_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= left.len())
    else {
        return Ok(false);
    };
    let Some(right_end) = right_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= right.len())
    else {
        return Ok(false);
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &left[left_end..],
        &right[right_end..],
        limits,
        report,
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(
    flags: u16,
    source: &[u8],
    candidate: &[u8],
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    if flags & 0x0008 == 0 {
        return equal_bytes(source, candidate, limits, report);
    }
    let left_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let right_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    Ok(left_prefix == right_prefix
        && source.len() == candidate.len()
        && source.len() >= left_prefix + 12
        && equal_bytes(
            &source[..left_prefix],
            &candidate[..right_prefix],
            limits,
            report,
        )?
        && equal_bytes(
            &source[left_prefix + 12..],
            &candidate[right_prefix + 12..],
            limits,
            report,
        )?)
}

fn selected_central_record_preserved(
    source: &[u8],
    candidate: &[u8],
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const OFFSET: std::ops::Range<usize> = 42..46;
    Ok(source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && equal_bytes(
            &source[..CRC_AND_SIZES.start],
            &candidate[..CRC_AND_SIZES.start],
            limits,
            report,
        )?
        && equal_bytes(
            &source[CRC_AND_SIZES.end..OFFSET.start],
            &candidate[CRC_AND_SIZES.end..OFFSET.start],
            limits,
            report,
        )?
        && equal_bytes(
            &source[OFFSET.end..],
            &candidate[OFFSET.end..],
            limits,
            report,
        )?)
}

fn message_info_preserved_except_length(source: &MessageInfo, candidate: &MessageInfo) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn message_info_matches_reference_transition(
    source: &MessageInfo,
    candidate: &MessageInfo,
    transition: DirectionalReferenceTransition<'_>,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
        && source.object_references == transition.source()
        && candidate.object_references == transition.target()
        && source.field_infos.len() == candidate.field_infos.len()
        && transition.fields().len() == source.field_infos.len()
        && source
            .field_infos
            .iter()
            .zip(&candidate.field_infos)
            .zip(transition.fields())
            .enumerate()
            .all(|(index, ((left, right), expected))| {
                expected.field_index() == index
                    && left.object_references == expected.source()
                    && right.object_references == expected.target()
                    && left.data_references == right.data_references
                    && left.path == right.path
                    && left.r#type == right.r#type
                    && left.unknown_field_rule == right.unknown_field_rule
                    && left.known_field_rule == right.known_field_rule
                    && left.known_field_version == right.known_field_version
                    && left.known_field_feature_identifier == right.known_field_feature_identifier
            })
}

fn equal_bytes(
    source: &[u8],
    candidate: &[u8],
    limits: Limits,
    report: &mut Report,
) -> Result<bool, Error> {
    let work = source.len().max(candidate.len());
    charge(report, limits, as_u64(work)?)?;
    report.bytes = report
        .bytes
        .checked_add(as_u64(work)?)
        .ok_or(Error::InvalidSource)?;
    Ok(source == candidate)
}

fn charge(report: &mut Report, limits: Limits, work: u64) -> Result<(), Error> {
    let observed = report.work.checked_add(work).ok_or(Error::LimitExceeded {
        observed: u64::MAX,
        maximum: limits.max_work,
    })?;
    if observed > limits.max_work {
        return Err(Error::LimitExceeded {
            observed,
            maximum: limits.max_work,
        });
    }
    report.work = observed;
    Ok(())
}

fn as_u64(value: usize) -> Result<u64, Error> {
    u64::try_from(value).map_err(|_error| Error::InvalidSource)
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_iwa_core::FieldInfo;

    use crate::table::cells::{
        DirectionalMessage as EvidenceMessage, EvidenceChangeKind, FieldReferenceRoute,
        MessageReferenceRoute, PatchEvidence, PhysicalLocation, ReferenceEvidence, ReferenceSpan,
    };

    use super::*;

    fn object(identifier: u64, type_: u32, payload: &[u8]) -> ArchiveObject {
        ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_,
                data: payload.to_vec(),
            }],
        )
        .expect("small canonical test object")
    }

    fn location(object_index: usize) -> MessageLocation {
        MessageLocation::new(
            ObjectLocation::new(ComponentLocation::new(0), object_index),
            0,
        )
    }

    fn verify_archives(
        source: Archive,
        target: Archive,
        messages: &[DirectionalMessage<'_>],
    ) -> Result<Report, Error> {
        let mut report = Report::default();
        verify_directional_archives(&source, &target, messages, Limits::unlimited(), &mut report)?;
        Ok(report)
    }

    fn reference_evidence(
        references: Vec<u64>,
        source: ReferenceSpan,
        target: ReferenceSpan,
        fields: Vec<FieldReferenceRoute>,
    ) -> PatchEvidence {
        let field_count = fields.len();
        let references = ReferenceEvidence::new(
            Arc::new(vec![MessageReferenceRoute::new(
                source,
                target,
                ReferenceSpan::new(0, field_count),
            )]),
            Arc::new(fields),
            Arc::new(references),
            crate::package::table_cells::Path::Package,
        )
        .expect("valid exact reference evidence");
        PatchEvidence::new(
            Arc::new(vec![
                EvidenceMessage::new(
                    Some(PhysicalLocation {
                        component: 0,
                        object: 0,
                        message: 0,
                    }),
                    Some(PhysicalLocation {
                        component: 0,
                        object: 0,
                        message: 0,
                    }),
                    2,
                    20,
                    EvidenceChangeKind::Replace,
                )
                .with_reference_transition(0),
            ]),
            Some(references),
            0,
            0,
            0,
            crate::package::table_cells::Path::Package,
        )
        .expect("valid transition evidence")
    }

    fn reference_transition(evidence: &PatchEvidence) -> DirectionalReferenceTransition<'_> {
        evidence
            .reference_transition(
                evidence
                    .directional_message(0)
                    .expect("message transition evidence"),
            )
            .expect("reference transition")
    }

    #[test]
    fn directional_delete_replace_append_accepts_shifted_persistent_object() {
        let source = Archive {
            objects: vec![object(1, 10, b"remove"), object(2, 20, b"before")],
        };
        let target = Archive {
            objects: vec![object(2, 20, b"after"), object(3, 30, b"append")],
        };
        let messages = [
            DirectionalMessage::new(
                Some(location(0)),
                None,
                1,
                10,
                DirectionalChange::Delete,
                None,
            ),
            DirectionalMessage::new(
                Some(location(1)),
                Some(location(0)),
                2,
                20,
                DirectionalChange::Replace,
                Some(b"after"),
            ),
            DirectionalMessage::new(
                None,
                Some(location(1)),
                3,
                30,
                DirectionalChange::Append,
                Some(b"append"),
            ),
        ];
        let report = verify_archives(source, target, &messages).expect("declared topology");
        assert_eq!(report.objects, 3);
        assert_eq!(report.messages, 3);
        assert!(report.work >= 5);
    }

    #[test]
    fn directional_inverse_restores_deleted_object_before_shifted_replacement() {
        let source = Archive {
            objects: vec![object(2, 20, b"after"), object(3, 30, b"append")],
        };
        let target = Archive {
            objects: vec![object(1, 10, b"remove"), object(2, 20, b"before")],
        };
        let messages = [
            DirectionalMessage::new(
                None,
                Some(location(0)),
                1,
                10,
                DirectionalChange::Append,
                Some(b"remove"),
            ),
            DirectionalMessage::new(
                Some(location(0)),
                Some(location(1)),
                2,
                20,
                DirectionalChange::Replace,
                Some(b"before"),
            ),
            DirectionalMessage::new(
                Some(location(1)),
                None,
                3,
                30,
                DirectionalChange::Delete,
                None,
            ),
        ];
        verify_archives(source, target, &messages).expect("inverse topology");
    }

    #[test]
    fn directional_shift_without_delete_evidence_is_rejected() {
        let source = Archive {
            objects: vec![object(1, 10, b"remove"), object(2, 20, b"same")],
        };
        let target = Archive {
            objects: vec![object(2, 20, b"same")],
        };
        assert_eq!(
            verify_archives(source, target, &[]),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_append_requires_every_message_to_be_evidenced() {
        let source = Archive::new();
        let target = Archive {
            objects: vec![object(3, 30, b"append")],
        };
        assert_eq!(
            verify_archives(source, target, &[]),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_append_rejects_noncanonical_archive_or_message_metadata() {
        let source = Archive::new();
        let mut appended = object(3, 30, b"append");
        appended.archive_info.should_merge = Some(false);
        let target = Archive {
            objects: vec![appended],
        };
        let messages = [DirectionalMessage::new(
            None,
            Some(location(0)),
            3,
            30,
            DirectionalChange::Append,
            Some(b"append"),
        )];
        assert_eq!(
            verify_archives(source, target, &messages),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_delete_rejects_noncanonical_inverse_provenance() {
        let mut deleted = object(3, 30, b"append");
        deleted.archive_info.message_infos[0]
            .object_references
            .push(77);
        let source = Archive {
            objects: vec![deleted],
        };
        let target = Archive::new();
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            None,
            3,
            30,
            DirectionalChange::Delete,
            None,
        )];
        assert_eq!(
            verify_archives(source, target, &messages),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_replace_rejects_type_or_payload_substitution() {
        let source = Archive {
            objects: vec![object(2, 20, b"before")],
        };
        let target = Archive {
            objects: vec![object(2, 21, b"after")],
        };
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            Some(location(0)),
            2,
            20,
            DirectionalChange::Replace,
            Some(b"after"),
        )];
        assert_eq!(
            verify_archives(source, target, &messages),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_replace_rejects_untyped_object_reference_drift() {
        let source = Archive {
            objects: vec![object(2, 20, b"before")],
        };
        let mut target = Archive {
            objects: vec![object(2, 20, b"after")],
        };
        target.objects[0].archive_info.message_infos[0]
            .object_references
            .push(77);
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            Some(location(0)),
            2,
            20,
            DirectionalChange::Replace,
            Some(b"after"),
        )];
        assert_eq!(
            verify_archives(source, target, &messages),
            Err(Error::Verification)
        );
    }

    #[test]
    fn directional_replace_accepts_exact_typed_object_reference_transition() {
        let source = Archive {
            objects: vec![object(2, 20, b"before")],
        };
        let mut target = Archive {
            objects: vec![object(2, 20, b"after")],
        };
        target.objects[0].archive_info.message_infos[0]
            .object_references
            .push(77);
        let evidence = reference_evidence(
            vec![77],
            ReferenceSpan::new(0, 0),
            ReferenceSpan::new(0, 1),
            vec![],
        );
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            Some(location(0)),
            2,
            20,
            DirectionalChange::Replace,
            Some(b"after"),
        )
        .with_reference_transition(reference_transition(&evidence))];
        verify_archives(source, target, &messages).expect("exact reference transition");
    }

    #[test]
    fn directional_replace_requires_complete_exact_field_reference_transition() {
        let mut source_object = object(2, 20, b"before");
        let mut field = FieldInfo::new(vec![4, 1]);
        field.object_references.push(10);
        source_object.archive_info.message_infos[0]
            .field_infos
            .push(field);
        let source = Archive {
            objects: vec![source_object],
        };
        let mut target = source.clone();
        target.objects[0].messages[0].data = b"after".to_vec();
        target.objects[0].archive_info.message_infos[0].length = 5;
        target.objects[0].archive_info.message_infos[0].field_infos[0].object_references = vec![20];
        let evidence = reference_evidence(
            vec![10, 20],
            ReferenceSpan::new(0, 0),
            ReferenceSpan::new(0, 0),
            vec![FieldReferenceRoute::new(
                0,
                ReferenceSpan::new(0, 1),
                ReferenceSpan::new(1, 1),
            )],
        );
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            Some(location(0)),
            2,
            20,
            DirectionalChange::Replace,
            Some(b"after"),
        )
        .with_reference_transition(reference_transition(&evidence))];
        verify_archives(source, target, &messages).expect("complete field transition");
    }

    #[test]
    fn directional_replace_rejects_data_reference_drift_even_with_object_transition() {
        let source = Archive {
            objects: vec![object(2, 20, b"before")],
        };
        let mut target = Archive {
            objects: vec![object(2, 20, b"after")],
        };
        let info = &mut target.objects[0].archive_info.message_infos[0];
        info.object_references.push(77);
        info.data_references.push(88);
        let evidence = reference_evidence(
            vec![77],
            ReferenceSpan::new(0, 0),
            ReferenceSpan::new(0, 1),
            vec![],
        );
        let messages = [DirectionalMessage::new(
            Some(location(0)),
            Some(location(0)),
            2,
            20,
            DirectionalChange::Replace,
            Some(b"after"),
        )
        .with_reference_transition(reference_transition(&evidence))];
        assert_eq!(
            verify_archives(source, target, &messages),
            Err(Error::Verification)
        );
    }

    #[test]
    fn preview_mask_rejects_partial_or_out_of_range_membership() {
        let names = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
        assert_eq!(
            validate_preview_mask(PreviewMask {
                names: &names,
                source_mask: 0b1000,
                target_mask: 0,
            }),
            Err(Error::InvalidPlan)
        );
        assert!(
            validate_preview_mask(PreviewMask {
                names: &names,
                source_mask: 0b101,
                target_mask: 0b010,
            })
            .is_ok()
        );
        assert!(!preview_membership_matches(true, false, true, true));
    }

    #[test]
    fn governed_usage_is_exactly_allocation_free() {
        let report = Report {
            objects: 7,
            bytes: 11,
            work: 13,
            ..Report::default()
        };
        let usage = report.usage();
        assert_eq!(usage.objects, 7);
        assert_eq!(usage.locality_bytes, 11);
        assert_eq!(usage.transaction_work, 13);
        assert_eq!(usage.retained_elements, 0);
        assert_eq!(usage.retained_bytes, 0);
        assert_eq!(usage.scratch_bytes, 0);
        assert_eq!(usage.peak_scratch_bytes, 0);
        assert_eq!(usage.allocation_events, 0);
    }
}
