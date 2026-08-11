use std::collections::{HashMap, HashSet};

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::numbers_sheet_order_codec::{
    self, DecodeOptions, DecodeReport, DocumentSheetOrderSnapshot, ReferenceSnapshot,
    TreeNodeSnapshot,
};

use super::super::{
    DOCUMENT_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE, Package, Resolved, SHEET_MESSAGE_TYPE,
};
use super::error::{map_read_error, map_sheet_order_codec_error};
use super::{Error, LimitKind};

const TREE_NODE_MESSAGE_TYPE: u32 = 205;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct MessageTarget {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) identifier: u64,
    pub(super) message_type: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct PreparedTarget {
    pub(super) document: MessageTarget,
    pub(super) sidebar_root: MessageTarget,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct NativeTarget {
    pub(super) document: MessageTarget,
    pub(super) sidebar_root: MessageTarget,
    pub(super) document_snapshot: DocumentSheetOrderSnapshot,
    pub(super) sidebar_snapshot: TreeNodeSnapshot,
    pub(super) document_fields: usize,
    pub(super) sidebar_fields: usize,
}

impl NativeTarget {
    pub(super) const fn prepared(&self) -> PreparedTarget {
        PreparedTarget {
            document: self.document,
            sidebar_root: self.sidebar_root,
        }
    }

    pub(super) fn sheet_identifier(&self, position: usize) -> Option<u64> {
        self.document_snapshot
            .sheet_references()
            .get(position)
            .copied()
            .map(ReferenceSnapshot::identifier)
    }
}

#[derive(Debug)]
pub(super) struct TransactionBudget {
    max_message_bytes: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_transaction_work: usize,
    remaining_fields: usize,
    remaining_work: usize,
    recursion_limit: u32,
    remaining_references: usize,
    remaining_transaction_work: usize,
}

impl TransactionBudget {
    pub(super) fn new(source: &Package) -> Self {
        let archive = source.state.options.archive();
        let wire = archive.max_iwa_stream_bytes().max(1);
        let aggregate_fields = wire.saturating_mul(5).max(1);
        let aggregate_work = wire.saturating_mul(12).max(1);
        let transaction = usize::try_from(archive.max_total_bytes()).unwrap_or(usize::MAX);
        Self {
            max_message_bytes: wire,
            maximum_fields: aggregate_fields,
            maximum_work: aggregate_work,
            maximum_references: source.state.options.semantic().max_references(),
            maximum_transaction_work: transaction,
            remaining_fields: aggregate_fields,
            remaining_work: aggregate_work,
            recursion_limit: 2,
            remaining_references: source.state.options.semantic().max_references(),
            remaining_transaction_work: transaction,
        }
    }

    const fn options(&self) -> DecodeOptions {
        DecodeOptions::new(
            self.max_message_bytes,
            self.remaining_fields,
            self.remaining_work,
            self.recursion_limit,
            self.remaining_references,
        )
    }

    fn consume(&mut self, report: DecodeReport) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_fields,
            self.maximum_fields,
            report.fields(),
            LimitKind::WireFields,
        )?;
        charge_remaining(
            &mut self.remaining_work,
            self.maximum_work,
            report.work_bytes(),
            LimitKind::WireWork,
        )?;
        charge_remaining(
            &mut self.remaining_references,
            self.maximum_references,
            report.references(),
            LimitKind::PayloadReferences,
        )?;
        Ok(())
    }

    pub(super) fn charge_transaction_work(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_transaction_work,
            self.maximum_transaction_work,
            amount,
            LimitKind::TransactionWork,
        )
    }

    pub(super) fn charge_wire(&mut self, fields: usize, work: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_fields,
            self.maximum_fields,
            fields,
            LimitKind::WireFields,
        )?;
        charge_remaining(
            &mut self.remaining_work,
            self.maximum_work,
            work,
            LimitKind::WireWork,
        )?;
        Ok(())
    }

    pub(super) fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_references,
            self.maximum_references,
            amount,
            LimitKind::PayloadReferences,
        )
    }

    fn charge_lookup(&mut self, source: &Package) -> Result<(), Error> {
        let objects = source.state.index.object_count().max(1);
        let comparisons = usize::try_from(usize::BITS - objects.leading_zeros())
            .map_err(|_conversion| Error::InvalidSource)?;
        self.charge_transaction_work(comparisons)
    }

    fn charge_object_metadata(&mut self, object: &ArchiveObject) -> Result<(), Error> {
        let mut work = object
            .messages
            .len()
            .checked_add(object.archive_info.message_infos.len())
            .ok_or(Error::InvalidSource)?;
        let mut references = 0usize;
        for info in &object.archive_info.message_infos {
            references = references
                .checked_add(info.object_references.len())
                .and_then(|value| value.checked_add(info.data_references.len()))
                .ok_or(Error::InvalidSource)?;
            work = work
                .checked_add(info.versions.len())
                .and_then(|value| value.checked_add(info.diff_merge_version.len()))
                .and_then(|value| value.checked_add(info.diff_read_version.len()))
                .and_then(|value| value.checked_add(info.fields_to_remove.len()))
                .and_then(|value| {
                    value.checked_add(
                        info.diff_field_path
                            .as_ref()
                            .map_or(0, |path| path.path.len()),
                    )
                })
                .ok_or(Error::InvalidSource)?;
            for removed in &info.fields_to_remove {
                work = work
                    .checked_add(removed.path.len())
                    .ok_or(Error::InvalidSource)?;
            }
            for field in &info.field_infos {
                references = references
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or(Error::InvalidSource)?;
                work = work
                    .checked_add(1)
                    .and_then(|value| value.checked_add(field.path.path.len()))
                    .and_then(|value| value.checked_add(field.known_field_version.len()))
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .known_field_feature_identifier
                                .as_ref()
                                .map_or(0, String::len),
                        )
                    })
                    .ok_or(Error::InvalidSource)?;
            }
        }
        self.charge_references(references)?;
        self.charge_transaction_work(work.checked_add(references).ok_or(Error::InvalidSource)?)
    }
}

fn charge_remaining(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: LimitKind,
) -> Result<(), Error> {
    let used = maximum
        .checked_sub(*remaining)
        .ok_or(Error::InvalidSource)?;
    let observed = used.checked_add(amount).ok_or(Error::LimitExceeded {
        kind,
        observed: u64::MAX,
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    })?;
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    *remaining = maximum - observed;
    Ok(())
}

pub(super) fn resolve_native_target(
    source: &Package,
    budget: &mut TransactionBudget,
) -> Result<NativeTarget, Error> {
    budget.charge_transaction_work(source.state.components.catalog().len())?;
    let document_component_index = source
        .state
        .components
        .catalog()
        .iter()
        .position(|component| component.name() == "Index/Document.iwa")
        .ok_or(Error::InvalidSource)?;
    let document_component = source
        .state
        .components
        .catalog()
        .get_index(document_component_index)
        .ok_or(Error::InvalidSource)?;
    let (document_object_index, document_object) = document_component
        .archive()
        .objects
        .iter()
        .enumerate()
        .find(|(_index, object)| object.archive_info.identifier == Some(1))
        .ok_or(Error::InvalidSource)?;
    budget.charge_transaction_work(document_component.archive().objects.len())?;
    budget.charge_object_metadata(document_object)?;
    let (document_message_index, document_message) =
        unique_message(&document_object.messages, DOCUMENT_MESSAGE_TYPE)?;
    validate_message_metadata(document_object, document_message_index)?;

    let (document_snapshot, document_report) =
        numbers_sheet_order_codec::decode_document_sheet_order_with_report(
            &document_message.data,
            budget.options(),
        )
        .map_err(map_sheet_order_codec_error)?;
    budget.consume(document_report)?;
    let sheet_references = document_snapshot.sheet_references();
    if sheet_references.len() != source.state.document.sheet_count() {
        return Err(Error::InvalidSource);
    }
    validate_order_metadata(document_object, document_message_index, sheet_references)?;

    let sidebar_reference = document_snapshot.sidebar_order();
    validate_reference(sidebar_reference)?;
    require_declared_reference(
        document_object,
        document_message_index,
        sidebar_reference.identifier(),
        &[5],
    )?;
    let sidebar = resolve_reference(source, sidebar_reference.identifier(), budget)?;
    require_same_component(document_component_index, sidebar.component_index)?;
    let sidebar_object = resolved_object(source, sidebar, sidebar_reference.identifier())?;
    budget.charge_object_metadata(sidebar_object)?;
    let (sidebar_message_index, sidebar_message) =
        unique_message(sidebar.messages, TREE_NODE_MESSAGE_TYPE)?;
    validate_message_metadata(sidebar_object, sidebar_message_index)?;
    let (sidebar_snapshot, sidebar_report) =
        numbers_sheet_order_codec::decode_tree_node_with_report(
            &sidebar_message.data,
            budget.options(),
        )
        .map_err(map_sheet_order_codec_error)?;
    budget.consume(sidebar_report)?;
    if sidebar_snapshot.object_reference().is_some() {
        return Err(Error::UnsupportedSource);
    }
    let child_references = sidebar_snapshot.child_references();
    if child_references.len() != sheet_references.len() {
        return Err(Error::InvalidSource);
    }
    let mut role_identifiers = validate_role_disjointness(
        sidebar_reference.identifier(),
        sheet_references,
        child_references,
    )?;
    validate_order_metadata(sidebar_object, sidebar_message_index, child_references)?;

    for (position, child_reference) in child_references.iter().copied().enumerate() {
        let child = resolve_reference(source, child_reference.identifier(), budget)?;
        require_same_component(document_component_index, child.component_index)?;
        let child_object = resolved_object(source, child, child_reference.identifier())?;
        budget.charge_object_metadata(child_object)?;
        let (message_index, message) = unique_message(child.messages, TREE_NODE_MESSAGE_TYPE)?;
        validate_message_metadata(child_object, message_index)?;
        let (child_snapshot, child_report) =
            numbers_sheet_order_codec::decode_tree_node_with_report(
                &message.data,
                budget.options(),
            )
            .map_err(map_sheet_order_codec_error)?;
        budget.consume(child_report)?;
        let sheet_reference = child_snapshot
            .object_reference()
            .ok_or(Error::UnsupportedSource)?;
        validate_reference(sheet_reference)?;
        if sheet_reference.identifier()
            != sheet_references
                .get(position)
                .copied()
                .ok_or(Error::InvalidSource)?
                .identifier()
        {
            return Err(Error::InvalidSource);
        }
        require_declared_reference(
            child_object,
            message_index,
            sheet_reference.identifier(),
            &[3],
        )?;
        validate_references(child_snapshot.child_references())?;
        require_declared_references(
            child_object,
            message_index,
            child_snapshot.child_references(),
            &[2],
        )?;
        for descendant in child_snapshot.child_references() {
            if !role_identifiers.insert(descendant.identifier()) {
                return Err(Error::InvalidSource);
            }
            let resolved = resolve_reference(source, descendant.identifier(), budget)?;
            require_same_component(document_component_index, resolved.component_index)?;
            let descendant_object = resolved_object(source, resolved, descendant.identifier())?;
            budget.charge_object_metadata(descendant_object)?;
            let (descendant_message_index, _message) =
                unique_message(resolved.messages, TREE_NODE_MESSAGE_TYPE)?;
            validate_message_metadata(descendant_object, descendant_message_index)?;
        }
        validate_plain_sheet(
            source,
            document_component_index,
            sheet_reference.identifier(),
            budget,
        )?;
    }

    Ok(NativeTarget {
        document: MessageTarget {
            component_index: document_component_index,
            object_index: document_object_index,
            message_index: document_message_index,
            identifier: 1,
            message_type: DOCUMENT_MESSAGE_TYPE,
        },
        sidebar_root: MessageTarget {
            component_index: sidebar.component_index,
            object_index: sidebar.object_index,
            message_index: sidebar_message_index,
            identifier: sidebar_reference.identifier(),
            message_type: TREE_NODE_MESSAGE_TYPE,
        },
        document_snapshot,
        sidebar_snapshot,
        document_fields: document_report.fields(),
        sidebar_fields: sidebar_report.fields(),
    })
}

fn resolve_reference<'a>(
    source: &'a Package,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<Resolved<'a>, Error> {
    budget.charge_lookup(source)?;
    source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource)
}

fn resolved_object<'a>(
    source: &'a Package,
    resolved: Resolved<'a>,
    identifier: u64,
) -> Result<&'a ArchiveObject, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .filter(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(Error::InvalidSource)
}

fn unique_message(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<(usize, &RawMessage), Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let first = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(first)
}

fn validate_plain_sheet(
    source: &Package,
    component: usize,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let resolved = resolve_reference(source, identifier, budget)?;
    require_same_component(component, resolved.component_index)?;
    let object = resolved_object(source, resolved, identifier)?;
    budget.charge_object_metadata(object)?;
    let ordinary = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHEET_MESSAGE_TYPE)
        .count();
    let forms = object
        .messages
        .iter()
        .filter(|message| message.type_ == FORM_BASED_SHEET_MESSAGE_TYPE)
        .count();
    match (ordinary, forms) {
        (1, 0) => Ok(()),
        (0, 1) => Err(Error::UnsupportedSource),
        _ => Err(Error::InvalidSource),
    }
}

fn require_same_component(expected: usize, actual: usize) -> Result<(), Error> {
    if expected == actual {
        Ok(())
    } else {
        Err(Error::UnsupportedSource)
    }
}

fn validate_reference(reference: ReferenceSnapshot) -> Result<(), Error> {
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_references(references: &[ReferenceSnapshot]) -> Result<(), Error> {
    let mut positions = HashMap::new();
    positions
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    for (position, reference) in references.iter().copied().enumerate() {
        validate_reference(reference)?;
        if positions.insert(reference.identifier(), position).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_role_disjointness(
    sidebar: u64,
    sheets: &[ReferenceSnapshot],
    children: &[ReferenceSnapshot],
) -> Result<HashSet<u64>, Error> {
    if sidebar == 1 {
        return Err(Error::InvalidSource);
    }
    let capacity = sheets
        .len()
        .checked_add(children.len())
        .and_then(|count| count.checked_add(2))
        .ok_or(Error::InvalidSource)?;
    let mut roles = HashSet::new();
    roles
        .try_reserve(capacity)
        .map_err(|_allocation| Error::Allocation { amount: capacity })?;
    roles.insert(1);
    roles.insert(sidebar);
    for reference in sheets.iter().chain(children) {
        validate_reference(*reference)?;
        if !roles.insert(reference.identifier()) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(roles)
}

fn validate_message_metadata(object: &ArchiveObject, message_index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_order_metadata(
    object: &ArchiveObject,
    message_index: usize,
    references: &[ReferenceSnapshot],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let mut positions = HashMap::new();
    positions
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    for (position, reference) in references.iter().copied().enumerate() {
        if positions.insert(reference.identifier(), position).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let mut next = 0usize;
    for identifier in &info.object_references {
        if let Some(position) = positions.get(identifier) {
            if *position != next {
                return Err(Error::InvalidSource);
            }
            next = next.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if next != references.len() {
        return Err(Error::InvalidSource);
    }
    if info.field_infos.iter().any(|field| {
        field
            .object_references
            .iter()
            .any(|identifier| positions.contains_key(identifier))
    }) {
        return Err(Error::UnsupportedSource);
    }
    Ok(())
}

fn require_declared_reference(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut field_occurrence = false;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count != 0 {
            if count != 1 || field_occurrence || field.path.as_slice() != accepted_path {
                return Err(Error::InvalidSource);
            }
            field_occurrence = true;
        }
    }
    Ok(())
}

fn require_declared_references(
    object: &ArchiveObject,
    message_index: usize,
    references: &[ReferenceSnapshot],
    accepted_path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let mut aggregate = HashMap::new();
    let mut fields = HashSet::new();
    aggregate
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    fields
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
        })?;
    for reference in references {
        if aggregate.insert(reference.identifier(), false).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    for identifier in &info.object_references {
        if let Some(seen) = aggregate.get_mut(identifier) {
            if *seen {
                return Err(Error::InvalidSource);
            }
            *seen = true;
        }
    }
    if aggregate.values().any(|seen| !seen) {
        return Err(Error::InvalidSource);
    }
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if aggregate.contains_key(identifier)
                && (field.path.as_slice() != accepted_path || !fields.insert(*identifier))
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}
