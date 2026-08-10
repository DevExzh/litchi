//! Bounded rooted Show-to-Slide resolution for placeholder visibility.

use std::collections::HashSet;

use litchi_core::Position;
use litchi_iwa_common::{WireLimits, wire::WireView};

use super::{
    Error, LimitKind, Package, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE, SlideSelector,
    TransactionBudget, map_read_error, map_wire_error, nested_unique_field, selected_message,
    strict_reference, validate_reference_metadata, validate_reference_metadata_set,
};

pub(super) struct FocusedSlide {
    pub(super) position: Position,
    pub(super) show_identifier: u64,
    pub(super) node_identifier: u64,
    pub(super) slide_identifier: u64,
    pub(super) rooted_node_identifiers: Vec<u64>,
    pub(super) rooted_slide_identifiers: Vec<u64>,
}

#[derive(Clone, Copy)]
struct RootedSlide {
    position: Position,
    node_identifier: u64,
    slide_identifier: u64,
}

pub(super) fn focused_slide(
    package: &Package,
    selector: SlideSelector<'_>,
    mutation: bool,
    budget: &mut TransactionBudget,
) -> Result<FocusedSlide, Error> {
    const DOCUMENT_MESSAGE_TYPE: u32 = 1;
    const SHOW_MESSAGE_TYPE: u32 = 2;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    budget.charge_reference()?;
    let show_identifier = package.root_show_identifier().map_err(map_read_error)?;
    if show_identifier == 1 {
        return Err(Error::InvalidSource);
    }
    let (_show_component, show) = package
        .object_with_component(show_identifier)
        .ok_or(Error::InvalidSource)?;
    let (show_index, show_payload) = selected_message(show, SHOW_MESSAGE_TYPE)?;
    budget.charge_work(show_payload.len())?;
    let tree = nested_unique_field(show_payload, &[3], limits)?.ok_or(Error::InvalidSource)?;
    let tree_view = WireView::parse_with_limits(tree.payload(), limits).map_err(map_wire_error)?;
    let count = tree_view
        .fields()
        .filter(|field| field.number() == 2)
        .count();
    let tree_work = count
        .checked_mul(2)
        .and_then(|field_work| tree.payload().len().checked_add(field_work))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(tree_work)?;
    if count > package.semantic_limits().max_slides() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::Slides,
            observed: count as u64,
            maximum: package.semantic_limits().max_slides() as u64,
        });
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(count)
        .map_err(|_allocation| Error::Allocation { amount: count })?;
    let mut rooted_nodes = HashSet::new();
    rooted_nodes
        .try_reserve(count)
        .map_err(|_allocation| Error::Allocation { amount: count })?;
    let mut rooted_slides = HashSet::new();
    rooted_slides
        .try_reserve(count)
        .map_err(|_allocation| Error::Allocation { amount: count })?;
    for field in tree_view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() != 2 {
            continue;
        }
        budget.charge_reference()?;
        let node_identifier = strict_reference(field, limits)?;
        if matches!(node_identifier, 1)
            || node_identifier == show_identifier
            || !rooted_nodes.insert(node_identifier)
        {
            return Err(Error::InvalidSource);
        }
        let (_node_component, node) = package
            .object_with_component(node_identifier)
            .ok_or(Error::InvalidSource)?;
        let (_node_index, node_payload) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
        budget.charge_work(node_payload.len())?;
        let slide = nested_unique_field(node_payload, &[2], limits)?.ok_or(Error::InvalidSource)?;
        budget.charge_reference()?;
        let slide_identifier = strict_reference(slide, limits)?;
        if matches!(slide_identifier, 1)
            || slide_identifier == show_identifier
            || !rooted_slides.insert(slide_identifier)
        {
            return Err(Error::InvalidSource);
        }
        records.push(RootedSlide {
            position: Position::new(records.len()),
            node_identifier,
            slide_identifier,
        });
    }
    if !rooted_nodes.is_disjoint(&rooted_slides) {
        return Err(Error::InvalidSource);
    }
    let selected = match selector {
        SlideSelector::Position(position) => records
            .get(position.get())
            .copied()
            .ok_or(Error::SlidePositionNotFound { position })?,
        SlideSelector::Name(name) => {
            let mut selected = None;
            for record in &records {
                let (_component, slide) = package
                    .object_with_component(record.slide_identifier)
                    .ok_or(Error::InvalidSource)?;
                let (_index, payload) = selected_message(slide, SLIDE_MESSAGE_TYPE)?;
                budget.charge_work(payload.len())?;
                if slide_name(payload, limits)? == Some(name) && selected.replace(*record).is_some()
                {
                    return Err(Error::AmbiguousSelector);
                }
            }
            selected.ok_or(Error::SlideNameNotFound)?
        },
    };
    if mutation {
        let (_root_component, root) = package
            .object_with_component(1)
            .ok_or(Error::InvalidSource)?;
        let (root_index, _root_payload) = selected_message(root, DOCUMENT_MESSAGE_TYPE)?;
        validate_reference_metadata(root, root_index, show_identifier, &[2])?;
        validate_reference_metadata_set(show, show_index, &rooted_nodes, &[3, 2])?;
        for record in &records {
            let (_component, node) = package
                .object_with_component(record.node_identifier)
                .ok_or(Error::InvalidSource)?;
            let (node_index, _payload) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
            validate_reference_metadata(node, node_index, record.slide_identifier, &[2])?;
        }
    }
    let mut rooted_slide_identifiers = Vec::new();
    rooted_slide_identifiers
        .try_reserve_exact(records.len())
        .map_err(|_allocation| Error::Allocation {
            amount: records.len(),
        })?;
    rooted_slide_identifiers.extend(records.iter().map(|record| record.slide_identifier));
    let mut rooted_node_identifiers = Vec::new();
    rooted_node_identifiers
        .try_reserve_exact(records.len())
        .map_err(|_allocation| Error::Allocation {
            amount: records.len(),
        })?;
    rooted_node_identifiers.extend(records.iter().map(|record| record.node_identifier));
    Ok(FocusedSlide {
        position: selected.position,
        show_identifier,
        node_identifier: selected.node_identifier,
        slide_identifier: selected.slide_identifier,
        rooted_node_identifiers,
        rooted_slide_identifiers,
    })
}

fn slide_name(source: &[u8], limits: WireLimits) -> Result<Option<&str>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut name = None;
    for field in view.fields() {
        if field.number() == 10 {
            field.validate_canonical_framing().map_err(map_wire_error)?;
            if field.wire_type() != 2 || name.replace(field.payload()).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    name.map(std::str::from_utf8)
        .transpose()
        .map_err(|_error| Error::InvalidSource)
}
