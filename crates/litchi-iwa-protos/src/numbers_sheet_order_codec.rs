//! Strict generated-free routing for the ordered Numbers sheet/sidebar seam.
//!
//! A handwritten two-pass router owns all untrusted traversal, canonical wire
//! validation, allocation, and aggregate accounting. Buffa is used only for a
//! forced borrowed cross-check of each selected `TSP.Reference`. The caller's
//! source bytes remain authoritative for preservation and raw-record splicing.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict routing helpers stay beside the public generated-free snapshots they build."
)]

use core::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_sheet_order_generated::LitchiIwaProjection as projection;

const DOCUMENT_SHEETS_FIELD: u32 = 1;
const DOCUMENT_SIDEBAR_ORDER_FIELD: u32 = 5;
const TREE_NODE_CHILDREN_FIELD: u32 = 2;
const TREE_NODE_OBJECT_FIELD: u32 = 3;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

/// Finite policy for one strict owner-payload decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
}

impl DecodeOptions {
    /// Construct an explicit finite decode profile.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Successful exact consumption for transaction-level budget merging.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
}

impl DecodeReport {
    /// Logical outer and nested reference fields validated.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Bytes inspected across both owner passes and strict/Buffa reference passes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest owner/reference depth reached.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Routed reference occurrences.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Aggregate nested reference payload bytes.
    #[must_use]
    pub const fn reference_bytes(self) -> usize {
        self.reference_bytes
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// One owner payload or configured Buffa message ceiling is too large.
    Bytes {
        /// Rejected source or configured value.
        observed: usize,
        /// Applied maximum.
        maximum: usize,
    },
    /// Aggregate selected-reference occurrence ceiling.
    References {
        /// First rejected count.
        observed: usize,
        /// Applied maximum.
        maximum: usize,
    },
    /// Aggregate logical field ceiling.
    Fields {
        /// First rejected count.
        observed: usize,
        /// Applied maximum.
        maximum: usize,
    },
    /// Aggregate traversal/cross-check work ceiling.
    Work {
        /// First rejected byte count.
        observed: usize,
        /// Applied maximum.
        maximum: usize,
    },
    /// Owner/reference nesting ceiling.
    Nesting {
        /// Rejected configured or traversed depth.
        observed: u32,
        /// Applied maximum.
        maximum: u32,
    },
}

/// Strict Numbers sheet-order decode failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Invalid,
    Limit(DecodeLimit),
    Allocation { amount: usize },
}

impl DecodeError {
    /// Return exact resource observations when a finite limit failed.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            DecodeErrorKind::Invalid | DecodeErrorKind::Allocation { .. } => None,
        }
    }

    /// Return the failed exact reservation size, when applicable.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            DecodeErrorKind::Invalid | DecodeErrorKind::Limit(_) => None,
        }
    }

    const fn invalid() -> Self {
        Self {
            kind: DecodeErrorKind::Invalid,
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid Numbers sheet-order payload")
    }
}

impl std::error::Error for DecodeError {}

/// Generated-free exact scalar projection of one `TSP.Reference`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    /// Required native identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    /// Optional legacy type with encoded presence preserved.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Optional legacy external marker with encoded presence preserved.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Ordered `TN.DocumentArchive.sheets` plus required sidebar-root reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentSheetOrderSnapshot {
    sheet_references: Box<[ReferenceSnapshot]>,
    sidebar_order: ReferenceSnapshot,
}

impl DocumentSheetOrderSnapshot {
    /// Borrow rooted sheet references in source order.
    #[must_use]
    pub fn sheet_references(&self) -> &[ReferenceSnapshot] {
        &self.sheet_references
    }

    /// Required sidebar-tree root reference.
    #[must_use]
    pub const fn sidebar_order(&self) -> ReferenceSnapshot {
        self.sidebar_order
    }
}

/// Ordered `TSK.TreeNode.children` and optional owned-object reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TreeNodeSnapshot {
    child_references: Box<[ReferenceSnapshot]>,
    object_reference: Option<ReferenceSnapshot>,
}

impl TreeNodeSnapshot {
    /// Borrow child-node references in source order.
    #[must_use]
    pub fn child_references(&self) -> &[ReferenceSnapshot] {
        &self.child_references
    }

    /// Optional object reference with encoded presence preserved.
    #[must_use]
    pub const fn object_reference(&self) -> Option<ReferenceSnapshot> {
        self.object_reference
    }
}

/// Decode ordered sheet and sidebar-root references from one Numbers document.
pub fn decode_document_sheet_order(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DocumentSheetOrderSnapshot, DecodeError> {
    Ok(decode_document_sheet_order_with_report(source, options)?.0)
}

/// Decode one Numbers document and return exact aggregate resource consumption.
pub fn decode_document_sheet_order_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DocumentSheetOrderSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let sheet_count = preflight_document(source, &mut budget)?;
    let mut sheets = allocate_exact(sheet_count)?;
    let mut sidebar = None;

    budget.charge_work(source.len())?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, &mut budget, 1, false)? {
        match field.number {
            DOCUMENT_SHEETS_FIELD => {
                sheets.push(project_reference(field.length_delimited()?, &mut budget)?);
            },
            DOCUMENT_SIDEBAR_ORDER_FIELD => {
                if sidebar.is_some() {
                    return Err(DecodeError::invalid());
                }
                sidebar = Some(project_reference(field.length_delimited()?, &mut budget)?);
            },
            _ => {},
        }
    }
    if sheets.len() != sheet_count {
        return Err(DecodeError::invalid());
    }
    Ok((
        DocumentSheetOrderSnapshot {
            sheet_references: sheets.into_boxed_slice(),
            sidebar_order: sidebar.ok_or_else(DecodeError::invalid)?,
        },
        budget.report(),
    ))
}

/// Decode one `TSK.TreeNode` reference envelope.
pub fn decode_tree_node(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TreeNodeSnapshot, DecodeError> {
    Ok(decode_tree_node_with_report(source, options)?.0)
}

/// Decode one `TSK.TreeNode` and return exact aggregate resource consumption.
pub fn decode_tree_node_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TreeNodeSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let child_count = preflight_tree_node(source, &mut budget)?;
    let mut children = allocate_exact(child_count)?;
    let mut object = None;

    budget.charge_work(source.len())?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, &mut budget, 1, false)? {
        match field.number {
            TREE_NODE_CHILDREN_FIELD => {
                children.push(project_reference(field.length_delimited()?, &mut budget)?);
            },
            TREE_NODE_OBJECT_FIELD => {
                if object.is_some() {
                    return Err(DecodeError::invalid());
                }
                object = Some(project_reference(field.length_delimited()?, &mut budget)?);
            },
            _ => {},
        }
    }
    if children.len() != child_count {
        return Err(DecodeError::invalid());
    }
    Ok((
        TreeNodeSnapshot {
            child_references: children.into_boxed_slice(),
            object_reference: object,
        },
        budget.report(),
    ))
}

fn preflight_document(source: &[u8], budget: &mut Budget) -> Result<usize, DecodeError> {
    budget.charge_work(source.len())?;
    let mut sheet_count = 0usize;
    let mut sidebar_seen = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1, true)? {
        match field.number {
            DOCUMENT_SHEETS_FIELD => {
                let payload = field.length_delimited()?;
                budget.charge_reference(payload.len())?;
                sheet_count = sheet_count
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
            },
            DOCUMENT_SIDEBAR_ORDER_FIELD => {
                if sidebar_seen {
                    return Err(DecodeError::invalid());
                }
                sidebar_seen = true;
                budget.charge_reference(field.length_delimited()?.len())?;
            },
            _ => {},
        }
    }
    if !sidebar_seen {
        return Err(DecodeError::invalid());
    }
    Ok(sheet_count)
}

fn preflight_tree_node(source: &[u8], budget: &mut Budget) -> Result<usize, DecodeError> {
    budget.charge_work(source.len())?;
    let mut child_count = 0usize;
    let mut object_seen = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1, true)? {
        match field.number {
            TREE_NODE_CHILDREN_FIELD => {
                let payload = field.length_delimited()?;
                budget.charge_reference(payload.len())?;
                child_count = child_count
                    .checked_add(1)
                    .ok_or_else(DecodeError::invalid)?;
            },
            TREE_NODE_OBJECT_FIELD => {
                if object_seen {
                    return Err(DecodeError::invalid());
                }
                object_seen = true;
                budget.charge_reference(field.length_delimited()?.len())?;
            },
            _ => {},
        }
    }
    Ok(child_count)
}

fn project_reference(source: &[u8], budget: &mut Budget) -> Result<ReferenceSnapshot, DecodeError> {
    budget.observe_depth(2)?;
    budget.charge_work(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 2, true)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid());
                }
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                deprecated_type = Some(canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::invalid());
                }
                deprecated_is_external = Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    let strict = ReferenceSnapshot {
        identifier: identifier.ok_or_else(DecodeError::invalid)?,
        deprecated_type,
        deprecated_is_external,
    };

    budget.charge_work(source.len())?;
    let view: projection::NumbersSheetReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    if !view.has_identifier()
        || view.identifier != strict.identifier
        || view.deprecated_type != strict.deprecated_type
        || view.deprecated_is_external != strict.deprecated_is_external
    {
        return Err(DecodeError::invalid());
    }
    Ok(strict)
}

fn allocate_exact<T>(capacity: usize) -> Result<Vec<T>, DecodeError> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|_allocation| DecodeError::allocation(capacity))?;
    Ok(values)
}

#[derive(Clone, Copy)]
struct StrictField<'source> {
    number: u32,
    wire_type: u8,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            StrictValue::LengthDelimited(value) if self.wire_type == 2 => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Fixed32 => Err(DecodeError::invalid()),
        }
    }

    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            StrictValue::Varint(value) if self.wire_type == 0 => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Fixed32 => Err(DecodeError::invalid()),
        }
    }
}

#[derive(Clone, Copy)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Fixed32,
}

fn next_field<'source>(
    remaining: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
    charge_field: bool,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    if remaining.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    if charge_field {
        budget.charge_field()?;
    }
    let tag = take_varint(remaining)?;
    let number = u32::try_from(tag >> 3).map_err(|_conversion| DecodeError::invalid())?;
    let wire_type = u8::try_from(tag & 7).map_err(|_conversion| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let value = match wire_type {
        0 => StrictValue::Varint(take_varint(remaining)?),
        1 => {
            take(remaining, 8)?;
            StrictValue::Fixed64
        },
        2 => {
            let length = usize::try_from(take_varint(remaining)?)
                .map_err(|_conversion| DecodeError::invalid())?;
            StrictValue::LengthDelimited(take(remaining, length)?)
        },
        5 => {
            take(remaining, 4)?;
            StrictValue::Fixed32
        },
        _ => return Err(DecodeError::invalid()),
    };
    Ok(Some(StrictField {
        number,
        wire_type,
        value,
    }))
}

fn take<'source>(
    remaining: &mut &'source [u8],
    amount: usize,
) -> Result<&'source [u8], DecodeError> {
    if remaining.len() < amount {
        return Err(DecodeError::invalid());
    }
    let (selected, rest) = remaining.split_at(amount);
    *remaining = rest;
    Ok(selected)
}

fn take_varint(remaining: &mut &[u8]) -> Result<u64, DecodeError> {
    let source = *remaining;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source.get(index).ok_or_else(DecodeError::invalid)?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let width = encoded_varint_len(value);
            if width != index + 1 {
                return Err(DecodeError::invalid());
            }
            *remaining = &source[index + 1..];
            return Ok(value);
        }
    }
    Err(DecodeError::invalid())
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if let Ok(positive) = i32::try_from(value) {
        return Ok(positive);
    }
    if value < MIN_SIGN_EXTENDED_I32 {
        return Err(DecodeError::invalid());
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes()))
        .map_err(|_conversion| DecodeError::invalid())
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

struct Budget {
    options: DecodeOptions,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    reference_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
            .map_err(|_conversion| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::limit(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limit(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        let mut budget = Self {
            options,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            reference_bytes: 0,
        };
        budget.observe_depth(1)?;
        Ok(budget)
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_reference(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.references.saturating_add(1);
        if observed > self.options.max_references {
            return Err(DecodeError::limit(DecodeLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references = observed;
        self.reference_bytes = self
            .reference_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes);
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            reference_bytes: self.reference_bytes,
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused unit-test fixtures require successful construction and exact errors"
)]
mod tests {
    use super::*;

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).unwrap();
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        push_varint(output, u64::from(number) << 3);
        push_varint(output, value);
    }

    fn push_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        push_varint(output, (u64::from(number) << 3) | 2);
        push_varint(output, u64::try_from(value.len()).unwrap());
        output.extend_from_slice(value);
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint_field(&mut output, REFERENCE_IDENTIFIER_FIELD, identifier);
        output
    }

    fn rich_reference(identifier: u64) -> Vec<u8> {
        let mut output = reference(identifier);
        push_varint_field(&mut output, REFERENCE_DEPRECATED_TYPE_FIELD, u64::MAX);
        push_varint_field(&mut output, REFERENCE_DEPRECATED_EXTERNAL_FIELD, 0);
        push_varint_field(&mut output, 99, 77);
        output
    }

    fn document(sheets: &[Vec<u8>], sidebar: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint_field(&mut output, 99, 1);
        for sheet in sheets {
            push_bytes_field(&mut output, DOCUMENT_SHEETS_FIELD, sheet);
        }
        push_bytes_field(&mut output, DOCUMENT_SIDEBAR_ORDER_FIELD, sidebar);
        output
    }

    fn wide_document(sheet_count: usize) -> Vec<u8> {
        let sheet = reference(7);
        let sidebar = reference(9);
        let mut output = Vec::with_capacity(
            sheet_count
                .saturating_mul(sheet.len().saturating_add(3))
                .saturating_add(sidebar.len().saturating_add(3)),
        );
        for _ in 0..sheet_count {
            push_bytes_field(&mut output, DOCUMENT_SHEETS_FIELD, &sheet);
        }
        push_bytes_field(&mut output, DOCUMENT_SIDEBAR_ORDER_FIELD, &sidebar);
        output
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), 128, 4_096, 2, 32)
    }

    #[test]
    fn document_and_tree_routes_force_reference_presence_and_bits() {
        let first = rich_reference(7);
        let second = reference(8);
        let sidebar = reference(9);
        let source = document(&[first.clone(), second.clone()], &sidebar);
        let (snapshot, report) =
            decode_document_sheet_order_with_report(&source, options(&source)).unwrap();
        assert_eq!(
            snapshot
                .sheet_references()
                .iter()
                .copied()
                .map(ReferenceSnapshot::identifier)
                .collect::<Vec<_>>(),
            [7, 8]
        );
        assert_eq!(snapshot.sheet_references()[0].deprecated_type(), Some(-1));
        assert_eq!(
            snapshot.sheet_references()[0].deprecated_is_external(),
            Some(false)
        );
        assert_eq!(snapshot.sidebar_order().identifier(), 9);
        assert_eq!(report.references(), 3);
        assert_eq!(
            report.reference_bytes(),
            first.len() + second.len() + sidebar.len()
        );
        assert_eq!(report.max_depth(), 2);
        assert_eq!(
            report.work_bytes(),
            source.len() * 2 + report.reference_bytes() * 2
        );

        let mut tree = Vec::new();
        push_bytes_field(&mut tree, TREE_NODE_CHILDREN_FIELD, &first);
        push_bytes_field(&mut tree, TREE_NODE_CHILDREN_FIELD, &second);
        push_bytes_field(&mut tree, TREE_NODE_OBJECT_FIELD, &sidebar);
        let tree_snapshot = decode_tree_node(&tree, options(&tree)).unwrap();
        assert_eq!(tree_snapshot.child_references().len(), 2);
        assert_eq!(
            tree_snapshot
                .object_reference()
                .map(ReferenceSnapshot::identifier),
            Some(9)
        );
    }

    #[test]
    fn distributed_limits_are_inclusive_and_exact() {
        let first = rich_reference(7);
        let second = rich_reference(8);
        let sidebar = rich_reference(9);
        let source = document(&[first, second], &sidebar);
        let (_, exact) = decode_document_sheet_order_with_report(
            &source,
            DecodeOptions::new(source.len(), 128, usize::MAX, 2, 32),
        )
        .unwrap();
        let exact_options = DecodeOptions::new(
            source.len(),
            exact.fields(),
            exact.work_bytes(),
            exact.max_depth(),
            exact.references(),
        );
        assert!(decode_document_sheet_order(&source, exact_options).is_ok());

        let fields = decode_document_sheet_order(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields() - 1,
                exact.work_bytes(),
                2,
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            fields.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: exact.fields(),
                maximum: exact.fields() - 1,
            })
        );

        let work = decode_document_sheet_order(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes() - 1,
                2,
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            work.resource_limit(),
            Some(DecodeLimit::Work {
                observed: exact.work_bytes(),
                maximum: exact.work_bytes() - 1,
            })
        );

        let references = decode_document_sheet_order(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes(),
                2,
                exact.references() - 1,
            ),
        )
        .unwrap_err();
        assert_eq!(
            references.resource_limit(),
            Some(DecodeLimit::References {
                observed: exact.references(),
                maximum: exact.references() - 1,
            })
        );

        let bytes = decode_document_sheet_order(
            &source,
            DecodeOptions::new(
                source.len() - 1,
                exact.fields(),
                exact.work_bytes(),
                2,
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            bytes.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );

        let nesting = decode_document_sheet_order(
            &source,
            DecodeOptions::new(
                source.len(),
                exact.fields(),
                exact.work_bytes(),
                1,
                exact.references(),
            ),
        )
        .unwrap_err();
        assert_eq!(
            nesting.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn wide_reference_routing_scales_linearly_and_max_minus_one_fails_in_preflight() {
        const SMALL_SHEETS: usize = 4_096;
        const LARGE_SHEETS: usize = 8_192;

        let small_source = wide_document(SMALL_SHEETS);
        let large_source = wide_document(LARGE_SHEETS);
        let (small, small_report) = decode_document_sheet_order_with_report(
            &small_source,
            DecodeOptions::new(
                small_source.len(),
                (SMALL_SHEETS + 1) * 2,
                usize::MAX,
                2,
                SMALL_SHEETS + 1,
            ),
        )
        .unwrap();
        let (large, large_report) = decode_document_sheet_order_with_report(
            &large_source,
            DecodeOptions::new(
                large_source.len(),
                (LARGE_SHEETS + 1) * 2,
                usize::MAX,
                2,
                LARGE_SHEETS + 1,
            ),
        )
        .unwrap();

        assert_eq!(small.sheet_references().len(), SMALL_SHEETS);
        assert_eq!(large.sheet_references().len(), LARGE_SHEETS);
        assert_eq!(small_report.references(), SMALL_SHEETS + 1);
        assert_eq!(large_report.references(), LARGE_SHEETS + 1);
        assert_eq!(small_report.fields(), (SMALL_SHEETS + 1) * 2);
        assert_eq!(large_report.fields(), (LARGE_SHEETS + 1) * 2);

        let small_output_bytes = small_report
            .references()
            .saturating_mul(size_of::<ReferenceSnapshot>());
        let large_output_bytes = large_report
            .references()
            .saturating_mul(size_of::<ReferenceSnapshot>());
        assert!(large_source.len() * 10 <= small_source.len() * 23);
        assert!(large_report.fields() * 10 <= small_report.fields() * 23);
        assert!(large_report.work_bytes() * 10 <= small_report.work_bytes() * 23);
        assert!(large_output_bytes * 10 <= small_output_bytes * 23);

        let error = decode_document_sheet_order(
            &large_source,
            DecodeOptions::new(
                large_source.len(),
                large_report.fields(),
                large_report.work_bytes(),
                2,
                large_report.references() - 1,
            ),
        )
        .unwrap_err();
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::References {
                observed: large_report.references(),
                maximum: large_report.references() - 1,
            })
        );
        assert_eq!(error.allocation_amount(), None);
    }

    #[test]
    fn malformed_known_fields_and_noncanonical_encodings_are_rejected() {
        let sidebar = reference(9);
        let valid = document(&[reference(7)], &sidebar);
        assert!(decode_document_sheet_order(&valid, options(&valid)).is_ok());

        let missing_sidebar = {
            let mut value = Vec::new();
            push_bytes_field(&mut value, DOCUMENT_SHEETS_FIELD, &reference(7));
            value
        };
        assert!(decode_document_sheet_order(&missing_sidebar, options(&missing_sidebar)).is_err());

        let duplicate_sidebar = {
            let mut value = document(&[], &sidebar);
            push_bytes_field(&mut value, DOCUMENT_SIDEBAR_ORDER_FIELD, &sidebar);
            value
        };
        assert!(
            decode_document_sheet_order(&duplicate_sidebar, options(&duplicate_sidebar)).is_err()
        );

        let wrong_outer_wire = [0x08, 1, 0x2a, 2, 0x08, 9];
        assert!(
            decode_document_sheet_order(&wrong_outer_wire, options(&wrong_outer_wire)).is_err()
        );

        let duplicate_identifier = [0x08, 1, 0x08, 2];
        let duplicate = document(&[duplicate_identifier.to_vec()], &sidebar);
        assert!(decode_document_sheet_order(&duplicate, options(&duplicate)).is_err());

        let overlong_identifier = [0x08, 0x81, 0x00];
        let overlong = document(&[overlong_identifier.to_vec()], &sidebar);
        assert!(decode_document_sheet_order(&overlong, options(&overlong)).is_err());

        let invalid_bool = [0x08, 7, 0x18, 2];
        let invalid_bool_document = document(&[invalid_bool.to_vec()], &sidebar);
        assert!(
            decode_document_sheet_order(&invalid_bool_document, options(&invalid_bool_document))
                .is_err()
        );

        let short_negative_i32 = [0x08, 7, 0x10, 0xff, 0xff, 0xff, 0xff, 0x0f];
        let invalid_i32 = document(&[short_negative_i32.to_vec()], &sidebar);
        assert!(decode_document_sheet_order(&invalid_i32, options(&invalid_i32)).is_err());

        let overlong_tag = [0x8a, 0x00, 0x02, 0x08, 0x07, 0x2a, 0x02, 0x08, 0x09];
        assert!(decode_document_sheet_order(&overlong_tag, options(&overlong_tag)).is_err());
        let overlong_length = [0x0a, 0x82, 0x00, 0x08, 0x07, 0x2a, 0x02, 0x08, 0x09];
        assert!(decode_document_sheet_order(&overlong_length, options(&overlong_length)).is_err());
        let group = [0x0b, 0x0c, 0x2a, 0x02, 0x08, 0x09];
        assert!(decode_document_sheet_order(&group, options(&group)).is_err());
    }

    #[test]
    fn unknown_scalar_and_length_delimited_records_remain_inert() {
        let mut selected = reference(7);
        push_varint_field(&mut selected, 91, 5);
        push_varint(&mut selected, (92_u64 << 3) | 1);
        selected.extend_from_slice(&[0; 8]);
        push_bytes_field(&mut selected, 93, &[0xff, 0x00]);
        push_varint(&mut selected, (94_u64 << 3) | 5);
        selected.extend_from_slice(&[0; 4]);
        let source = document(&[selected], &reference(9));
        assert_eq!(
            decode_document_sheet_order(&source, options(&source))
                .unwrap()
                .sheet_references()[0]
                .identifier(),
            7
        );
    }

    #[test]
    fn empty_tree_is_valid_and_reports_two_zero_byte_owner_passes() {
        let (snapshot, report) =
            decode_tree_node_with_report(&[], DecodeOptions::new(0, 0, 0, 1, 0)).unwrap();
        assert!(snapshot.child_references().is_empty());
        assert_eq!(snapshot.object_reference(), None);
        assert_eq!(report.fields(), 0);
        assert_eq!(report.work_bytes(), 0);
        assert_eq!(report.max_depth(), 1);
        assert_eq!(report.references(), 0);
        assert_eq!(report.reference_bytes(), 0);
    }

    #[test]
    fn errors_are_content_free_and_invalid_profiles_are_typed() {
        let error = decode_tree_node(&[0xff], DecodeOptions::new(1, 1, 1, 1, 1)).unwrap_err();
        assert_eq!(error.to_string(), "invalid Numbers sheet-order payload");
        assert_eq!(error.resource_limit(), None);
        assert_eq!(error.allocation_amount(), None);

        let too_deep = decode_tree_node(&[], DecodeOptions::new(0, 0, 0, 65, 0)).unwrap_err();
        assert_eq!(
            too_deep.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 65,
                maximum: 64,
            })
        );
    }
}
