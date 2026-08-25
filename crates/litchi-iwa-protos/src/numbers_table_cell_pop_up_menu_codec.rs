//! Strict generated-free projection for Numbers control-cell popup menus.
//!
//! The native `PopUpMenuModel` contains a deprecated repeated field and a
//! repeated `TSCE.CellValueArchive` field.  This module deliberately parses
//! those repeated records by borrowing the caller's bytes; the private Buffa
//! sidecar contains only singular parity envelopes and cannot materialize
//! input-width collections.  Unknown fields/groups are retained by raw source
//! spans, while every known field is checked for canonical wire shape before
//! a semantic snapshot is published.
//! Control-list refcount validation remains in
//! [`crate::numbers_table_cell_storage_codec`]; this codec owns the selected
//! `PopUpMenuModel` and `CellSpecArchive` payloads.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire helpers stay beside the snapshots they construct."
)]

use core::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_table_cell_pop_up_menu_generated::LitchiIwaNumbersTableCellPopUpMenuProjection as projection;

const POPUP_ITEM_FIELD: u32 = 2;
const POPUP_DEPRECATED_ITEM_FIELD: u32 = 1;
const CELL_VALUE_TYPE_FIELD: u32 = 1;
const CELL_VALUE_STRING_FIELD: u32 = 5;
const STRING_VALUE_FIELD: u32 = 1;
const STRING_FORMAT_FIELD: u32 = 2;
const STRING_IMPLICIT_FIELD: u32 = 3;
const STRING_EXPLICIT_FIELD: u32 = 4;
const STRING_REGEX_FIELD: u32 = 5;
const STRING_CASE_SENSITIVE_REGEX_FIELD: u32 = 6;
const FORMAT_TYPE_FIELD: u32 = 1;
const CELL_SPEC_INTERACTION_FIELD: u32 = 1;
const CELL_SPEC_MODEL_FIELD: u32 = 6;
const CELL_SPEC_FIRST_FIELD: u32 = 7;
const CELL_SPEC_FORMULA_FIELD: u32 = 2;
const CELL_SPEC_DEPRECATED_LABEL_FIELD: u32 = 8;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;
const POPUP_INTERACTION_TYPE: u32 = 7;
const FORMAT_TYPE_TEXT: u64 = 260;
const STRING_TYPE: u64 = 5;
const NIL_TYPE: u64 = 1;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MAX_RECURSION: u32 = 64;

/// Finite aggregate limits for one popup-model or cell-spec operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
    max_items: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    /// Construct all finite decode and rewrite ceilings explicitly.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
        max_items: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
            max_items,
            max_text_bytes,
        }
    }

    /// Conservative finite limits derived from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.checked_mul(2).unwrap_or(usize::MAX),
            bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            bytes.checked_mul(16).unwrap_or(usize::MAX).max(1),
            MAX_RECURSION,
            bytes.max(1),
            bytes.max(1),
            bytes,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the item-count ceiling.
    #[must_use]
    pub const fn with_max_items(mut self, maximum: usize) -> Self {
        self.max_items = maximum;
        self
    }

    /// Replace the selected text-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }
}

/// Typed finite failure for strict popup decoding or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input payload exceeded the configured source ceiling.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate payload exceeded the configured output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Known and unknown field records exceeded the aggregate ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate inspected bytes exceeded the work ceiling.
    Work { observed: usize, maximum: usize },
    /// Nested message/group depth exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Reference envelopes exceeded the aggregate ceiling.
    References { observed: usize, maximum: usize },
    /// Repeated popup values exceeded the aggregate ceiling.
    Items { observed: usize, maximum: usize },
    /// Selected UTF-8 text exceeded the aggregate ceiling.
    Text { observed: usize, maximum: usize },
    /// A fallible output or scratch reservation was refused.
    Allocation { requested: usize },
    /// Retained candidate bytes exceeded the execution ceiling.
    Retained { observed: usize, maximum: usize },
    /// Temporary scratch bytes exceeded the execution ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// Strict popup codec error.  Native identifiers are intentionally omitted
/// from diagnostics to avoid leaking object-routing data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    const fn invalid() -> Self {
        Self { limit: None }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }

    /// Return a typed resource failure, when this error is a limit failure.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    /// Return the refused allocation size, if applicable.
    #[must_use]
    pub const fn allocation_requested(self) -> Option<usize> {
        match self.limit {
            Some(DecodeLimit::Allocation { requested }) => Some(requested),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid Numbers table-cell popup-menu payload")
    }
}

impl std::error::Error for DecodeError {}

/// Exact aggregate consumption for one strict operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn items(self) -> usize {
        self.items
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// A strict canonical `TSP.Reference` projection.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl fmt::Debug for ReferenceSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReferenceSnapshot")
            .field("identifier", &"<redacted>")
            .field("deprecated_type", &self.deprecated_type)
            .field("deprecated_is_external", &self.deprecated_is_external)
            .finish()
    }
}

impl ReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// One borrowed popup string item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuItem<'source> {
    value: &'source str,
}

impl<'source> PopUpMenuItem<'source> {
    #[must_use]
    pub const fn value(self) -> &'source str {
        self.value
    }
}

/// Borrowed semantic facts for a strict `PopUpMenuModel` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuModelSnapshot<'source> {
    source: &'source [u8],
    first_nil: bool,
    item_count: usize,
    text_bytes: usize,
}

impl<'source> PopUpMenuModelSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn has_nil_sentinel(self) -> bool {
        self.first_nil
    }
    #[must_use]
    pub const fn item_count(self) -> usize {
        self.item_count
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Iterate values without allocating an input-width collection.
    pub fn items(self) -> PopUpMenuItems<'source> {
        PopUpMenuItems {
            source: self.source,
            offset: 0,
            index: 0,
        }
    }
}

/// Borrowed iterator over popup string values.
#[derive(Debug, Clone, Copy)]
pub struct PopUpMenuItems<'source> {
    source: &'source [u8],
    offset: usize,
    index: usize,
}

impl<'source> Iterator for PopUpMenuItems<'source> {
    type Item = PopUpMenuItem<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.offset < self.source.len() {
            let field = parse_one_field(self.source, self.offset, 0).ok()?;
            self.offset = field.end;
            if field.number != POPUP_ITEM_FIELD {
                continue;
            }
            let payload = field.payload?;
            let value = match parse_cell_value_for_iterator(payload) {
                Ok(Some(value)) => value,
                Ok(None) => {
                    self.index = 1;
                    continue;
                },
                Err(_error) => return None,
            };
            self.index = self.index.saturating_add(1);
            return Some(PopUpMenuItem { value });
        }
        None
    }
}

/// Borrowed facts for one popup `CellSpecArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellSpecSnapshot<'source> {
    source: &'source [u8],
    interaction_type: u32,
    popup_model: ReferenceSnapshot,
    starts_with_first: bool,
}

impl<'source> CellSpecSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn interaction_type(self) -> u32 {
        self.interaction_type
    }
    #[must_use]
    pub const fn popup_model(self) -> ReferenceSnapshot {
        self.popup_model
    }
    #[must_use]
    pub const fn starts_with_first(self) -> bool {
        self.starts_with_first
    }
}

/// Visitor for streaming popup item values without allocating a vector.
pub trait PopUpMenuVisitor {
    /// Observe one source-borrowed string item.
    fn visit_item(&mut self, item: PopUpMenuItem<'_>) -> Result<(), DecodeError>;
}

/// Decode a popup model without exposing generated Buffa values.
pub fn decode_popup_menu_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PopUpMenuModelSnapshot<'_>, DecodeError> {
    Ok(decode_popup_menu_model_with_report(source, options)?.0)
}

/// Decode a popup model and return its aggregate strict resource report.
pub fn decode_popup_menu_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(PopUpMenuModelSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_popup_model_with_visitor(source, &mut budget, None)?;
    Ok((snapshot, budget.finish(0)))
}

/// Decode a popup model and stream each item to a caller-owned visitor.
pub fn decode_popup_menu_model_with_visitor<V: PopUpMenuVisitor>(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let _snapshot = parse_popup_model_with_visitor(source, &mut budget, Some(visitor))?;
    Ok(budget.finish(0))
}

/// Decode a popup cell-spec archive.
pub fn decode_cell_spec(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CellSpecSnapshot<'_>, DecodeError> {
    Ok(decode_cell_spec_with_report(source, options)?.0)
}

/// Decode a popup cell-spec archive with aggregate resource accounting.
pub fn decode_cell_spec_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellSpecSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_cell_spec(source, &mut budget)?;
    if !budget.unknown_fields {
        buffa_cell_spec_parity(source, &mut budget)?;
    }
    Ok((snapshot, budget.finish(0)))
}

fn buffa_cell_spec_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::CellSpecArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_message_bytes)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

fn buffa_cell_value_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::CellValueArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_message_bytes)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

fn buffa_string_value_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::StringCellValueArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_message_bytes)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    raw: &'source [u8],
    payload: Option<&'source [u8]>,
    varint: Option<u64>,
    varint_canonical: bool,
    nested_fields: usize,
    nested_work_bytes: usize,
    nested_max_depth: u32,
    end: usize,
}

impl Field<'_> {
    fn known_varint(self) -> Result<u64, DecodeError> {
        if self.wire != 0 || !self.varint_canonical {
            return Err(DecodeError::invalid());
        }
        self.varint.ok_or_else(DecodeError::invalid)
    }
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    options: DecodeOptions,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    unknown_fields: bool,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::InputBytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION.min(options.recursion_limit.max(1)),
            }));
        }
        Ok(Self {
            options,
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            items: 0,
            text_bytes: 0,
            unknown_fields: false,
        })
    }

    fn field(&mut self, amount: usize, depth: u32) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        self.max_depth = self.max_depth.max(depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }

    fn reference(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.references = self.references.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::References {
                observed: usize::MAX,
                maximum: self.options.max_references,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        if self.references > self.options.max_references {
            return Err(DecodeError::limited(DecodeLimit::References {
                observed: self.references,
                maximum: self.options.max_references,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn item(&mut self, text: usize) -> Result<(), DecodeError> {
        self.items = self.items.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Items {
                observed: usize::MAX,
                maximum: self.options.max_items,
            })
        })?;
        self.text_bytes = self.text_bytes.checked_add(text).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Text {
                observed: usize::MAX,
                maximum: self.options.max_text_bytes,
            })
        })?;
        if self.items > self.options.max_items {
            return Err(DecodeError::limited(DecodeLimit::Items {
                observed: self.items,
                maximum: self.options.max_items,
            }));
        }
        if self.text_bytes > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: self.text_bytes,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }

    fn mark_unknown(&mut self) {
        self.unknown_fields = true;
    }

    fn nested_fields(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
    ) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(fields).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(work_bytes).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        self.max_depth = self.max_depth.max(max_depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        if max_depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: max_depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }

    fn finish(self, output_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            items: self.items,
            text_bytes: self.text_bytes,
            allocations: 0,
            retained_bytes: output_bytes,
            scratch_bytes: 0,
        }
    }
}

fn parse_popup_model_with_visitor<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    mut visitor: Option<&mut dyn PopUpMenuVisitor>,
) -> Result<PopUpMenuModelSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut first_nil = false;
    let mut saw_item = false;
    let mut text_bytes = 0usize;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            POPUP_DEPRECATED_ITEM_FIELD => return Err(DecodeError::invalid()),
            POPUP_ITEM_FIELD => {
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                let (value, is_nil, value_text_bytes) = parse_cell_value(payload, budget, 1)?;
                if !saw_item {
                    if !is_nil {
                        return Err(DecodeError::invalid());
                    }
                    first_nil = true;
                    saw_item = true;
                } else if is_nil {
                    return Err(DecodeError::invalid());
                } else {
                    text_bytes = text_bytes
                        .checked_add(value_text_bytes)
                        .ok_or_else(DecodeError::invalid)?;
                    budget.item(value_text_bytes)?;
                    if value.is_none() {
                        return Err(DecodeError::invalid());
                    }
                    if let Some(visitor) = visitor.as_deref_mut() {
                        visitor.visit_item(PopUpMenuItem {
                            value: value.ok_or_else(DecodeError::invalid)?,
                        })?;
                    }
                }
            },
            _ => budget.mark_unknown(),
        }
    }
    if !saw_item || !first_nil || budget.items == 0 {
        return Err(DecodeError::invalid());
    }
    Ok(PopUpMenuModelSnapshot {
        source,
        first_nil,
        item_count: budget.items,
        text_bytes,
    })
}

fn parse_cell_value<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(Option<&'source str>, bool, usize), DecodeError> {
    let mut offset = 0usize;
    let mut value_type = None;
    let mut string_payload = None;
    let mut wrapper_seen = [false; 5];
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            CELL_VALUE_TYPE_FIELD => {
                if field.wire != 0 || value_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                if !matches!(value, 1..=5) {
                    return Err(DecodeError::invalid());
                }
                value_type = Some(value);
            },
            2..=6 => {
                if field.wire != 2 {
                    return Err(DecodeError::invalid());
                }
                let slot = usize::try_from(field.number - 2).map_err(|_| DecodeError::invalid())?;
                if wrapper_seen[slot] {
                    return Err(DecodeError::invalid());
                }
                wrapper_seen[slot] = true;
                if field.number == CELL_VALUE_STRING_FIELD {
                    string_payload = field.payload;
                }
            },
            _ => budget.mark_unknown(),
        }
    }
    let value_type = value_type.ok_or_else(DecodeError::invalid)?;
    match value_type {
        NIL_TYPE if wrapper_seen.iter().all(|seen| !seen) => {
            if !budget.unknown_fields {
                buffa_cell_value_parity(source, budget)?;
            }
            Ok((None, true, 0))
        },
        STRING_TYPE
            if wrapper_seen[3]
                && !wrapper_seen[..3].iter().any(|seen| *seen)
                && !wrapper_seen[4] =>
        {
            let payload = string_payload.ok_or_else(DecodeError::invalid)?;
            let (value, text_bytes) = parse_string_value(payload, budget, depth + 1)?;
            if !budget.unknown_fields {
                buffa_cell_value_parity(source, budget)?;
                buffa_string_value_parity(payload, budget)?;
            }
            Ok((Some(value), false, text_bytes))
        },
        _ => Err(DecodeError::invalid()),
    }
}

fn parse_string_value<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(&'source str, usize), DecodeError> {
    let mut offset = 0usize;
    let mut value = None;
    let mut format = None;
    let mut explicit = None;
    let mut regex = None;
    let mut case_sensitive = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            STRING_VALUE_FIELD => {
                if field.wire != 2 || value.is_some() {
                    return Err(DecodeError::invalid());
                }
                let bytes = field.payload.ok_or_else(DecodeError::invalid)?;
                let text = str::from_utf8(bytes).map_err(|_| DecodeError::invalid())?;
                if text.chars().any(char::is_control) {
                    return Err(DecodeError::invalid());
                }
                value = Some((text, bytes.len()));
            },
            STRING_FORMAT_FIELD => {
                if field.wire != 2 || format.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                parse_text_format(payload, budget, depth + 1)?;
                format = Some(());
            },
            STRING_IMPLICIT_FIELD => return Err(DecodeError::invalid()),
            STRING_EXPLICIT_FIELD => set_bool(&mut explicit, field, false)?,
            STRING_REGEX_FIELD => set_bool(&mut regex, field, false)?,
            STRING_CASE_SENSITIVE_REGEX_FIELD => set_bool(&mut case_sensitive, field, false)?,
            _ => budget.mark_unknown(),
        }
    }
    let (text, text_bytes) = value.ok_or_else(DecodeError::invalid)?;
    if format.is_none()
        || explicit != Some(false)
        || regex != Some(false)
        || case_sensitive != Some(false)
    {
        return Err(DecodeError::invalid());
    }
    Ok((text, text_bytes))
}

fn parse_text_format(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    let mut format_type = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            FORMAT_TYPE_FIELD => {
                if field.wire != 0 || format_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                format_type = Some(field.known_varint()?);
            },
            _ => budget.mark_unknown(),
        }
    }
    if format_type != Some(FORMAT_TYPE_TEXT) {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parse_cell_spec<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<CellSpecSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut interaction = None;
    let mut model = None;
    let mut starts = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            CELL_SPEC_INTERACTION_FIELD => {
                if field.wire != 0 || interaction.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                interaction = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?);
            },
            CELL_SPEC_FORMULA_FIELD | 3..=5 | CELL_SPEC_DEPRECATED_LABEL_FIELD => {
                return Err(DecodeError::invalid());
            },
            CELL_SPEC_MODEL_FIELD => {
                if field.wire != 2 || model.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                model = Some(parse_reference(payload, budget, 1)?);
            },
            CELL_SPEC_FIRST_FIELD => set_bool(&mut starts, field, true)?,
            _ => budget.mark_unknown(),
        }
    }
    if interaction != Some(POPUP_INTERACTION_TYPE) {
        return Err(DecodeError::invalid());
    }
    Ok(CellSpecSnapshot {
        source,
        interaction_type: interaction.ok_or_else(DecodeError::invalid)?,
        popup_model: model.ok_or_else(DecodeError::invalid)?,
        starts_with_first: starts.ok_or_else(DecodeError::invalid)?,
    })
}

fn parse_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.reference(source.len())?;
    let mut offset = 0usize;
    let mut identifier = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if field.wire != 0 || identifier.is_some() {
                    return Err(DecodeError::invalid());
                }
                identifier = Some(field.known_varint()?);
            },
            // These legacy fields are omitted by the canonical writer.  A
            // source-preserving CellSpec rewrite has no safe way to retain
            // their exact presence/framing, so reject them even when their
            // values are the benign defaults (zero/false).
            REFERENCE_TYPE_FIELD | REFERENCE_EXTERNAL_FIELD => {
                return Err(DecodeError::invalid());
            },
            _ => budget.mark_unknown(),
        }
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(DecodeError::invalid)?;
    Ok(ReferenceSnapshot {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    })
}

fn set_bool(
    target: &mut Option<bool>,
    field: Field<'_>,
    expected_presence: bool,
) -> Result<(), DecodeError> {
    if field.wire != 0 || target.is_some() {
        return Err(DecodeError::invalid());
    }
    let value = field.known_varint()?;
    if value > 1 {
        return Err(DecodeError::invalid());
    }
    if expected_presence && value > 1 {
        return Err(DecodeError::invalid());
    }
    *target = Some(value != 0);
    Ok(())
}

fn parse_cell_value_for_iterator(source: &[u8]) -> Result<Option<&str>, DecodeError> {
    let mut budget = Budget::new(source, DecodeOptions::for_source(source))?;
    let (value, is_nil, _) = parse_cell_value(source, &mut budget, 1)?;
    if is_nil {
        return Ok(None);
    }
    value.map(Some).ok_or_else(DecodeError::invalid)
}

fn parse_one_field<'source>(
    source: &'source [u8],
    offset: usize,
    depth: u32,
) -> Result<Field<'source>, DecodeError> {
    parse_one_field_limited(source, offset, depth, MAX_RECURSION)
}

fn parse_one_field_limited<'source>(
    source: &'source [u8],
    offset: usize,
    depth: u32,
    recursion_limit: u32,
) -> Result<Field<'source>, DecodeError> {
    let start = offset;
    let (key, key_len) = read_varint(source, offset)?;
    let number = u32::try_from(key >> 3).map_err(|_| DecodeError::invalid())?;
    let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let mut cursor = offset
        .checked_add(key_len)
        .ok_or_else(DecodeError::invalid)?;
    let mut payload = None;
    let mut varint = None;
    let mut varint_canonical = true;
    match wire {
        0 => {
            let (value, length, canonical) = read_varint_relaxed(source, cursor)?;
            varint = Some(value);
            varint_canonical = canonical;
            cursor = cursor
                .checked_add(length)
                .ok_or_else(DecodeError::invalid)?;
        },
        1 => {
            cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
            if cursor > source.len() {
                return Err(DecodeError::invalid());
            }
        },
        2 => {
            let (length, length_bytes) = read_varint(source, cursor)?;
            cursor = cursor
                .checked_add(length_bytes)
                .ok_or_else(DecodeError::invalid)?;
            let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
            let end = cursor
                .checked_add(length)
                .ok_or_else(DecodeError::invalid)?;
            if end > source.len() {
                return Err(DecodeError::invalid());
            }
            payload = Some(&source[cursor..end]);
            cursor = end;
        },
        3 => {
            if depth >= recursion_limit {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: depth.saturating_add(1),
                    maximum: recursion_limit,
                }));
            }
            let group = skip_group(
                source,
                cursor,
                number,
                depth.saturating_add(1),
                recursion_limit,
            )?;
            cursor = group.end;
            return Ok(Field {
                number,
                wire,
                raw: &source[start..cursor],
                payload,
                varint,
                varint_canonical,
                nested_fields: group.fields,
                nested_work_bytes: group.work_bytes,
                nested_max_depth: group.max_depth,
                end: cursor,
            });
        },
        4 => return Err(DecodeError::invalid()),
        5 => {
            cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?;
            if cursor > source.len() {
                return Err(DecodeError::invalid());
            }
        },
        _ => return Err(DecodeError::invalid()),
    }
    Ok(Field {
        number,
        wire,
        raw: &source[start..cursor],
        payload,
        varint,
        varint_canonical,
        nested_fields: 0,
        nested_work_bytes: 0,
        nested_max_depth: 0,
        end: cursor,
    })
}

#[derive(Debug, Clone, Copy)]
struct GroupScan {
    end: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

fn skip_group(
    source: &[u8],
    mut cursor: usize,
    root_number: u32,
    depth: u32,
    recursion_limit: u32,
) -> Result<GroupScan, DecodeError> {
    let mut stack = [0u32; 64];
    let mut stack_len = 1usize;
    stack[0] = root_number;
    let mut fields = 0usize;
    let mut work_bytes = 0usize;
    let mut max_depth = depth;
    while cursor < source.len() {
        let field_start = cursor;
        let (key, key_len) = read_varint(source, cursor)?;
        let number = u32::try_from(key >> 3).map_err(|_| DecodeError::invalid())?;
        let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid())?;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(DecodeError::invalid());
        }
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        match wire {
            0 => {
                let (_, length, _) = read_varint_relaxed(source, cursor)?;
                cursor = cursor
                    .checked_add(length)
                    .ok_or_else(DecodeError::invalid)?;
            },
            1 => cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?,
            2 => {
                let (length, length_bytes) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(length_bytes)
                    .ok_or_else(DecodeError::invalid)?;
                let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
                cursor = cursor
                    .checked_add(length)
                    .ok_or_else(DecodeError::invalid)?;
            },
            3 => {
                if stack_len >= stack.len()
                    || u32::try_from(stack_len).unwrap_or(u32::MAX) >= recursion_limit
                {
                    return Err(DecodeError::limited(DecodeLimit::Nesting {
                        observed: depth.saturating_add(stack_len as u32),
                        maximum: recursion_limit,
                    }));
                }
                stack[stack_len] = number;
                stack_len += 1;
            },
            4 => {
                if stack_len == 0 || stack[stack_len - 1] != number {
                    return Err(DecodeError::invalid());
                }
                stack_len -= 1;
                if stack_len == 0 {
                    let raw_len = cursor
                        .checked_sub(field_start)
                        .ok_or_else(DecodeError::invalid)?;
                    fields = fields.checked_add(1).ok_or_else(|| {
                        DecodeError::limited(DecodeLimit::Fields {
                            observed: usize::MAX,
                            maximum: usize::MAX,
                        })
                    })?;
                    work_bytes = work_bytes.checked_add(raw_len).ok_or_else(|| {
                        DecodeError::limited(DecodeLimit::Work {
                            observed: usize::MAX,
                            maximum: usize::MAX,
                        })
                    })?;
                    return Ok(GroupScan {
                        end: cursor,
                        fields,
                        work_bytes,
                        max_depth,
                    });
                }
            },
            5 => cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?,
            _ => return Err(DecodeError::invalid()),
        }
        if cursor > source.len() {
            return Err(DecodeError::invalid());
        }
        let raw_len = cursor
            .checked_sub(field_start)
            .ok_or_else(DecodeError::invalid)?;
        fields = fields.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
        work_bytes = work_bytes.checked_add(raw_len).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
        max_depth = max_depth.max(depth.saturating_add(stack_len as u32).saturating_sub(1));
    }
    Err(DecodeError::invalid())
}

fn read_varint(source: &[u8], offset: usize) -> Result<(u64, usize), DecodeError> {
    let (value, consumed, canonical) = read_varint_relaxed(source, offset)?;
    if !canonical {
        return Err(DecodeError::invalid());
    }
    Ok((value, consumed))
}

fn read_varint_relaxed(source: &[u8], offset: usize) -> Result<(u64, usize, bool), DecodeError> {
    let mut value = 0u64;
    let mut shift = 0u32;
    let mut index = offset;
    while index < source.len() && index - offset < 10 {
        let byte = source[index];
        let part = u64::from(byte & 0x7f);
        if shift == 63 && part > 1 {
            return Err(DecodeError::invalid());
        }
        value |= part.checked_shl(shift).ok_or_else(DecodeError::invalid)?;
        index += 1;
        if byte & 0x80 == 0 {
            let consumed = index - offset;
            return Ok((value, consumed, encoded_varint_len(value) == consumed));
        }
        shift += 7;
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

/// Exact preflight requirements for a prepared popup write.
///
/// The field/work/depth/item/reference values are the strict candidate-decode
/// values used by `execute` after emission; execution therefore cannot publish
/// a candidate whose wire shape diverges from the prepared projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn items(self) -> usize {
        self.items
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Caller-provided execution ceilings for a prepared write.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from a prepared requirement set.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        Self {
            output_bytes: requirements.output_bytes,
            fields: requirements.fields,
            work_bytes: requirements.work_bytes,
            max_depth: requirements.max_depth,
            references: requirements.references,
            items: requirements.items,
            text_bytes: requirements.text_bytes,
            allocations: requirements.allocations,
            retained_bytes: requirements.retained_bytes,
            scratch_bytes: requirements.scratch_bytes,
        }
    }

    /// Start with independent outer limits set to their largest values.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            references: usize::MAX,
            items: usize::MAX,
            text_bytes: usize::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }
    #[must_use]
    pub const fn with_references(mut self, value: usize) -> Self {
        self.references = value;
        self
    }
    #[must_use]
    pub const fn with_items(mut self, value: usize) -> Self {
        self.items = value;
        self
    }
    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Output of a prepared popup write. Candidates remain private to the package
/// transaction until locality and reopen checks complete.
#[derive(Debug, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: DecodeReport,
}

impl RewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
    #[must_use]
    pub const fn report(&self) -> DecodeReport {
        self.report
    }
}

/// Prepared canonical `PopUpMenuModel` write. Preparation allocates no
/// candidate bytes and only borrows caller-owned item text.
#[derive(Debug, Clone, Copy)]
pub struct PreparedPopUpMenuModelWrite<'items> {
    items: &'items [&'items str],
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl<'items> PreparedPopUpMenuModelWrite<'items> {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_popup_model(&mut bytes, self.items)?;
        verify_popup_model_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical popup `CellSpecArchive` write.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCellSpecWrite {
    model_identifier: u64,
    starts_with_first: bool,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCellSpecWrite {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_cell_spec(&mut bytes, self.model_identifier, self.starts_with_first)?;
        verify_cell_spec_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepare a canonical popup model payload from borrowed item text.
pub fn prepare_popup_menu_model_write<'items>(
    items: &'items [&'items str],
    options: DecodeOptions,
) -> Result<PreparedPopUpMenuModelWrite<'items>, DecodeError> {
    if items.is_empty() || items.len() > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Items {
            observed: items.len(),
            maximum: options.max_items,
        }));
    }
    let mut text_bytes = 0usize;
    for text in items {
        validate_text(text, options.max_text_bytes)?;
        text_bytes = text_bytes
            .checked_add(text.len())
            .ok_or_else(DecodeError::invalid)?;
    }
    let output_bytes = popup_model_output_len(items)?;
    let fields = 2usize
        .checked_add(
            items
                .len()
                .checked_mul(9)
                .ok_or_else(DecodeError::invalid)?,
        )
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: popup_model_work_bytes(items)?
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 3,
        references: 0,
        items: items.len(),
        text_bytes,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedPopUpMenuModelWrite {
        items,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical popup model encoding wrapper.
pub fn canonical_popup_menu_model(
    items: &[&str],
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_popup_menu_model_write(items, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for package owners that describe canonical model
/// construction as a rewrite operation.
///
/// This is a new-payload constructor, not a source patch: it intentionally
/// does not accept an existing payload and therefore does not claim to carry
/// unknown fields or deprecated-but-valid source framing into the result.
pub fn rewrite_popup_menu_model(
    items: &[&str],
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_popup_menu_model(items, options)
}

/// Prepare a canonical popup `CellSpecArchive` payload.
pub fn prepare_cell_spec_write(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<PreparedCellSpecWrite, DecodeError> {
    if model_identifier == 0 {
        return Err(DecodeError::invalid());
    }
    let output_bytes = cell_spec_output_len(model_identifier)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: 4,
        work_bytes: output_bytes
            .checked_add(1 + encoded_varint_len(model_identifier))
            .and_then(|work| work.checked_add(1 + encoded_varint_len(model_identifier)))
            .and_then(|work| work.checked_add(output_bytes))
            .and_then(|work| work.checked_add(output_bytes))
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 1,
        references: 1,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedCellSpecWrite {
        model_identifier,
        starts_with_first,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical popup cell-spec encoding wrapper.
pub fn canonical_cell_spec(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_cell_spec_write(model_identifier, starts_with_first, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for canonical control-cell spec publication.
///
/// This is likewise creation-only.  Existing `CellSpecArchive` bytes must be
/// retained by the package owner unless a future source-preserving patch API
/// is added; this function never silently drops source unknowns because it
/// never accepts source bytes.
pub fn rewrite_cell_spec(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_cell_spec(model_identifier, starts_with_first, options)
}

fn report_from_requirements(requirements: RewriteExecutionRequirements) -> DecodeReport {
    DecodeReport {
        input_bytes: 0,
        output_bytes: requirements.output_bytes,
        fields: requirements.fields,
        work_bytes: requirements.work_bytes,
        max_depth: requirements.max_depth,
        references: requirements.references,
        items: requirements.items,
        text_bytes: requirements.text_bytes,
        allocations: requirements.allocations,
        retained_bytes: requirements.retained_bytes,
        scratch_bytes: requirements.scratch_bytes,
    }
}

fn verify_popup_model_candidate(
    source: &[u8],
    mut options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    options.max_message_bytes = options.max_message_bytes.max(source.len());
    let (_, report) = decode_popup_menu_model_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.items != requirements.items
        || report.text_bytes != requirements.text_bytes
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn verify_cell_spec_candidate(
    source: &[u8],
    mut options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    options.max_message_bytes = options.max_message_bytes.max(source.len());
    let (_, report) = decode_cell_spec_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.references != requirements.references
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn check_options(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    if requirements.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if requirements.references > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: requirements.references,
            maximum: options.max_references,
        }));
    }
    if requirements.items > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Items {
            observed: requirements.items,
            maximum: options.max_items,
        }));
    }
    if requirements.text_bytes > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: requirements.text_bytes,
            maximum: options.max_text_bytes,
        }));
    }
    Ok(())
}

fn check_requirements(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    macro_rules! check {
        ($observed:expr, $maximum:expr, $kind:ident) => {
            if $observed > $maximum {
                return Err(DecodeError::limited(DecodeLimit::$kind {
                    observed: $observed,
                    maximum: $maximum,
                }));
            }
        };
    }
    check!(requirements.output_bytes, limits.output_bytes, OutputBytes);
    check!(requirements.fields, limits.fields, Fields);
    check!(requirements.work_bytes, limits.work_bytes, Work);
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    check!(requirements.references, limits.references, References);
    check!(requirements.items, limits.items, Items);
    check!(requirements.text_bytes, limits.text_bytes, Text);
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limited(DecodeLimit::Allocation {
            requested: requirements.allocations,
        }));
    }
    check!(requirements.retained_bytes, limits.retained_bytes, Retained);
    check!(requirements.scratch_bytes, limits.scratch_bytes, Scratch);
    Ok(())
}

fn validate_text(text: &str, maximum: usize) -> Result<(), DecodeError> {
    if text.len() > maximum || text.chars().any(char::is_control) {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: text.len(),
            maximum,
        }));
    }
    Ok(())
}

fn popup_model_output_len(items: &[&str]) -> Result<usize, DecodeError> {
    let nil_cell = 2usize;
    let nil_outer = length_field_len(POPUP_ITEM_FIELD, nil_cell)?;
    let mut length = nil_outer;
    for text in items {
        let format_payload_len = 1usize
            .checked_add(encoded_varint_len(FORMAT_TYPE_TEXT))
            .ok_or_else(DecodeError::invalid)?;
        let string_len = length_field_len(STRING_VALUE_FIELD, text.len())?
            .checked_add(length_field_len(STRING_FORMAT_FIELD, format_payload_len)?)
            .and_then(|value| value.checked_add(2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(length_field_len(CELL_VALUE_STRING_FIELD, string_len)?)
            .ok_or_else(DecodeError::invalid)?;
        length = length
            .checked_add(length_field_len(POPUP_ITEM_FIELD, cell_len)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn popup_model_work_bytes(items: &[&str]) -> Result<usize, DecodeError> {
    let mut nested = 4usize; // nil payload + one Buffa parity scan
    for text in items {
        let string_len = 1usize
            .checked_add(encoded_varint_len(
                u64::try_from(text.len()).map_err(|_| DecodeError::invalid())?,
            ))
            .and_then(|value| value.checked_add(text.len()))
            .and_then(|value| value.checked_add(5 + 2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(
                1 + encoded_varint_len(
                    u64::try_from(string_len).map_err(|_| DecodeError::invalid())?,
                ) + string_len,
            )
            .ok_or_else(DecodeError::invalid)?;
        nested = nested
            .checked_add(cell_len)
            .and_then(|value| value.checked_add(string_len))
            .and_then(|value| value.checked_add(3))
            .and_then(|value| value.checked_add(cell_len))
            .and_then(|value| value.checked_add(string_len))
            .ok_or_else(DecodeError::invalid)?;
    }
    popup_model_output_len(items)?
        .checked_add(nested)
        .ok_or_else(DecodeError::invalid)
}

fn cell_spec_output_len(model_identifier: u64) -> Result<usize, DecodeError> {
    let reference_len = 1usize
        .checked_add(encoded_varint_len(model_identifier))
        .ok_or_else(DecodeError::invalid)?;
    2usize
        .checked_add(length_field_len(CELL_SPEC_MODEL_FIELD, reference_len)?)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(DecodeError::invalid)
}

fn length_field_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let key = encoded_varint_len(u64::from(number) << 3);
    let payload =
        encoded_varint_len(u64::try_from(payload_len).map_err(|_| DecodeError::invalid())?);
    key.checked_add(payload)
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(DecodeError::invalid)
}

fn emit_popup_model(output: &mut Vec<u8>, items: &[&str]) -> Result<(), DecodeError> {
    emit_len_field(output, POPUP_ITEM_FIELD, &[0x08, 0x01])?;
    for text in items {
        let string_len = 1usize
            .checked_add(encoded_varint_len(
                u64::try_from(text.len()).map_err(|_| DecodeError::invalid())?,
            ))
            .and_then(|value| value.checked_add(text.len()))
            .and_then(|value| value.checked_add(5 + 2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(
                1 + encoded_varint_len(
                    u64::try_from(string_len).map_err(|_| DecodeError::invalid())?,
                ) + string_len,
            )
            .ok_or_else(DecodeError::invalid)?;
        emit_len_field_header(output, POPUP_ITEM_FIELD, cell_len)?;
        emit_varint_field(output, CELL_VALUE_TYPE_FIELD, STRING_TYPE)?;
        emit_len_field_header(output, CELL_VALUE_STRING_FIELD, string_len)?;
        emit_len_field(output, STRING_VALUE_FIELD, text.as_bytes())?;
        emit_len_field(output, STRING_FORMAT_FIELD, &[0x08, 0x84, 0x02])?;
        emit_varint_field(output, STRING_EXPLICIT_FIELD, 0)?;
        emit_varint_field(output, STRING_REGEX_FIELD, 0)?;
        emit_varint_field(output, STRING_CASE_SENSITIVE_REGEX_FIELD, 0)?;
    }
    Ok(())
}

fn emit_cell_spec(
    output: &mut Vec<u8>,
    model_identifier: u64,
    starts_with_first: bool,
) -> Result<(), DecodeError> {
    emit_varint_field(
        output,
        CELL_SPEC_INTERACTION_FIELD,
        u64::from(POPUP_INTERACTION_TYPE),
    )?;
    let reference_len = 1usize
        .checked_add(encoded_varint_len(model_identifier))
        .ok_or_else(DecodeError::invalid)?;
    emit_len_field_header(output, CELL_SPEC_MODEL_FIELD, reference_len)?;
    emit_varint_field(output, REFERENCE_IDENTIFIER_FIELD, model_identifier)?;
    emit_varint_field(output, CELL_SPEC_FIRST_FIELD, u64::from(starts_with_first))?;
    Ok(())
}

fn emit_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    write_varint(output, u64::from(number) << 3)?;
    write_varint(output, value)
}

fn emit_len_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), DecodeError> {
    emit_len_field_header(output, number, payload.len())?;
    output.extend_from_slice(payload);
    Ok(())
}

fn emit_len_field_header(
    output: &mut Vec<u8>,
    number: u32,
    payload_len: usize,
) -> Result<(), DecodeError> {
    write_varint(output, u64::from(number) << 3 | 2)?;
    write_varint(
        output,
        u64::try_from(payload_len).map_err(|_| DecodeError::invalid())?,
    )
}

fn write_varint(output: &mut Vec<u8>, mut value: u64) -> Result<(), DecodeError> {
    while value >= 0x80 {
        output.push(value as u8 | 0x80);
        value >>= 7;
    }
    output.push(u8::try_from(value).map_err(|_| DecodeError::invalid())?);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 16 * 1024, 64 * 1024, 64, 16, 64, 4096)
    }

    #[test]
    fn canonical_model_roundtrips_nil_and_strings() {
        let items = ["Low", "High"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let (snapshot, report) =
            decode_popup_menu_model_with_report(output.bytes(), options()).expect("decode model");
        assert!(snapshot.has_nil_sentinel());
        assert_eq!(snapshot.item_count(), 2);
        assert_eq!(
            snapshot
                .items()
                .map(PopUpMenuItem::value)
                .collect::<Vec<_>>(),
            items
        );
        assert_eq!(report.input_bytes(), output.bytes().len());
        assert_eq!(report.text_bytes(), 7);
    }

    #[test]
    fn canonical_cell_spec_roundtrips_reference_and_selection() {
        let output = canonical_cell_spec(41, true, options()).expect("canonical cell spec");
        let (snapshot, report) =
            decode_cell_spec_with_report(output.bytes(), options()).expect("decode cell spec");
        assert_eq!(snapshot.interaction_type(), POPUP_INTERACTION_TYPE);
        assert_eq!(snapshot.popup_model().identifier(), 41);
        assert!(snapshot.starts_with_first());
        assert_eq!(report.references(), 1);
    }

    #[test]
    fn cell_spec_unknown_fields_are_raw_preserved_without_buffa_rejection() {
        let output = canonical_cell_spec(41, true, options()).expect("canonical cell spec");
        let mut source = output.bytes().to_vec();
        source.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
        let snapshot = decode_cell_spec(&source, options()).expect("opaque unknown field");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.popup_model().identifier(), 41);
    }

    #[test]
    fn strict_model_rejects_deprecated_item_and_missing_sentinel() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut deprecated = output.bytes().to_vec();
        deprecated.extend_from_slice(&[0x0a, 0x00]);
        assert!(decode_popup_menu_model(&deprecated, options()).is_err());

        let mut missing = output.bytes().to_vec();
        // The first field is the two-byte nil payload wrapped in field 2.
        assert_eq!(&missing[..4], &[0x12, 0x02, 0x08, 0x01]);
        missing.drain(..4);
        assert!(decode_popup_menu_model(&missing, options()).is_err());
    }

    #[test]
    fn strict_cell_spec_rejects_zero_or_duplicate_reference() {
        let mut zero = vec![0x08, 0x07, 0x32, 0x02, 0x08, 0x00, 0x38, 0x01];
        assert!(decode_cell_spec(&zero, options()).is_err());
        zero[5] = 0x01;
        zero.extend_from_slice(&[0x32, 0x02, 0x08, 0x02]);
        assert!(decode_cell_spec(&zero, options()).is_err());
    }

    #[test]
    fn strict_cell_spec_rejects_deprecated_reference_presence_even_at_defaults() {
        let output = canonical_cell_spec(41, false, options()).expect("canonical cell spec");

        let mut deprecated_type = output.bytes().to_vec();
        deprecated_type[3] += 2;
        deprecated_type.splice(6..6, [REFERENCE_TYPE_FIELD as u8 * 2, 0]);
        assert!(decode_cell_spec(&deprecated_type, options()).is_err());

        let mut deprecated_external = output.bytes().to_vec();
        deprecated_external[3] += 2;
        deprecated_external.splice(6..6, [REFERENCE_EXTERNAL_FIELD as u8 * 2, 0]);
        assert!(decode_cell_spec(&deprecated_external, options()).is_err());
    }

    #[test]
    fn unknown_group_is_structurally_checked_and_raw_retained() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut source = output.bytes().to_vec();
        // Unknown field 100, balanced start/end group with an unknown scalar.
        source.extend_from_slice(&[
            0xa0, 0x06, 0x81, 0x00, // unknown scalar value 1, deliberately overlong
            0xa3, 0x06, 0x08, 0x81, 0x00, 0xa4, 0x06,
        ]);
        let snapshot = decode_popup_menu_model(&source, options()).expect("unknown group");
        assert_eq!(snapshot.raw(), source.as_slice());

        let mut unterminated = source;
        unterminated.pop();
        assert!(decode_popup_menu_model(&unterminated, options()).is_err());
    }

    #[test]
    fn unknown_group_fields_work_and_nesting_are_budgeted() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut source = output.bytes().to_vec();
        source.extend_from_slice(&[
            0xa3, 0x06, // unknown group 100
            0xa3, 0x06, 0x08, 0x81, 0x00, 0xa4, 0x06, // nested group + scalar
            0xa4, 0x06,
        ]);
        let (_, report) =
            decode_popup_menu_model_with_report(&source, options()).expect("metered unknown group");
        assert!(report.fields() > 0);
        assert!(report.work_bytes() >= source.len());
        assert!(
            decode_popup_menu_model(&source, options().with_max_output_bytes(usize::MAX),).is_ok()
        );
        let field_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            report.fields() - 1,
            options().max_work_bytes,
            options().recursion_limit,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, field_limited)
                .expect_err("nested field ceiling")
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            options().max_fields,
            report.work_bytes() - 1,
            options().recursion_limit,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, work_limited)
                .expect_err("nested work ceiling")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let nesting_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            options().max_fields,
            options().max_work_bytes,
            1,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, nesting_limited)
                .expect_err("nested depth ceiling")
                .resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn visitor_streams_items_without_a_second_reported_parse() {
        struct Visitor {
            values: Vec<String>,
        }

        impl PopUpMenuVisitor for Visitor {
            fn visit_item(&mut self, item: PopUpMenuItem<'_>) -> Result<(), DecodeError> {
                self.values.push(item.value().to_owned());
                Ok(())
            }
        }

        let output =
            canonical_popup_menu_model(&["Low", "High"], options()).expect("canonical model");
        let mut visitor = Visitor { values: Vec::new() };
        let visitor_report =
            decode_popup_menu_model_with_visitor(output.bytes(), options(), &mut visitor)
                .expect("stream visitor");
        let (_, report) = decode_popup_menu_model_with_report(output.bytes(), options())
            .expect("ordinary report");
        assert_eq!(visitor.values, ["Low", "High"]);
        assert_eq!(visitor_report, report);
    }

    #[test]
    fn prepared_model_exact_replay_and_minus_one_axes_fail_before_output() {
        let items = ["A", "B"];
        let prepared = prepare_popup_menu_model_write(&items, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let exact = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact execute");
        assert_eq!(exact.bytes().len(), requirements.output_bytes());
        assert_eq!(exact.report().fields(), requirements.fields());
        assert_eq!(exact.report().work_bytes(), requirements.work_bytes());
        let (_, decoded_report) = decode_popup_menu_model_with_report(exact.bytes(), options())
            .expect("strict candidate verification");
        assert_eq!(decoded_report.fields(), requirements.fields());
        assert_eq!(
            decoded_report.work_bytes(),
            requirements.work_bytes() - requirements.output_bytes()
        );
        assert_eq!(decoded_report.max_depth(), requirements.max_depth());
        assert_eq!(decoded_report.items(), requirements.items());
        assert_eq!(decoded_report.text_bytes(), requirements.text_bytes());
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .expect_err("output ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::OutputBytes { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1)
                )
                .expect_err("field ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Fields { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .expect_err("work ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Work { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_max_depth(requirements.max_depth() - 1)
                )
                .expect_err("depth ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Nesting { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_items(requirements.items() - 1)
                )
                .expect_err("item ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Items { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_text_bytes(requirements.text_bytes() - 1)
                )
                .expect_err("text ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Text { .. }))
        );
        assert!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_allocations(0))
                .expect_err("allocation ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Allocation { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_retained_bytes(requirements.retained_bytes() - 1)
                )
                .expect_err("retained ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Retained { .. }))
        );
    }

    #[test]
    fn prepared_cell_spec_exact_replay_and_reference_limit() {
        let prepared = prepare_cell_spec_write(91, false, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        let snapshot = decode_cell_spec(output.bytes(), options()).expect("decode");
        assert_eq!(snapshot.popup_model().identifier(), 91);
        assert!(!snapshot.starts_with_first());
        let (_, decoded_report) = decode_cell_spec_with_report(output.bytes(), options())
            .expect("strict candidate verification");
        assert_eq!(decoded_report.fields(), requirements.fields());
        assert_eq!(
            decoded_report.work_bytes(),
            requirements.work_bytes() - requirements.output_bytes()
        );
        assert_eq!(decoded_report.max_depth(), requirements.max_depth());
        assert_eq!(decoded_report.references(), requirements.references());
        assert!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_references(0))
                .expect_err("reference ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::References { .. }))
        );
    }
}
