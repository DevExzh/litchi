//! Bounded protobuf wire mutations that retain untouched fields byte-for-byte.

#![allow(
    clippy::missing_errors_doc,
    reason = "All fallible wire operations return the shared structured Error; callers can match its variants"
)]

use std::cell::Cell;
use std::collections::HashMap;

use crate::{Error, LimitKind, Result, WireLimits};

/// A bounded, source-bound view of one protobuf-style IWA wire message.
///
/// The view borrows the input once and retains only compact private spans for
/// its fields. Field views are consequently tied to this source and cannot
/// accidentally be used with a different byte slice after parsing.
#[allow(
    clippy::module_name_repetitions,
    reason = "WireView is the explicit borrowed container name used at format boundaries"
)]
#[derive(Debug)]
pub struct WireView<'a> {
    source: &'a [u8],
    spans: Vec<WireSpan>,
}

impl<'a> WireView<'a> {
    /// Parse one borrowed wire message under the default resource profile.
    pub fn parse(source: &'a [u8]) -> Result<Self> {
        Self::parse_with_limits(source, WireLimits::default())
    }

    /// Parse one borrowed wire message under an explicit finite resource
    /// profile.
    pub fn parse_with_limits(source: &'a [u8], limits: WireLimits) -> Result<Self> {
        let spans = parse_wire_items_with_limits(source, limits, "wire view spans", |field| {
            span_from_field(field)
        })?;
        Ok(Self { source, spans })
    }

    /// Return the one source byte slice retained by this view.
    #[must_use]
    pub fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Return the source bytes without copying them.
    #[must_use]
    pub fn as_bytes(&self) -> &'a [u8] {
        self.source
    }

    /// Number of parsed fields.
    #[must_use]
    pub fn len(&self) -> usize {
        self.spans.len()
    }

    /// Whether the source contains no fields.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.spans.is_empty()
    }

    /// Borrow a field by its zero-based source order.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<WireFieldView<'a>> {
        self.spans.get(index).copied().map(|span| WireFieldView {
            source: self.source,
            span,
        })
    }

    /// Borrow all fields in their original source order.
    pub fn fields(&self) -> impl Iterator<Item = WireFieldView<'a>> + '_ {
        self.spans.iter().copied().map(|span| WireFieldView {
            source: self.source,
            span,
        })
    }

    /// Alias for [`Self::fields`] for callers that prefer iterator terminology.
    pub fn iter(&self) -> impl Iterator<Item = WireFieldView<'a>> + '_ {
        self.fields()
    }
}

/// A borrowed view of one field in a [`WireView`].
///
/// This value contains a compact span and a borrow of the parsed view's
/// source, not an independently supplied byte slice. It is cheap to copy and
/// remains valid only while that source is valid.
#[allow(
    clippy::module_name_repetitions,
    reason = "WireFieldView is the explicit borrowed counterpart to WireField"
)]
#[derive(Debug, Clone, Copy)]
pub struct WireFieldView<'a> {
    source: &'a [u8],
    span: WireSpan,
}

impl<'a> WireFieldView<'a> {
    /// Protobuf field number.
    #[must_use]
    pub const fn number(self) -> u32 {
        self.span.number
    }

    /// Protobuf wire type tag.
    #[must_use]
    pub const fn wire_type(self) -> u8 {
        self.span.wire_type
    }

    /// Return the complete encoded field, including key and length prefix.
    #[must_use]
    pub fn raw(self) -> &'a [u8] {
        &self.source[self.span.start as usize..self.span.end as usize]
    }

    /// Return the encoded field key.
    #[must_use]
    pub fn key(self) -> &'a [u8] {
        &self.source[self.span.start as usize..self.span.key_end as usize]
    }

    /// Return the field payload, excluding a length prefix when present.
    #[must_use]
    pub fn payload(self) -> &'a [u8] {
        &self.source[self.span.payload_start as usize..self.span.end as usize]
    }

    /// Require the field key to use its canonical varint representation.
    pub fn validate_canonical_key(self) -> Result<()> {
        validate_canonical_key_bytes(self.number(), self.wire_type(), self.key())
    }

    /// Return the field key after requiring canonical framing.
    pub fn canonical_key(self) -> Result<&'a [u8]> {
        self.validate_canonical_key()?;
        Ok(self.key())
    }

    /// Require a length-delimited field's length prefix to be canonical.
    pub fn validate_canonical_length(self) -> Result<()> {
        if self.wire_type() != 2 {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} is not length-delimited",
                self.number()
            )));
        }
        validate_canonical_length_bytes(self.number(), self.payload(), self.length_prefix())
    }

    /// Require canonical key framing and, when applicable, a canonical length
    /// prefix.
    pub fn validate_canonical_framing(self) -> Result<()> {
        self.validate_canonical_key()?;
        if self.wire_type() == 2 {
            self.validate_canonical_length()?;
        }
        Ok(())
    }

    /// Return the payload after requiring canonical field framing.
    pub fn canonical_payload(self) -> Result<&'a [u8]> {
        self.validate_canonical_framing()?;
        Ok(self.payload())
    }

    /// Return the encoded length prefix of a length-delimited field.
    fn length_prefix(self) -> &'a [u8] {
        &self.source[self.span.key_end as usize..self.span.payload_start as usize]
    }
}

/// Decision returned by a bounded wire-tree preflight visitor.
///
/// Length-delimited protobuf fields are ambiguous at the wire level: their
/// payload may be a message, a string, bytes, or packed scalars. The visitor
/// supplies the missing schema knowledge by selecting only payloads that are
/// deferred submessages in the caller's generated model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "WireDescent remains explicit when imported by schema adapters"
)]
pub enum WireDescent {
    /// Treat this field as an opaque scalar payload.
    Skip,
    /// Validate this length-delimited payload as a nested wire message.
    Descend,
}

/// One parser-produced field presented during a bounded wire-tree preflight.
///
/// The containing-message path contains the field numbers followed from the
/// root to reach this field's message. It is empty for root fields. The path
/// is borrowed only for the visitor call. The field may be retained, but it
/// remains bound to the exact source slice parsed by the scanner and cannot be
/// paired with caller-supplied bytes.
#[derive(Debug, Clone, Copy)]
#[allow(
    clippy::module_name_repetitions,
    reason = "WireVisit names the wire-specific callback value at schema boundaries"
)]
pub struct WireVisit<'source, 'path> {
    path: &'path [u32],
    field: WireFieldView<'source>,
}

impl<'source, 'path> WireVisit<'source, 'path> {
    /// Field-number path of the message containing this field.
    #[must_use]
    pub const fn path(self) -> &'path [u32] {
        self.path
    }

    /// Nested-message depth of the message containing this field.
    #[must_use]
    pub const fn depth(self) -> usize {
        self.path.len()
    }

    /// Source-bound field currently being visited.
    #[must_use]
    pub const fn field(self) -> WireFieldView<'source> {
        self.field
    }
}

/// Resource totals from a completed bounded wire-tree preflight.
///
/// `scanned_bytes` counts the bytes inspected in every selected message. A
/// nested payload is therefore charged again when the visitor asks to scan it,
/// matching the aggregate work performed by subsequent deferred decodes
/// rather than merely checking the size of the root allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "WirePreflight distinguishes this report from format-level package preflights"
)]
pub struct WirePreflight {
    scanned_bytes: usize,
    fields: usize,
    messages: usize,
    max_depth: usize,
}

impl WirePreflight {
    /// Aggregate message bytes inspected, including selected descendants.
    #[must_use]
    pub const fn scanned_bytes(self) -> usize {
        self.scanned_bytes
    }

    /// Aggregate fields parsed across the root and selected descendants.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Number of messages scanned, including the root message.
    #[must_use]
    pub const fn messages(self) -> usize {
        self.messages
    }

    /// Deepest selected message, where the root message has depth zero.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
}

struct WirePreflightContext {
    limits: WireLimits,
    scanned_bytes: usize,
    fields: usize,
    messages: usize,
    max_depth: usize,
}

impl WirePreflightContext {
    const fn new(limits: WireLimits) -> Self {
        Self {
            limits,
            scanned_bytes: 0,
            fields: 0,
            messages: 0,
            max_depth: 0,
        }
    }

    fn finish(self) -> WirePreflight {
        WirePreflight {
            scanned_bytes: self.scanned_bytes,
            fields: self.fields,
            messages: self.messages,
            max_depth: self.max_depth,
        }
    }

    fn charge_message(&mut self, bytes: usize, depth: usize) -> Result<()> {
        if depth > self.limits.max_nesting() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: depth,
                limit: self.limits.max_nesting(),
            });
        }
        let observed = self.scanned_bytes.saturating_add(bytes);
        if observed > self.limits.max_input_bytes() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit: self.limits.max_input_bytes(),
            });
        }
        self.scanned_bytes = observed;
        self.messages = self.messages.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("protobuf preflight message count overflow".to_owned())
        })?;
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    fn charge_field(&mut self) -> Result<()> {
        let observed = self.fields.saturating_add(1);
        if observed > self.limits.max_fields() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: self.limits.max_fields(),
            });
        }
        self.fields = observed;
        Ok(())
    }
}

/// Compact private field metadata retained by [`WireView`].
///
/// The configured input ceiling is below `u32::MAX`, so four-byte offsets are
/// sufficient while keeping the parsed representation free of per-field byte
/// references.
#[derive(Debug, Clone, Copy)]
struct WireSpan {
    number: u32,
    wire_type: u8,
    start: u32,
    key_end: u32,
    payload_start: u32,
    end: u32,
}

#[allow(
    clippy::module_name_repetitions,
    reason = "WireField is the stable, explicit name of a parsed protobuf field"
)]
#[derive(Debug, Clone, Copy)]
pub struct WireField {
    number: u32,
    wire_type: u8,
    start: usize,
    key_end: usize,
    payload_start: usize,
    end: usize,
}

impl WireField {
    /// Protobuf field number.
    #[must_use]
    pub const fn number(self) -> u32 {
        self.number
    }

    /// Protobuf wire type tag.
    #[must_use]
    pub const fn wire_type(self) -> u8 {
        self.wire_type
    }

    /// Byte offset at which the encoded field key begins.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Byte offset immediately after the encoded field key.
    #[must_use]
    pub const fn key_end(self) -> usize {
        self.key_end
    }

    /// Byte offset at which a length-delimited payload begins.
    #[must_use]
    pub const fn payload_start(self) -> usize {
        self.payload_start
    }

    /// Byte offset immediately after the encoded field.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// Return the complete encoded field through checked slice boundaries.
    ///
    /// The field does not retain a source lifetime, so the caller must supply
    /// the exact bytes from which it was parsed. A mismatched or truncated
    /// source is rejected instead of being sliced unchecked.
    pub fn raw(self, data: &[u8]) -> Result<&[u8]> {
        checked_wire_range(data, self.start, self.end, "field")
    }

    /// Return the encoded field key through checked slice boundaries.
    pub fn key(self, data: &[u8]) -> Result<&[u8]> {
        checked_wire_range(data, self.start, self.key_end, "field key")
    }

    /// Return this field's payload through checked slice boundaries.
    ///
    /// For a length-delimited field, the returned slice excludes the encoded
    /// length prefix. For the other supported wire types, it contains the
    /// complete encoded scalar payload. This method validates the offsets
    /// against `data` but intentionally does not require canonical varints;
    /// callers handling opaque fields can therefore continue to preserve
    /// their original representation.
    ///
    /// # Errors
    ///
    /// Returns an error when this field's recorded offsets are inconsistent
    /// or do not fit within `data`.
    pub fn checked_payload(self, data: &[u8]) -> Result<&[u8]> {
        if self.start > self.key_end
            || self.key_end > self.payload_start
            || self.payload_start > self.end
        {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} has invalid byte offsets",
                self.number
            )));
        }
        data.get(self.payload_start..self.end).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "protobuf field {} extends beyond the supplied data",
                self.number
            ))
        })
    }

    /// Return the field payload through checked slice boundaries.
    pub fn payload(self, data: &[u8]) -> Result<&[u8]> {
        self.checked_payload(data)
    }

    /// Require the encoded field key to use its canonical varint form.
    ///
    /// The parser remains permissive so unknown fields can be retained
    /// byte-for-byte. Use this method when a schema-recognized field is about
    /// to be interpreted or rewritten.
    ///
    /// # Errors
    ///
    /// Returns an error when the field offsets are invalid, the key is
    /// malformed, or the key uses an overlong varint representation.
    pub fn validate_canonical_key(self, data: &[u8]) -> Result<()> {
        let key = checked_wire_range(data, self.start, self.key_end, "field key")?;
        let (encoded_key, key_length) =
            crate::varint::decode_varint_from_bytes(key).map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid protobuf key for field {}: {error}",
                    self.number
                ))
            })?;
        let expected_key = (u64::from(self.number) << 3) | u64::from(self.wire_type);
        if encoded_key != expected_key
            || key_length != key.len()
            || key_length != crate::varint::encoded_len(expected_key)
        {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} has a noncanonical key",
                self.number
            )));
        }
        Ok(())
    }

    /// Require a length-delimited field's length prefix to be canonical.
    ///
    /// This validates both the prefix encoding and its correspondence with
    /// the checked payload range. It does not validate the field key; callers
    /// that need both invariants should call [`Self::validate_canonical_key`]
    /// as well.
    ///
    /// # Errors
    ///
    /// Returns an error when the field is not length-delimited, its offsets
    /// are invalid, or its length prefix is malformed, overlong, or does not
    /// describe the payload range.
    pub fn validate_canonical_length(self, data: &[u8]) -> Result<()> {
        if self.wire_type != 2 {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} is not length-delimited",
                self.number
            )));
        }
        let payload = self.checked_payload(data)?;
        let length_prefix =
            checked_wire_range(data, self.key_end, self.payload_start, "length prefix")?;
        let (encoded_length, prefix_length) =
            crate::varint::decode_varint_from_bytes(length_prefix).map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid protobuf length for field {}: {error}",
                    self.number
                ))
            })?;
        let expected_length = u64::try_from(payload.len()).map_err(|_conversion| {
            Error::InvalidFormat("protobuf payload length exceeds u64".to_owned())
        })?;
        if encoded_length != expected_length
            || prefix_length != length_prefix.len()
            || prefix_length != crate::varint::encoded_len(expected_length)
        {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} has a noncanonical length prefix",
                self.number
            )));
        }
        Ok(())
    }

    /// Require the complete key and length framing for this field to be
    /// canonical.
    ///
    /// The parser remains permissive so unknown fields can be retained
    /// byte-for-byte. Callers can use this helper when a schema-recognized
    /// field is about to be interpreted or rewritten.
    pub fn validate_canonical_framing(self, data: &[u8]) -> Result<()> {
        self.validate_canonical_key(data)?;
        if self.wire_type == 2 {
            self.validate_canonical_length(data)?;
        }
        Ok(())
    }
}

/// Typed replacement for one leaf in a batched nested-field rewrite.
///
/// `None` clears a present leaf, or validates that an absent leaf remains
/// absent. `Some` replaces a present leaf or appends an absent leaf. Keeping
/// the wire type outside the option means clearing still verifies the exact
/// schema-selected wire type before removing bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NestedFieldReplacement<'a> {
    /// Optional protobuf varint.
    Varint(Option<u64>),
    /// Optional little-endian protobuf fixed32.
    Fixed32(Option<u32>),
    /// Optional little-endian protobuf fixed64.
    Fixed64(Option<u64>),
    /// Optional protobuf length-delimited payload.
    LengthDelimited(Option<&'a [u8]>),
}

impl NestedFieldReplacement<'_> {
    const fn wire_type(self) -> u8 {
        match self {
            Self::Varint(_) => 0,
            Self::Fixed64(_) => 1,
            Self::LengthDelimited(_) => 2,
            Self::Fixed32(_) => 5,
        }
    }

    #[allow(
        clippy::match_same_arms,
        reason = "The variants carry three different Option scalar types plus one borrowed slice"
    )]
    const fn is_some(self) -> bool {
        match self {
            Self::Varint(value) => value.is_some(),
            Self::Fixed32(value) => value.is_some(),
            Self::Fixed64(value) => value.is_some(),
            Self::LengthDelimited(value) => value.is_some(),
        }
    }
}

/// One schema-selected leaf edit for [`patch_nested_fields_batched_with_limits`].
///
/// `path` is a nonempty sequence of canonical protobuf field numbers. Every
/// non-leaf component must identify exactly one canonical length-delimited
/// field. The leaf must occur exactly once when `expected_present` is true and
/// not at all when it is false. Paths may share ancestors, but duplicate paths
/// and leaf/ancestor path collisions are rejected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NestedFieldEdit<'a> {
    path: &'a [u32],
    expected_present: bool,
    replacement: NestedFieldReplacement<'a>,
}

impl<'a> NestedFieldEdit<'a> {
    /// Construct one batched nested-field edit.
    #[must_use]
    pub const fn new(
        path: &'a [u32],
        expected_present: bool,
        replacement: NestedFieldReplacement<'a>,
    ) -> Self {
        Self {
            path,
            expected_present,
            replacement,
        }
    }

    /// Canonical field-number path from the root message to the leaf.
    #[must_use]
    pub const fn path(self) -> &'a [u32] {
        self.path
    }

    /// Whether the selected leaf must occur exactly once.
    #[must_use]
    pub const fn expected_present(self) -> bool {
        self.expected_present
    }

    /// Typed optional replacement for the selected leaf.
    #[must_use]
    pub const fn replacement(self) -> NestedFieldReplacement<'a> {
        self.replacement
    }
}

#[derive(Debug)]
struct BatchTrieNode {
    children: Vec<(u32, usize)>,
    edit: Option<usize>,
}

impl BatchTrieNode {
    const fn new() -> Self {
        Self {
            children: Vec::new(),
            edit: None,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum BatchFieldAction {
    Copy,
    Leaf(usize),
    Nested(usize),
}

#[derive(Debug, Clone, Copy)]
struct BatchPlannedField {
    start: usize,
    key_end: usize,
    end: usize,
    action: BatchFieldAction,
}

#[derive(Debug)]
struct BatchMessagePlan {
    fields: Vec<BatchPlannedField>,
    appended: Vec<usize>,
    output_len: usize,
    changed: bool,
}

#[derive(Debug)]
struct BatchPlanContext {
    limits: WireLimits,
    source_fields: usize,
    removed_fields: usize,
    appended_fields: usize,
    plans: Vec<BatchMessagePlan>,
}

impl BatchPlanContext {
    const fn new(limits: WireLimits) -> Self {
        Self {
            limits,
            source_fields: 0,
            removed_fields: 0,
            appended_fields: 0,
            plans: Vec::new(),
        }
    }

    fn charge_field(&mut self) -> Result<()> {
        self.source_fields = self
            .source_fields
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
        ensure_field_count(self.source_fields, self.limits)
    }

    fn push_plan(&mut self, plan: BatchMessagePlan) -> Result<usize> {
        let index = self.plans.len();
        self.plans
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource: "batched wire message plans",
                amount: index + 1,
            })?;
        self.plans.push(plan);
        Ok(index)
    }
}

fn checked_wire_range<'a>(
    data: &'a [u8],
    start: usize,
    end: usize,
    name: &str,
) -> Result<&'a [u8]> {
    data.get(start..end).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "protobuf {name} range {start}..{end} is outside the supplied data"
        ))
    })
}

fn validate_canonical_key_bytes(number: u32, wire_type: u8, key: &[u8]) -> Result<()> {
    let (encoded_key, key_length) =
        crate::varint::decode_varint_from_bytes(key).map_err(|error| {
            Error::InvalidFormat(format!("invalid protobuf key for field {number}: {error}"))
        })?;
    let expected_key = (u64::from(number) << 3) | u64::from(wire_type);
    if encoded_key != expected_key
        || key_length != key.len()
        || key_length != crate::varint::encoded_len(expected_key)
    {
        return Err(Error::InvalidFormat(format!(
            "protobuf field {number} has a noncanonical key"
        )));
    }
    Ok(())
}

fn validate_canonical_length_bytes(
    number: u32,
    payload: &[u8],
    length_prefix: &[u8],
) -> Result<()> {
    let (encoded_length, prefix_length) = crate::varint::decode_varint_from_bytes(length_prefix)
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid protobuf length for field {number}: {error}"
            ))
        })?;
    let expected_length = u64::try_from(payload.len()).map_err(|_conversion| {
        Error::InvalidFormat("protobuf payload length exceeds u64".to_owned())
    })?;
    if encoded_length != expected_length
        || prefix_length != length_prefix.len()
        || prefix_length != crate::varint::encoded_len(expected_length)
    {
        return Err(Error::InvalidFormat(format!(
            "protobuf field {number} has a noncanonical length prefix"
        )));
    }
    Ok(())
}

pub fn parse_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    parse_wire_fields_with_limits(data, WireLimits::default())
}

/// Parse protobuf wire fields under an explicit finite resource profile.
pub fn parse_wire_fields_with_limits(data: &[u8], limits: WireLimits) -> Result<Vec<WireField>> {
    if data.len() > limits.max_input_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: data.len(),
            limit: limits.max_input_bytes(),
        });
    }
    let mut items = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let field = parse_wire_field(data, offset)?;
        offset = field.end;
        if items.len() >= limits.max_fields() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: items.len() + 1,
                limit: limits.max_fields(),
            });
        }
        items
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource: "wire fields",
                amount: items.len() + 1,
            })?;
        items.push(field);
    }
    Ok(items)
}

fn parse_wire_field(data: &[u8], start: usize) -> Result<WireField> {
    let (key, key_length) = crate::varint::decode_varint_from_bytes(&data[start..])
        .map_err(|error| Error::InvalidFormat(format!("invalid protobuf key: {error}")))?;
    let key_end = start
        .checked_add(key_length)
        .ok_or_else(|| Error::InvalidFormat("protobuf key offset overflow".to_owned()))?;
    let number = u32::try_from(key >> 3).map_err(|error| {
        Error::InvalidFormat(format!("protobuf field number does not fit u32: {error}"))
    })?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(Error::InvalidFormat(format!(
            "invalid protobuf field number {number}"
        )));
    }
    let wire_type = u8::try_from(key & 7).map_err(|error| {
        Error::InvalidFormat(format!("protobuf wire type does not fit u8: {error}"))
    })?;
    let mut payload_start = key_end;
    let end = match wire_type {
        0 => {
            let (_, length) =
                crate::varint::decode_varint_from_bytes(&data[key_end..]).map_err(|error| {
                    Error::InvalidFormat(format!("invalid protobuf varint value: {error}"))
                })?;
            key_end
                .checked_add(length)
                .ok_or_else(|| Error::InvalidFormat("protobuf varint offset overflow".to_owned()))?
        },
        1 => key_end
            .checked_add(8)
            .ok_or_else(|| Error::InvalidFormat("protobuf fixed64 offset overflow".to_owned()))?,
        2 => {
            let (encoded_length, prefix_length) =
                crate::varint::decode_varint_from_bytes(&data[key_end..]).map_err(|error| {
                    Error::InvalidFormat(format!("invalid protobuf length: {error}"))
                })?;
            let payload = key_end.checked_add(prefix_length).ok_or_else(|| {
                Error::InvalidFormat("protobuf length-prefix overflow".to_owned())
            })?;
            payload_start = payload;
            let length = usize::try_from(encoded_length).map_err(|error| {
                Error::InvalidFormat(format!("protobuf field length exceeds usize: {error}"))
            })?;
            payload
                .checked_add(length)
                .ok_or_else(|| Error::InvalidFormat("protobuf field range overflow".to_owned()))?
        },
        5 => key_end
            .checked_add(4)
            .ok_or_else(|| Error::InvalidFormat("protobuf fixed32 offset overflow".to_owned()))?,
        3 | 4 => {
            return Err(Error::InvalidFormat(
                "deprecated protobuf groups are not supported".to_owned(),
            ));
        },
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid protobuf wire type {wire_type}"
            )));
        },
    };
    if end > data.len() {
        return Err(Error::InvalidFormat("truncated protobuf field".to_owned()));
    }
    Ok(WireField {
        number,
        wire_type,
        start,
        key_end,
        payload_start,
        end,
    })
}

fn parse_wire_items_with_limits<T, F>(
    data: &[u8],
    limits: WireLimits,
    resource: &'static str,
    mut map: F,
) -> Result<Vec<T>>
where
    F: FnMut(WireField) -> Result<T>,
{
    if data.len() > limits.max_input_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: data.len(),
            limit: limits.max_input_bytes(),
        });
    }
    let mut items = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let field = parse_wire_field(data, offset)?;
        offset = field.end;
        if items.len() >= limits.max_fields() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: items.len() + 1,
                limit: limits.max_fields(),
            });
        }
        items
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource,
                amount: items.len() + 1,
            })?;
        items.push(map(field)?);
    }
    Ok(items)
}

fn span_from_field(field: WireField) -> Result<WireSpan> {
    Ok(WireSpan {
        number: field.number,
        wire_type: field.wire_type,
        start: u32::try_from(field.start).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        key_end: u32::try_from(field.key_end).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        payload_start: u32::try_from(field.payload_start).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        end: u32::try_from(field.end).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
    })
}

/// Parse a borrowed source into a source-bound wire view.
pub fn parse_wire_view(source: &[u8]) -> Result<WireView<'_>> {
    WireView::parse(source)
}

/// Parse a borrowed source into a source-bound wire view under finite limits.
pub fn parse_wire_view_with_limits(source: &[u8], limits: WireLimits) -> Result<WireView<'_>> {
    WireView::parse_with_limits(source, limits)
}

/// Preflight a root message and every schema-selected deferred submessage
/// under the default finite resource profile.
///
/// The visitor runs once for each structurally valid field in depth-first wire
/// order. Return [`WireDescent::Descend`] only for fields that the schema
/// declares to be messages. The containing-message [`WireVisit::path`] lets a
/// caller distinguish the same field number in different message types.
///
/// All selected messages share one aggregate byte, field, and nesting budget.
/// This closes the budget reset that would result from independently parsing
/// each deferred payload after a lazy top-level decode.
pub fn preflight_wire_tree<'source, F>(source: &'source [u8], visitor: F) -> Result<WirePreflight>
where
    F: for<'path> FnMut(WireVisit<'source, 'path>) -> Result<WireDescent>,
{
    preflight_wire_tree_with_limits(source, WireLimits::default(), visitor)
}

/// Preflight a root message and every schema-selected deferred submessage
/// under one explicit finite resource profile.
///
/// The scanner does not expose offsets or accept caller-created fields. Each
/// [`WireVisit`] is built directly from the message slice currently being
/// parsed, and descent is possible only during that visit. Nested byte ranges
/// and depths therefore cannot be forged or detached from their source.
///
/// Input bytes are charged for every scanned message, including overlapping
/// ancestor and descendant slices. Field count is likewise aggregate across
/// the whole selected tree. The path stack is grown with fallible allocation
/// and never exceeds [`WireLimits::max_nesting`].
pub fn preflight_wire_tree_with_limits<'source, F>(
    source: &'source [u8],
    limits: WireLimits,
    mut visitor: F,
) -> Result<WirePreflight>
where
    F: for<'path> FnMut(WireVisit<'source, 'path>) -> Result<WireDescent>,
{
    let mut context = WirePreflightContext::new(limits);
    let mut path = Vec::new();
    preflight_wire_message(source, 0, &mut path, &mut context, &mut visitor)?;
    Ok(context.finish())
}

fn preflight_wire_message<'source, F>(
    source: &'source [u8],
    depth: usize,
    path: &mut Vec<u32>,
    context: &mut WirePreflightContext,
    visitor: &mut F,
) -> Result<()>
where
    F: for<'path> FnMut(WireVisit<'source, 'path>) -> Result<WireDescent>,
{
    context.charge_message(source.len(), depth)?;
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_wire_field(source, offset)?;
        offset = field.end;
        context.charge_field()?;
        let field_view = WireFieldView {
            source,
            span: span_from_field(field)?,
        };
        if visitor(WireVisit {
            path: path.as_slice(),
            field: field_view,
        })? == WireDescent::Skip
        {
            continue;
        }
        if field_view.wire_type() != 2 {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} cannot be descended because it is not length-delimited",
                field_view.number()
            )));
        }
        let child_depth = depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf nesting depth overflow".to_owned()))?;
        if child_depth > context.limits.max_nesting() {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: child_depth,
                limit: context.limits.max_nesting(),
            });
        }
        path.try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource: "wire preflight path",
                amount: path.len() + 1,
            })?;
        path.push(field_view.number());
        let child_result =
            preflight_wire_message(field_view.payload(), child_depth, path, context, visitor);
        path.pop();
        child_result?;
    }
    Ok(())
}

/// Overlay singular protobuf fields while retaining every untouched byte.
///
/// Each field present in `overlay` replaces the field with the same number in
/// `base`, or is appended when the base does not contain it. Duplicate fields
/// and wire-type changes are rejected because this helper is only suitable for
/// singular schema fields.
pub fn overlay_singular_wire_fields(base: &[u8], overlay: &[u8]) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let base_fields = parse_wire_fields_with_limits(base, limits)?;
    let overlay_fields = parse_wire_fields_with_limits(overlay, limits)?;
    ensure_rewrite_work(overlay_fields.len(), limits)?;
    if overlay_fields.is_empty() {
        return clone_output(base, limits);
    }

    // Index both messages once. Reparsing the growing output for every
    // overlay field made sparse overlays quadratic in parsing and byte-scan
    // work.
    let mut overlay_indexes = HashMap::new();
    overlay_indexes
        .try_reserve(overlay_fields.len())
        .map_err(|_allocation| Error::Allocation {
            resource: "overlay field numbers",
            amount: overlay_fields.len(),
        })?;
    for (index, field) in overlay_fields.iter().enumerate() {
        if overlay_indexes.insert(field.number, index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "singular protobuf overlay field {} occurs multiple times",
                field.number
            )));
        }
    }

    let mut base_entries = HashMap::new();
    base_entries
        .try_reserve(base_fields.len())
        .map_err(|_allocation| Error::Allocation {
            resource: "base field numbers",
            amount: base_fields.len(),
        })?;
    for (index, field) in base_fields.iter().enumerate() {
        let entry = base_entries.entry(field.number).or_insert((index, 0usize));
        entry.1 = entry
            .1
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    }

    let mut appended_fields = 0usize;
    for overlay_field in &overlay_fields {
        let Some(&(base_index, base_count)) = base_entries.get(&overlay_field.number) else {
            appended_fields = appended_fields
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
            continue;
        };
        if base_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "singular protobuf base field {} occurs multiple times",
                overlay_field.number
            )));
        }
        if base_fields[base_index].wire_type != overlay_field.wire_type {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} changes wire type during overlay",
                overlay_field.number
            )));
        }
    }

    let output_field_count = base_fields
        .len()
        .checked_add(appended_fields)
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(output_field_count, limits)?;

    let mut output_size = 0usize;
    for field in &base_fields {
        let field_size = overlay_indexes.get(&field.number).map_or_else(
            || field.end - field.start,
            |&index| {
                let replacement = &overlay_fields[index];
                replacement.end - replacement.start
            },
        );
        output_size = output_size
            .checked_add(field_size)
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    }
    for field in &overlay_fields {
        if !base_entries.contains_key(&field.number) {
            output_size = output_size
                .checked_add(field.end - field.start)
                .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
        }
    }

    let mut output = output_with_capacity(output_size, limits)?;
    for field in &base_fields {
        if let Some(&index) = overlay_indexes.get(&field.number) {
            let replacement = &overlay_fields[index];
            output.extend_from_slice(&overlay[replacement.start..replacement.end]);
        } else {
            output.extend_from_slice(&base[field.start..field.end]);
        }
    }
    for field in &overlay_fields {
        if !base_entries.contains_key(&field.number) {
            output.extend_from_slice(&overlay[field.start..field.end]);
        }
    }
    debug_assert_eq!(output.len(), output_size);
    Ok(output)
}

pub fn patch_nested_length_delimited_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    ensure_nesting(path.len(), WireLimits::default())?;
    patch_nested_field(data, path, &|nested_data, field_number| {
        patch_length_delimited_field(nested_data, field_number, expected_leaf, replacement)
    })
}

pub fn patch_nested_varint_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    ensure_nesting(path.len(), WireLimits::default())?;
    patch_nested_field(data, path, &|nested_data, field_number| {
        patch_varint_field(nested_data, field_number, expected_leaf, replacement)
    })
}

pub fn patch_nested_fixed32_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    ensure_nesting(path.len(), WireLimits::default())?;
    patch_nested_field(data, path, &|nested_data, field_number| {
        patch_fixed32_field(nested_data, field_number, expected_leaf, replacement)
    })
}

pub fn patch_nested_fixed64_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    ensure_nesting(path.len(), WireLimits::default())?;
    patch_nested_field(data, path, &|nested_data, field_number| {
        patch_fixed64_field(nested_data, field_number, expected_leaf, replacement)
    })
}

/// Rewrite many nested singular leaves with one bounded parse/sizing pass and
/// one final emission.
///
/// The edits are compiled into a field-number trie, so shared ancestors are
/// decoded and rebuilt once even when many leaves change below them. Selected
/// ancestors and leaves must have canonical keys; selected length-delimited
/// framing and selected varint values must also be canonical. All unselected
/// records remain opaque and retain their exact bytes and order, including
/// unknown fields, duplicate unknowns, noncanonical unknown varints, and
/// matched protobuf groups. New leaves are appended to their containing
/// message in edit-slice order.
///
/// The shared input, field, output, and nesting budgets cover the whole
/// selected tree. Rewrite work is exactly the root source length plus final
/// emitted length: every retained or replaced source byte is scanned at most
/// once by the planning pass and every output byte is emitted once. The final
/// output uses one exact, fallible allocation; planning storage is also grown
/// fallibly.
pub fn patch_nested_fields_batched_with_limits(
    data: &[u8],
    edits: &[NestedFieldEdit<'_>],
    limits: WireLimits,
) -> Result<Vec<u8>> {
    if data.len() > limits.max_input_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: data.len(),
            limit: limits.max_input_bytes(),
        });
    }

    let trie = build_batch_edit_trie(edits, limits)?;
    let mut context = BatchPlanContext::new(limits);
    let root_plan = plan_batch_message(data, 0, data.len(), 0, 0, &trie, edits, &mut context)?;
    let root = context
        .plans
        .get(root_plan)
        .ok_or_else(|| Error::InvalidFormat("protobuf batch root plan is missing".to_owned()))?;
    let output_fields = context
        .source_fields
        .checked_sub(context.removed_fields)
        .and_then(|count| count.checked_add(context.appended_fields))
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(output_fields, limits)?;
    let output_size = root.output_len;
    let work = data
        .len()
        .checked_add(output_size)
        .ok_or_else(|| Error::InvalidFormat("protobuf rewrite-work overflow".to_owned()))?;
    ensure_rewrite_work(work, limits)?;

    let mut output = output_with_capacity(output_size, limits)?;
    emit_batch_message(root_plan, data, edits, &context.plans, &mut output)?;
    debug_assert_eq!(output.len(), output_size);
    Ok(output)
}

/// Rewrite many nested singular leaves under the default finite resource
/// profile. See [`patch_nested_fields_batched_with_limits`] for invariants and
/// byte-preservation guarantees.
pub fn patch_nested_fields_batched(data: &[u8], edits: &[NestedFieldEdit<'_>]) -> Result<Vec<u8>> {
    patch_nested_fields_batched_with_limits(data, edits, WireLimits::default())
}

fn build_batch_edit_trie(
    edits: &[NestedFieldEdit<'_>],
    limits: WireLimits,
) -> Result<Vec<BatchTrieNode>> {
    ensure_field_count(edits.len(), limits)?;
    let mut trie = Vec::new();
    trie.try_reserve(1)
        .map_err(|_allocation| Error::Allocation {
            resource: "batched wire edit trie",
            amount: 1,
        })?;
    trie.push(BatchTrieNode::new());

    for (edit_index, edit) in edits.iter().enumerate() {
        if edit.path.is_empty() {
            return Err(Error::InvalidFormat(
                "protobuf field path cannot be empty".to_owned(),
            ));
        }
        ensure_nesting(edit.path.len() - 1, limits)?;
        let mut node_index = 0usize;
        for &field_number in edit.path {
            validate_field_number(field_number)?;
            if trie[node_index].edit.is_some() {
                return Err(Error::InvalidFormat(
                    "protobuf batch edit path crosses an edited leaf".to_owned(),
                ));
            }
            let existing = trie[node_index]
                .children
                .binary_search_by_key(&field_number, |&(number, _)| number)
                .ok()
                .map(|position| trie[node_index].children[position].1);
            node_index = if let Some(child) = existing {
                child
            } else {
                let child = trie.len();
                // The root is not a field-path component, so the next node's
                // index is also the resulting number of selected components.
                ensure_field_count(child, limits)?;
                trie.try_reserve(1)
                    .map_err(|_allocation| Error::Allocation {
                        resource: "batched wire edit trie",
                        amount: child + 1,
                    })?;
                trie.push(BatchTrieNode::new());
                let position = trie[node_index]
                    .children
                    .binary_search_by_key(&field_number, |&(number, _)| number)
                    .unwrap_or_else(|position| position);
                trie[node_index]
                    .children
                    .try_reserve(1)
                    .map_err(|_allocation| Error::Allocation {
                        resource: "batched wire trie children",
                        amount: trie[node_index].children.len() + 1,
                    })?;
                trie[node_index]
                    .children
                    .insert(position, (field_number, child));
                child
            };
        }
        if trie[node_index].edit.is_some() {
            return Err(Error::InvalidFormat(
                "protobuf batch edit path occurs more than once".to_owned(),
            ));
        }
        if !trie[node_index].children.is_empty() {
            return Err(Error::InvalidFormat(
                "protobuf batch edited leaf is also an ancestor".to_owned(),
            ));
        }
        trie[node_index].edit = Some(edit_index);
    }
    Ok(trie)
}

#[allow(
    clippy::too_many_arguments,
    reason = "The recursive planner carries one explicit source range, trie node, depth, and shared bounded state"
)]
fn plan_batch_message(
    data: &[u8],
    start: usize,
    end: usize,
    trie_node: usize,
    depth: usize,
    trie: &[BatchTrieNode],
    edits: &[NestedFieldEdit<'_>],
    context: &mut BatchPlanContext,
) -> Result<usize> {
    ensure_nesting(depth, context.limits)?;
    let node = trie
        .get(trie_node)
        .ok_or_else(|| Error::InvalidFormat("protobuf batch trie node is missing".to_owned()))?;
    let mut seen = Vec::new();
    seen.try_reserve_exact(node.children.len())
        .map_err(|_allocation| Error::Allocation {
            resource: "batched wire occurrence counters",
            amount: node.children.len(),
        })?;
    seen.resize(node.children.len(), 0usize);
    let mut fields = Vec::new();
    let mut output_len = 0usize;
    let mut changed = false;
    let mut offset = start;

    while offset < end {
        let field = parse_batch_field(data, offset, end, depth, context)?;
        offset = field.end;
        let selected = node
            .children
            .binary_search_by_key(&field.number, |&(number, _)| number)
            .ok();
        let (action, field_output_len, field_changed) = if let Some(position) = selected {
            seen[position] = seen[position]
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("protobuf occurrence overflow".to_owned()))?;
            if seen[position] > 1 {
                return Err(Error::InvalidFormat(format!(
                    "singular protobuf field {} occurs more than once",
                    field.number
                )));
            }
            let child_index = node.children[position].1;
            let child = &trie[child_index];
            if let Some(edit_index) = child.edit {
                plan_batch_leaf(data, field, edit_index, edits, context)?
            } else {
                if field.wire_type != 2 {
                    return Err(Error::InvalidFormat(format!(
                        "protobuf field {} is not length-delimited",
                        field.number
                    )));
                }
                field.validate_canonical_framing(data)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("protobuf nesting depth overflow".to_owned())
                })?;
                ensure_nesting(child_depth, context.limits)?;
                let nested_plan = plan_batch_message(
                    data,
                    field.payload_start,
                    field.end,
                    child_index,
                    child_depth,
                    trie,
                    edits,
                    context,
                )?;
                let nested = context.plans.get(nested_plan).ok_or_else(|| {
                    Error::InvalidFormat("protobuf nested batch plan is missing".to_owned())
                })?;
                if nested.changed {
                    let payload_length =
                        u64::try_from(nested.output_len).map_err(|_conversion| {
                            Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
                        })?;
                    let length = field
                        .key_end
                        .checked_sub(field.start)
                        .and_then(|key| key.checked_add(crate::varint::encoded_len(payload_length)))
                        .and_then(|framing| framing.checked_add(nested.output_len))
                        .ok_or_else(|| {
                            Error::InvalidFormat("protobuf output size overflow".to_owned())
                        })?;
                    (BatchFieldAction::Nested(nested_plan), length, true)
                } else {
                    (BatchFieldAction::Copy, field.end - field.start, false)
                }
            }
        } else {
            (BatchFieldAction::Copy, field.end - field.start, false)
        };
        output_len = output_len
            .checked_add(field_output_len)
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
        changed |= field_changed;
        fields
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource: "batched wire field plans",
                amount: fields.len() + 1,
            })?;
        fields.push(BatchPlannedField {
            start: field.start,
            key_end: field.key_end,
            end: field.end,
            action,
        });
    }

    let mut appended = Vec::new();
    for (position, &(field_number, child_index)) in node.children.iter().enumerate() {
        let child = &trie[child_index];
        if seen[position] != 0 {
            continue;
        }
        let Some(edit_index) = child.edit else {
            return Err(Error::InvalidFormat(format!(
                "singular protobuf ancestor field {field_number} changed during mutation"
            )));
        };
        let edit = edits[edit_index];
        if edit.expected_present {
            return Err(Error::InvalidFormat(format!(
                "singular protobuf field {field_number} changed during mutation"
            )));
        }
        if !edit.replacement.is_some() {
            continue;
        }
        let appended_len = encoded_batch_replacement_len(field_number, edit.replacement, true)?;
        output_len = output_len
            .checked_add(appended_len)
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
        context.appended_fields = context
            .appended_fields
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
        appended
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation {
                resource: "batched appended wire edits",
                amount: appended.len() + 1,
            })?;
        appended.push(edit_index);
        changed = true;
    }
    appended.sort_unstable();

    context.push_plan(BatchMessagePlan {
        fields,
        appended,
        output_len,
        changed,
    })
}

fn plan_batch_leaf(
    data: &[u8],
    field: WireField,
    edit_index: usize,
    edits: &[NestedFieldEdit<'_>],
    context: &mut BatchPlanContext,
) -> Result<(BatchFieldAction, usize, bool)> {
    let edit = edits[edit_index];
    if !edit.expected_present {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {} changed during mutation",
            field.number
        )));
    }
    let expected_wire_type = edit.replacement.wire_type();
    if field.wire_type != expected_wire_type {
        return Err(Error::InvalidFormat(format!(
            "protobuf field {} has wire type {}, expected {expected_wire_type}",
            field.number, field.wire_type
        )));
    }
    validate_batch_selected_leaf(data, field)?;
    if !edit.replacement.is_some() {
        context.removed_fields = context
            .removed_fields
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
        return Ok((BatchFieldAction::Leaf(edit_index), 0, true));
    }
    if batch_replacement_matches(data, field, edit.replacement)? {
        return Ok((BatchFieldAction::Copy, field.end - field.start, false));
    }
    let length = encoded_batch_replacement_len(field.number, edit.replacement, false)?
        .checked_add(field.key_end - field.start)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    Ok((BatchFieldAction::Leaf(edit_index), length, true))
}

fn validate_batch_selected_leaf(data: &[u8], field: WireField) -> Result<()> {
    field.validate_canonical_key(data)?;
    if field.wire_type == 2 {
        field.validate_canonical_length(data)?;
    } else if field.wire_type == 0 {
        let payload = field.payload(data)?;
        let (value, width) = crate::varint::decode_varint_from_bytes(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid protobuf varint value for field {}: {error}",
                field.number
            ))
        })?;
        if width != payload.len() || width != crate::varint::encoded_len(value) {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {} has a noncanonical varint value",
                field.number
            )));
        }
    }
    Ok(())
}

fn batch_replacement_matches(
    data: &[u8],
    field: WireField,
    replacement: NestedFieldReplacement<'_>,
) -> Result<bool> {
    let payload = field.payload(data)?;
    Ok(match replacement {
        NestedFieldReplacement::Varint(Some(value)) => {
            let (existing, width) =
                crate::varint::decode_varint_from_bytes(payload).map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid protobuf varint value for field {}: {error}",
                        field.number
                    ))
                })?;
            width == payload.len() && existing == value
        },
        NestedFieldReplacement::Fixed32(Some(value)) => payload == value.to_le_bytes(),
        NestedFieldReplacement::Fixed64(Some(value)) => payload == value.to_le_bytes(),
        NestedFieldReplacement::LengthDelimited(Some(value)) => payload == value,
        NestedFieldReplacement::Varint(None)
        | NestedFieldReplacement::Fixed32(None)
        | NestedFieldReplacement::Fixed64(None)
        | NestedFieldReplacement::LengthDelimited(None) => false,
    })
}

fn encoded_batch_replacement_len(
    field_number: u32,
    replacement: NestedFieldReplacement<'_>,
    include_key: bool,
) -> Result<usize> {
    let payload_len = match replacement {
        NestedFieldReplacement::Varint(Some(value)) => crate::varint::encoded_len(value),
        NestedFieldReplacement::Fixed32(Some(_)) => 4,
        NestedFieldReplacement::Fixed64(Some(_)) => 8,
        NestedFieldReplacement::LengthDelimited(Some(payload)) => {
            let length = u64::try_from(payload.len()).map_err(|_conversion| {
                Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
            })?;
            crate::varint::encoded_len(length)
                .checked_add(payload.len())
                .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?
        },
        NestedFieldReplacement::Varint(None)
        | NestedFieldReplacement::Fixed32(None)
        | NestedFieldReplacement::Fixed64(None)
        | NestedFieldReplacement::LengthDelimited(None) => 0,
    };
    if !include_key {
        return Ok(payload_len);
    }
    let key = (u64::from(field_number) << 3) | u64::from(replacement.wire_type());
    crate::varint::encoded_len(key)
        .checked_add(payload_len)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))
}

fn parse_batch_field(
    data: &[u8],
    start: usize,
    end: usize,
    depth: usize,
    context: &mut BatchPlanContext,
) -> Result<WireField> {
    let (number, wire_type, key_end) = parse_batch_key(data, start, end)?;
    context.charge_field()?;
    if wire_type == 4 {
        return Err(Error::InvalidFormat(format!(
            "unexpected protobuf end-group field {number}"
        )));
    }
    let mut payload_start = key_end;
    let field_end = if wire_type == 3 {
        let group_depth = depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf nesting depth overflow".to_owned()))?;
        ensure_nesting(group_depth, context.limits)?;
        skip_batch_group(data, key_end, end, number, group_depth, context)?
    } else {
        let (payload, parsed_end) = parse_batch_payload(data, key_end, end, wire_type)?;
        payload_start = payload;
        parsed_end
    };
    Ok(WireField {
        number,
        wire_type,
        start,
        key_end,
        payload_start,
        end: field_end,
    })
}

fn parse_batch_key(data: &[u8], start: usize, end: usize) -> Result<(u32, u8, usize)> {
    let source = data.get(start..end).ok_or_else(|| {
        Error::InvalidFormat("protobuf field range is outside the supplied data".to_owned())
    })?;
    let (key, key_length) = crate::varint::decode_varint_from_bytes(source)
        .map_err(|error| Error::InvalidFormat(format!("invalid protobuf key: {error}")))?;
    let key_end = start
        .checked_add(key_length)
        .ok_or_else(|| Error::InvalidFormat("protobuf key offset overflow".to_owned()))?;
    let number = u32::try_from(key >> 3).map_err(|error| {
        Error::InvalidFormat(format!("protobuf field number does not fit u32: {error}"))
    })?;
    validate_field_number(number)?;
    let wire_type = u8::try_from(key & 7).map_err(|error| {
        Error::InvalidFormat(format!("protobuf wire type does not fit u8: {error}"))
    })?;
    if wire_type > 5 {
        return Err(Error::InvalidFormat(format!(
            "invalid protobuf wire type {wire_type}"
        )));
    }
    Ok((number, wire_type, key_end))
}

fn parse_batch_payload(
    data: &[u8],
    key_end: usize,
    message_end: usize,
    wire_type: u8,
) -> Result<(usize, usize)> {
    let mut payload_start = key_end;
    let end = match wire_type {
        0 => {
            let source = data.get(key_end..message_end).ok_or_else(|| {
                Error::InvalidFormat("protobuf varint range is outside the message".to_owned())
            })?;
            let (_, width) = crate::varint::decode_varint_from_bytes(source).map_err(|error| {
                Error::InvalidFormat(format!("invalid protobuf varint value: {error}"))
            })?;
            key_end
                .checked_add(width)
                .ok_or_else(|| Error::InvalidFormat("protobuf varint offset overflow".to_owned()))?
        },
        1 => key_end
            .checked_add(8)
            .ok_or_else(|| Error::InvalidFormat("protobuf fixed64 offset overflow".to_owned()))?,
        2 => {
            let source = data.get(key_end..message_end).ok_or_else(|| {
                Error::InvalidFormat("protobuf length range is outside the message".to_owned())
            })?;
            let (encoded_length, width) =
                crate::varint::decode_varint_from_bytes(source).map_err(|error| {
                    Error::InvalidFormat(format!("invalid protobuf length: {error}"))
                })?;
            payload_start = key_end.checked_add(width).ok_or_else(|| {
                Error::InvalidFormat("protobuf length-prefix overflow".to_owned())
            })?;
            let length = usize::try_from(encoded_length).map_err(|error| {
                Error::InvalidFormat(format!("protobuf field length exceeds usize: {error}"))
            })?;
            payload_start
                .checked_add(length)
                .ok_or_else(|| Error::InvalidFormat("protobuf field range overflow".to_owned()))?
        },
        5 => key_end
            .checked_add(4)
            .ok_or_else(|| Error::InvalidFormat("protobuf fixed32 offset overflow".to_owned()))?,
        3 | 4 => {
            return Err(Error::InvalidFormat(
                "protobuf group was parsed as a scalar payload".to_owned(),
            ));
        },
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid protobuf wire type {wire_type}"
            )));
        },
    };
    if end > message_end {
        return Err(Error::InvalidFormat("truncated protobuf field".to_owned()));
    }
    Ok((payload_start, end))
}

fn skip_batch_group(
    data: &[u8],
    mut offset: usize,
    message_end: usize,
    group_number: u32,
    depth: usize,
    context: &mut BatchPlanContext,
) -> Result<usize> {
    while offset < message_end {
        let (number, wire_type, key_end) = parse_batch_key(data, offset, message_end)?;
        context.charge_field()?;
        if wire_type == 4 {
            if number != group_number {
                return Err(Error::InvalidFormat(format!(
                    "protobuf end-group field {number} does not match start-group field {group_number}"
                )));
            }
            return Ok(key_end);
        }
        offset = if wire_type == 3 {
            let nested_depth = depth.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("protobuf nesting depth overflow".to_owned())
            })?;
            ensure_nesting(nested_depth, context.limits)?;
            skip_batch_group(data, key_end, message_end, number, nested_depth, context)?
        } else {
            parse_batch_payload(data, key_end, message_end, wire_type)?.1
        };
    }
    Err(Error::InvalidFormat(format!(
        "protobuf start-group field {group_number} is not terminated"
    )))
}

fn emit_batch_message(
    plan_index: usize,
    data: &[u8],
    edits: &[NestedFieldEdit<'_>],
    plans: &[BatchMessagePlan],
    output: &mut Vec<u8>,
) -> Result<()> {
    let plan = plans.get(plan_index).ok_or_else(|| {
        Error::InvalidFormat("protobuf batch emission plan is missing".to_owned())
    })?;
    for field in &plan.fields {
        match field.action {
            BatchFieldAction::Copy => output.extend_from_slice(&data[field.start..field.end]),
            BatchFieldAction::Leaf(edit_index) => {
                let edit = edits[edit_index];
                if edit.replacement.is_some() {
                    output.extend_from_slice(&data[field.start..field.key_end]);
                    emit_batch_replacement_payload(edit.replacement, output)?;
                }
            },
            BatchFieldAction::Nested(nested_plan) => {
                output.extend_from_slice(&data[field.start..field.key_end]);
                let nested = plans.get(nested_plan).ok_or_else(|| {
                    Error::InvalidFormat("protobuf nested emission plan is missing".to_owned())
                })?;
                let length = u64::try_from(nested.output_len).map_err(|_conversion| {
                    Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
                })?;
                crate::varint::encode_varint_into(output, length);
                emit_batch_message(nested_plan, data, edits, plans, output)?;
            },
        }
    }
    for &edit_index in &plan.appended {
        let edit = edits[edit_index];
        let field_number = *edit.path.last().ok_or_else(|| {
            Error::InvalidFormat("protobuf field path cannot be empty".to_owned())
        })?;
        let key = (u64::from(field_number) << 3) | u64::from(edit.replacement.wire_type());
        crate::varint::encode_varint_into(output, key);
        emit_batch_replacement_payload(edit.replacement, output)?;
    }
    Ok(())
}

fn emit_batch_replacement_payload(
    replacement: NestedFieldReplacement<'_>,
    output: &mut Vec<u8>,
) -> Result<()> {
    match replacement {
        NestedFieldReplacement::Varint(Some(value)) => {
            crate::varint::encode_varint_into(output, value);
        },
        NestedFieldReplacement::Fixed32(Some(value)) => {
            output.extend_from_slice(&value.to_le_bytes());
        },
        NestedFieldReplacement::Fixed64(Some(value)) => {
            output.extend_from_slice(&value.to_le_bytes());
        },
        NestedFieldReplacement::LengthDelimited(Some(payload)) => {
            let length = u64::try_from(payload.len()).map_err(|_conversion| {
                Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
            })?;
            crate::varint::encode_varint_into(output, length);
            output.extend_from_slice(payload);
        },
        NestedFieldReplacement::Varint(None)
        | NestedFieldReplacement::Fixed32(None)
        | NestedFieldReplacement::Fixed64(None)
        | NestedFieldReplacement::LengthDelimited(None) => {},
    }
    Ok(())
}

pub fn patch_length_delimited_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    let existing_field = singular_field(data, field_number, expected_present)?;
    let Some(field) = existing_field else {
        let Some(replacement_payload) = replacement else {
            return clone_output(data, WireLimits::default());
        };
        let limits = WireLimits::default();
        ensure_one_more_field(data, limits)?;
        let mut output = clone_output(data, limits)?;
        append_length_delimited_field_unchecked(
            &mut output,
            field_number,
            replacement_payload,
            limits,
        )?;
        return Ok(output);
    };
    require_length_delimited(&field)?;
    match replacement {
        Some(replacement_payload) => {
            replace_existing_length_delimited_field(data, &field, replacement_payload)
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Patch one optional protobuf varint while retaining every other byte.
///
/// `expected_present` is checked against the raw wire representation so a
/// concurrent schema mismatch, duplicate singular field, or wrong wire type
/// fails instead of being silently normalized.
pub fn patch_varint_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    let existing_field = singular_field(data, field_number, expected_present)?;
    let Some(field) = existing_field else {
        let Some(replacement_value) = replacement else {
            return clone_output(data, WireLimits::default());
        };
        let limits = WireLimits::default();
        ensure_one_more_field(data, limits)?;
        let mut output = clone_output(data, limits)?;
        append_varint_field_unchecked(&mut output, field_number, replacement_value, limits)?;
        return Ok(output);
    };
    require_wire_type(&field, 0, "varint")?;
    match replacement {
        Some(replacement_value) => {
            let mut buffer = [0u8; crate::varint::MAX_BYTES];
            let encoded = crate::varint::encode_varint_to_buffer(replacement_value, &mut buffer);
            replace_existing_scalar_field(data, &field, encoded)
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Patch one optional protobuf fixed32 value while retaining every other byte.
pub fn patch_fixed32_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    let existing_field = singular_field(data, field_number, expected_present)?;
    let Some(field) = existing_field else {
        let Some(replacement_value) = replacement else {
            return clone_output(data, WireLimits::default());
        };
        let limits = WireLimits::default();
        ensure_one_more_field(data, limits)?;
        let mut output = clone_output(data, limits)?;
        append_scalar_field_unchecked(
            &mut output,
            field_number,
            5,
            &replacement_value.to_le_bytes(),
            limits,
        )?;
        return Ok(output);
    };
    require_wire_type(&field, 5, "fixed32")?;
    match replacement {
        Some(replacement_value) => {
            replace_existing_scalar_field(data, &field, &replacement_value.to_le_bytes())
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Patch one optional protobuf fixed64 value while retaining every other byte.
pub fn patch_fixed64_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    let existing_field = singular_field(data, field_number, expected_present)?;
    let Some(field) = existing_field else {
        let Some(replacement_value) = replacement else {
            return clone_output(data, WireLimits::default());
        };
        let limits = WireLimits::default();
        ensure_one_more_field(data, limits)?;
        let mut output = clone_output(data, limits)?;
        append_scalar_field_unchecked(
            &mut output,
            field_number,
            1,
            &replacement_value.to_le_bytes(),
            limits,
        )?;
        return Ok(output);
    };
    require_wire_type(&field, 1, "fixed64")?;
    match replacement {
        Some(replacement_value) => {
            replace_existing_scalar_field(data, &field, &replacement_value.to_le_bytes())
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Transform one singular length-delimited field.
///
/// The callback may use a caller-owned error type. Shared wire failures are
/// converted through `From<Error>`, so format facades can retain their own
/// structured error taxonomy without duplicating the mutation algorithm.
pub fn transform_length_delimited_field<F, E>(
    data: &[u8],
    field_number: u32,
    transform: F,
) -> std::result::Result<Vec<u8>, E>
where
    F: FnOnce(&[u8]) -> std::result::Result<Vec<u8>, E>,
    E: From<Error>,
{
    let field = singular_field(data, field_number, true)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "singular protobuf field {field_number} must occur exactly once"
        ))
    })?;
    require_length_delimited(&field)?;
    let replacement = transform(&data[field.payload_start..field.end])?;
    Ok(replace_existing_length_delimited_field(
        data,
        &field,
        &replacement,
    )?)
}

pub fn append_repeated_length_delimited_field(
    data: &[u8],
    field_number: u32,
    payload: &[u8],
) -> Result<Vec<u8>> {
    // Parse first so malformed existing data cannot be normalized accidentally.
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_field_count(
        fields
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?,
        limits,
    )?;
    let mut output = clone_output(data, limits)?;
    append_length_delimited_field_unchecked(&mut output, field_number, payload, limits)?;
    Ok(output)
}

/// Return raw payloads for every occurrence of one repeated length-delimited
/// field, rejecting a same-number field with an incompatible wire type.
pub fn repeated_length_delimited_payloads(data: &[u8], field_number: u32) -> Result<Vec<&[u8]>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .count();
    let mut payloads = Vec::new();
    payloads
        .try_reserve(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "repeated length-delimited payloads",
            amount: matches,
        })?;
    for field in &fields {
        if field.number != field_number {
            continue;
        }
        require_length_delimited(field)?;
        payloads.push(&data[field.payload_start..field.end]);
    }
    Ok(payloads)
}

/// Replace all occurrences of a repeated length-delimited field while keeping
/// unrelated fields in their original byte positions. Existing field slots
/// retain their original key bytes; additional values are inserted directly
/// after the final existing occurrence so an append followed by removal is
/// byte-exact.
pub fn rewrite_repeated_length_delimited_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[Vec<u8>],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_rewrite_work(replacements.len(), limits)?;
    let field_count = fields.len();
    let matches = matching_fields(&fields, field_number)?;
    for field in &matches {
        require_length_delimited(field)?;
    }
    let retained_count = field_count
        .checked_sub(matches.len())
        .and_then(|count| count.checked_add(replacements.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(retained_count, limits)?;
    if matches.is_empty() {
        if !replacements.is_empty() {
            validate_field_number(field_number)?;
        }
        let mut output = clone_output(data, limits)?;
        for replacement in replacements {
            append_length_delimited_field_unchecked(
                &mut output,
                field_number,
                replacement,
                limits,
            )?;
        }
        return Ok(output);
    }

    // A semantic no-op must remain byte-exact, including overlong length
    // prefixes and noncanonical keys on otherwise unknown-to-this-helper
    // fields. Re-encoding equal payloads would silently normalize them.
    if replacements.len() == matches.len()
        && matches
            .iter()
            .zip(replacements)
            .all(|(field, replacement)| {
                &data[field.payload_start..field.end] == replacement.as_slice()
            })
    {
        return clone_output(data, limits);
    }

    if !replacements.is_empty() {
        validate_field_number(field_number)?;
    }
    let removed_length = matches.iter().try_fold(0usize, |total, field| {
        total.checked_add(field.end - field.start)
    });
    let mut capacity = data
        .len()
        .checked_sub(removed_length.ok_or_else(|| {
            Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    for (index, replacement) in replacements.iter().enumerate() {
        let key_length = matches.get(index).map_or_else(
            || field_key_size(field_number, 2),
            |field| Ok(field.key_end - field.start),
        )?;
        capacity = capacity
            .checked_add(
                key_length
                    .checked_add(encoded_length_size(replacement.len())?)
                    .and_then(|length| length.checked_add(replacement.len()))
                    .ok_or_else(|| {
                        Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
                    })?,
            )
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            crate::varint::encode_varint_into(
                &mut output,
                u64::try_from(replacement.len()).map_err(|_conversion| {
                    Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
                })?,
            );
            output.extend_from_slice(replacement);
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for replacement in &replacements[matches.len().min(replacements.len())..] {
                append_length_delimited_field_unchecked(
                    &mut output,
                    field_number,
                    replacement,
                    limits,
                )?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Return every occurrence of an unpacked repeated varint field.
///
/// Proto2 iWork archives commonly use unpacked integer arrays. Rejecting a
/// packed or otherwise mismatched occurrence keeps bounded mutations from
/// silently normalizing an unexpected representation.
pub fn repeated_varint_values(data: &[u8], field_number: u32) -> Result<Vec<u64>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "repeated varint values",
            amount: matches,
        })?;
    for field in fields {
        if field.number != field_number {
            continue;
        }
        require_wire_type(&field, 0, "varint")?;
        let (value, length) = crate::varint::decode_varint_from_bytes(
            &data[field.key_end..field.end],
        )
        .map_err(|error| Error::InvalidFormat(format!("invalid protobuf varint value: {error}")))?;
        let value_end = field
            .key_end
            .checked_add(length)
            .ok_or_else(|| Error::InvalidFormat("protobuf varint offset overflow".to_owned()))?;
        if value_end != field.end {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {field_number} has trailing varint bytes"
            )));
        }
        values.push(value);
    }
    Ok(values)
}

/// Replace an unpacked repeated varint field while preserving unrelated bytes
/// and the original key bytes and positions of retained slots.
pub fn rewrite_repeated_varint_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_rewrite_work(replacements.len(), limits)?;
    let field_count = fields.len();
    let matches = matching_fields(&fields, field_number)?;
    for field in &matches {
        require_wire_type(field, 0, "varint")?;
    }
    let retained_count = field_count
        .checked_sub(matches.len())
        .and_then(|count| count.checked_add(replacements.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(retained_count, limits)?;
    if matches.is_empty() {
        if !replacements.is_empty() {
            validate_field_number(field_number)?;
        }
        let mut output = clone_output(data, limits)?;
        for &replacement in replacements {
            append_varint_field_unchecked(&mut output, field_number, replacement, limits)?;
        }
        return Ok(output);
    }

    if replacements.len() == matches.len()
        && matches
            .iter()
            .zip(replacements)
            .all(|(field, replacement)| {
                is_same_varint_value(*replacement, &data[field.key_end..field.end])
            })
    {
        return clone_output(data, limits);
    }

    if !replacements.is_empty() {
        validate_field_number(field_number)?;
    }
    let removed_length = matches.iter().try_fold(0usize, |total, field| {
        total.checked_add(field.end - field.start)
    });
    let mut capacity = data
        .len()
        .checked_sub(removed_length.ok_or_else(|| {
            Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    for (index, replacement) in replacements.iter().enumerate() {
        let key_length = matches.get(index).map_or_else(
            || field_key_size(field_number, 0),
            |field| Ok(field.key_end - field.start),
        )?;
        capacity = capacity
            .checked_add(
                key_length
                    .checked_add(crate::varint::encoded_len(*replacement))
                    .ok_or_else(|| {
                        Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
                    })?,
            )
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(&replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            crate::varint::encode_varint_into(&mut output, replacement);
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for &replacement in &replacements[matches.len().min(replacements.len())..] {
                append_varint_field_unchecked(&mut output, field_number, replacement, limits)?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Return every occurrence of an unpacked repeated fixed64 field.
pub fn repeated_fixed64_values(data: &[u8], field_number: u32) -> Result<Vec<u64>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "repeated fixed64 values",
            amount: matches,
        })?;
    for field in fields {
        if field.number != field_number {
            continue;
        }
        require_wire_type(&field, 1, "fixed64")?;
        let bytes: [u8; 8] = data[field.payload_start..field.end]
            .try_into()
            .map_err(|_conversion| Error::InvalidFormat("truncated protobuf fixed64".to_owned()))?;
        values.push(u64::from_le_bytes(bytes));
    }
    Ok(values)
}

/// Return every occurrence of an unpacked repeated fixed32 field.
pub fn repeated_fixed32_values(data: &[u8], field_number: u32) -> Result<Vec<u32>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "repeated fixed32 values",
            amount: matches,
        })?;
    for field in fields {
        if field.number != field_number {
            continue;
        }
        require_wire_type(&field, 5, "fixed32")?;
        let bytes: [u8; 4] = data[field.payload_start..field.end]
            .try_into()
            .map_err(|_conversion| Error::InvalidFormat("truncated protobuf fixed32".to_owned()))?;
        values.push(u32::from_le_bytes(bytes));
    }
    Ok(values)
}

/// Replace an unpacked repeated fixed32 field while preserving unrelated bytes.
pub fn rewrite_repeated_fixed32_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u32],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_rewrite_work(replacements.len(), limits)?;
    let field_count = fields.len();
    let matches = matching_fields(&fields, field_number)?;
    for field in &matches {
        require_wire_type(field, 5, "fixed32")?;
    }
    let retained_count = field_count
        .checked_sub(matches.len())
        .and_then(|count| count.checked_add(replacements.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(retained_count, limits)?;
    if matches.is_empty() {
        if !replacements.is_empty() {
            validate_field_number(field_number)?;
        }
        let mut output = clone_output(data, limits)?;
        for replacement in replacements {
            append_scalar_field_unchecked(
                &mut output,
                field_number,
                5,
                &replacement.to_le_bytes(),
                limits,
            )?;
        }
        return Ok(output);
    }

    if replacements.len() == matches.len()
        && matches
            .iter()
            .zip(replacements)
            .all(|(field, replacement)| {
                data[field.payload_start..field.end] == replacement.to_le_bytes()
            })
    {
        return clone_output(data, limits);
    }

    if !replacements.is_empty() {
        validate_field_number(field_number)?;
    }
    let removed_length = matches.iter().try_fold(0usize, |total, field| {
        total.checked_add(field.end - field.start)
    });
    let mut capacity = data
        .len()
        .checked_sub(removed_length.ok_or_else(|| {
            Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    for index in 0..replacements.len() {
        let key_length = matches.get(index).map_or_else(
            || field_key_size(field_number, 5),
            |field| Ok(field.key_end - field.start),
        )?;
        capacity = capacity
            .checked_add(key_length.checked_add(4).ok_or_else(|| {
                Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
            })?)
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            output.extend_from_slice(&replacement.to_le_bytes());
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for replacement in &replacements[matches.len().min(replacements.len())..] {
                append_scalar_field_unchecked(
                    &mut output,
                    field_number,
                    5,
                    &replacement.to_le_bytes(),
                    limits,
                )?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Replace an unpacked repeated fixed64 field while preserving unrelated bytes.
pub fn rewrite_repeated_fixed64_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_rewrite_work(replacements.len(), limits)?;
    let field_count = fields.len();
    let matches = matching_fields(&fields, field_number)?;
    for field in &matches {
        require_wire_type(field, 1, "fixed64")?;
    }
    let retained_count = field_count
        .checked_sub(matches.len())
        .and_then(|count| count.checked_add(replacements.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(retained_count, limits)?;
    if matches.is_empty() {
        if !replacements.is_empty() {
            validate_field_number(field_number)?;
        }
        let mut output = clone_output(data, limits)?;
        for replacement in replacements {
            append_scalar_field_unchecked(
                &mut output,
                field_number,
                1,
                &replacement.to_le_bytes(),
                limits,
            )?;
        }
        return Ok(output);
    }

    if replacements.len() == matches.len()
        && matches
            .iter()
            .zip(replacements)
            .all(|(field, replacement)| {
                data[field.payload_start..field.end] == replacement.to_le_bytes()
            })
    {
        return clone_output(data, limits);
    }

    if !replacements.is_empty() {
        validate_field_number(field_number)?;
    }
    let removed_length = matches.iter().try_fold(0usize, |total, field| {
        total.checked_add(field.end - field.start)
    });
    let mut capacity = data
        .len()
        .checked_sub(removed_length.ok_or_else(|| {
            Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    for index in 0..replacements.len() {
        let key_length = matches.get(index).map_or_else(
            || field_key_size(field_number, 1),
            |field| Ok(field.key_end - field.start),
        )?;
        capacity = capacity
            .checked_add(key_length.checked_add(8).ok_or_else(|| {
                Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
            })?)
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            output.extend_from_slice(&replacement.to_le_bytes());
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for replacement in &replacements[matches.len().min(replacements.len())..] {
                append_scalar_field_unchecked(
                    &mut output,
                    field_number,
                    1,
                    &replacement.to_le_bytes(),
                    limits,
                )?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Transform every occurrence of a repeated length-delimited field while
/// preserving its original key bytes and the position of unrelated fields.
/// Callback errors may use any type that converts shared wire errors through
/// `From<Error>`.
pub fn transform_repeated_length_delimited_fields<F, E>(
    data: &[u8],
    field_number: u32,
    mut transform: F,
) -> std::result::Result<Vec<u8>, E>
where
    F: FnMut(&[u8]) -> std::result::Result<Vec<u8>, E>,
    E: From<Error>,
{
    let limits = WireLimits::default();
    let work = Cell::new(0usize);
    transform_repeated_length_delimited_fields_with_budget(
        data,
        field_number,
        &mut transform,
        limits,
        &work,
    )
}

fn transform_repeated_length_delimited_fields_with_budget<F, E>(
    data: &[u8],
    field_number: u32,
    mut transform: F,
    limits: WireLimits,
    work: &Cell<usize>,
) -> std::result::Result<Vec<u8>, E>
where
    F: FnMut(&[u8]) -> std::result::Result<Vec<u8>, E>,
    E: From<Error>,
{
    let payloads = repeated_length_delimited_payloads(data, field_number)?;
    let total_work = work
        .get()
        .checked_add(payloads.len())
        .ok_or_else(|| Error::InvalidFormat("protobuf rewrite-work overflow".to_owned()))?;
    ensure_rewrite_work(total_work, limits)?;
    work.set(total_work);
    let mut replacements = Vec::new();
    replacements
        .try_reserve(payloads.len())
        .map_err(|_allocation| Error::Allocation {
            resource: "transformed length-delimited payloads",
            amount: payloads.len(),
        })?;
    for payload in payloads {
        replacements.push(transform(payload)?);
    }
    Ok(rewrite_repeated_length_delimited_fields(
        data,
        field_number,
        &replacements,
    )?)
}

/// Transform all length-delimited leaves at a protobuf field path.
///
/// Every path component may occur zero, one, or many times. This makes the
/// operation suitable for schema-known paths that cross optional and repeated
/// messages without normalizing untouched siblings or unknown fields. Callback
/// errors may use any type that converts shared wire errors through
/// `From<Error>`.
pub fn transform_length_delimited_fields_at_path<F, E>(
    data: &[u8],
    path: &[u32],
    mut transform: F,
) -> std::result::Result<Vec<u8>, E>
where
    F: FnMut(&[u8]) -> std::result::Result<Vec<u8>, E>,
    E: From<Error>,
{
    fn visit<E>(
        data: &[u8],
        path: &[u32],
        transform: &mut dyn FnMut(&[u8]) -> std::result::Result<Vec<u8>, E>,
        limits: WireLimits,
        work: &Cell<usize>,
    ) -> std::result::Result<Vec<u8>, E>
    where
        E: From<Error>,
    {
        ensure_nesting(path.len(), limits)?;
        let (&field_number, remainder) = path.first().zip(path.get(1..)).ok_or_else(|| {
            Error::InvalidFormat("protobuf field path cannot be empty".to_owned())
        })?;
        if remainder.is_empty() {
            return transform_repeated_length_delimited_fields_with_budget(
                data,
                field_number,
                transform,
                limits,
                work,
            );
        }
        transform_repeated_length_delimited_fields_with_budget(
            data,
            field_number,
            |nested| visit(nested, remainder, transform, limits, work),
            limits,
            work,
        )
    }

    let limits = WireLimits::default();
    ensure_nesting(path.len(), limits)?;
    let work = Cell::new(0usize);
    visit(data, path, &mut transform, limits, &work)
}

/// Remove exactly one repeated field selected by a caller-owned predicate.
/// Predicate errors may use any type that converts shared wire errors through
/// `From<Error>`.
pub fn remove_repeated_length_delimited_field_where<F, E>(
    data: &[u8],
    field_number: u32,
    mut remove: F,
) -> std::result::Result<Vec<u8>, E>
where
    F: FnMut(&[u8]) -> std::result::Result<bool, E>,
    E: From<Error>,
{
    let limits = WireLimits::default();
    let fields = parse_wire_fields_with_limits(data, limits)?;
    ensure_rewrite_work(fields.len(), limits)?;
    let mut removed = Vec::new();
    removed
        .try_reserve(fields.len())
        .map_err(|_allocation| Error::Allocation {
            resource: "removed wire fields",
            amount: fields.len(),
        })?;
    for field in &fields {
        if field.number != field_number {
            continue;
        }
        require_length_delimited(field)?;
        if remove(&data[field.payload_start..field.end])? {
            removed.push(*field);
        }
    }
    if removed.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "expected to remove one protobuf field {field_number}, matched {}",
            removed.len()
        ))
        .into());
    }
    Ok(remove_fields(data, removed)?)
}

fn require_length_delimited(field: &WireField) -> Result<()> {
    require_wire_type(field, 2, "length-delimited")
}

fn patch_nested_field<F>(data: &[u8], path: &[u32], patch_leaf: &F) -> Result<Vec<u8>>
where
    F: Fn(&[u8], u32) -> Result<Vec<u8>>,
{
    ensure_nesting(path.len(), WireLimits::default())?;
    let (&field_number, remainder) = path
        .split_first()
        .ok_or_else(|| Error::InvalidFormat("protobuf field path cannot be empty".to_owned()))?;
    if remainder.is_empty() {
        return patch_leaf(data, field_number);
    }
    transform_length_delimited_field(data, field_number, |nested| {
        patch_nested_field(nested, remainder, patch_leaf)
    })
}

fn singular_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
) -> Result<Option<WireField>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields
        .into_iter()
        .filter(|field| field.number == field_number);
    let field = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} occurs more than once"
        )));
    }
    if field.is_none() == expected_present {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} changed during mutation"
        )));
    }
    Ok(field)
}

fn require_wire_type(field: &WireField, expected: u8, name: &str) -> Result<()> {
    if field.wire_type != expected {
        return Err(Error::InvalidFormat(format!(
            "protobuf field {} is not {name}",
            field.number
        )));
    }
    Ok(())
}

pub fn append_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
) -> Result<()> {
    append_length_delimited_field_with_limits(output, field_number, payload, WireLimits::default())
}

/// Append one length-delimited field under an explicit output budget.
pub fn append_length_delimited_field_with_limits(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    let field_count = parse_wire_fields_with_limits(output, limits)?.len();
    ensure_field_count(
        field_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?,
        limits,
    )?;
    append_length_delimited_field_unchecked(output, field_number, payload, limits)
}

fn append_length_delimited_field_unchecked(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    validate_field_number(field_number)?;
    let payload_length = u64::try_from(payload.len()).map_err(|_conversion| {
        Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
    })?;
    let key = (u64::from(field_number) << 3) | 2;
    let additional = crate::varint::encoded_len(key)
        .checked_add(crate::varint::encoded_len(payload_length))
        .and_then(|length| length.checked_add(payload.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    reserve_output(output, additional, limits)?;
    crate::varint::encode_varint_into(output, key);
    crate::varint::encode_varint_into(output, payload_length);
    output.extend_from_slice(payload);
    Ok(())
}

pub fn append_varint_field(output: &mut Vec<u8>, field_number: u32, value: u64) -> Result<()> {
    let limits = WireLimits::default();
    let field_count = parse_wire_fields_with_limits(output, limits)?.len();
    ensure_field_count(
        field_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?,
        limits,
    )?;
    append_varint_field_unchecked(output, field_number, value, limits)
}

fn append_varint_field_unchecked(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
    limits: WireLimits,
) -> Result<()> {
    validate_field_number(field_number)?;
    let key = u64::from(field_number) << 3;
    let additional = crate::varint::encoded_len(key)
        .checked_add(crate::varint::encoded_len(value))
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    reserve_output(output, additional, limits)?;
    crate::varint::encode_varint_into(output, key);
    crate::varint::encode_varint_into(output, value);
    Ok(())
}

#[cfg(test)]
fn append_scalar_field(
    output: &mut Vec<u8>,
    field_number: u32,
    wire_type: u8,
    payload: &[u8],
) -> Result<()> {
    append_scalar_field_unchecked(
        output,
        field_number,
        wire_type,
        payload,
        WireLimits::default(),
    )
}

fn append_scalar_field_unchecked(
    output: &mut Vec<u8>,
    field_number: u32,
    wire_type: u8,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    validate_field_number(field_number)?;
    validate_wire_type(wire_type)?;
    let key = (u64::from(field_number) << 3) | u64::from(wire_type);
    let additional = crate::varint::encoded_len(key)
        .checked_add(payload.len())
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    reserve_output(output, additional, limits)?;
    crate::varint::encode_varint_into(output, key);
    output.extend_from_slice(payload);
    Ok(())
}

fn validate_field_number(field_number: u32) -> Result<()> {
    if field_number == 0 || field_number > 0x1fff_ffff {
        return Err(Error::InvalidFormat(format!(
            "invalid protobuf field number {field_number}"
        )));
    }
    Ok(())
}

fn validate_wire_type(wire_type: u8) -> Result<()> {
    if matches!(wire_type, 0 | 1 | 2 | 5) {
        return Ok(());
    }
    Err(Error::InvalidFormat(format!(
        "invalid protobuf wire type {wire_type}"
    )))
}

fn reserve_output(output: &mut Vec<u8>, additional: usize, limits: WireLimits) -> Result<()> {
    let requested = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    if requested > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: requested,
            limit: limits.max_output_bytes(),
        });
    }
    output
        .try_reserve(additional)
        .map_err(|_allocation| Error::Allocation {
            resource: "wire output",
            amount: requested,
        })?;
    Ok(())
}

fn output_with_capacity(capacity: usize, limits: WireLimits) -> Result<Vec<u8>> {
    if capacity > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: capacity,
            limit: limits.max_output_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Error::Allocation {
            resource: "wire output",
            amount: capacity,
        })?;
    Ok(output)
}

fn clone_output(data: &[u8], limits: WireLimits) -> Result<Vec<u8>> {
    let mut output = output_with_capacity(data.len(), limits)?;
    output.extend_from_slice(data);
    Ok(output)
}

fn ensure_field_count(count: usize, limits: WireLimits) -> Result<()> {
    if count > limits.max_fields() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: count,
            limit: limits.max_fields(),
        });
    }
    Ok(())
}

fn ensure_one_more_field(data: &[u8], limits: WireLimits) -> Result<()> {
    let fields = parse_wire_fields_with_limits(data, limits)?;
    let count = fields
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("protobuf field-count overflow".to_owned()))?;
    ensure_field_count(count, limits)
}

fn ensure_nesting(depth: usize, limits: WireLimits) -> Result<()> {
    if depth > limits.max_nesting() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::Nesting,
            observed: depth,
            limit: limits.max_nesting(),
        });
    }
    Ok(())
}

fn ensure_rewrite_work(work: usize, limits: WireLimits) -> Result<()> {
    if work > limits.max_rewrite_work() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::RewriteWork,
            observed: work,
            limit: limits.max_rewrite_work(),
        });
    }
    Ok(())
}

fn matching_fields(fields: &[WireField], field_number: u32) -> Result<Vec<WireField>> {
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .count();
    let mut matching = Vec::new();
    matching
        .try_reserve(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "matching wire fields",
            amount: matches,
        })?;
    matching.extend(
        fields
            .iter()
            .copied()
            .filter(|field| field.number == field_number),
    );
    Ok(matching)
}

fn encoded_length_size(length_bytes: usize) -> Result<usize> {
    let length = u64::try_from(length_bytes)
        .map_err(|_conversion| Error::InvalidFormat("protobuf length exceeds u64".to_owned()))?;
    Ok(crate::varint::encoded_len(length))
}

fn is_same_varint_value(value: u64, encoded: &[u8]) -> bool {
    crate::varint::decode_varint_from_bytes(encoded)
        .is_ok_and(|(decoded, length)| decoded == value && length == encoded.len())
}

fn field_key_size(field_number: u32, wire_type: u8) -> Result<usize> {
    validate_field_number(field_number)?;
    validate_wire_type(wire_type)?;
    Ok(crate::varint::encoded_len(
        (u64::from(field_number) << 3) | u64::from(wire_type),
    ))
}

fn replace_existing_length_delimited_field(
    data: &[u8],
    field: &WireField,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default();
    let old_payload = &data[field.payload_start..field.end];
    if old_payload == replacement {
        return clone_output(data, limits);
    }
    let replacement_length = u64::try_from(replacement.len()).map_err(|_conversion| {
        Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
    })?;
    let length_prefix = if old_payload.len() == replacement.len() {
        &data[field.key_end..field.payload_start]
    } else {
        &[]
    };
    let encoded_length = if length_prefix.is_empty() {
        crate::varint::encoded_len(replacement_length)
    } else {
        length_prefix.len()
    };
    let old_length = field.end - field.start;
    let new_length = field.key_end - field.start + encoded_length + replacement.len();
    let capacity = data
        .len()
        .checked_sub(old_length)
        .and_then(|length| length.checked_add(new_length))
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    let mut output = output_with_capacity(capacity, limits)?;
    output.extend_from_slice(&data[..field.start]);
    output.extend_from_slice(&data[field.start..field.key_end]);
    if length_prefix.is_empty() {
        crate::varint::encode_varint_into(&mut output, replacement_length);
    } else {
        output.extend_from_slice(length_prefix);
    }
    output.extend_from_slice(replacement);
    output.extend_from_slice(&data[field.end..]);
    Ok(output)
}

fn replace_existing_scalar_field(
    data: &[u8],
    field: &WireField,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if &data[field.key_end..field.end] == replacement {
        return clone_output(data, WireLimits::default());
    }
    let old_payload_length = field.end - field.key_end;
    let capacity = data
        .len()
        .checked_sub(old_payload_length)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    let mut output = output_with_capacity(capacity, WireLimits::default())?;
    output.extend_from_slice(&data[..field.key_end]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&data[field.end..]);
    Ok(output)
}

fn remove_fields(data: &[u8], mut fields: Vec<WireField>) -> Result<Vec<u8>> {
    fields.sort_by_key(|field| field.start);
    let removed_length = fields.iter().try_fold(0usize, |total, field| {
        total.checked_add(field.end - field.start)
    });
    let capacity = data
        .len()
        .checked_sub(removed_length.ok_or_else(|| {
            Error::InvalidFormat("protobuf removed-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf removed-field range overflow".to_owned()))?;
    let mut output = output_with_capacity(capacity, WireLimits::default())?;
    let mut copied = 0usize;
    for field in fields {
        if field.start < copied || field.end > data.len() {
            return Err(Error::InvalidFormat(
                "protobuf fields to remove overlap or exceed the payload".to_owned(),
            ));
        }
        output.extend_from_slice(&data[copied..field.start]);
        copied = field.end;
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

#[cfg(test)]
#[allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "Wire tests intentionally use concise fixture assertions"
)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[derive(Debug, PartialEq, Eq)]
    enum CallbackError {
        Wire(Error),
        Sentinel,
    }

    impl From<Error> for CallbackError {
        fn from(error: Error) -> Self {
            Self::Wire(error)
        }
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut field = crate::varint::encode_varint(u64::from(number) << 3);
        field.extend(crate::varint::encode_varint(value));
        field
    }

    fn length_delimited_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut field = Vec::new();
        append_length_delimited_field(&mut field, number, payload).unwrap();
        field
    }

    #[test]
    fn generic_callback_boundary_preserves_caller_errors() {
        let mut repeated = Vec::new();
        append_length_delimited_field(&mut repeated, 2, b"first").unwrap();
        append_length_delimited_field(&mut repeated, 2, b"second").unwrap();
        let repeated_error = transform_repeated_length_delimited_fields(&repeated, 2, |_| {
            Err::<Vec<u8>, _>(CallbackError::Sentinel)
        })
        .unwrap_err();
        assert_eq!(repeated_error, CallbackError::Sentinel);

        let mut nested = Vec::new();
        append_length_delimited_field(&mut nested, 2, b"leaf").unwrap();
        let mut outer = Vec::new();
        append_length_delimited_field(&mut outer, 3, &nested).unwrap();
        let path_error = transform_length_delimited_fields_at_path(&outer, &[3, 2], |_| {
            Err::<Vec<u8>, _>(CallbackError::Sentinel)
        })
        .unwrap_err();
        assert_eq!(path_error, CallbackError::Sentinel);

        let removal_error = remove_repeated_length_delimited_field_where(&repeated, 2, |_| {
            Err::<bool, _>(CallbackError::Sentinel)
        })
        .unwrap_err();
        assert_eq!(removal_error, CallbackError::Sentinel);
    }

    #[test]
    fn singular_wire_overlay_replaces_and_appends_without_touching_siblings() {
        let mut base = varint_field(1, 10);
        base.extend(varint_field(2, 20));
        let mut overlay = varint_field(2, 21);
        overlay.extend(varint_field(3, 30));
        let mut expected = varint_field(1, 10);
        expected.extend(varint_field(2, 21));
        expected.extend(varint_field(3, 30));
        assert_eq!(
            overlay_singular_wire_fields(&base, &overlay).unwrap(),
            expected
        );

        let mut duplicate = varint_field(2, 21);
        duplicate.extend(varint_field(2, 22));
        assert!(overlay_singular_wire_fields(&base, &duplicate).is_err());
    }

    #[test]
    fn scalar_patches_preserve_unknown_fields_and_restore_exact_bytes() {
        let mut original = varint_field(99, 9001);
        original.extend([0xaa, 0x01, 0x03, b'a', b'b', b'c']);

        let with_varint = patch_varint_field(&original, 39, false, Some(1)).unwrap();
        assert!(with_varint.starts_with(&original));
        let with_fixed =
            patch_fixed32_field(&with_varint, 30, false, Some(612.0_f32.to_bits())).unwrap();
        assert!(with_fixed.starts_with(&original));
        let with_double =
            patch_fixed64_field(&with_fixed, 31, false, Some(2.5_f64.to_bits())).unwrap();
        assert!(with_double.starts_with(&original));

        let without_varint = patch_varint_field(&with_double, 39, true, None).unwrap();
        let without_fixed = patch_fixed32_field(&without_varint, 30, true, None).unwrap();
        let restored = patch_fixed64_field(&without_fixed, 31, true, None).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn nested_scalar_patch_retains_unknown_ancestors() {
        let mut inner = varint_field(99, 1234);
        let key = crate::varint::encode_varint((u64::from(4_u32) << 3) | 2);
        let mut outer = key;
        outer.extend(crate::varint::encode_varint(inner.len() as u64));
        outer.append(&mut inner);
        outer.extend(varint_field(100, 5678));
        let baseline = outer.clone();

        let changed =
            patch_nested_fixed32_field(&outer, &[4, 1], false, Some(42.0_f32.to_bits())).unwrap();
        let restored = patch_nested_fixed32_field(&changed, &[4, 1], true, None).unwrap();
        assert_eq!(restored, baseline);
    }

    #[test]
    fn batched_nested_edits_match_sequential_typed_patches() {
        let mut leaf = varint_field(1, 7);
        append_scalar_field(&mut leaf, 2, 5, &10_u32.to_le_bytes()).unwrap();
        append_scalar_field(&mut leaf, 3, 1, &20_u64.to_le_bytes()).unwrap();
        append_length_delimited_field(&mut leaf, 4, b"old").unwrap();
        leaf.extend(varint_field(99, 9001));
        let child = length_delimited_field(2, &leaf);
        let mut source = varint_field(90, 1);
        source.extend(length_delimited_field(4, &child));
        source.extend(varint_field(91, 2));

        let edits = [
            NestedFieldEdit::new(&[4, 2, 1], true, NestedFieldReplacement::Varint(Some(8))),
            NestedFieldEdit::new(&[4, 2, 2], true, NestedFieldReplacement::Fixed32(None)),
            NestedFieldEdit::new(&[4, 2, 3], true, NestedFieldReplacement::Fixed64(Some(21))),
            NestedFieldEdit::new(
                &[4, 2, 4],
                true,
                NestedFieldReplacement::LengthDelimited(Some(b"new value")),
            ),
            NestedFieldEdit::new(&[4, 2, 5], false, NestedFieldReplacement::Varint(Some(55))),
        ];
        let batch = patch_nested_fields_batched(&source, &edits).unwrap();

        let mut sequential = patch_nested_varint_field(&source, &[4, 2, 1], true, Some(8)).unwrap();
        sequential = patch_nested_fixed32_field(&sequential, &[4, 2, 2], true, None).unwrap();
        sequential = patch_nested_fixed64_field(&sequential, &[4, 2, 3], true, Some(21)).unwrap();
        sequential =
            patch_nested_length_delimited_field(&sequential, &[4, 2, 4], true, Some(b"new value"))
                .unwrap();
        sequential = patch_nested_varint_field(&sequential, &[4, 2, 5], false, Some(55)).unwrap();
        assert_eq!(batch, sequential);
    }

    #[test]
    fn batched_nested_clears_and_appends_every_leaf_wire_type_exactly() {
        let mut source = varint_field(99, 9001);
        source.extend(varint_field(1, 7));
        append_scalar_field(&mut source, 2, 5, &10_u32.to_le_bytes()).unwrap();
        append_scalar_field(&mut source, 3, 1, &20_u64.to_le_bytes()).unwrap();
        append_length_delimited_field(&mut source, 4, b"opaque").unwrap();
        let cleared = patch_nested_fields_batched(
            &source,
            &[
                NestedFieldEdit::new(&[1], true, NestedFieldReplacement::Varint(None)),
                NestedFieldEdit::new(&[2], true, NestedFieldReplacement::Fixed32(None)),
                NestedFieldEdit::new(&[3], true, NestedFieldReplacement::Fixed64(None)),
                NestedFieldEdit::new(&[4], true, NestedFieldReplacement::LengthDelimited(None)),
            ],
        )
        .unwrap();
        assert_eq!(cleared, varint_field(99, 9001));

        let restored = patch_nested_fields_batched(
            &cleared,
            &[
                NestedFieldEdit::new(&[1], false, NestedFieldReplacement::Varint(Some(7))),
                NestedFieldEdit::new(&[2], false, NestedFieldReplacement::Fixed32(Some(10))),
                NestedFieldEdit::new(&[3], false, NestedFieldReplacement::Fixed64(Some(20))),
                NestedFieldEdit::new(
                    &[4],
                    false,
                    NestedFieldReplacement::LengthDelimited(Some(b"opaque")),
                ),
            ],
        )
        .unwrap();
        assert_eq!(restored, source);
    }

    #[test]
    fn batched_nested_noop_is_exact_and_appends_in_edit_order() {
        let mut inner = vec![0xa0, 0x81, 0x00, 0x81, 0x00];
        inner.extend(varint_field(1, 7));
        inner.extend([0xa5, 0x01, 1, 2, 3, 4]);
        inner.extend([0x92, 0x00, 0x81, 0x00, b'z']);
        let source = length_delimited_field(4, &inner);
        assert_eq!(patch_nested_fields_batched(&source, &[]).unwrap(), source);
        let noops = [
            NestedFieldEdit::new(&[4, 1], true, NestedFieldReplacement::Varint(Some(7))),
            NestedFieldEdit::new(&[4, 3], false, NestedFieldReplacement::Fixed32(None)),
        ];
        assert_eq!(
            patch_nested_fields_batched(&source, &noops).unwrap(),
            source
        );

        let appends = [
            NestedFieldEdit::new(&[4, 7], false, NestedFieldReplacement::Varint(Some(1))),
            NestedFieldEdit::new(&[4, 5], false, NestedFieldReplacement::Fixed32(Some(2))),
        ];
        let changed = patch_nested_fields_batched(&source, &appends).unwrap();
        let mut expected_inner = inner;
        expected_inner.extend(varint_field(7, 1));
        append_scalar_field(&mut expected_inner, 5, 5, &2_u32.to_le_bytes()).unwrap();
        assert_eq!(changed, length_delimited_field(4, &expected_inner));
    }

    #[test]
    fn batched_nested_preserves_all_unknown_wire_types_and_groups() {
        let mut group = vec![0x53];
        group.extend([0x88, 0x00, 0x81, 0x00]);
        group.push(0x5b);
        group.extend([0x15, 1, 2, 3, 4]);
        group.extend([0x5c, 0x54]);

        let mut source = vec![0x91, 0x01, 1, 2, 3, 4, 5, 6, 7, 8];
        source.extend(&group);
        source.extend(varint_field(1, 7));
        source.extend([0x9d, 0x01, 9, 8, 7, 6]);
        source.extend([0xa2, 0x01, 0x81, 0x00, b'x']);
        let edit = [NestedFieldEdit::new(
            &[1],
            true,
            NestedFieldReplacement::Varint(Some(8)),
        )];
        let changed = patch_nested_fields_batched(&source, &edit).unwrap();

        let mut expected = vec![0x91, 0x01, 1, 2, 3, 4, 5, 6, 7, 8];
        expected.extend(&group);
        expected.extend(varint_field(1, 8));
        expected.extend([0x9d, 0x01, 9, 8, 7, 6]);
        expected.extend([0xa2, 0x01, 0x81, 0x00, b'x']);
        assert_eq!(changed, expected);
    }

    #[test]
    fn batched_nested_rejects_malformed_groups_and_bounds_group_depth() {
        let edit = [NestedFieldEdit::new(
            &[1],
            true,
            NestedFieldReplacement::Varint(Some(2)),
        )];
        for malformed in [
            vec![0x53, 0x08, 0x01],
            vec![0x53, 0x08, 0x01, 0x5c],
            vec![0x54, 0x08, 0x01],
        ] {
            assert!(patch_nested_fields_batched(&malformed, &edit).is_err());
        }

        let nested_groups = [0x53, 0x5b, 0x5c, 0x54, 0x08, 0x01];
        let shallow = WireLimits::default().with_nesting(1).unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&nested_groups, &edit, shallow),
            Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: 2,
                limit: 1,
            })
        ));
        let exact = WireLimits::default().with_nesting(2).unwrap();
        assert!(patch_nested_fields_batched_with_limits(&nested_groups, &edit, exact).is_ok());
    }

    #[test]
    fn batched_nested_recalculates_each_ancestor_varint_width() {
        for (old_payload, new_payload, old_inner, new_inner) in
            [(125, 126, 127, 128), (16_380, 16_381, 16_383, 16_384)]
        {
            let old = vec![b'a'; old_payload];
            let new = vec![b'b'; new_payload];
            let inner = length_delimited_field(1, &old);
            assert_eq!(inner.len(), old_inner);
            let source = length_delimited_field(4, &inner);
            let edit = [NestedFieldEdit::new(
                &[4, 1],
                true,
                NestedFieldReplacement::LengthDelimited(Some(&new)),
            )];
            let changed = patch_nested_fields_batched(&source, &edit).unwrap();
            let expected_inner = length_delimited_field(1, &new);
            assert_eq!(expected_inner.len(), new_inner);
            assert_eq!(changed, length_delimited_field(4, &expected_inner));
        }
    }

    #[test]
    fn batched_nested_enforces_presence_type_and_path_uniqueness() {
        let source = varint_field(1, 7);
        let duplicate_source = [source.as_slice(), source.as_slice()].concat();
        let replace = NestedFieldEdit::new(&[1], true, NestedFieldReplacement::Varint(Some(8)));
        assert!(patch_nested_fields_batched(&duplicate_source, &[replace]).is_err());
        assert!(
            patch_nested_fields_batched(
                &source,
                &[NestedFieldEdit::new(
                    &[1],
                    false,
                    NestedFieldReplacement::Varint(Some(8)),
                )],
            )
            .is_err()
        );
        assert!(
            patch_nested_fields_batched(
                &source,
                &[NestedFieldEdit::new(
                    &[1],
                    true,
                    NestedFieldReplacement::Fixed32(Some(8)),
                )],
            )
            .is_err()
        );
        assert!(patch_nested_fields_batched(&source, &[replace, replace]).is_err());
        assert!(
            patch_nested_fields_batched(
                &source,
                &[
                    replace,
                    NestedFieldEdit::new(&[1, 2], true, NestedFieldReplacement::Varint(Some(3)),),
                ],
            )
            .is_err()
        );
    }

    #[test]
    fn batched_nested_rejects_noncanonical_selected_framing_only() {
        let varint_edit = [NestedFieldEdit::new(
            &[1],
            true,
            NestedFieldReplacement::Varint(Some(2)),
        )];
        assert!(patch_nested_fields_batched(&[0x88, 0x00, 0x01], &varint_edit).is_err());
        assert!(patch_nested_fields_batched(&[0x08, 0x81, 0x00], &varint_edit).is_err());

        let length_edit = [NestedFieldEdit::new(
            &[1],
            true,
            NestedFieldReplacement::LengthDelimited(Some(b"y")),
        )];
        assert!(patch_nested_fields_batched(&[0x0a, 0x81, 0x00, b'x'], &length_edit).is_err());

        let nested_edit = [NestedFieldEdit::new(
            &[2, 1],
            true,
            NestedFieldReplacement::Varint(Some(2)),
        )];
        assert!(
            patch_nested_fields_batched(&[0x92, 0x00, 0x02, 0x08, 0x01], &nested_edit).is_err()
        );
        assert!(
            patch_nested_fields_batched(&[0x12, 0x82, 0x00, 0x08, 0x01], &nested_edit).is_err()
        );
    }

    #[test]
    fn batched_nested_limit_boundaries_are_exact() {
        let source = varint_field(1, 1);
        let replace = [NestedFieldEdit::new(
            &[1],
            true,
            NestedFieldReplacement::Varint(Some(300)),
        )];
        let expected = varint_field(1, 300);
        let work = source.len() + expected.len();
        let exact = WireLimits::default()
            .with_input_bytes(source.len())
            .unwrap()
            .with_output_bytes(expected.len())
            .unwrap()
            .with_fields(1)
            .unwrap()
            .with_rewrite_work(work)
            .unwrap();
        assert_eq!(
            patch_nested_fields_batched_with_limits(&source, &replace, exact).unwrap(),
            expected
        );

        let input_short = WireLimits::default()
            .with_input_bytes(source.len() - 1)
            .unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&source, &replace, input_short),
            Err(Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                ..
            })
        ));
        let output_short = WireLimits::default()
            .with_output_bytes(expected.len() - 1)
            .unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&source, &replace, output_short),
            Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
        let work_short = WireLimits::default().with_rewrite_work(work - 1).unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&source, &replace, work_short),
            Err(Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed,
                limit,
            }) if observed == work && limit == work - 1
        ));

        let append = [NestedFieldEdit::new(
            &[2],
            false,
            NestedFieldReplacement::Varint(Some(3)),
        )];
        let one_field = WireLimits::default().with_fields(1).unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&source, &append, one_field),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));

        let validators = [
            NestedFieldEdit::new(&[2], false, NestedFieldReplacement::Varint(None)),
            NestedFieldEdit::new(&[3], false, NestedFieldReplacement::Fixed32(None)),
        ];
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&[], &validators, one_field),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));
    }

    #[test]
    fn batched_nested_depth_limit_applies_to_selected_paths() {
        let leaf = varint_field(1, 1);
        let middle = length_delimited_field(2, &leaf);
        let source = length_delimited_field(3, &middle);
        let edit = [NestedFieldEdit::new(
            &[3, 2, 1],
            true,
            NestedFieldReplacement::Varint(Some(2)),
        )];
        let too_shallow = WireLimits::default().with_nesting(1).unwrap();
        assert!(matches!(
            patch_nested_fields_batched_with_limits(&source, &edit, too_shallow),
            Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: 2,
                limit: 1,
            })
        ));
        let exact = WireLimits::default().with_nesting(2).unwrap();
        assert!(patch_nested_fields_batched_with_limits(&source, &edit, exact).is_ok());
    }

    #[test]
    fn repeated_rewrite_preserves_slots_and_append_remove_is_exact() {
        let mut original = varint_field(90, 1);
        append_length_delimited_field(&mut original, 2, b"first").unwrap();
        original.extend(varint_field(91, 2));
        append_length_delimited_field(&mut original, 2, b"second").unwrap();
        original.extend(varint_field(92, 3));

        let reordered = rewrite_repeated_length_delimited_fields(
            &original,
            2,
            &[b"second".to_vec(), b"first".to_vec()],
        )
        .unwrap();
        assert_eq!(
            repeated_length_delimited_payloads(&reordered, 2).unwrap(),
            [b"second".as_slice(), b"first".as_slice()]
        );
        let restored = rewrite_repeated_length_delimited_fields(
            &reordered,
            2,
            &[b"first".to_vec(), b"second".to_vec()],
        )
        .unwrap();
        assert_eq!(restored, original);

        let appended = rewrite_repeated_length_delimited_fields(
            &original,
            2,
            &[b"first".to_vec(), b"second".to_vec(), b"third".to_vec()],
        )
        .unwrap();
        let removed = rewrite_repeated_length_delimited_fields(
            &appended,
            2,
            &[b"first".to_vec(), b"second".to_vec()],
        )
        .unwrap();
        assert_eq!(removed, original);
    }

    #[test]
    fn repeated_noops_preserve_noncanonical_framing_and_nested_absence() {
        let length_delimited = [0x12, 0x81, 0x00, b'x'];
        assert_eq!(
            rewrite_repeated_length_delimited_fields(&length_delimited, 2, &[b"x".to_vec()])
                .unwrap(),
            length_delimited
        );

        let varint = [0x10, 0x80, 0x00];
        assert_eq!(
            rewrite_repeated_varint_fields(&varint, 2, &[0]).unwrap(),
            varint
        );

        let nested = [0x1a, 0x82, 0x00, 0x08, 0x01];
        let calls = Cell::new(0usize);
        let unchanged = transform_length_delimited_fields_at_path(&nested, &[3, 2], |_| {
            calls.set(calls.get() + 1);
            Ok::<Vec<u8>, Error>(Vec::new())
        })
        .unwrap();
        assert_eq!(unchanged, nested);
        assert_eq!(calls.get(), 0);
    }

    #[test]
    fn repeated_varint_rewrite_preserves_slots_and_append_remove_is_exact() {
        let mut original = varint_field(90, 1);
        original.extend(varint_field(2, 10));
        original.extend(varint_field(91, 2));
        original.extend(varint_field(2, 20));
        original.extend(varint_field(92, 3));

        let changed = rewrite_repeated_varint_fields(&original, 2, &[11, 21, 31]).unwrap();
        assert_eq!(repeated_varint_values(&changed, 2).unwrap(), [11, 21, 31]);
        let restored = rewrite_repeated_varint_fields(&changed, 2, &[10, 20]).unwrap();
        assert_eq!(restored, original);

        let mut packed = original.clone();
        append_length_delimited_field(&mut packed, 2, &[1, 2]).unwrap();
        assert!(rewrite_repeated_varint_fields(&packed, 2, &[10, 20]).is_err());
    }

    #[test]
    fn repeated_fixed64_rewrite_preserves_slots_and_append_remove_is_exact() {
        let mut original = varint_field(90, 1);
        append_scalar_field(&mut original, 2, 1, &10_u64.to_le_bytes()).unwrap();
        original.extend(varint_field(91, 2));
        append_scalar_field(&mut original, 2, 1, &20_u64.to_le_bytes()).unwrap();
        original.extend(varint_field(92, 3));

        let changed = rewrite_repeated_fixed64_fields(&original, 2, &[11, 21, 31]).unwrap();
        assert_eq!(repeated_fixed64_values(&changed, 2).unwrap(), [11, 21, 31]);
        let restored = rewrite_repeated_fixed64_fields(&changed, 2, &[10, 20]).unwrap();
        assert_eq!(restored, original);

        let mut wrong_wire = original.clone();
        wrong_wire.extend(varint_field(2, 30));
        assert!(rewrite_repeated_fixed64_fields(&wrong_wire, 2, &[10, 20]).is_err());
    }

    #[test]
    fn repeated_fixed32_rewrite_preserves_slots_and_append_remove_is_exact() {
        let mut original = varint_field(90, 1);
        append_scalar_field(&mut original, 2, 5, &10_u32.to_le_bytes()).unwrap();
        original.extend(varint_field(91, 2));
        append_scalar_field(&mut original, 2, 5, &20_u32.to_le_bytes()).unwrap();
        original.extend(varint_field(92, 3));

        let changed = rewrite_repeated_fixed32_fields(&original, 2, &[11, 21, 31]).unwrap();
        assert_eq!(repeated_fixed32_values(&changed, 2).unwrap(), [11, 21, 31]);
        let restored = rewrite_repeated_fixed32_fields(&changed, 2, &[10, 20]).unwrap();
        assert_eq!(restored, original);

        let mut wrong_wire = original.clone();
        wrong_wire.extend(varint_field(2, 30));
        assert!(rewrite_repeated_fixed32_fields(&wrong_wire, 2, &[10, 20]).is_err());
    }

    #[test]
    fn repeated_nested_transform_preserves_unknown_ancestors_and_siblings() {
        let mut first_leaf = varint_field(1, 10);
        first_leaf.extend(varint_field(90, 900));
        let mut second_leaf = varint_field(1, 20);
        second_leaf.extend(varint_field(91, 901));
        let mut nested = varint_field(80, 800);
        append_length_delimited_field(&mut nested, 2, &first_leaf).unwrap();
        nested.extend(varint_field(81, 801));
        append_length_delimited_field(&mut nested, 2, &second_leaf).unwrap();
        let mut original = varint_field(70, 700);
        append_length_delimited_field(&mut original, 3, &nested).unwrap();
        original.extend(varint_field(71, 701));

        let changed = transform_length_delimited_fields_at_path(&original, &[3, 2], |leaf| {
            let fields = parse_wire_fields(leaf)?;
            let identifier = fields
                .iter()
                .find(|field| field.number == 1)
                .ok_or_else(|| Error::InvalidFormat("missing identifier".to_owned()))?;
            let (value, _) = crate::varint::decode_varint_from_bytes(
                &leaf[identifier.payload_start..identifier.end],
            )
            .map_err(|error| Error::InvalidFormat(format!("invalid identifier: {error}")))?;
            patch_varint_field(leaf, 1, true, Some(value + 100))
        })
        .unwrap();
        assert_ne!(changed, original);

        let restored = transform_length_delimited_fields_at_path(&changed, &[3, 2], |leaf| {
            let fields = parse_wire_fields(leaf)?;
            let identifier = fields
                .iter()
                .find(|field| field.number == 1)
                .ok_or_else(|| Error::InvalidFormat("missing identifier".to_owned()))?;
            let (value, _) = crate::varint::decode_varint_from_bytes(
                &leaf[identifier.payload_start..identifier.end],
            )
            .map_err(|error| Error::InvalidFormat(format!("invalid identifier: {error}")))?;
            patch_varint_field(leaf, 1, true, Some(value - 100))
        })
        .unwrap();
        assert_eq!(restored, original);
        assert!(
            transform_length_delimited_fields_at_path(&original, &[], |_| {
                Ok::<Vec<u8>, Error>(Vec::new())
            })
            .is_err()
        );
    }

    #[test]
    fn scalar_patches_reject_duplicates_wrong_types_and_truncation() {
        let mut duplicate = varint_field(39, 0);
        duplicate.extend(varint_field(39, 1));
        assert!(patch_varint_field(&duplicate, 39, true, Some(0)).is_err());

        let wrong_type = varint_field(30, 1);
        assert!(patch_fixed32_field(&wrong_type, 30, true, Some(0)).is_err());
        assert!(patch_nested_varint_field(&wrong_type, &[], true, Some(0)).is_err());
        assert!(patch_varint_field(&[0x98, 0x02, 0x80], 35, true, Some(1)).is_err());
    }

    #[test]
    fn parser_enforces_input_and_field_budgets() {
        let mut data = varint_field(1, 1);
        data.extend(varint_field(2, 2));
        let limits = WireLimits::default().with_fields(1).unwrap();
        assert!(matches!(
            parse_wire_fields_with_limits(&data, limits),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));

        let limits = WireLimits::default().with_input_bytes(1).unwrap();
        assert!(matches!(
            parse_wire_fields_with_limits(&data, limits),
            Err(Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: 4,
                limit: 1,
            })
        ));
    }

    #[test]
    fn length_delimited_append_checks_field_numbers_and_output_budget() {
        let mut output = Vec::new();
        assert!(append_length_delimited_field(&mut output, 0, b"x").is_err());
        assert!(append_length_delimited_field(&mut output, 0x2000_0000, b"x").is_err());

        let limits = WireLimits::default().with_output_bytes(2).unwrap();
        assert!(matches!(
            append_length_delimited_field_with_limits(&mut output, 1, b"x", limits),
            Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn wire_field_offsets_are_exposed_through_checked_slices() {
        let data = varint_field(7, 42);
        let field = parse_wire_fields(&data).unwrap().pop().unwrap();
        assert_eq!(field.number(), 7);
        assert_eq!(field.wire_type(), 0);
        assert_eq!(field.start(), 0);
        assert_eq!(field.key_end(), 1);
        assert_eq!(field.payload_start(), 1);
        assert_eq!(field.end(), data.len());
        assert_eq!(field.raw(&data).unwrap(), data);
        assert_eq!(field.key(&data).unwrap(), &data[..1]);
        assert_eq!(field.payload(&data).unwrap(), &data[1..]);
        assert!(field.raw(&[]).is_err());
    }

    #[test]
    fn wire_field_helpers_validate_canonical_framing_without_tightening_parser() {
        let overlong_key = [0x88, 0x00, 0x07];
        let key_field = parse_wire_fields(&overlong_key).unwrap().pop().unwrap();
        assert_eq!(key_field.checked_payload(&overlong_key).unwrap(), [0x07]);
        assert!(key_field.validate_canonical_key(&overlong_key).is_err());
        assert!(
            key_field
                .validate_canonical_key(&overlong_key[..2])
                .is_err()
        );

        let canonical_key = varint_field(1, 7);
        let canonical_key_field = parse_wire_fields(&canonical_key).unwrap().pop().unwrap();
        canonical_key_field
            .validate_canonical_key(&canonical_key)
            .unwrap();
        assert!(
            canonical_key_field
                .validate_canonical_length(&canonical_key)
                .is_err()
        );

        let fixed32 = [0x1d, 0x78, 0x56, 0x34, 0x12];
        let fixed32_field = parse_wire_fields(&fixed32).unwrap().pop().unwrap();
        assert_eq!(
            fixed32_field.checked_payload(&fixed32).unwrap(),
            &fixed32[1..]
        );

        let fixed64 = [0x09, 0xef, 0xcd, 0xab, 0x89, 0x67, 0x45, 0x23, 0x01];
        let fixed64_field = parse_wire_fields(&fixed64).unwrap().pop().unwrap();
        assert_eq!(
            fixed64_field.checked_payload(&fixed64).unwrap(),
            &fixed64[1..]
        );

        let overlong_length = [0x0a, 0x81, 0x00, b'x'];
        let length_field = parse_wire_fields(&overlong_length).unwrap().pop().unwrap();
        assert_eq!(
            length_field.checked_payload(&overlong_length).unwrap(),
            b"x"
        );
        length_field
            .validate_canonical_key(&overlong_length)
            .unwrap();
        assert!(
            length_field
                .validate_canonical_length(&overlong_length)
                .is_err()
        );

        let canonical_length = [0x0a, 0x01, b'x'];
        let canonical_length_field = parse_wire_fields(&canonical_length).unwrap().pop().unwrap();
        canonical_length_field
            .validate_canonical_key(&canonical_length)
            .unwrap();
        canonical_length_field
            .validate_canonical_length(&canonical_length)
            .unwrap();

        assert!(length_field.checked_payload(&overlong_length[..3]).is_err());
        assert!(canonical_key_field.checked_payload(&[]).is_err());
    }

    #[test]
    fn source_bound_wire_view_borrows_one_source_and_exposes_field_views() {
        let mut source = varint_field(1, 7);
        append_length_delimited_field(&mut source, 2, b"payload").unwrap();
        let other_source = [0x08, 0x63, 0x12, 0x01, b'x'];
        let view = parse_wire_view(&source).unwrap();
        let fields: Vec<_> = view.fields().collect();

        assert_eq!(view.source().as_ptr(), source.as_ptr());
        assert_eq!(view.as_bytes(), source);
        assert_eq!(view.len(), 2);
        assert!(!view.is_empty());
        assert!(view.get(2).is_none());

        assert_eq!(fields[0].raw(), &source[..2]);
        assert_eq!(fields[0].key(), &source[..1]);
        assert_eq!(fields[0].payload(), &source[1..2]);
        assert_eq!(fields[0].number(), 1);
        assert_eq!(fields[0].wire_type(), 0);
        assert_eq!(fields[1].raw(), &source[2..]);
        assert_eq!(fields[1].key(), &source[2..3]);
        assert_eq!(fields[1].payload(), b"payload");

        // There is no source argument on a field view. Its payload remains
        // tied to `source` even when another same-shaped message is present.
        assert_ne!(other_source, fields[0].raw());
        assert_eq!(fields[0].canonical_payload().unwrap(), [0x07]);
        assert_eq!(fields[1].canonical_payload().unwrap(), b"payload");
    }

    #[test]
    fn source_bound_wire_views_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<WireView<'static>>();
        assert_send_sync::<WireFieldView<'static>>();
    }

    #[test]
    fn source_bound_wire_view_helpers_reject_noncanonical_framing() {
        let overlong_key = [0x88, 0x00, 0x07];
        let key_field = parse_wire_view(&overlong_key).unwrap().get(0).unwrap();
        assert_eq!(key_field.raw(), overlong_key);
        assert_eq!(key_field.key(), &overlong_key[..2]);
        assert_eq!(key_field.payload(), &overlong_key[2..]);
        assert!(key_field.validate_canonical_key().is_err());
        assert!(key_field.canonical_key().is_err());

        let overlong_length = [0x0a, 0x81, 0x00, b'x'];
        let length_field = parse_wire_view(&overlong_length).unwrap().get(0).unwrap();
        length_field.validate_canonical_key().unwrap();
        assert!(length_field.validate_canonical_length().is_err());
        assert!(length_field.validate_canonical_framing().is_err());
        assert!(length_field.canonical_payload().is_err());

        let canonical = [0x0a, 0x01, b'x'];
        let canonical_field = parse_wire_view(&canonical).unwrap().get(0).unwrap();
        assert_eq!(canonical_field.canonical_key().unwrap(), &canonical[..1]);
        assert_eq!(canonical_field.canonical_payload().unwrap(), b"x");
        assert!(
            parse_wire_view(&[0x08, 0x00])
                .unwrap()
                .get(0)
                .unwrap()
                .validate_canonical_length()
                .is_err()
        );
    }

    #[test]
    fn source_bound_wire_view_rejects_truncation_and_limits_before_publication() {
        for truncated in [[0x08].as_slice(), &[0x09, 0x01], &[0x0a, 0x02, b'x']] {
            assert!(parse_wire_view(truncated).is_err());
        }

        let mut data = varint_field(1, 1);
        data.extend(varint_field(2, 2));
        let one_field = WireLimits::default().with_fields(1).unwrap();
        assert!(matches!(
            parse_wire_view_with_limits(&data, one_field),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));

        let one_byte = WireLimits::default().with_input_bytes(1).unwrap();
        assert!(matches!(
            parse_wire_view_with_limits(&data, one_byte),
            Err(Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: 4,
                limit: 1,
            })
        ));
    }

    #[test]
    fn source_bound_wire_view_retains_compact_spans_with_fallible_growth() {
        assert!(size_of::<WireSpan>() <= 32);
        let data = [0x08, 0x00, 0x10, 0x01];
        let limits = WireLimits::default().with_fields(1).unwrap();
        assert!(matches!(
            WireView::parse_with_limits(&data, limits),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));
    }

    #[test]
    fn append_rejects_invalid_fields_without_mutating_output() {
        let mut output = Vec::new();
        assert!(append_length_delimited_field(&mut output, 0, b"x").is_err());
        assert!(output.is_empty());
        assert!(append_length_delimited_field(&mut output, 0x2000_0000, b"x").is_err());
        assert!(output.is_empty());

        append_length_delimited_field(&mut output, 0x1fff_ffff, b"x").unwrap();
        assert_eq!(parse_wire_fields(&output).unwrap()[0].number(), 0x1fff_ffff);
    }

    #[test]
    fn output_field_budget_applies_to_append_and_overlay() {
        let at_limit = [0x08, 0x00].repeat(WireLimits::default().max_fields());
        let mut appended = at_limit.clone();
        assert!(matches!(
            append_length_delimited_field(&mut appended, 1, b"x"),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                ..
            })
        ));
        assert_eq!(appended, at_limit);

        assert!(matches!(
            overlay_singular_wire_fields(&at_limit, &[0x12, 0x01, b'x']),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                ..
            })
        ));
    }

    #[test]
    fn overlay_validates_an_empty_overlay_and_preserves_noncanonical_noops() {
        assert!(overlay_singular_wire_fields(&[0x80], &[]).is_err());

        let data = [0x0a, 0x82, 0x00, b'a', b'b'];
        assert_eq!(
            patch_length_delimited_field(&data, 1, true, Some(b"ab")).unwrap(),
            data
        );
        assert_eq!(
            patch_length_delimited_field(&data, 1, true, Some(b"cd")).unwrap(),
            [0x0a, 0x82, 0x00, b'c', b'd']
        );
    }

    fn descend_test_schema(path: &[u32], field: WireFieldView<'_>) -> bool {
        field.wire_type() == 2
            && ((path.is_empty() && field.number() == 2) || (path == [2] && field.number() == 4))
    }

    #[test]
    fn wire_tree_preflight_visits_selected_messages_with_shared_totals() {
        let leaf = varint_field(7, 70);
        let mut child = varint_field(3, 30);
        append_length_delimited_field(&mut child, 4, &leaf).unwrap();
        let mut source = varint_field(1, 10);
        append_length_delimited_field(&mut source, 2, &child).unwrap();
        append_length_delimited_field(&mut source, 9, &[0x80]).unwrap();

        let mut visited = Vec::new();
        let report = preflight_wire_tree(&source, |visit| {
            let field = visit.field();
            visited.push((visit.path().to_vec(), field.number(), field.wire_type()));
            Ok(if descend_test_schema(visit.path(), field) {
                WireDescent::Descend
            } else {
                WireDescent::Skip
            })
        })
        .unwrap();

        assert_eq!(
            visited,
            vec![
                (vec![], 1, 0),
                (vec![], 2, 2),
                (vec![2], 3, 0),
                (vec![2], 4, 2),
                (vec![2, 4], 7, 0),
                (vec![], 9, 2),
            ]
        );
        assert_eq!(report.fields(), 6);
        assert_eq!(report.messages(), 3);
        assert_eq!(report.max_depth(), 2);
        assert_eq!(
            report.scanned_bytes(),
            source.len() + child.len() + leaf.len()
        );
    }

    #[test]
    fn wire_tree_preflight_shares_field_budget_across_siblings() {
        let child = varint_field(1, 1);
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 2, &child).unwrap();
        append_length_delimited_field(&mut source, 2, &child).unwrap();
        let limits = WireLimits::default().with_fields(3).unwrap();

        assert!(matches!(
            preflight_wire_tree_with_limits(&source, limits, |visit| Ok(if visit.depth() == 0 {
                WireDescent::Descend
            } else {
                WireDescent::Skip
            })),
            Err(Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 4,
                limit: 3,
            })
        ));
    }

    #[test]
    fn wire_tree_preflight_shares_aggregate_scan_byte_budget() {
        let child = varint_field(1, 1);
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 2, &child).unwrap();
        let aggregate = source.len() + child.len();
        let limits = WireLimits::default()
            .with_input_bytes(aggregate - 1)
            .unwrap();

        assert!(matches!(
            preflight_wire_tree_with_limits(&source, limits, |visit| Ok(
                if visit.depth() == 0 {
                    WireDescent::Descend
                } else {
                    WireDescent::Skip
                }
            )),
            Err(Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit,
            }) if observed == aggregate && limit == aggregate - 1
        ));
    }

    #[test]
    fn wire_tree_preflight_rejects_adversarial_depth_before_descent() {
        let mut source = varint_field(9, 1);
        for _ in 0..4 {
            let mut parent = Vec::new();
            append_length_delimited_field(&mut parent, 1, &source).unwrap();
            source = parent;
        }
        let limits = WireLimits::default().with_nesting(3).unwrap();

        assert!(matches!(
            preflight_wire_tree_with_limits(&source, limits, |visit| Ok(
                if visit.field().wire_type() == 2 {
                    WireDescent::Descend
                } else {
                    WireDescent::Skip
                }
            )),
            Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: 4,
                limit: 3,
            })
        ));
    }

    #[test]
    fn wire_tree_preflight_validates_only_selected_deferred_messages() {
        let mut source = Vec::new();
        append_length_delimited_field(&mut source, 2, &[0x08]).unwrap();

        let skipped = preflight_wire_tree(&source, |_| Ok(WireDescent::Skip)).unwrap();
        assert_eq!(skipped.fields(), 1);
        assert!(preflight_wire_tree(&source, |_| Ok(WireDescent::Descend)).is_err());

        let scalar = varint_field(1, 7);
        assert!(matches!(
            preflight_wire_tree(&scalar, |_| Ok(WireDescent::Descend)),
            Err(Error::InvalidFormat(message))
                if message.contains("not length-delimited")
        ));
    }

    #[test]
    fn wire_tree_preflight_matches_recursive_legacy_field_parsing() {
        fn collect_legacy(
            source: &[u8],
            path: &mut Vec<u32>,
            collected: &mut Vec<(Vec<u32>, u32, u8, Vec<u8>)>,
        ) {
            for field in parse_wire_fields(source).unwrap() {
                let field_view = WireView::parse(field.raw(source).unwrap())
                    .unwrap()
                    .get(0)
                    .unwrap();
                collected.push((
                    path.clone(),
                    field.number(),
                    field.wire_type(),
                    field.raw(source).unwrap().to_vec(),
                ));
                if descend_test_schema(path, field_view) {
                    path.push(field.number());
                    collect_legacy(field.payload(source).unwrap(), path, collected);
                    path.pop();
                }
            }
        }

        let leaf = varint_field(8, 800);
        let mut child = varint_field(3, 300);
        append_length_delimited_field(&mut child, 4, &leaf).unwrap();
        child.extend([0x95, 0x00, 1, 2, 3, 4]);
        let mut source = vec![0x88, 0x00, 0x01];
        append_length_delimited_field(&mut source, 2, &child).unwrap();
        source.extend([0x29, 1, 2, 3, 4, 5, 6, 7, 8]);

        let mut legacy = Vec::new();
        collect_legacy(&source, &mut Vec::new(), &mut legacy);
        let mut preflight = Vec::new();
        let report = preflight_wire_tree(&source, |visit| {
            let field = visit.field();
            preflight.push((
                visit.path().to_vec(),
                field.number(),
                field.wire_type(),
                field.raw().to_vec(),
            ));
            Ok(if descend_test_schema(visit.path(), field) {
                WireDescent::Descend
            } else {
                WireDescent::Skip
            })
        })
        .unwrap();

        assert_eq!(preflight, legacy);
        assert_eq!(report.fields(), legacy.len());

        for malformed in [
            vec![0x08],
            vec![0x0a, 0x02, 0x08],
            vec![0x0b, 0x0c],
            vec![0x0e],
        ] {
            assert_eq!(
                preflight_wire_tree(&malformed, |_| Ok(WireDescent::Skip)).is_ok(),
                parse_wire_fields(&malformed).is_ok()
            );
        }
    }

    #[test]
    fn nested_mutations_enforce_depth_budget() {
        let path = vec![1; WireLimits::default().max_nesting() + 1];
        assert!(matches!(
            transform_length_delimited_fields_at_path(&[], &path, |_| Ok(Vec::new())),
            Err(Error::LimitExceeded {
                kind: LimitKind::Nesting,
                ..
            })
        ));
        assert!(matches!(
            WireLimits::default().with_nesting(WireLimits::MAX_NESTING + 1),
            Err(Error::InvalidLimit { .. })
        ));
    }
}
