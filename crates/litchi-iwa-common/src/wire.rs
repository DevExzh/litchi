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
        let spans = parse_wire_items_with_limits(source, limits, "wire view spans", Ok)?;
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

/// Compact private field metadata retained by [`WireView`].
///
/// The configured input ceiling is below `u32::MAX`, so four-byte offsets are
/// sufficient while keeping the parsed representation free of per-field slice
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
    parse_wire_items_with_limits(data, limits, "wire fields", |span| {
        Ok(WireField {
            number: span.number,
            wire_type: span.wire_type,
            start: span.start as usize,
            key_end: span.key_end as usize,
            payload_start: span.payload_start as usize,
            end: span.end as usize,
        })
    })
}

fn parse_wire_items_with_limits<T, F>(
    data: &[u8],
    limits: WireLimits,
    resource: &'static str,
    mut map: F,
) -> Result<Vec<T>>
where
    F: FnMut(WireSpan) -> Result<T>,
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
        let span = parse_wire_span(data, offset)?;
        offset = span.end as usize;
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
        items.push(map(span)?);
    }
    Ok(items)
}

fn parse_wire_span(data: &[u8], start: usize) -> Result<WireSpan> {
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
    Ok(WireSpan {
        number,
        wire_type,
        start: u32::try_from(start).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        key_end: u32::try_from(key_end).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        payload_start: u32::try_from(payload_start).map_err(|_conversion| {
            Error::InvalidFormat("protobuf offset exceeds u32".to_owned())
        })?,
        end: u32::try_from(end).map_err(|_conversion| {
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
    fn wire_field_offsets_are_exposed_through_checked_views() {
        let data = varint_field(7, 42);
        let field = parse_wire_fields(&data).unwrap().pop().unwrap();
        assert_eq!(field.number(), 7);
        assert_eq!(field.wire_type(), 0);
        assert_eq!(field.start(), 0);
        assert_eq!(field.key_end(), 1);
        assert_eq!(field.payload_start(), 1);
        assert_eq!(field.end(), data.len());
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
