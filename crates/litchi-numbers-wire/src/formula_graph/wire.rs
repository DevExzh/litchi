//! Private bounded wire primitives for the formula-graph rewrites.
//!
//! Every helper takes the caller's residual [`WireLimits`] explicitly: the
//! formula-graph ledger narrows the field and work budgets before each call,
//! so a nested or repeated payload can never reset the aggregate ceilings.
//! The helpers build on the public, limit-taking `litchi-iwa-common` wire
//! scanners and never normalise untouched bytes: keys, length prefixes and
//! the position of every retained field are copied verbatim.

use std::cell::Cell;

use litchi_iwa_common::varint::{
    decode_varint_from_bytes, encode_varint_into, encode_varint_to_buffer, encoded_len,
};
use litchi_iwa_common::wire::{WireDescent, WireField, parse_wire_fields, preflight_wire_tree};
use litchi_iwa_common::{Error, LimitKind, Result, WireLimits};

/// Count the root fields of `source` without allocating a field table.
pub(super) fn root_field_count(source: &[u8], limits: WireLimits) -> Result<usize> {
    preflight_wire_tree(source, limits, |_visit| Ok(WireDescent::Skip))
        .map(|report| report.fields())
}

/// Parse root fields; the shared parser enforces `limits.max_fields()`
/// while the table grows.
pub(super) fn parse_fields_within(source: &[u8], limits: WireLimits) -> Result<Vec<WireField>> {
    parse_wire_fields(source, limits)
}

/// Patch one optional singular length-delimited field.
pub(super) fn patch_length_delimited_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<&[u8]>,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data, limits)?;
    let selected = singular(&fields, field_number, expected_present)?;
    let Some(field) = selected else {
        let Some(payload) = replacement else {
            return clone_output(data, limits);
        };
        ensure_field_count(fields.len().saturating_add(1), limits)?;
        let mut output = clone_output(data, limits)?;
        append_length_delimited_unchecked(&mut output, field_number, payload, limits)?;
        return Ok(output);
    };
    require_wire_type(field, 2)?;
    let old_payload = field.payload(data)?;
    let Some(payload) = replacement else {
        return without_field(data, field, limits);
    };
    if old_payload == payload {
        return clone_output(data, limits);
    }
    let replacement_length = u64::try_from(payload.len()).map_err(|_| overflow())?;
    let preserve_prefix = old_payload.len() == payload.len();
    let length_size = if preserve_prefix {
        field.payload_start() - field.key_end()
    } else {
        encoded_len(replacement_length)
    };
    let output_len = data
        .len()
        .checked_sub(field.end() - field.start())
        .and_then(|value| value.checked_add(field.key_end() - field.start()))
        .and_then(|value| value.checked_add(length_size))
        .and_then(|value| value.checked_add(payload.len()))
        .ok_or_else(overflow)?;
    let mut output = output_with_capacity(output_len, limits)?;
    output.extend_from_slice(&data[..field.key_end()]);
    if preserve_prefix {
        output.extend_from_slice(&data[field.key_end()..field.payload_start()]);
    } else {
        encode_varint_into(&mut output, replacement_length);
    }
    output.extend_from_slice(payload);
    output.extend_from_slice(&data[field.end()..]);
    Ok(output)
}

/// Patch one optional singular varint field.
pub(super) fn patch_varint_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data, limits)?;
    let selected = singular(&fields, field_number, expected_present)?;
    let Some(field) = selected else {
        let Some(value) = replacement else {
            return clone_output(data, limits);
        };
        ensure_field_count(fields.len().saturating_add(1), limits)?;
        let mut output = clone_output(data, limits)?;
        append_varint_unchecked(&mut output, field_number, value, limits)?;
        return Ok(output);
    };
    require_wire_type(field, 0)?;
    let Some(value) = replacement else {
        return without_field(data, field, limits);
    };
    let mut bytes = [0u8; litchi_iwa_common::MAX_BYTES];
    let encoded = encode_varint_to_buffer(value, &mut bytes);
    let output_len = data
        .len()
        .checked_sub(field.end() - field.start())
        .and_then(|length| length.checked_add(field.key_end() - field.start()))
        .and_then(|length| length.checked_add(encoded.len()))
        .ok_or_else(overflow)?;
    let mut output = output_with_capacity(output_len, limits)?;
    output.extend_from_slice(&data[..field.key_end()]);
    output.extend_from_slice(encoded);
    output.extend_from_slice(&data[field.end()..]);
    Ok(output)
}

/// Every occurrence of an unpacked repeated varint field.
pub(super) fn repeated_varint_values(
    data: &[u8],
    field_number: u32,
    limits: WireLimits,
) -> Result<Vec<u64>> {
    let fields = parse_wire_fields(data, limits)?;
    let matches = fields
        .iter()
        .filter(|field| field.number() == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve_exact(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "repeated varint values",
            amount: matches,
        })?;
    for field in fields {
        if field.number() != field_number {
            continue;
        }
        require_wire_type(field, 0)?;
        let encoded = &data[field.key_end()..field.end()];
        let (value, length) = decode_varint_from_bytes(encoded)
            .map_err(|error| Error::InvalidFormat(format!("invalid protobuf varint: {error}")))?;
        if length != encoded.len() {
            return Err(Error::InvalidFormat(format!(
                "protobuf field {field_number} has trailing varint bytes"
            )));
        }
        values.push(value);
    }
    Ok(values)
}

/// Replace an unpacked repeated varint field, keeping unrelated bytes and the
/// key bytes and positions of retained slots.
pub(super) fn rewrite_repeated_varint_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data, limits)?;
    ensure_rewrite_work(replacements.len(), limits)?;
    let matches = fields
        .iter()
        .copied()
        .filter(|field| field.number() == field_number)
        .collect::<Vec<_>>();
    for field in &matches {
        require_wire_type(*field, 0)?;
    }
    let retained = fields
        .len()
        .checked_sub(matches.len())
        .and_then(|count| count.checked_add(replacements.len()))
        .ok_or_else(overflow)?;
    ensure_field_count(retained, limits)?;
    if matches.is_empty() {
        let mut output = clone_output(data, limits)?;
        for &replacement in replacements {
            append_varint_unchecked(&mut output, field_number, replacement, limits)?;
        }
        return Ok(output);
    }
    if replacements.len() == matches.len()
        && matches
            .iter()
            .zip(replacements)
            .all(|(field, replacement)| {
                is_same_varint(*replacement, &data[field.key_end()..field.end()])
            })
    {
        return clone_output(data, limits);
    }
    let removed = matches
        .iter()
        .try_fold(0usize, |total, field| {
            total.checked_add(field.end() - field.start())
        })
        .ok_or_else(overflow)?;
    let mut capacity = data.len().checked_sub(removed).ok_or_else(overflow)?;
    for (index, replacement) in replacements.iter().enumerate() {
        let key_length = match matches.get(index) {
            Some(field) => field.key_end() - field.start(),
            None => key_size(field_number, 0)?,
        };
        capacity = capacity
            .checked_add(key_length)
            .and_then(|value| value.checked_add(encoded_len(*replacement)))
            .ok_or_else(overflow)?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start()]);
        if let Some(&replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start()..field.key_end()]);
            encode_varint_into(&mut output, replacement);
        }
        copied = field.end();
        if index + 1 == matches.len() {
            for &replacement in &replacements[matches.len().min(replacements.len())..] {
                append_varint_unchecked(&mut output, field_number, replacement, limits)?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Transform every length-delimited leaf at a field path; each component may
/// occur zero, one or many times. One work counter is shared across the
/// whole path so repeated parents cannot reset the rewrite budget.
pub(super) fn transform_length_delimited_fields_at_path<F, E>(
    data: &[u8],
    path: &[u32],
    mut transform: F,
    limits: WireLimits,
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
        let Some((&field_number, remainder)) = path.split_first() else {
            return Err(
                Error::InvalidFormat("protobuf field path cannot be empty".to_owned()).into(),
            );
        };
        if remainder.is_empty() {
            return transform_repeated(data, field_number, transform, limits, work);
        }
        transform_repeated(
            data,
            field_number,
            &mut |nested: &[u8]| visit(nested, remainder, transform, limits, work),
            limits,
            work,
        )
    }

    ensure_nesting(path.len(), limits)?;
    let work = Cell::new(0usize);
    visit(data, path, &mut transform, limits, &work)
}

fn transform_repeated<E>(
    data: &[u8],
    field_number: u32,
    transform: &mut dyn FnMut(&[u8]) -> std::result::Result<Vec<u8>, E>,
    limits: WireLimits,
    work: &Cell<usize>,
) -> std::result::Result<Vec<u8>, E>
where
    E: From<Error>,
{
    let fields = parse_wire_fields(data, limits)?;
    let matches = fields
        .iter()
        .filter(|field| field.number() == field_number)
        .count();
    let total_work = work.get().checked_add(matches).ok_or_else(overflow)?;
    ensure_rewrite_work(total_work, limits)?;
    work.set(total_work);
    if matches == 0 {
        return Ok(clone_output(data, limits)?);
    }
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(matches)
        .map_err(|_allocation| Error::Allocation {
            resource: "transformed length-delimited payloads",
            amount: matches,
        })?;
    for field in &fields {
        if field.number() != field_number {
            continue;
        }
        require_wire_type(*field, 2)?;
        replacements.push(transform(field.payload(data)?)?);
    }
    // A semantic no-op stays byte-exact, including overlong length prefixes;
    // otherwise every matched slot keeps its key and takes a canonical length
    // prefix, exactly as the shared repeated rewrite does.
    if fields
        .iter()
        .filter(|field| field.number() == field_number)
        .zip(&replacements)
        .all(|(field, replacement)| {
            &data[field.payload_start()..field.end()] == replacement.as_slice()
        })
    {
        return Ok(clone_output(data, limits)?);
    }
    let mut capacity = data.len();
    for (field, replacement) in fields
        .iter()
        .filter(|field| field.number() == field_number)
        .zip(&replacements)
    {
        let new_length = u64::try_from(replacement.len()).map_err(|_| overflow())?;
        capacity = capacity
            .checked_sub(field.end() - field.key_end())
            .and_then(|value| value.checked_add(encoded_len(new_length)))
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(overflow)?;
    }
    let mut output = output_with_capacity(capacity, limits)?;
    let mut copied = 0usize;
    for (field, replacement) in fields
        .iter()
        .filter(|field| field.number() == field_number)
        .zip(&replacements)
    {
        output.extend_from_slice(&data[copied..field.key_end()]);
        let length = u64::try_from(replacement.len()).map_err(|_| overflow())?;
        encode_varint_into(&mut output, length);
        output.extend_from_slice(replacement);
        copied = field.end();
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Append one length-delimited field under an explicit field and output budget.
pub(super) fn append_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    let count = root_field_count(output, limits)?;
    ensure_field_count(count.checked_add(1).ok_or_else(overflow)?, limits)?;
    append_length_delimited_unchecked(output, field_number, payload, limits)
}

/// The root-field count of a buffer that only grows by appends.
///
/// Counting once and tracking the appends keeps a run of `k` appends linear
/// in the buffer instead of rescanning it per append; each append still
/// checks the residual field and output budgets it is given.
#[derive(Debug, Clone, Copy)]
pub(super) struct Appender {
    fields: usize,
}

impl Appender {
    /// Count the root fields `output` already holds.
    pub(super) fn of(output: &[u8], limits: WireLimits) -> Result<Self> {
        Ok(Self {
            fields: root_field_count(output, limits)?,
        })
    }

    /// Append one length-delimited field.
    pub(super) fn length_delimited(
        &mut self,
        output: &mut Vec<u8>,
        field_number: u32,
        payload: &[u8],
        limits: WireLimits,
    ) -> Result<()> {
        let fields = self.fields.checked_add(1).ok_or_else(overflow)?;
        ensure_field_count(fields, limits)?;
        append_length_delimited_unchecked(output, field_number, payload, limits)?;
        self.fields = fields;
        Ok(())
    }
}

/// Append one varint field under an explicit field and output budget.
pub(super) fn append_varint_field(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
    limits: WireLimits,
) -> Result<()> {
    let count = root_field_count(output, limits)?;
    ensure_field_count(count.checked_add(1).ok_or_else(overflow)?, limits)?;
    append_varint_unchecked(output, field_number, value, limits)
}

fn append_length_delimited_unchecked(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    let key = u64::from(field_number) << 3 | 2;
    validate_field_number(field_number)?;
    let length = u64::try_from(payload.len()).map_err(|_| overflow())?;
    let additional = encoded_len(key)
        .checked_add(encoded_len(length))
        .and_then(|value| value.checked_add(payload.len()))
        .ok_or_else(overflow)?;
    reserve_output(output, additional, limits)?;
    encode_varint_into(output, key);
    encode_varint_into(output, length);
    output.extend_from_slice(payload);
    Ok(())
}

fn append_varint_unchecked(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
    limits: WireLimits,
) -> Result<()> {
    validate_field_number(field_number)?;
    let key = u64::from(field_number) << 3;
    let additional = encoded_len(key)
        .checked_add(encoded_len(value))
        .ok_or_else(overflow)?;
    reserve_output(output, additional, limits)?;
    encode_varint_into(output, key);
    encode_varint_into(output, value);
    Ok(())
}

fn singular(
    fields: &[WireField],
    field_number: u32,
    expected_present: bool,
) -> Result<Option<WireField>> {
    let mut selected = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == field_number)
    {
        if selected.is_some() {
            return Err(Error::InvalidFormat(format!(
                "singular protobuf field {field_number} occurs more than once"
            )));
        }
        selected = Some(field);
    }
    if selected.is_some() != expected_present {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} changed during mutation"
        )));
    }
    Ok(selected)
}

fn without_field(data: &[u8], field: WireField, limits: WireLimits) -> Result<Vec<u8>> {
    let mut output = output_with_capacity(data.len() - (field.end() - field.start()), limits)?;
    output.extend_from_slice(&data[..field.start()]);
    output.extend_from_slice(&data[field.end()..]);
    Ok(output)
}

fn require_wire_type(field: WireField, expected: u8) -> Result<()> {
    if field.wire_type() != expected {
        return Err(Error::InvalidFormat(format!(
            "protobuf field {} has wire type {}, expected {expected}",
            field.number(),
            field.wire_type()
        )));
    }
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

fn key_size(field_number: u32, wire_type: u8) -> Result<usize> {
    validate_field_number(field_number)?;
    Ok(encoded_len(
        u64::from(field_number) << 3 | u64::from(wire_type),
    ))
}

fn is_same_varint(value: u64, encoded: &[u8]) -> bool {
    decode_varint_from_bytes(encoded)
        .is_ok_and(|(decoded, length)| decoded == value && length == encoded.len())
}

fn reserve_output(output: &mut Vec<u8>, additional: usize, limits: WireLimits) -> Result<()> {
    let requested = output.len().checked_add(additional).ok_or_else(overflow)?;
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
            resource: "formula graph wire output",
            amount: requested,
        })
}

fn output_with_capacity(capacity: usize, limits: WireLimits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    reserve_output(&mut output, capacity, limits)?;
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

fn overflow() -> Error {
    Error::InvalidFormat("protobuf output size overflow".to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, field: u32, value: u64) {
        append_varint_field(output, field, value, WireLimits::DEFAULT).unwrap();
    }

    #[test]
    fn repeated_varints_keep_order_and_unrelated_fields() {
        let mut source = Vec::new();
        varint(&mut source, 5, 1);
        varint(&mut source, 7, 99);
        varint(&mut source, 5, 2);
        let values = repeated_varint_values(&source, 5, WireLimits::DEFAULT).unwrap();
        assert_eq!(values, [1, 2]);
        let rewritten =
            rewrite_repeated_varint_fields(&source, 5, &[300, 2], WireLimits::DEFAULT).unwrap();
        let mut expected = Vec::new();
        varint(&mut expected, 5, 300);
        varint(&mut expected, 7, 99);
        varint(&mut expected, 5, 2);
        assert_eq!(rewritten, expected);
        assert_eq!(
            rewrite_repeated_varint_fields(&source, 5, &[1, 2], WireLimits::DEFAULT).unwrap(),
            source
        );
    }

    #[test]
    fn nested_transform_preserves_framing_and_counts_work() {
        let mut leaf = Vec::new();
        varint(&mut leaf, 1, 4);
        let mut parent = Vec::new();
        append_length_delimited_field(&mut parent, 2, &leaf, WireLimits::DEFAULT).unwrap();
        append_length_delimited_field(&mut parent, 2, &leaf, WireLimits::DEFAULT).unwrap();
        let mut root = Vec::new();
        append_length_delimited_field(&mut root, 3, &parent, WireLimits::DEFAULT).unwrap();
        varint(&mut root, 9, 1);
        let output = transform_length_delimited_fields_at_path::<_, Error>(
            &root,
            &[3, 2],
            |payload| patch_varint_field(payload, 1, true, Some(400), WireLimits::DEFAULT),
            WireLimits::DEFAULT,
        )
        .unwrap();
        let mut new_leaf = Vec::new();
        varint(&mut new_leaf, 1, 400);
        let mut new_parent = Vec::new();
        append_length_delimited_field(&mut new_parent, 2, &new_leaf, WireLimits::DEFAULT).unwrap();
        append_length_delimited_field(&mut new_parent, 2, &new_leaf, WireLimits::DEFAULT).unwrap();
        let mut expected = Vec::new();
        append_length_delimited_field(&mut expected, 3, &new_parent, WireLimits::DEFAULT).unwrap();
        varint(&mut expected, 9, 1);
        assert_eq!(output, expected);

        let narrow = WireLimits::DEFAULT.with_rewrite_work(2).unwrap();
        let error = transform_length_delimited_fields_at_path::<_, Error>(
            &root,
            &[3, 2],
            |payload| Ok(payload.to_vec()),
            narrow,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed: 3,
                limit: 2
            }
        ));
    }

    #[test]
    fn singular_patches_reject_presence_mismatch_and_respect_output_limits() {
        let mut source = Vec::new();
        varint(&mut source, 2, 1);
        assert!(patch_varint_field(&source, 2, false, Some(3), WireLimits::DEFAULT).is_err());
        assert!(patch_length_delimited_field(&source, 4, true, None, WireLimits::DEFAULT).is_err());
        let narrow = WireLimits::DEFAULT.with_output_bytes(2).unwrap();
        assert!(matches!(
            patch_length_delimited_field(&source, 4, false, Some(&[1, 2]), narrow),
            Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                ..
            })
        ));
        let patched = patch_varint_field(&source, 2, true, Some(300), WireLimits::DEFAULT).unwrap();
        let mut expected = Vec::new();
        varint(&mut expected, 2, 300);
        assert_eq!(patched, expected);
    }
}
