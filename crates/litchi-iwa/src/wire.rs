//! Bounded protobuf wire mutations that retain untouched fields byte-for-byte.

use crate::{Error, Result};

#[derive(Debug, Clone, Copy)]
pub(crate) struct WireField {
    pub(crate) number: u32,
    pub(crate) wire_type: u8,
    pub(crate) start: usize,
    pub(crate) key_end: usize,
    pub(crate) payload_start: usize,
    pub(crate) end: usize,
}

pub(crate) fn parse_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    let mut fields = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let start = offset;
        let (key, key_length) = crate::varint::decode_varint_from_bytes(&data[offset..])
            .map_err(|error| Error::InvalidFormat(format!("invalid protobuf key: {error}")))?;
        offset = offset
            .checked_add(key_length)
            .ok_or_else(|| Error::InvalidFormat("protobuf key offset overflow".to_owned()))?;
        let number = key >> 3;
        if number == 0 || number > 0x1fff_ffff {
            return Err(Error::InvalidFormat(format!(
                "invalid protobuf field number {number}"
            )));
        }
        let wire_type = (key & 7) as u8;
        let key_end = offset;
        let mut payload_start = offset;
        match wire_type {
            0 => {
                let (_, length) = crate::varint::decode_varint_from_bytes(&data[offset..])
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid protobuf varint value: {error}"))
                    })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("protobuf varint offset overflow".to_owned())
                })?;
            },
            1 => {
                offset = offset.checked_add(8).ok_or_else(|| {
                    Error::InvalidFormat("protobuf fixed64 offset overflow".to_owned())
                })?;
            },
            2 => {
                let (length, prefix_length) =
                    crate::varint::decode_varint_from_bytes(&data[offset..]).map_err(|error| {
                        Error::InvalidFormat(format!("invalid protobuf length: {error}"))
                    })?;
                offset = offset.checked_add(prefix_length).ok_or_else(|| {
                    Error::InvalidFormat("protobuf length-prefix overflow".to_owned())
                })?;
                payload_start = offset;
                let length = usize::try_from(length).map_err(|_| {
                    Error::InvalidFormat("protobuf field length exceeds usize".to_owned())
                })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("protobuf field range overflow".to_owned())
                })?;
            },
            5 => {
                offset = offset.checked_add(4).ok_or_else(|| {
                    Error::InvalidFormat("protobuf fixed32 offset overflow".to_owned())
                })?;
            },
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
        }
        if offset > data.len() {
            return Err(Error::InvalidFormat("truncated protobuf field".to_owned()));
        }
        fields.push(WireField {
            number: number as u32,
            wire_type,
            start,
            key_end,
            payload_start,
            end: offset,
        });
    }
    Ok(fields)
}

pub(crate) fn patch_nested_length_delimited_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    patch_nested_field(data, path, &|data, field_number| {
        patch_length_delimited_field(data, field_number, expected_leaf, replacement)
    })
}

pub(crate) fn patch_nested_varint_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    patch_nested_field(data, path, &|data, field_number| {
        patch_varint_field(data, field_number, expected_leaf, replacement)
    })
}

pub(crate) fn patch_nested_fixed32_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    patch_nested_field(data, path, &|data, field_number| {
        patch_fixed32_field(data, field_number, expected_leaf, replacement)
    })
}

pub(crate) fn patch_nested_fixed64_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    patch_nested_field(data, path, &|data, field_number| {
        patch_fixed64_field(data, field_number, expected_leaf, replacement)
    })
}

pub(crate) fn patch_length_delimited_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} occurs {} times",
            matches.len()
        )));
    }
    if matches.is_empty() == expected_present {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} changed during mutation"
        )));
    }
    let Some(field) = matches.first().copied() else {
        let Some(replacement) = replacement else {
            return Ok(data.to_vec());
        };
        let mut output = data.to_vec();
        append_length_delimited_field(&mut output, field_number, replacement)?;
        return Ok(output);
    };
    require_length_delimited(field)?;
    match replacement {
        Some(replacement) => replace_existing_length_delimited_field(data, field, replacement),
        None => remove_fields(data, vec![*field]),
    }
}

/// Patch one optional protobuf varint while retaining every other byte.
///
/// `expected_present` is checked against the raw wire representation so a
/// concurrent schema mismatch, duplicate singular field, or wrong wire type
/// fails instead of being silently normalized.
pub(crate) fn patch_varint_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    let field = singular_field(data, field_number, expected_present)?;
    let Some(field) = field else {
        let Some(replacement) = replacement else {
            return Ok(data.to_vec());
        };
        let mut output = data.to_vec();
        append_scalar_field(
            &mut output,
            field_number,
            0,
            &crate::varint::encode_varint(replacement),
        )?;
        return Ok(output);
    };
    require_wire_type(&field, 0, "varint")?;
    match replacement {
        Some(replacement) => {
            replace_existing_scalar_field(data, &field, &crate::varint::encode_varint(replacement))
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Patch one optional protobuf fixed32 value while retaining every other byte.
pub(crate) fn patch_fixed32_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    let field = singular_field(data, field_number, expected_present)?;
    let Some(field) = field else {
        let Some(replacement) = replacement else {
            return Ok(data.to_vec());
        };
        let mut output = data.to_vec();
        append_scalar_field(&mut output, field_number, 5, &replacement.to_le_bytes())?;
        return Ok(output);
    };
    require_wire_type(&field, 5, "fixed32")?;
    match replacement {
        Some(replacement) => {
            replace_existing_scalar_field(data, &field, &replacement.to_le_bytes())
        },
        None => remove_fields(data, vec![field]),
    }
}

/// Patch one optional protobuf fixed64 value while retaining every other byte.
pub(crate) fn patch_fixed64_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    let field = singular_field(data, field_number, expected_present)?;
    let Some(field) = field else {
        let Some(replacement) = replacement else {
            return Ok(data.to_vec());
        };
        let mut output = data.to_vec();
        append_scalar_field(&mut output, field_number, 1, &replacement.to_le_bytes())?;
        return Ok(output);
    };
    require_wire_type(&field, 1, "fixed64")?;
    match replacement {
        Some(replacement) => {
            replace_existing_scalar_field(data, &field, &replacement.to_le_bytes())
        },
        None => remove_fields(data, vec![field]),
    }
}

pub(crate) fn transform_length_delimited_field<F>(
    data: &[u8],
    field_number: u32,
    transform: F,
) -> Result<Vec<u8>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {field_number} must occur exactly once, found {}",
            matches.len()
        )));
    }
    let field = matches[0];
    require_length_delimited(field)?;
    let replacement = transform(&data[field.payload_start..field.end])?;
    replace_existing_length_delimited_field(data, field, &replacement)
}

pub(crate) fn append_repeated_length_delimited_field(
    data: &[u8],
    field_number: u32,
    payload: &[u8],
) -> Result<Vec<u8>> {
    // Parse first so malformed existing data cannot be normalized accidentally.
    parse_wire_fields(data)?;
    let mut output = data.to_vec();
    append_length_delimited_field(&mut output, field_number, payload)?;
    Ok(output)
}

/// Return raw payloads for every occurrence of one repeated length-delimited
/// field, rejecting a same-number field with an incompatible wire type.
pub(crate) fn repeated_length_delimited_payloads(
    data: &[u8],
    field_number: u32,
) -> Result<Vec<&[u8]>> {
    let fields = parse_wire_fields(data)?;
    fields
        .iter()
        .filter(|field| field.number == field_number)
        .map(|field| {
            require_length_delimited(field)?;
            Ok(&data[field.payload_start..field.end])
        })
        .collect()
}

/// Replace all occurrences of a repeated length-delimited field while keeping
/// unrelated fields in their original byte positions. Existing field slots
/// retain their original key bytes; additional values are inserted directly
/// after the final existing occurrence so an append followed by removal is
/// byte-exact.
pub(crate) fn rewrite_repeated_length_delimited_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[Vec<u8>],
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .into_iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    for field in &matches {
        require_length_delimited(field)?;
    }
    if matches.is_empty() {
        let mut output = data.to_vec();
        for replacement in replacements {
            append_length_delimited_field(&mut output, field_number, replacement)?;
        }
        return Ok(output);
    }

    let replacement_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        total
            .checked_add(replacement.len())
            .and_then(|length| length.checked_add(10))
    });
    let capacity = data
        .len()
        .checked_add(replacement_bytes.ok_or_else(|| {
            Error::InvalidFormat("protobuf repeated-field size overflow".to_owned())
        })?)
        .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?;
    let mut output = Vec::with_capacity(capacity);
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            output.extend(crate::varint::encode_varint(
                u64::try_from(replacement.len()).map_err(|_| {
                    Error::InvalidFormat("protobuf replacement exceeds u64".to_owned())
                })?,
            ));
            output.extend_from_slice(replacement);
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for replacement in &replacements[matches.len().min(replacements.len())..] {
                append_length_delimited_field(&mut output, field_number, replacement)?;
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
pub(crate) fn repeated_varint_values(data: &[u8], field_number: u32) -> Result<Vec<u64>> {
    parse_wire_fields(data)?
        .into_iter()
        .filter(|field| field.number == field_number)
        .map(|field| {
            require_wire_type(&field, 0, "varint")?;
            let (value, length) =
                crate::varint::decode_varint_from_bytes(&data[field.key_end..field.end]).map_err(
                    |error| Error::InvalidFormat(format!("invalid protobuf varint value: {error}")),
                )?;
            if field.key_end + length != field.end {
                return Err(Error::InvalidFormat(format!(
                    "protobuf field {field_number} has trailing varint bytes"
                )));
            }
            Ok(value)
        })
        .collect()
}

/// Replace an unpacked repeated varint field while preserving unrelated bytes
/// and the original key bytes and positions of retained slots.
pub(crate) fn rewrite_repeated_varint_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .into_iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    for field in &matches {
        require_wire_type(field, 0, "varint")?;
    }
    if matches.is_empty() {
        let mut output = data.to_vec();
        for &replacement in replacements {
            append_scalar_field(
                &mut output,
                field_number,
                0,
                &crate::varint::encode_varint(replacement),
            )?;
        }
        return Ok(output);
    }

    let mut output = Vec::with_capacity(
        data.len()
            .checked_add(replacements.len().saturating_mul(11))
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?,
    );
    let mut copied = 0usize;
    for (index, field) in matches.iter().enumerate() {
        output.extend_from_slice(&data[copied..field.start]);
        if let Some(&replacement) = replacements.get(index) {
            output.extend_from_slice(&data[field.start..field.key_end]);
            output.extend(crate::varint::encode_varint(replacement));
        }
        copied = field.end;
        if index + 1 == matches.len() {
            for &replacement in &replacements[matches.len().min(replacements.len())..] {
                append_scalar_field(
                    &mut output,
                    field_number,
                    0,
                    &crate::varint::encode_varint(replacement),
                )?;
            }
        }
    }
    output.extend_from_slice(&data[copied..]);
    Ok(output)
}

/// Transform every occurrence of a repeated length-delimited field while
/// preserving its original key bytes and the position of unrelated fields.
pub(crate) fn transform_repeated_length_delimited_fields<F>(
    data: &[u8],
    field_number: u32,
    mut transform: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<Vec<u8>>,
{
    let replacements = repeated_length_delimited_payloads(data, field_number)?
        .into_iter()
        .map(&mut transform)
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

/// Transform all length-delimited leaves at a protobuf field path.
///
/// Every path component may occur zero, one, or many times. This makes the
/// operation suitable for schema-known paths that cross optional and repeated
/// messages without normalizing untouched siblings or unknown fields.
pub(crate) fn transform_length_delimited_fields_at_path<F>(
    data: &[u8],
    path: &[u32],
    mut transform: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<Vec<u8>>,
{
    fn visit(
        data: &[u8],
        path: &[u32],
        transform: &mut dyn FnMut(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<Vec<u8>> {
        let (&field_number, remainder) = path.first().zip(path.get(1..)).ok_or_else(|| {
            Error::InvalidFormat("protobuf field path cannot be empty".to_owned())
        })?;
        if remainder.is_empty() {
            return transform_repeated_length_delimited_fields(data, field_number, transform);
        }
        transform_repeated_length_delimited_fields(data, field_number, |nested| {
            visit(nested, remainder, transform)
        })
    }

    visit(data, path, &mut transform)
}

pub(crate) fn remove_repeated_length_delimited_field_where<F>(
    data: &[u8],
    field_number: u32,
    mut remove: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<bool>,
{
    let fields = parse_wire_fields(data)?;
    let mut removed = Vec::new();
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
        )));
    }
    remove_fields(data, removed)
}

fn require_length_delimited(field: &WireField) -> Result<()> {
    require_wire_type(field, 2, "length-delimited")
}

fn patch_nested_field<F>(data: &[u8], path: &[u32], patch_leaf: &F) -> Result<Vec<u8>>
where
    F: Fn(&[u8], u32) -> Result<Vec<u8>>,
{
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

pub(crate) fn append_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
) -> Result<()> {
    output.extend(crate::varint::encode_varint(
        (u64::from(field_number) << 3) | 2,
    ));
    output.extend(crate::varint::encode_varint(
        u64::try_from(payload.len())
            .map_err(|_| Error::InvalidFormat("protobuf replacement exceeds u64".to_owned()))?,
    ));
    output.extend_from_slice(payload);
    Ok(())
}

pub(crate) fn append_varint_field(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
) -> Result<()> {
    append_scalar_field(
        output,
        field_number,
        0,
        &crate::varint::encode_varint(value),
    )
}

fn append_scalar_field(
    output: &mut Vec<u8>,
    field_number: u32,
    wire_type: u8,
    payload: &[u8],
) -> Result<()> {
    if field_number == 0 || field_number > 0x1fff_ffff {
        return Err(Error::InvalidFormat(format!(
            "invalid protobuf field number {field_number}"
        )));
    }
    output.extend(crate::varint::encode_varint(
        (u64::from(field_number) << 3) | u64::from(wire_type),
    ));
    output.extend_from_slice(payload);
    Ok(())
}

fn replace_existing_length_delimited_field(
    data: &[u8],
    field: &WireField,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let old_length = field.end - field.start;
    let mut output = Vec::with_capacity(
        data.len()
            .checked_sub(old_length)
            .and_then(|length| length.checked_add(replacement.len()))
            .and_then(|length| length.checked_add(10))
            .ok_or_else(|| Error::InvalidFormat("protobuf output size overflow".to_owned()))?,
    );
    output.extend_from_slice(&data[..field.start]);
    output.extend_from_slice(&data[field.start..field.key_end]);
    output.extend(crate::varint::encode_varint(
        u64::try_from(replacement.len())
            .map_err(|_| Error::InvalidFormat("protobuf replacement exceeds u64".to_owned()))?,
    ));
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
    let mut output = Vec::with_capacity(capacity);
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
    let mut output = Vec::with_capacity(
        data.len()
            .checked_sub(removed_length.ok_or_else(|| {
                Error::InvalidFormat("protobuf removed-field size overflow".to_owned())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormat("protobuf removed-field range overflow".to_owned())
            })?,
    );
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
mod tests {
    use super::*;

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut field = crate::varint::encode_varint(u64::from(number) << 3);
        field.extend(crate::varint::encode_varint(value));
        field
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
            transform_length_delimited_fields_at_path(&original, &[], |_| Ok(Vec::new())).is_err()
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
}
