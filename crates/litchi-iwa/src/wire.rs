//! Facade-local adapters for the shared bounded protobuf wire kernel.
//!
//! The wire parser and scalar/repeated mutation algorithms live in
//! `litchi-iwa-common`. This module keeps only the application crate's
//! visibility boundary and the few callbacks that return the facade error
//! type. It is intentionally not a second wire implementation; the callbacks
//! delegate to the common generic error boundary below.

use litchi_iwa_common::wire as common_wire;

#[cfg(test)]
use crate::Error;
use crate::Result;

pub(crate) use common_wire::WireField;

pub(crate) fn parse_wire_view(data: &[u8]) -> Result<common_wire::WireView<'_>> {
    Ok(common_wire::parse_wire_view(data)?)
}

pub(crate) fn parse_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    Ok(common_wire::parse_wire_fields(data)?)
}

pub(crate) fn overlay_singular_wire_fields(base: &[u8], overlay: &[u8]) -> Result<Vec<u8>> {
    Ok(common_wire::overlay_singular_wire_fields(base, overlay)?)
}

pub(crate) fn patch_nested_length_delimited_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_nested_length_delimited_field(
        data,
        path,
        expected_leaf,
        replacement,
    )?)
}

pub(crate) fn patch_nested_varint_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_nested_varint_field(
        data,
        path,
        expected_leaf,
        replacement,
    )?)
}

pub(crate) fn patch_nested_fixed32_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_nested_fixed32_field(
        data,
        path,
        expected_leaf,
        replacement,
    )?)
}

pub(crate) fn patch_nested_fixed64_field(
    data: &[u8],
    path: &[u32],
    expected_leaf: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_nested_fixed64_field(
        data,
        path,
        expected_leaf,
        replacement,
    )?)
}

pub(crate) fn patch_length_delimited_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_length_delimited_field(
        data,
        field_number,
        expected_present,
        replacement,
    )?)
}

pub(crate) fn patch_varint_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_varint_field(
        data,
        field_number,
        expected_present,
        replacement,
    )?)
}

pub(crate) fn patch_fixed32_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u32>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_fixed32_field(
        data,
        field_number,
        expected_present,
        replacement,
    )?)
}

pub(crate) fn patch_fixed64_field(
    data: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    Ok(common_wire::patch_fixed64_field(
        data,
        field_number,
        expected_present,
        replacement,
    )?)
}

pub(crate) fn transform_length_delimited_field<F>(
    data: &[u8],
    field_number: u32,
    transform: F,
) -> Result<Vec<u8>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    common_wire::transform_length_delimited_field(data, field_number, transform)
}

pub(crate) fn append_repeated_length_delimited_field(
    data: &[u8],
    field_number: u32,
    payload: &[u8],
) -> Result<Vec<u8>> {
    Ok(common_wire::append_repeated_length_delimited_field(
        data,
        field_number,
        payload,
    )?)
}

pub(crate) fn repeated_length_delimited_payloads(
    data: &[u8],
    field_number: u32,
) -> Result<Vec<&[u8]>> {
    Ok(common_wire::repeated_length_delimited_payloads(
        data,
        field_number,
    )?)
}

pub(crate) fn rewrite_repeated_length_delimited_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[Vec<u8>],
) -> Result<Vec<u8>> {
    Ok(common_wire::rewrite_repeated_length_delimited_fields(
        data,
        field_number,
        replacements,
    )?)
}

pub(crate) fn repeated_varint_values(data: &[u8], field_number: u32) -> Result<Vec<u64>> {
    Ok(common_wire::repeated_varint_values(data, field_number)?)
}

pub(crate) fn rewrite_repeated_varint_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
) -> Result<Vec<u8>> {
    Ok(common_wire::rewrite_repeated_varint_fields(
        data,
        field_number,
        replacements,
    )?)
}

pub(crate) fn repeated_fixed64_values(data: &[u8], field_number: u32) -> Result<Vec<u64>> {
    Ok(common_wire::repeated_fixed64_values(data, field_number)?)
}

#[cfg(test)]
pub(crate) fn repeated_fixed32_values(data: &[u8], field_number: u32) -> Result<Vec<u32>> {
    Ok(common_wire::repeated_fixed32_values(data, field_number)?)
}

pub(crate) fn rewrite_repeated_fixed32_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u32],
) -> Result<Vec<u8>> {
    Ok(common_wire::rewrite_repeated_fixed32_fields(
        data,
        field_number,
        replacements,
    )?)
}

pub(crate) fn rewrite_repeated_fixed64_fields(
    data: &[u8],
    field_number: u32,
    replacements: &[u64],
) -> Result<Vec<u8>> {
    Ok(common_wire::rewrite_repeated_fixed64_fields(
        data,
        field_number,
        replacements,
    )?)
}

pub(crate) fn transform_repeated_length_delimited_fields<F>(
    data: &[u8],
    field_number: u32,
    transform: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<Vec<u8>>,
{
    common_wire::transform_repeated_length_delimited_fields(data, field_number, transform)
}

pub(crate) fn transform_length_delimited_fields_at_path<F>(
    data: &[u8],
    path: &[u32],
    transform: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<Vec<u8>>,
{
    common_wire::transform_length_delimited_fields_at_path(data, path, transform)
}

pub(crate) fn remove_repeated_length_delimited_field_where<F>(
    data: &[u8],
    field_number: u32,
    remove: F,
) -> Result<Vec<u8>>
where
    F: FnMut(&[u8]) -> Result<bool>,
{
    common_wire::remove_repeated_length_delimited_field_where(data, field_number, remove)
}

pub(crate) fn append_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
) -> Result<()> {
    Ok(common_wire::append_length_delimited_field(
        output,
        field_number,
        payload,
    )?)
}

pub(crate) fn append_varint_field(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
) -> Result<()> {
    Ok(common_wire::append_varint_field(
        output,
        field_number,
        value,
    )?)
}

#[cfg(test)]
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
    if !matches!(wire_type, 1 | 5) {
        return Err(Error::InvalidFormat(format!(
            "invalid protobuf wire type {wire_type}"
        )));
    }
    litchi_iwa_common::varint::encode_varint_into(
        output,
        (u64::from(field_number) << 3) | u64::from(wire_type),
    );
    output.extend_from_slice(payload);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut field = litchi_iwa_common::varint::encode_varint(u64::from(number) << 3);
        field.extend(litchi_iwa_common::varint::encode_varint(value));
        field
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
    fn varint_rewrites_cover_full_u64_range() {
        let original = varint_field(7, 0);
        let replaced = patch_varint_field(&original, 7, true, Some(u64::MAX)).unwrap();
        assert_eq!(repeated_varint_values(&replaced, 7).unwrap(), [u64::MAX]);

        let appended = rewrite_repeated_varint_fields(&original, 7, &[127, 128, u64::MAX]).unwrap();
        assert_eq!(
            repeated_varint_values(&appended, 7).unwrap(),
            [127, 128, u64::MAX]
        );
    }

    #[test]
    fn append_rejects_invalid_fields_without_mutating_output() {
        let prefix = varint_field(9, 1);
        let mut output = prefix.clone();

        assert!(append_length_delimited_field(&mut output, 0, b"x").is_err());
        assert_eq!(output, prefix);

        assert!(append_length_delimited_field(&mut output, 0x2000_0000, b"x").is_err());
        assert_eq!(output, prefix);

        append_length_delimited_field(&mut output, 0x1fff_ffff, b"x").unwrap();
        assert_eq!(
            parse_wire_fields(&output).unwrap().last().unwrap().number(),
            0x1fff_ffff
        );
    }

    #[test]
    fn nested_scalar_patch_retains_unknown_ancestors() {
        let mut inner = varint_field(99, 1234);
        let key = litchi_iwa_common::varint::encode_varint((u64::from(4_u32) << 3) | 2);
        let mut outer = key;
        outer.extend(litchi_iwa_common::varint::encode_varint(inner.len() as u64));
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
                .find(|field| field.number() == 1)
                .ok_or_else(|| Error::InvalidFormat("missing identifier".to_owned()))?;
            let (value, _) = litchi_iwa_common::varint::decode_varint_from_bytes(
                &leaf[identifier.payload_start()..identifier.end()],
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
                .find(|field| field.number() == 1)
                .ok_or_else(|| Error::InvalidFormat("missing identifier".to_owned()))?;
            let (value, _) = litchi_iwa_common::varint::decode_varint_from_bytes(
                &leaf[identifier.payload_start()..identifier.end()],
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
    fn common_callback_boundary_preserves_facade_errors() {
        let mut data = Vec::new();
        append_length_delimited_field(&mut data, 2, b"payload").unwrap();
        let error = transform_length_delimited_field(&data, 2, |_| {
            Err(Error::InvalidFormat("callback sentinel".to_owned()))
        })
        .unwrap_err();
        assert!(matches!(error, Error::InvalidFormat(message) if message == "callback sentinel"));
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
