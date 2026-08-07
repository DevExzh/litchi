//! Facade error adapters for the dependency-neutral bounded IWA wire kernel.
//!
//! `litchi-iwa-common` owns the parsed [`WireField`] representation and every
//! bounded scalar/repeated mutation algorithm. This module contains no wire
//! state or parser of its own; the small forwarding functions only convert the
//! common error into the facade error expected by existing archive callers.

use litchi_iwa_common::wire as common_wire;

use crate::Result;

pub(crate) use common_wire::WireField;

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

pub(crate) fn append_length_delimited_field(
    output: &mut Vec<u8>,
    field_number: u32,
    payload: &[u8],
) -> Result<()> {
    common_wire::append_length_delimited_field(output, field_number, payload)?;
    Ok(())
}

pub(crate) fn append_varint_field(
    output: &mut Vec<u8>,
    field_number: u32,
    value: u64,
) -> Result<()> {
    common_wire::append_varint_field(output, field_number, value)?;
    Ok(())
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
