//! Typed VARIANT semantic codec for Property Set values.

use super::super::super::model::*;
use super::wire::{
    ValueReader, append_bytes, append_u16, append_u32, append_u64, encode_ansi, pad4,
    read_codepage_string, read_u16, read_unicode_string, reserve_bytes,
};
use chrono::{DateTime, Duration, Utc};
use litchi_cfb::OleError;
use litchi_cfb::consts::*;

const MAX_VECTOR_ELEMENTS: usize = 1_000_000;

pub(super) fn serialize_typed(value: &Value, codepage: u16) -> Result<Vec<u8>, OleError> {
    let mut out = try_vec_with_capacity(4, "serialized property value")?;
    append_typed(&mut out, value, codepage)?;
    Ok(out)
}
fn append_typed(out: &mut Vec<u8>, value: &Value, codepage: u16) -> Result<(), OleError> {
    let vt = variant_type(value);
    append_u16(out, vt, "serialized property value")?;
    append_u16(out, 0, "serialized property value")?;
    append_body(out, value, codepage)?;
    Ok(())
}
fn variant_type(value: &Value) -> u16 {
    match value {
        Value::Empty => VT_EMPTY,
        Value::Null => VT_NULL,
        Value::I1(_) => VT_I1,
        Value::UI1(_) => VT_UI1,
        Value::I2(_) => VT_I2,
        Value::UI2(_) => VT_UI2,
        Value::I4(_) => VT_I4,
        Value::UI4(_) => VT_UI4,
        Value::I8(_) => VT_I8,
        Value::UI8(_) => VT_UI8,
        Value::Int(_) => VT_INT,
        Value::UInt(_) => VT_UINT,
        Value::R4(_) => VT_R4,
        Value::R8(_) => VT_R8,
        Value::Currency(_) => VT_CY,
        Value::Date(_) => VT_DATE,
        Value::Bstr(_) => VT_BSTR,
        Value::Error(_) => VT_ERROR,
        Value::Bool(_) => VT_BOOL,
        Value::Decimal(_) => VT_DECIMAL,
        Value::Lpstr(_) => VT_LPSTR,
        Value::Lpwstr(_) => VT_LPWSTR,
        Value::Filetime(_) => VT_FILETIME,
        Value::Blob(_) => VT_BLOB,
        Value::Clipboard { .. } => VT_CF,
        Value::Clsid(_) => VT_CLSID,
        Value::Vector(vector) => VT_VECTOR | vector.scalar().raw(),
        Value::Array(array) => VT_ARRAY | array.scalar().raw(),
        Value::Unknown { variant_type, .. } => *variant_type,
    }
}
fn append_body(out: &mut Vec<u8>, value: &Value, codepage: u16) -> Result<(), OleError> {
    match value {
        Value::Empty | Value::Null => {},
        Value::I1(v) => append_bytes(out, &[*v as u8], "serialized property value")?,
        Value::UI1(v) => append_bytes(out, &[*v], "serialized property value")?,
        Value::I2(v) => append_bytes(out, &v.to_le_bytes(), "serialized property value")?,
        Value::UI2(v) => append_bytes(out, &v.to_le_bytes(), "serialized property value")?,
        Value::I4(v) | Value::Int(v) => {
            append_bytes(out, &v.to_le_bytes(), "serialized property value")?
        },
        Value::UI4(v) | Value::UInt(v) | Value::Error(v) => {
            append_u32(out, *v, "serialized property value")?
        },
        Value::I8(v) | Value::Currency(v) => {
            append_bytes(out, &v.to_le_bytes(), "serialized property value")?
        },
        Value::UI8(v) | Value::Filetime(v) => append_u64(out, *v, "serialized property value")?,
        Value::R4(v) => append_u32(out, v.to_bits(), "serialized property value")?,
        Value::R8(v) | Value::Date(v) => append_u64(out, v.to_bits(), "serialized property value")?,
        Value::Bool(v) => append_bytes(
            out,
            &(if *v { -1i16 } else { 0 }).to_le_bytes(),
            "serialized property value",
        )?,
        Value::Decimal(v) => append_bytes(out, v, "serialized property value")?,
        Value::Bstr(v) | Value::Lpstr(v) => append_codepage_string(out, v, codepage)?,
        Value::Lpwstr(v) => {
            let units = checked_add(v.encode_utf16().count(), 1, "LPWSTR length")?;
            let byte_len = checked_mul(units, 2, "LPWSTR length")?;
            append_u32(
                out,
                checked_u32(units, "LPWSTR length")?,
                "serialized property value",
            )?;
            reserve_bytes(out, byte_len, "serialized property value")?;
            for unit in v.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes());
            }
            out.extend_from_slice(&0u16.to_le_bytes());
        },
        Value::Blob(v) => {
            append_u32(
                out,
                checked_u32(v.len(), "Blob size")?,
                "serialized property value",
            )?;
            append_bytes(out, v, "serialized property value")?;
            pad4(out)?;
        },
        Value::Clipboard { format, data } => {
            let size = checked_add(data.len(), 4, "Clipboard size")?;
            append_u32(
                out,
                checked_u32(size, "Clipboard size")?,
                "serialized property value",
            )?;
            append_bytes(out, &format.to_le_bytes(), "serialized property value")?;
            append_bytes(out, data, "serialized property value")?;
            pad4(out)?;
        },
        Value::Clsid(v) => append_bytes(out, v.as_bytes(), "serialized property value")?,
        Value::Vector(vector) => {
            vector.validate()?;
            append_u32(
                out,
                checked_u32(vector.values().len(), "Vector element count")?,
                "serialized property value",
            )?;
            for value in vector.values() {
                if vector.scalar() == Scalar::Variant {
                    append_typed(out, value, codepage)?;
                    pad4(out)?;
                } else {
                    append_body(out, value, codepage)?;
                }
            }
            pad4(out)?;
        },
        Value::Array(array) => {
            array.validate()?;
            append_u32(
                out,
                u32::from(array.scalar().raw()),
                "serialized array scalar type",
            )?;
            append_u32(
                out,
                checked_u32(array.dimensions().len(), "Array dimension count")?,
                "serialized array dimension count",
            )?;
            for dimension in array.dimensions() {
                append_u32(out, dimension.size(), "serialized array dimension")?;
                append_bytes(
                    out,
                    &dimension.index_offset().to_le_bytes(),
                    "serialized array dimension",
                )?;
            }
            for value in array.values() {
                if array.scalar() == Scalar::Variant {
                    append_typed(out, value, codepage)?;
                    pad4(out)?;
                } else {
                    append_body(out, value, codepage)?;
                }
            }
            pad4(out)?;
        },
        Value::Unknown { data, .. } => append_bytes(out, data, "serialized property value")?,
    }
    Ok(())
}
fn append_codepage_string(out: &mut Vec<u8>, value: &str, codepage: u16) -> Result<(), OleError> {
    if codepage == UNICODE_CODEPAGE {
        let units = checked_add(value.encode_utf16().count(), 1, "CodePageString length")?;
        let byte_len = checked_mul(units, 2, "CodePageString length")?;
        append_u32(
            out,
            checked_u32(byte_len, "CodePageString length")?,
            "serialized CodePageString",
        )?;
        reserve_bytes(out, byte_len, "serialized CodePageString")?;
        for unit in value.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out.extend_from_slice(&[0, 0]);
    } else {
        let bytes = encode_ansi(value, codepage)?;
        let byte_len = checked_add(bytes.len(), 1, "CodePageString length")?;
        append_u32(
            out,
            checked_u32(byte_len, "CodePageString length")?,
            "serialized CodePageString",
        )?;
        reserve_bytes(out, byte_len, "serialized CodePageString")?;
        out.extend_from_slice(&bytes);
        out.push(0);
    }
    pad4(out)?;
    Ok(())
}
pub(crate) fn parse_typed_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<Value, OleError> {
    if data.len() < 4 {
        return Err(invalid("Typed property value is truncated"));
    }
    let variant_type = read_u16(data, 0, "property variant type")?;
    if read_u16(data, 2, "property reserved field")? != 0 {
        return Err(invalid("Typed property reserved field must be zero"));
    }
    let value_offset = property_offset
        .checked_add(4)
        .ok_or_else(|| invalid("Property value offset overflow"))?;
    let mut reader = ValueReader::new(&data[4..], value_offset);
    let value = parse_value_body(&mut reader, variant_type, codepage, 0)?;
    let allows_opaque_tail = matches!(
        variant_type,
        VT_EMPTY | VT_NULL | VT_BSTR | VT_LPSTR | VT_LPWSTR
    );
    if reader.remaining_len() <= 3 && allows_opaque_tail {
        // Some producers retain up to three opaque bytes after a top-level
        // value, especially when a non-DWORD property offset is used. The
        // property offset table still bounds the value.
        reader.take_remaining();
    } else {
        reader
            .finish_zero_padding("property padding")
            .map_err(|error| {
                invalid(format!(
                    "variant 0x{variant_type:04x} has {} trailing bytes: {error}",
                    reader.remaining_len()
                ))
            })?;
    }
    Ok(value)
}

fn parse_value_body(
    reader: &mut ValueReader<'_>,
    variant_type: u16,
    codepage: u16,
    depth: usize,
) -> Result<Value, OleError> {
    if depth > 8 {
        return Err(invalid("Property value nesting exceeds the safety limit"));
    }
    if variant_type & VT_VECTOR != 0 {
        let base_type = variant_type & !VT_VECTOR;
        let Some(scalar) = Scalar::from_raw(base_type) else {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested vector type 0x{variant_type:04x}"
                )));
            }
            return Ok(Value::Unknown {
                variant_type,
                data: try_copy_bytes(reader.take_remaining(), "unknown vector property")?,
            });
        };
        if !scalar.allows_vector() {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested vector type 0x{variant_type:04x}"
                )));
            }
            return Ok(Value::Unknown {
                variant_type,
                data: try_copy_bytes(reader.take_remaining(), "unknown vector property")?,
            });
        }
        return parse_vector(reader, scalar, codepage, depth + 1);
    }
    if variant_type & VT_ARRAY != 0 {
        let base_type = variant_type & !VT_ARRAY;
        let Some(scalar) = Scalar::from_raw(base_type) else {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested array type 0x{variant_type:04x}"
                )));
            }
            return Ok(Value::Unknown {
                variant_type,
                data: try_copy_bytes(reader.take_remaining(), "unknown array property")?,
            });
        };
        if !scalar.allows_array() {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested array type 0x{variant_type:04x}"
                )));
            }
            return Ok(Value::Unknown {
                variant_type,
                data: try_copy_bytes(reader.take_remaining(), "unknown array property")?,
            });
        }
        return parse_array(reader, scalar, codepage, depth + 1);
    }
    match variant_type {
        VT_EMPTY => Ok(Value::Empty),
        VT_NULL => Ok(Value::Null),
        VT_I1 => Ok(Value::I1(reader.read_i8("I1 value")?)),
        VT_UI1 => Ok(Value::UI1(reader.read_u8("UI1 value")?)),
        VT_I2 => Ok(Value::I2(reader.read_i16("I2 value")?)),
        VT_UI2 => Ok(Value::UI2(reader.read_u16("UI2 value")?)),
        VT_I4 => Ok(Value::I4(reader.read_i32("I4 value")?)),
        VT_UI4 => Ok(Value::UI4(reader.read_u32("UI4 value")?)),
        VT_I8 => Ok(Value::I8(reader.read_i64("I8 value")?)),
        VT_UI8 => Ok(Value::UI8(reader.read_u64("UI8 value")?)),
        VT_INT => Ok(Value::Int(reader.read_i32("INT value")?)),
        VT_UINT => Ok(Value::UInt(reader.read_u32("UINT value")?)),
        VT_R4 => Ok(Value::R4(f32::from_bits(reader.read_u32("R4 value")?))),
        VT_R8 => Ok(Value::R8(f64::from_bits(reader.read_u64("R8 value")?))),
        VT_CY => Ok(Value::Currency(reader.read_i64("CY value")?)),
        VT_DATE => Ok(Value::Date(f64::from_bits(reader.read_u64("DATE value")?))),
        VT_BSTR => Ok(Value::Bstr(read_codepage_string(
            reader,
            codepage,
            "BSTR value",
            depth == 0,
        )?)),
        VT_ERROR => Ok(Value::Error(reader.read_u32("ERROR value")?)),
        // Office producers sometimes use 1 for TRUE even though VARIANT_BOOL
        // conventionally uses -1. Preserve the established nonzero mapping.
        VT_BOOL => Ok(Value::Bool(reader.read_i16("BOOL value")? != 0)),
        VT_DECIMAL => {
            let raw = reader.take(16, "DECIMAL value")?;
            let mut value = [0u8; 16];
            value.copy_from_slice(raw);
            Ok(Value::Decimal(value))
        },
        VT_LPSTR => Ok(Value::Lpstr(read_codepage_string(
            reader,
            codepage,
            "LPSTR value",
            depth == 0,
        )?)),
        VT_LPWSTR => Ok(Value::Lpwstr(read_unicode_string(
            reader,
            "LPWSTR value",
            depth == 0,
        )?)),
        VT_FILETIME => Ok(Value::Filetime(reader.read_u64("FILETIME value")?)),
        VT_BLOB | VT_BLOB_OBJECT => {
            let size = usize::try_from(reader.read_u32("blob size")?)
                .map_err(|_| invalid("Blob size is too large"))?;
            let blob = try_copy_bytes(reader.take(size, "blob data")?, "blob data")?;
            reader.align4(depth == 0, "blob padding")?;
            Ok(Value::Blob(blob))
        },
        VT_CF => {
            let size = usize::try_from(reader.read_u32("clipboard size")?)
                .map_err(|_| invalid("Clipboard size is too large"))?;
            if size < 4 {
                return Err(invalid("Clipboard size must include its format field"));
            }
            let format = reader.read_i32("clipboard format")?;
            let payload =
                try_copy_bytes(reader.take(size - 4, "clipboard data")?, "clipboard data")?;
            reader.align4(depth == 0, "clipboard padding")?;
            Ok(Value::Clipboard {
                format,
                data: payload,
            })
        },
        VT_CLSID => {
            let raw = reader.take(16, "CLSID value")?;
            let mut value = [0u8; 16];
            value.copy_from_slice(raw);
            Ok(Value::Clsid(Guid::from_bytes(value)))
        },
        _ if depth == 0 => Ok(Value::Unknown {
            variant_type,
            data: try_copy_bytes(reader.take_remaining(), "unknown property value")?,
        }),
        _ => Err(invalid(format!(
            "Unsupported nested property variant 0x{variant_type:04x}"
        ))),
    }
}

fn parse_vector(
    reader: &mut ValueReader<'_>,
    scalar: Scalar,
    codepage: u16,
    depth: usize,
) -> Result<Value, OleError> {
    let count = usize::try_from(reader.read_u32("vector element count")?)
        .map_err(|_| invalid("Vector element count is too large"))?;
    if count > MAX_VECTOR_ELEMENTS {
        return Err(invalid(format!(
            "Vector element count {count} exceeds the safety limit"
        )));
    }
    let minimum = minimum_value_size(scalar.raw()).unwrap_or(0);
    if minimum != 0
        && count
            .checked_mul(minimum)
            .is_none_or(|size| size > reader.remaining_len())
    {
        return Err(invalid("Vector element count exceeds its property range"));
    }
    if scalar == Scalar::Variant && count > reader.remaining_len() / 4 {
        return Err(invalid("Variant vector count exceeds its property range"));
    }
    let mut values = try_vec_with_capacity(count, "property vector")?;
    for _ in 0..count {
        let value = if scalar == Scalar::Variant {
            let nested_type = reader.read_u16("variant vector element type")?;
            if reader.read_u16("variant vector reserved field")? != 0 {
                return Err(invalid("Variant vector reserved field must be zero"));
            }
            let value = parse_value_body(reader, nested_type, codepage, depth + 1)?;
            reader.align4(false, "variant vector element padding")?;
            value
        } else {
            parse_value_body(reader, scalar.raw(), codepage, depth + 1)?
        };
        values.push(value);
    }
    if scalar != Scalar::Variant {
        reader.align4(false, "vector element padding")?;
    }
    Ok(Value::Vector(Vector::new(scalar, values)?))
}

fn parse_array(
    reader: &mut ValueReader<'_>,
    scalar: Scalar,
    codepage: u16,
    depth: usize,
) -> Result<Value, OleError> {
    let declared_type = reader.read_u32("array scalar type")?;
    if declared_type != u32::from(scalar.raw()) {
        return Err(invalid(
            "Array header scalar type does not match its VARIANT type",
        ));
    }
    let dimensions = usize::try_from(reader.read_u32("array dimension count")?)
        .map_err(|_| invalid("Array dimension count is too large"))?;
    if !(1..=31).contains(&dimensions) {
        return Err(invalid("Array dimension count must be between 1 and 31"));
    }
    let mut shape = try_vec_with_capacity(dimensions, "property array dimensions")?;
    let mut count = 1usize;
    for _ in 0..dimensions {
        let size = reader.read_u32("array dimension size")?;
        let index_offset = reader.read_i32("array dimension index offset")?;
        count = count
            .checked_mul(
                usize::try_from(size).map_err(|_| invalid("Array dimension is too large"))?,
            )
            .ok_or_else(|| invalid("Array element count overflows usize"))?;
        if count > MAX_VECTOR_ELEMENTS {
            return Err(invalid("Array exceeds the element safety limit"));
        }
        shape.push(Dimension::new(size, index_offset));
    }
    let mut values = try_vec_with_capacity(count, "property array values")?;
    for _ in 0..count {
        let value = if scalar == Scalar::Variant {
            let nested_type = reader.read_u16("array variant element type")?;
            if reader.read_u16("array variant reserved field")? != 0 {
                return Err(invalid("Array variant reserved field must be zero"));
            }
            let value = parse_value_body(reader, nested_type, codepage, depth + 1)?;
            reader.align4(false, "array variant element padding")?;
            value
        } else {
            parse_value_body(reader, scalar.raw(), codepage, depth + 1)?
        };
        values.push(value);
    }
    if scalar != Scalar::Variant {
        reader.align4(false, "array element padding")?;
    }
    Ok(Value::Array(Array::new(scalar, shape, values)?))
}

fn minimum_value_size(variant_type: u16) -> Option<usize> {
    match variant_type {
        VT_I1 | VT_UI1 => Some(1),
        VT_I2 | VT_UI2 | VT_BOOL => Some(2),
        VT_I4 | VT_UI4 | VT_INT | VT_UINT | VT_R4 | VT_ERROR => Some(4),
        VT_I8 | VT_UI8 | VT_R8 | VT_CY | VT_DATE | VT_FILETIME => Some(8),
        VT_CLSID => Some(16),
        VT_VARIANT | VT_BSTR | VT_LPSTR | VT_LPWSTR | VT_CF => Some(4),
        _ => None,
    }
}

pub(crate) fn filetime_to_date(filetime: u64) -> Option<DateTime<Utc>> {
    const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
    let doc_epoch = i64::try_from(filetime).ok()?;
    let nanos = doc_epoch.checked_sub(EPOCH_DIFF)?.checked_mul(100)?;
    Some(DateTime::from_timestamp_nanos(nanos))
}

pub(crate) fn filetime_to_duration(filetime: u64) -> Option<Duration> {
    let nanos = filetime.checked_mul(100)?;
    Some(Duration::nanoseconds(i64::try_from(nanos).ok()?))
}
