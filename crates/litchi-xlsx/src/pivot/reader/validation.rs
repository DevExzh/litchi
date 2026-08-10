//! Shared scalar and structural validation for pivot XML codecs.

use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;

use litchi_core::sheet::Result as SheetResult;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::Rect;

pub(super) fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

pub(super) fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_source| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

pub(super) fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

pub(super) fn required_i32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<i32> {
    let value = required_string(element, name, decoder, description)?;
    value
        .parse::<i32>()
        .map_err(|_source| format!("invalid {description} '{value}'").into())
}

pub(super) fn optional_i32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<i32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_source| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

pub(super) fn optional_u8(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<u8>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_source| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

pub(super) fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(format!("invalid {description} '{value}'").into()),
        })
        .transpose()
}

pub(super) fn required_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<bool> {
    optional_bool(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

pub(super) fn optional_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<f64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            let parsed = value
                .parse::<f64>()
                .map_err(|_source| format!("invalid {description} '{value}'"))?;
            if !parsed.is_finite() {
                return Err(format!("{description} must be finite").into());
            }
            Ok(parsed)
        })
        .transpose()
}

pub(super) fn required_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<f64> {
    optional_f64(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

pub(super) fn mark_once(seen: &mut bool, description: &str) -> SheetResult<()> {
    if *seen {
        return Err(format!("duplicate {description} element").into());
    }
    *seen = true;
    Ok(())
}

pub(super) fn validate_count(
    expected: Option<u32>,
    actual: usize,
    description: &str,
) -> SheetResult<()> {
    if let Some(expected) = expected
        && usize::try_from(expected) != Ok(actual)
    {
        return Err(
            format!("{description} count is {expected}, but {actual} elements were found").into(),
        );
    }
    Ok(())
}

pub(super) fn validate_cell_range(range: &str, description: &str) -> SheetResult<()> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| format!("empty {description}"))?;
    Rect::from_a1(&first.replace('$', ""))?;
    if let Some(second) = references.next() {
        Rect::from_a1(&second.replace('$', ""))?;
    }
    if references.next().is_some() {
        return Err(format!("invalid {description} range '{range}'").into());
    }
    Ok(())
}
