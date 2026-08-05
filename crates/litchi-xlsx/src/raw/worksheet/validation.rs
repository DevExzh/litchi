//! Worksheet attribute, ordering, and coordinate validation.

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::{COLUMNS, ROWS};
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{NamespaceResolver, ResolveResult};

use super::model::{Context, TextTarget};
use crate::error::{Result, invalid};
use crate::layout::{self, Defaults};
use crate::row;

pub(crate) fn parse_defaults_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    descent: Option<layout::Descent>,
) -> Result<Defaults> {
    validate_defaults_attributes(element, resolver)?;
    let base_width = optional_u32(
        element,
        b"baseColWidth",
        decoder,
        "worksheet base column width",
    )?
    .map(|value| {
        u8::try_from(value)
            .map_err(|_| invalid("worksheet base column width exceeds Office maximum 255"))
    })
    .transpose()?;
    let width = optional_f64(
        element,
        b"defaultColWidth",
        decoder,
        "worksheet default column width",
    )?
    .map(layout::Width::new)
    .transpose()?;
    let height = optional_f64(
        element,
        b"defaultRowHeight",
        decoder,
        "worksheet default row height",
    )?
    .ok_or_else(|| invalid("worksheet sheetFormatPr is missing defaultRowHeight"))
    .and_then(|value| layout::Height::new(value).map_err(Into::into))?;
    let row_outline = optional_u32(
        element,
        b"outlineLevelRow",
        decoder,
        "worksheet row outline summary",
    )?
    .map(row::OutlineAt::from)
    .map(row::OutlineAt::resolve)
    .transpose()?;
    let column_outline = optional_u32(
        element,
        b"outlineLevelCol",
        decoder,
        "worksheet column outline summary",
    )?
    .map(row::OutlineAt::from)
    .map(row::OutlineAt::resolve)
    .transpose()?;
    let mut flags = layout::Flags::empty();
    let mut present = layout::Flags::empty();
    for (attribute, flag, field) in [
        (
            b"customHeight".as_slice(),
            layout::Flags::CUSTOM_HEIGHT,
            "customHeight",
        ),
        (
            b"zeroHeight".as_slice(),
            layout::Flags::HIDDEN,
            "zeroHeight",
        ),
        (b"thickTop".as_slice(), layout::Flags::THICK_TOP, "thickTop"),
        (
            b"thickBottom".as_slice(),
            layout::Flags::THICK_BOTTOM,
            "thickBottom",
        ),
    ] {
        if let Some(value) = optional_bool(element, attribute, decoder, field)? {
            present.insert(flag);
            flags.set(flag, value);
        }
    }
    Ok(Defaults {
        base_width,
        width,
        height,
        descent,
        row_outline,
        column_outline,
        flags,
        present,
    })
}

fn validate_defaults_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let supported = match namespace {
            ResolveResult::Unbound => matches!(
                local.as_ref(),
                b"baseColWidth"
                    | b"defaultColWidth"
                    | b"defaultRowHeight"
                    | b"customHeight"
                    | b"zeroHeight"
                    | b"thickTop"
                    | b"thickBottom"
                    | b"outlineLevelRow"
                    | b"outlineLevelCol"
            ),
            // Namespaced attributes belong to extension owners. They are not
            // interpreted here, but must survive a preserve-mode edit.
            ResolveResult::Bound(_) => true,
            ResolveResult::Unknown(_) => false,
        };
        if !supported {
            return Err(invalid(format!(
                "unknown worksheet sheetFormatPr attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
    }
    Ok(())
}

pub(crate) fn parse_one_based_row(value: &str) -> Result<u32> {
    let row = value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid worksheet row '{value}'")))?;
    if !(1..=ROWS).contains(&row) {
        return Err(invalid(format!("worksheet row {row} exceeds the grid")));
    }
    Ok(row)
}

pub(crate) fn parse_a1(value: &str) -> Result<(u32, u32)> {
    let bytes = value.as_bytes();
    let split = bytes
        .iter()
        .position(u8::is_ascii_digit)
        .ok_or_else(|| invalid(format!("invalid cell reference '{value}'")))?;
    if split == 0 || split == bytes.len() {
        return Err(invalid(format!("invalid cell reference '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[..split] {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid(format!("invalid cell reference '{value}'")));
        }
        column = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid(format!("cell reference '{value}' overflows")))?;
    }
    if column == 0 || column > COLUMNS {
        return Err(invalid(format!(
            "cell reference '{value}' exceeds the column grid"
        )));
    }
    let row = std::str::from_utf8(&bytes[split..])
        .ok()
        .and_then(|row| row.parse::<u32>().ok())
        .filter(|row| (1..=ROWS).contains(row))
        .ok_or_else(|| invalid(format!("cell reference '{value}' exceeds the row grid")))?;
    Ok((row, column))
}

pub(crate) fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

pub(crate) fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description}")))
}

pub(crate) fn optional_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<f64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

pub(crate) fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid {description} '{value}'"))),
        })
        .transpose()
}

pub(super) fn current(stack: &[Context]) -> Result<Context> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("worksheet XML is missing its root context"))
}

pub(super) fn text_target(stack: &[Context]) -> Option<TextTarget> {
    match stack.last() {
        Some(Context::Formula) => Some(TextTarget::Formula),
        Some(Context::Value) => Some(TextTarget::Value),
        Some(Context::Text(target)) => Some(*target),
        _ => None,
    }
}
