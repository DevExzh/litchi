//! Cross-field and resource validation for XLSB XML map bindings.

use std::collections::HashSet;

use super::model::{Limits, MAX_COLUMN, MAX_ROW, MappedTable, SingleCellBinding};
use super::{XmlMapInfo, XmlMapLimits};
use crate::package::error::{Error, Result};

pub(crate) fn invalid(typ: impl Into<String>, value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: value.into(),
    }
}

pub(crate) fn nonzero_id(value: u32, typ: &'static str) -> Result<()> {
    if value == 0 {
        Err(invalid(typ, "zero"))
    } else {
        Ok(())
    }
}

pub(crate) fn list_id(value: u32, typ: &'static str) -> Result<()> {
    if value == 0 || value == u32::MAX {
        Err(invalid(typ, format!("{value}, expected 1..=4294967294")))
    } else {
        Ok(())
    }
}

pub(crate) fn xml_data_type(value: u32) -> Result<()> {
    if (1..=0x2D).contains(&value) {
        Ok(())
    } else {
        Err(invalid("XmlDataType", format!("0x{value:08X}")))
    }
}

pub(crate) fn xpath(value: &str, max_units: usize) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > max_units {
        return Err(invalid(
            "XmlMappedXpath",
            format!("{units} UTF-16 units, expected 1..={max_units}"),
        ));
    }
    if !value.starts_with('/') {
        return Err(invalid("XmlMappedXpath", "XPath is not absolute"));
    }
    if value
        .chars()
        .any(|ch| ch == '\0' || (ch.is_control() && !matches!(ch, '\t' | '\n' | '\r')))
    {
        return Err(invalid(
            "XmlMappedXpath",
            "XPath contains forbidden XML text",
        ));
    }
    validate_xpath_steps(value)
}

/// Recognize only the non-evaluating `XmlMappedXpath` subset we can prove:
/// QName child steps, one immediate named-attribute equality predicate per
/// element step, and an optional final named-attribute step.
fn validate_xpath_steps(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut offset = 1usize;
    if offset == bytes.len() {
        return Err(invalid("XmlMappedXpath", "XPath has no element step"));
    }
    while offset < bytes.len() {
        if bytes[offset] == b'/' {
            return Err(invalid("XmlMappedXpath", "empty or descendant step"));
        }
        let attribute_step = bytes[offset] == b'@';
        if attribute_step {
            offset += 1;
        }
        offset = parse_qname(value, offset)?;
        if attribute_step {
            if offset != bytes.len() {
                return Err(invalid(
                    "XmlMappedXpath",
                    "attribute step is not the final step",
                ));
            }
            return Ok(());
        }
        if bytes.get(offset) == Some(&b'[') {
            offset = parse_attribute_predicate(value, offset)?;
        }
        if offset == bytes.len() {
            return Ok(());
        }
        if bytes[offset] != b'/' {
            return Err(invalid("XmlMappedXpath", "unsupported XPath syntax"));
        }
        offset += 1;
        if offset == bytes.len() {
            return Err(invalid("XmlMappedXpath", "trailing child separator"));
        }
    }
    Ok(())
}

fn parse_attribute_predicate(value: &str, mut offset: usize) -> Result<usize> {
    let bytes = value.as_bytes();
    offset += 1;
    offset = skip_xpath_whitespace(bytes, offset);
    if bytes.get(offset) != Some(&b'@') {
        return Err(invalid(
            "XmlMappedXpath",
            "predicate is not a named-attribute comparison",
        ));
    }
    offset = parse_qname(value, offset + 1)?;
    offset = skip_xpath_whitespace(bytes, offset);
    if bytes.get(offset) != Some(&b'=') {
        return Err(invalid(
            "XmlMappedXpath",
            "attribute predicate is missing equality",
        ));
    }
    offset = skip_xpath_whitespace(bytes, offset + 1);
    let quote = bytes
        .get(offset)
        .copied()
        .ok_or_else(|| invalid("XmlMappedXpath", "attribute predicate is missing a literal"))?;
    if !matches!(quote, b'\'' | b'"') {
        return Err(invalid(
            "XmlMappedXpath",
            "attribute predicate literal is not quoted",
        ));
    }
    offset += 1;
    while let Some(&byte) = bytes.get(offset) {
        if byte == quote {
            offset = skip_xpath_whitespace(bytes, offset + 1);
            if bytes.get(offset) != Some(&b']') {
                return Err(invalid(
                    "XmlMappedXpath",
                    "predicate contains unsupported compound syntax",
                ));
            }
            return Ok(offset + 1);
        }
        offset += 1;
    }
    Err(invalid(
        "XmlMappedXpath",
        "attribute predicate literal is unterminated",
    ))
}

fn skip_xpath_whitespace(bytes: &[u8], mut offset: usize) -> usize {
    while bytes
        .get(offset)
        .is_some_and(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
    {
        offset += 1;
    }
    offset
}

fn parse_qname(value: &str, offset: usize) -> Result<usize> {
    let tail = value
        .get(offset..)
        .ok_or_else(|| invalid("XmlMappedXpath", "invalid UTF-8 boundary"))?;
    let mut end = offset;
    let mut colon = false;
    let mut component_start = true;
    for ch in tail.chars() {
        let valid = if component_start {
            ch == '_' || ch.is_alphabetic()
        } else {
            ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric()
        };
        if valid {
            component_start = false;
            end += ch.len_utf8();
            continue;
        }
        if ch == ':' && !colon && !component_start {
            colon = true;
            component_start = true;
            end += 1;
            continue;
        }
        break;
    }
    if end == offset || component_start {
        Err(invalid("XmlMappedXpath", "invalid QName step"))
    } else {
        Ok(end)
    }
}

pub(crate) fn cell(row: u32, column: u32) -> Result<()> {
    if row > MAX_ROW {
        return Err(invalid(
            "single-cell row",
            format!("{row} exceeds {MAX_ROW}"),
        ));
    }
    if column > MAX_COLUMN {
        return Err(invalid(
            "single-cell column",
            format!("{column} exceeds {MAX_COLUMN}"),
        ));
    }
    Ok(())
}

pub(crate) fn mapped_table(value: &MappedTable, limits: Limits) -> Result<()> {
    list_id(value.table_id(), "mapped table ID")?;
    if value.columns().len() > limits.max_bindings {
        return Err(Error::InvalidLength {
            expected: limits.max_bindings,
            found: value.columns().len(),
        });
    }
    let mut ids = HashSet::with_capacity(value.columns().len());
    for binding in value.columns() {
        if !ids.insert(binding.column_id()) {
            return Err(invalid(
                "mapped table",
                format!("duplicate column ID {}", binding.column_id()),
            ));
        }
        xpath(binding.xpath().as_str(), limits.max_xpath_units)?;
    }
    Ok(())
}

pub(crate) fn single_cells(values: &[SingleCellBinding], limits: Limits) -> Result<()> {
    if values.len() > limits.max_bindings {
        return Err(Error::InvalidLength {
            expected: limits.max_bindings,
            found: values.len(),
        });
    }
    let mut table_ids = HashSet::with_capacity(values.len());
    let mut cells = HashSet::with_capacity(values.len());
    for value in values {
        nonzero_id(value.table_id(), "single-cell table ID")?;
        cell(value.cell().row(), value.cell().column())?;
        xpath(value.xpath().as_str(), limits.max_xpath_units)?;
        if !table_ids.insert(value.table_id()) {
            return Err(invalid(
                "single-cell tables",
                format!("duplicate table ID {}", value.table_id()),
            ));
        }
        if !cells.insert(value.cell()) {
            return Err(invalid(
                "single-cell tables",
                format!(
                    "duplicate cell ({}, {})",
                    value.cell().row(),
                    value.cell().column()
                ),
            ));
        }
    }
    Ok(())
}

/// Validate the shared catalog under caller-selected resource ceilings.
pub fn validate_catalog(info: &XmlMapInfo, limits: XmlMapLimits) -> Result<()> {
    litchi_ooxml_common::spreadsheet_xml_maps::validate_xml_map_info_with_limits(info, &limits)?;
    Ok(())
}

/// Prove that every BIFF binding references a map present in the catalog.
pub fn validate_binding_map_ids(
    catalog: &XmlMapInfo,
    tables: &[MappedTable],
    single_cells: &[SingleCellBinding],
) -> Result<()> {
    let map_ids: HashSet<u32> = catalog.maps.iter().map(|value| value.id).collect();
    for binding in tables
        .iter()
        .flat_map(MappedTable::columns)
        .chain(single_cells.iter().map(SingleCellBinding::column_binding))
    {
        if !map_ids.contains(&binding.map_id()) {
            return Err(invalid(
                "XML map binding",
                format!("unknown map ID {}", binding.map_id()),
            ));
        }
    }
    Ok(())
}
