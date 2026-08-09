//! Validation for bounded `SpreadsheetML` Custom XML Maps.

use std::collections::HashSet;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::invalid;
use super::model::{XmlMapInfo, XmlMapInfoRef, XmlMapLimits};
use crate::Result;

/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_xml_map_info(info: &XmlMapInfo) -> Result<()> {
    validate_xml_map_info_with_limits(info, &XmlMapLimits::DEFAULT)
}

/// Validate a `MapInfo` value using caller-selected resource ceilings.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_xml_map_info_with_limits(info: &XmlMapInfo, limits: &XmlMapLimits) -> Result<()> {
    let info = XmlMapInfoRef::from_owned_with_limits(info, limits)?;
    validate_xml_map_info_ref_with_limits(&info, limits)
}

/// Validate a borrowed `MapInfo` projection using default resource ceilings.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_xml_map_info_ref(info: &XmlMapInfoRef<'_>) -> Result<()> {
    validate_xml_map_info_ref_with_limits(info, &XmlMapLimits::DEFAULT)
}

/// Validate a borrowed `MapInfo` projection using caller-selected ceilings.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn validate_xml_map_info_ref_with_limits(
    info: &XmlMapInfoRef<'_>,
    limits: &XmlMapLimits,
) -> Result<()> {
    bounded(info.selection_namespaces, limits)?;
    if info.schemas.is_empty() || info.schemas.len() > limits.max_schemas {
        return Err(invalid("MapInfo requires 1..4096 Schema children"));
    }
    if info.maps.is_empty() || info.maps.len() > limits.max_maps {
        return Err(invalid("MapInfo requires 1..65536 Map children"));
    }
    let mut schema_ids = HashSet::new();
    for schema in &info.schemas {
        bounded_office_string(schema.id, limits)?;
        if !schema_ids.insert(schema.id) {
            return Err(invalid("duplicate Schema ID"));
        }
        optional_bounded_office_string(schema.schema_reference, limits)?;
        optional_bounded_office_string(schema.namespace, limits)?;
        if let Some(payload) = &schema.payload_xml {
            validate_opaque(payload, limits)?;
        }
    }
    let mut map_ids = HashSet::new();
    for map in &info.maps {
        if map.id == 0 || map.id > i32::MAX as u32 {
            return Err(invalid("Map ID must be in 1..=2147483647"));
        }
        if !map_ids.insert(map.id) {
            return Err(invalid("duplicate Map ID"));
        }
        for value in [&map.name, &map.root_element, &map.schema_id] {
            bounded_office_string(value, limits)?;
        }
        if !schema_ids.contains(map.schema_id) {
            return Err(invalid("Map references an unknown SchemaID"));
        }
        if let Some(binding) = &map.data_binding {
            optional_bounded_office_string(binding.data_binding_name, limits)?;
            optional_bounded_office_string(binding.file_binding_name, limits)?;
            if let Some(true) = binding.file_binding {
                let connection_id = binding
                    .connection_id
                    .ok_or_else(|| invalid("ConnectionID is required when FileBinding is true"))?;
                if connection_id > i32::MAX as u32 {
                    return Err(invalid("ConnectionID must be at most 2147483647"));
                }
            } else {
                if binding.connection_id.is_some() {
                    return Err(invalid(
                        "ConnectionID is only permitted when FileBinding is true",
                    ));
                }
                if binding.file_binding_name.is_some() {
                    return Err(invalid(
                        "FileBindingName is absent when FileBinding is false",
                    ));
                }
            }
            if let Some(payload) = &binding.payload_xml {
                validate_opaque(payload, limits)?;
            }
        }
    }
    Ok(())
}
fn validate_opaque(xml: &[u8], limits: &XmlMapLimits) -> Result<()> {
    if xml.len() > limits.max_opaque_bytes {
        return Err(invalid("opaque XML payload exceeds 16 MiB"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut events = 0usize;
    loop {
        events += 1;
        if events > limits.max_events {
            return Err(invalid("opaque XML event limit exceeded"));
        }
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                if depth == 0 {
                    roots += 1;
                }
                depth += 1;
                if depth > limits.max_depth {
                    return Err(invalid("opaque XML depth limit exceeded"));
                }
            },
            Ok(Event::Empty(_)) if depth == 0 => {
                if limits.max_depth == 0 {
                    return Err(invalid("opaque XML depth limit exceeded"));
                }
                roots += 1;
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid opaque XML nesting"))?;
            },
            Ok(Event::DocType(_) | Event::PI(_) | Event::Decl(_)) => {
                return Err(invalid(
                    "DTD, declarations, and processing instructions are rejected in opaque XML",
                ));
            },
            Ok(Event::Text(t))
                if depth == 0 && !t.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("text outside opaque XML root"));
            },
            Ok(Event::CData(_) | Event::GeneralRef(_)) if depth == 0 => {
                return Err(invalid("data outside opaque XML root"));
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
            _ => {},
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid(
            "opaque XML must contain exactly one complete element",
        ));
    }
    Ok(())
}

fn bounded(value: &str, limits: &XmlMapLimits) -> Result<()> {
    if value.len() > limits.max_string_bytes {
        Err(invalid("custom XML maps string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn bounded_office_string(value: &str, limits: &XmlMapLimits) -> Result<()> {
    bounded(value, limits)?;
    if value.chars().count() > 65_535 {
        return Err(invalid("custom XML maps string exceeds 65535 characters"));
    }
    Ok(())
}

fn optional_bounded_office_string(value: Option<&str>, limits: &XmlMapLimits) -> Result<()> {
    if let Some(v) = value {
        bounded_office_string(v, limits)
    } else {
        Ok(())
    }
}
fn xml_error(error: impl std::fmt::Display) -> crate::Error {
    invalid(error.to_string())
}
