//! XLSX compatibility forwarding for the shared XML Maps codec.

use litchi_core::sheet::Result;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::model::{XmlMapConformance, XmlMapInfo};

impl XmlMapInfo {
    /// Parse a bounded SpreadsheetML XML Maps part.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        match litchi_ooxml_common::spreadsheet_xml_maps::parse_xml_map_info(xml) {
            Ok(value) => Ok(value.into()),
            Err(error) if error.to_string().starts_with("invalid boolean attribute '") => {
                // Excel producers historically emitted the XML Schema-compatible
                // `0`/`1` spellings here. Keep that XLSX reader compatibility at
                // this boundary while common and XLSB parsing remain strict.
                let Some(normalized) = normalize_legacy_boolean_attributes(xml) else {
                    return Err(error.into());
                };
                Ok(
                    litchi_ooxml_common::spreadsheet_xml_maps::parse_xml_map_info(&normalized)?
                        .into(),
                )
            },
            Err(error) => Err(error.into()),
        }
    }

    /// Serialize this XML Maps catalog for Transitional or Strict SpreadsheetML.
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        let conformance = if strict {
            litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Strict
        } else {
            litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Transitional
        };
        Ok(
            litchi_ooxml_common::spreadsheet_xml_maps::serialize_xml_map_info_ref(
                &self.to_common_ref()?,
                conformance,
            )?,
        )
    }
}

#[derive(Clone, Copy)]
enum Context {
    Other,
    MapInfo,
    Map,
}

fn normalize_legacy_boolean_attributes(xml: &[u8]) -> Option<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut replacements = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().ok()?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let context = classify_context(
                    element.local_name().as_ref(),
                    element.name().as_ref(),
                    stack.last().copied(),
                );
                normalize_element(
                    xml,
                    start,
                    end,
                    element.local_name().as_ref(),
                    element.name().as_ref(),
                    stack.last().copied(),
                    &mut replacements,
                );
                stack.push(context);
            },
            Event::Empty(element) => normalize_element(
                xml,
                start,
                end,
                element.local_name().as_ref(),
                element.name().as_ref(),
                stack.last().copied(),
                &mut replacements,
            ),
            Event::End(_) => {
                stack.pop()?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if replacements.is_empty() {
        return None;
    }
    let extra = replacements
        .iter()
        .map(
            |(range, replacement): &(std::ops::Range<usize>, &'static [u8])| {
                replacement.len() - range.len()
            },
        )
        .sum::<usize>();
    let mut normalized = Vec::with_capacity(xml.len().saturating_add(extra));
    let mut cursor = 0;
    for (range, replacement) in replacements {
        normalized.extend_from_slice(&xml[cursor..range.start]);
        normalized.extend_from_slice(replacement);
        cursor = range.end;
    }
    normalized.extend_from_slice(&xml[cursor..]);
    Some(normalized)
}

fn classify_context(local_name: &[u8], qualified_name: &[u8], parent: Option<Context>) -> Context {
    if parent.is_none() && local_name == b"MapInfo" {
        Context::MapInfo
    } else if matches!(parent, Some(Context::MapInfo))
        && local_name == b"Map"
        && qualified_name == local_name
    {
        Context::Map
    } else {
        Context::Other
    }
}

fn normalize_element(
    xml: &[u8],
    event_start: usize,
    end: usize,
    local_name: &[u8],
    qualified_name: &[u8],
    parent: Option<Context>,
    replacements: &mut Vec<(std::ops::Range<usize>, &'static [u8])>,
) {
    let attributes: &[&[u8]] = if matches!(parent, Some(Context::MapInfo))
        && local_name == b"Map"
        && qualified_name == local_name
    {
        &[
            b"ShowImportExportValidationErrors",
            b"AutoFit",
            b"Append",
            b"PreserveSortAFLayout",
            b"PreserveFormat",
        ]
    } else if matches!(parent, Some(Context::Map))
        && local_name == b"DataBinding"
        && qualified_name == local_name
    {
        &[b"FileBinding"]
    } else {
        return;
    };
    let Some(start) = xml[event_start..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .map(|position| event_start + position)
    else {
        return;
    };
    find_numeric_boolean_attributes(&xml[start..end], start, attributes, replacements);
}

fn find_numeric_boolean_attributes(
    tag: &[u8],
    offset: usize,
    allowed: &[&[u8]],
    replacements: &mut Vec<(std::ops::Range<usize>, &'static [u8])>,
) {
    let mut position = 1;
    while position < tag.len()
        && !tag[position].is_ascii_whitespace()
        && !matches!(tag[position], b'/' | b'>')
    {
        position += 1;
    }
    while position < tag.len() {
        while position < tag.len() && tag[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= tag.len() || matches!(tag[position], b'/' | b'>') {
            break;
        }
        let name_start = position;
        while position < tag.len()
            && !tag[position].is_ascii_whitespace()
            && !matches!(tag[position], b'=' | b'/' | b'>')
        {
            position += 1;
        }
        let name = &tag[name_start..position];
        let local_name = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
        while position < tag.len() && tag[position].is_ascii_whitespace() {
            position += 1;
        }
        if tag.get(position) != Some(&b'=') {
            break;
        }
        position += 1;
        while position < tag.len() && tag[position].is_ascii_whitespace() {
            position += 1;
        }
        let Some(&quote @ (b'\'' | b'"')) = tag.get(position) else {
            break;
        };
        position += 1;
        let value_start = position;
        while position < tag.len() && tag[position] != quote {
            position += 1;
        }
        if position >= tag.len() {
            break;
        }
        if allowed.contains(&local_name) {
            match &tag[value_start..position] {
                b"0" => replacements.push((offset + value_start..offset + position, b"false")),
                b"1" => replacements.push((offset + value_start..offset + position, b"true")),
                _ => {},
            }
        }
        position += 1;
    }
}

/// Parse a bounded SpreadsheetML XML Maps part using the historical XLSX
/// boxed-error result surface.
pub fn parse_xml_map_info(xml: &[u8]) -> Result<XmlMapInfo> {
    XmlMapInfo::parse(xml)
}

/// Serialize XML Maps using the historical XLSX boxed-error result surface.
pub fn serialize_xml_map_info(
    info: &XmlMapInfo,
    conformance: XmlMapConformance,
) -> Result<Vec<u8>> {
    info.to_xml(conformance.is_strict())
}

pub(super) fn patch_source(
    source: &[u8],
    before: &XmlMapInfo,
    after: &XmlMapInfo,
    before_strict: bool,
    after_strict: bool,
) -> Result<Vec<u8>> {
    let before_conformance = if before_strict {
        litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Strict
    } else {
        litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Transitional
    };
    let after_conformance = if after_strict {
        litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Strict
    } else {
        litchi_ooxml_common::spreadsheet_xml_maps::XmlMapConformance::Transitional
    };
    Ok(
        litchi_ooxml_common::spreadsheet_xml_maps::patch_xml_map_info_source_ref(
            source,
            &before.to_common_ref()?,
            &after.to_common_ref()?,
            before_conformance,
            after_conformance,
        )?,
    )
}
