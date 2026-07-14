//! Shared XML decoding helpers.

use crate::error::{OoxmlError, Result};
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const OMML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const STRICT_OMML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";

/// Decode a numeric or predefined XML entity reference.
pub(crate) fn decode_xml_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| OoxmlError::Xml(error.to_string()))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "unsupported XML entity reference '&{name};'"
        ))),
    }
}

fn is_omml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == OMML_NAMESPACE || *value == STRICT_OMML_NAMESPACE
        },
        // Paragraphs and runs are often slices whose namespace declarations live on an
        // ancestor. Accept the conventional prefix only while it is unresolved; an explicit
        // foreign binding reaches the Bound branch and is rejected.
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"m",
        ResolveResult::Unbound => false,
    }
}

/// Locate exact `<oMath>` byte ranges in transitional or strict OMML XML.
pub(crate) fn scan_omml_formula_ranges(
    xml_bytes: &[u8],
    mut emit: impl FnMut(u32, u32) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start,
        NestedStart,
        Empty,
        End,
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    let mut capture: Option<(usize, usize)> = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("OMML offset does not fit usize".to_string()))?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element) if is_omml_name(&namespace, element.name(), b"oMath") => {
                    ScanEvent::Start
                },
                Event::Empty(element)
                    if capture.is_none() && is_omml_name(&namespace, element.name(), b"oMath") =>
                {
                    ScanEvent::Empty
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                _ => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("OMML offset does not fit usize".to_string()))?;

        match event {
            ScanEvent::Start => capture = Some((event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured OMML formula".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("OMML nesting is too deep".to_string())
                })?;
            },
            ScanEvent::Empty => emit_omml_range(event_start, event_end, &mut emit)?,
            ScanEvent::End => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(OoxmlError::InvalidFormat(
                        "missing captured OMML formula".to_string(),
                    ));
                };
                *depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| OoxmlError::InvalidFormat("invalid OMML nesting".to_string()))?;
                if *depth == 0 {
                    let Some((start, _)) = capture.take() else {
                        return Err(OoxmlError::InvalidFormat(
                            "missing OMML formula range".to_string(),
                        ));
                    };
                    emit_omml_range(start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated OMML formula".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    Ok(())
}

fn emit_omml_range(
    start: usize,
    end: usize,
    emit: &mut impl FnMut(u32, u32) -> Result<()>,
) -> Result<()> {
    let length = end
        .checked_sub(start)
        .ok_or_else(|| OoxmlError::InvalidFormat("invalid OMML formula range".to_string()))?;
    emit(
        u32::try_from(start)
            .map_err(|_| OoxmlError::InvalidFormat("OMML offset exceeds u32".to_string()))?,
        u32::try_from(length)
            .map_err(|_| OoxmlError::InvalidFormat("OMML length exceeds u32".to_string()))?,
    )
}

/// Copy exact OMML formula XML into the string-based public API representation.
pub(crate) fn extract_omml_formulas(xml_bytes: &[u8]) -> Result<Vec<String>> {
    let mut formulas = Vec::new();
    scan_omml_formula_ranges(xml_bytes, |start, length| {
        formulas.push(omml_formula_xml(xml_bytes, start, length)?);
        Ok(())
    })?;
    Ok(formulas)
}

pub(crate) fn omml_formula_xml(xml_bytes: &[u8], start: u32, length: u32) -> Result<String> {
    let start = start as usize;
    let end = start
        .checked_add(length as usize)
        .ok_or_else(|| OoxmlError::InvalidFormat("OMML formula range overflows".to_string()))?;
    let formula = xml_bytes
        .get(start..end)
        .ok_or_else(|| OoxmlError::InvalidFormat("invalid OMML formula range".to_string()))?;
    Ok(std::str::from_utf8(formula)
        .map_err(|_| OoxmlError::InvalidFormat("OMML formula is not UTF-8".to_string()))?
        .to_owned())
}
