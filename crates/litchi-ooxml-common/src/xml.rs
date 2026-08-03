//! Shared, namespace-aware OOXML decoding helpers.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use thiserror::Error;

/// Result of a shared OOXML decoding operation.
pub type Result<T> = std::result::Result<T, XmlError>;

/// Failure to decode or structurally scan shared OOXML markup.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum XmlError {
    /// The XML stream is not well formed or contains an invalid encoded value.
    #[error("malformed OOXML markup: {0}")]
    Malformed(String),
    /// The XML is well formed but violates a required OOXML invariant.
    #[error("invalid OOXML structure: {0}")]
    Invalid(String),
}

/// Transitional DrawingML main namespace.
pub const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
/// Strict DrawingML main namespace.
pub const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
/// Transitional DrawingML chart namespace.
pub const DRAWINGML_CHART_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/chart";
/// Strict DrawingML chart namespace.
pub const STRICT_DRAWINGML_CHART_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
/// Transitional Office Math namespace URI.
pub const OMML_NAMESPACE_URI: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const OMML_NAMESPACE: &[u8] = OMML_NAMESPACE_URI.as_bytes();
const STRICT_OMML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";

/// Decode a numeric or predefined XML entity reference.
pub fn decode_xml_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| XmlError::Malformed(error.to_string()))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| XmlError::Malformed(error.to_string()))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(XmlError::Invalid(format!(
            "unsupported XML entity reference '&{name};'"
        ))),
    }
}

pub fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        )
}

pub fn is_drawingml_chart_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == DRAWINGML_CHART_NAMESPACE
                    || *value == STRICT_DRAWINGML_CHART_NAMESPACE
        )
}

pub fn unqualified_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| XmlError::Malformed(error.to_string()))?;
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if value.is_some() {
                return Err(XmlError::Invalid(format!(
                    "duplicate XML attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                    .map_err(|error| XmlError::Malformed(error.to_string()))?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

/// Return whether `value` is an XML 1.0 Fifth Edition NCName.
///
/// Relationship IDs and namespace prefixes use this Unicode-aware grammar;
/// ASCII-only approximations reject valid producer documents.
pub fn is_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_character)
}

fn is_ncname_start(character: char) -> bool {
    character != ':' && is_name_start(character)
}

fn is_ncname_character(character: char) -> bool {
    character != ':'
        && (is_name_start(character)
            || matches!(
                character,
                '-' | '.' | '0'..='9' | '\u{B7}' | '\u{300}'..='\u{36F}' | '\u{203F}'..='\u{2040}'
            ))
}

fn is_name_start(character: char) -> bool {
    matches!(
        character,
        ':' | 'A'..='Z' | '_' | 'a'..='z'
            | '\u{C0}'..='\u{D6}'
            | '\u{D8}'..='\u{F6}'
            | '\u{F8}'..='\u{2FF}'
            | '\u{370}'..='\u{37D}'
            | '\u{37F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

pub fn is_omml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
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
pub fn scan_omml_formula_ranges<E>(
    xml_bytes: &[u8],
    mut emit: impl FnMut(u32, u32) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<XmlError>,
{
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
            .map_err(|_| XmlError::Invalid("OMML offset does not fit usize".to_string()))?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| XmlError::Malformed(error.to_string()))?;
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
            .map_err(|_| XmlError::Invalid("OMML offset does not fit usize".to_string()))?;

        match event {
            ScanEvent::Start => capture = Some((event_start, 1)),
            ScanEvent::NestedStart => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(
                        XmlError::Invalid("missing captured OMML formula".to_string()).into(),
                    );
                };
                *depth = depth
                    .checked_add(1)
                    .ok_or_else(|| XmlError::Invalid("OMML nesting is too deep".to_string()))?;
            },
            ScanEvent::Empty => emit_omml_range(event_start, event_end, &mut emit)?,
            ScanEvent::End => {
                let Some((_, depth)) = capture.as_mut() else {
                    return Err(
                        XmlError::Invalid("missing captured OMML formula".to_string()).into(),
                    );
                };
                *depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| XmlError::Invalid("invalid OMML nesting".to_string()))?;
                if *depth == 0 {
                    let Some((start, _)) = capture.take() else {
                        return Err(
                            XmlError::Invalid("missing OMML formula range".to_string()).into()
                        );
                    };
                    emit_omml_range(start, event_end, &mut emit)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(XmlError::Invalid("unterminated OMML formula".to_string()).into());
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }
    Ok(())
}

fn emit_omml_range<E>(
    start: usize,
    end: usize,
    emit: &mut impl FnMut(u32, u32) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<XmlError>,
{
    let length = end
        .checked_sub(start)
        .ok_or_else(|| XmlError::Invalid("invalid OMML formula range".to_string()))?;
    emit(
        u32::try_from(start)
            .map_err(|_| XmlError::Invalid("OMML offset exceeds u32".to_string()))?,
        u32::try_from(length)
            .map_err(|_| XmlError::Invalid("OMML length exceeds u32".to_string()))?,
    )
}

/// Copy exact OMML formula XML into the string-based public API representation.
pub fn extract_omml_formulas(xml_bytes: &[u8]) -> Result<Vec<String>> {
    let mut formulas = Vec::new();
    scan_omml_formula_ranges(xml_bytes, |start, length| {
        formulas.push(omml_formula_xml(xml_bytes, start, length)?);
        Ok(())
    })?;
    Ok(formulas)
}

pub fn omml_formula_xml(xml_bytes: &[u8], start: u32, length: u32) -> Result<String> {
    let start = start as usize;
    let end = start
        .checked_add(length as usize)
        .ok_or_else(|| XmlError::Invalid("OMML formula range overflows".to_string()))?;
    let formula = xml_bytes
        .get(start..end)
        .ok_or_else(|| XmlError::Invalid("invalid OMML formula range".to_string()))?;
    Ok(std::str::from_utf8(formula)
        .map_err(|_| XmlError::Invalid("OMML formula is not UTF-8".to_string()))?
        .to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::Reader;

    #[test]
    fn duplicate_unqualified_attributes_are_rejected() {
        let mut reader = Reader::from_str(r#"<node value="one" value="two"/>"#);
        let Event::Empty(element) = reader.read_event().expect("well-formed event") else {
            panic!("expected an empty element");
        };
        let error = unqualified_attribute_value(&element, b"value", reader.decoder())
            .expect_err("duplicate is invalid");
        assert!(matches!(
            error,
            XmlError::Malformed(_) | XmlError::Invalid(_)
        ));
    }

    #[test]
    fn ncname_accepts_unicode_and_rejects_colons_or_invalid_starts() {
        assert!(is_ncname("rId1"));
        assert!(is_ncname("关系一"));
        assert!(is_ncname("éclair.一"));
        assert!(!is_ncname(""));
        assert!(!is_ncname("1relationship"));
        assert!(!is_ncname("r:id"));
        assert!(!is_ncname("relationship id"));
    }

    #[test]
    fn scans_exact_transitional_and_strict_omml_ranges() {
        let xml = concat!(
            r#"<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" "#,
            r#"xmlns:s="http://purl.oclc.org/ooxml/officeDocument/math">"#,
            r#"<m:oMath><m:r/></m:oMath><s:oMath/></root>"#
        );
        let mut ranges = Vec::new();
        scan_omml_formula_ranges::<XmlError>(xml.as_bytes(), |start, length| {
            ranges.push((start, length));
            Ok(())
        })
        .expect("valid formulas");

        let formulas = ranges
            .into_iter()
            .map(|(start, length)| omml_formula_xml(xml.as_bytes(), start, length))
            .collect::<Result<Vec<_>>>()
            .expect("valid ranges");
        assert_eq!(formulas, ["<m:oMath><m:r/></m:oMath>", "<s:oMath/>"]);
    }

    #[test]
    fn rejects_unterminated_omml_without_emitting_partial_content() {
        let xml = br#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r>"#;
        let mut emitted = 0usize;
        let error = scan_omml_formula_ranges::<XmlError>(xml, |_, _| {
            emitted += 1;
            Ok(())
        })
        .expect_err("unterminated formula");
        assert!(matches!(
            error,
            XmlError::Malformed(_) | XmlError::Invalid(_)
        ));
        assert_eq!(emitted, 0);
    }
}
