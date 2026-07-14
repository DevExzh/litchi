/// Text frame for accessing text content in shapes.
use crate::common::xml::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};

const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => {
            prefix.as_slice() == b"a"
                || fragment_prefix
                    .as_ref()
                    .and_then(|prefix| prefix.as_deref())
                    == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
    }
}

pub(crate) fn extract_drawingml_text(
    xml_bytes: &[u8],
    paragraph_separator: Option<char>,
) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut text_depth = None;
    let mut seen_paragraph = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("DrawingML nesting is too deep".to_string())
                })?;
                if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                    if seen_paragraph
                        && !result.is_empty()
                        && let Some(separator) = paragraph_separator
                        && !result.ends_with(separator)
                    {
                        result.push(separator);
                    }
                    seen_paragraph = true;
                } else if text_depth.is_none()
                    && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml_name(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Empty(element) => {
                if is_drawingml_name(&namespace, element.name(), b"p", &fragment_prefix) {
                    if seen_paragraph
                        && !result.is_empty()
                        && let Some(separator) = paragraph_separator
                        && !result.ends_with(separator)
                    {
                        result.push(separator);
                    }
                    seen_paragraph = true;
                } else if is_drawingml_name(&namespace, element.name(), b"br", &fragment_prefix) {
                    result.push('\n');
                } else if is_drawingml_name(&namespace, element.name(), b"tab", &fragment_prefix) {
                    result.push('\t');
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                result.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Event::CData(text) if text_depth.is_some() => {
                result.push_str(
                    &text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_drawingml_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid DrawingML nesting".to_string())
                })?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated DrawingML XML".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

/// A text frame containing text content.
///
/// Text frames are found in shape objects and provide access to the
/// paragraphs and text within the shape.
///
/// # Examples
///
/// ```rust,ignore
/// let text_frame = shape.text_frame()?;
/// println!("Text: {}", text_frame.text()?);
///
/// for para in text_frame.paragraphs()? {
///     println!("Paragraph: {}", para.text()?);
/// }
///
/// // Check for embedded formulas
/// for formula in text_frame.omml_formulas()? {
///     println!("Found OMML formula: {}", formula);
/// }
/// ```
#[derive(Debug, Clone)]
pub struct TextFrame {
    /// Raw XML bytes
    xml_bytes: Vec<u8>,
}

impl TextFrame {
    /// Create a TextFrame from XML bytes.
    pub(crate) fn from_xml(xml_bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            xml_bytes: xml_bytes.to_vec(),
        })
    }

    /// Extract all text from this text frame.
    ///
    /// Returns all text content concatenated together.
    pub fn text(&self) -> Result<String> {
        extract_drawingml_text(&self.xml_bytes, Some('\n'))
    }

    /// Get paragraphs in this text frame.
    ///
    /// Returns a vector of Paragraph objects.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut paragraphs = Vec::new();
        let mut current_para_xml = Vec::new();
        let mut in_para = false;
        let mut depth = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    // DrawingML paragraphs are <a:p>
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        current_para_xml.extend_from_slice(b"<a:p>");
                    } else if in_para {
                        depth += 1;
                        current_para_xml.push(b'<');
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_para_xml.push(b' ');
                            current_para_xml.extend_from_slice(attr.key.as_ref());
                            current_para_xml.extend_from_slice(b"=\"");
                            current_para_xml.extend_from_slice(&attr.value);
                            current_para_xml.push(b'"');
                        }
                        current_para_xml.push(b'>');
                    }
                },
                Ok(Event::End(e)) if in_para => {
                    current_para_xml.extend_from_slice(b"</");
                    current_para_xml.extend_from_slice(e.name().as_ref());
                    current_para_xml.push(b'>');

                    depth -= 1;
                    if depth == 0 && e.local_name().as_ref() == b"p" {
                        paragraphs.push(Paragraph::new(current_para_xml.clone()));
                        in_para = false;
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_para => {
                    current_para_xml.push(b'<');
                    current_para_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_para_xml.push(b' ');
                        current_para_xml.extend_from_slice(attr.key.as_ref());
                        current_para_xml.extend_from_slice(b"=\"");
                        current_para_xml.extend_from_slice(&attr.value);
                        current_para_xml.push(b'"');
                    }
                    current_para_xml.extend_from_slice(b"/>");
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(paragraphs)
    }

    /// Extract all OMML formulas from this text frame.
    ///
    /// Returns a vector of OMML formula strings found in any paragraph within this text frame.
    pub fn omml_formulas(&self) -> Result<Vec<String>> {
        let mut formulas = Vec::new();
        for para in self.paragraphs()? {
            // For PPTX, we need to check if the paragraph contains OMML formulas
            // This is a simplified approach - in a full implementation, we would
            // need to parse the paragraph XML for OMML content similar to how
            // we do it for DOCX runs
            if let Ok(text) = para.text() {
                // Look for OMML-like patterns in the text (simplified heuristic)
                if text.contains("oMath") || text.contains("m:oMath") {
                    // In a full implementation, we would extract the actual OMML XML
                    formulas.push(text);
                }
            }
        }
        Ok(formulas)
    }
}

/// A paragraph in a text frame.
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// Raw XML bytes for this paragraph
    xml_bytes: Vec<u8>,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Extract all text from this paragraph.
    pub fn text(&self) -> Result<String> {
        extract_drawingml_text(&self.xml_bytes, None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn drawingml_text_preserves_runs_whitespace_and_paragraphs() {
        let xml = br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml">
            <d:txBody>
                <d:p><d:r><d:t xml:space="preserve"> A &amp; </d:t></d:r><d:r><false:t>ignored</false:t><d:t><![CDATA[B < C]]></d:t><d:tab/><d:br/></d:r></d:p>
                <d:p><d:r><d:t>Second</d:t></d:r></d:p>
            </d:txBody>
        </p:sp>"#;
        let frame = TextFrame::from_xml(xml).unwrap();
        assert_eq!(frame.text().unwrap(), " A & B < C\t\nSecond");
    }

    #[test]
    fn drawingml_paragraph_text_accepts_inherited_conventional_prefix() {
        let paragraph = Paragraph::new(
            br#"<a:p><a:r><a:t>one</a:t></a:r><a:r><a:t>two</a:t></a:r></a:p>"#.to_vec(),
        );
        assert_eq!(paragraph.text().unwrap(), "onetwo");
    }

    #[test]
    fn drawingml_text_rejects_foreign_lookalikes_and_truncation() {
        let foreign = Paragraph::new(
            br#"<x:p xmlns:x="urn:not-drawingml"><x:r><x:t>ignored</x:t></x:r></x:p>"#.to_vec(),
        );
        assert_eq!(foreign.text().unwrap(), "");

        let truncated = Paragraph::new(
            br#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>bad</a:t>"#
                .to_vec(),
        );
        assert!(truncated.text().is_err());
    }
}
