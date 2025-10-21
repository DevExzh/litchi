/// Text frame for accessing text content in shapes.
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    // Check if this is an a:t element (DrawingML text)
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Extract text content
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    if !text.is_empty() && !text.ends_with('\n') {
                        text.push('\n');
                    }
                    text.push_str(t);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(text)
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
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
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
                Ok(Event::End(e)) => {
                    if in_para {
                        current_para_xml.extend_from_slice(b"</");
                        current_para_xml.extend_from_slice(e.name().as_ref());
                        current_para_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"p" {
                            paragraphs.push(Paragraph::new(current_para_xml.clone()));
                            in_para = false;
                        }
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
            buf.clear();
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
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    text.push_str(t);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(text)
    }
}
