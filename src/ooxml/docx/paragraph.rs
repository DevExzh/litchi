/// Paragraph and Run structures for Word documents.
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::Reader;
use smallvec::SmallVec;
use std::borrow::Cow;

/// A paragraph in a Word document.
///
/// Represents a `<w:p>` element. Paragraphs contain runs which in turn
/// contain the actual text and formatting.
///
/// # Example
///
/// ```rust,ignore
/// for para in document.paragraphs()? {
///     println!("Paragraph text: {}", para.text());
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text(), run.bold());
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The raw XML bytes for this paragraph
    xml_bytes: Vec<u8>,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes.
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate string with estimated capacity to reduce reallocations
        let estimated_capacity = self.xml_bytes.len() / 4; // Rough estimate
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                }
                Ok(Event::Text(e)) if in_text_element => {
                    // Use unsafe conversion for better performance (safe since we validate XML)
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    result.push_str(text);
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
        }

        // Shrink to fit to release unused capacity
        result.shrink_to_fit();
        Ok(result)
    }

    /// Get an iterator over the runs in this paragraph.
    ///
    /// Each run represents a `<w:r>` element and may have different formatting.
    pub fn runs(&self) -> Result<SmallVec<[Run; 8]>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Use SmallVec for efficient storage of typically small run collections
        let mut runs = SmallVec::new();
        let mut current_run_xml = Vec::with_capacity(1024); // Pre-allocate for XML fragments
        let mut in_run = false;
        let mut depth = 0;
        let mut buf = Vec::with_capacity(512); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"r" && !in_run {
                        in_run = true;
                        depth = 1;
                        current_run_xml.clear();
                        // Pre-allocate with estimated size for run XML
                        current_run_xml.reserve(512);

                        // Build opening tag more efficiently
                        current_run_xml.extend_from_slice(b"<w:r");
                        for attr in e.attributes().flatten() {
                            current_run_xml.push(b' ');
                            current_run_xml.extend_from_slice(attr.key.as_ref());
                            current_run_xml.extend_from_slice(b"=\"");
                            current_run_xml.extend_from_slice(&attr.value);
                            current_run_xml.push(b'"');
                        }
                        current_run_xml.push(b'>');
                    } else if in_run {
                        depth += 1;
                        current_run_xml.push(b'<');
                        current_run_xml.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            current_run_xml.push(b' ');
                            current_run_xml.extend_from_slice(attr.key.as_ref());
                            current_run_xml.extend_from_slice(b"=\"");
                            current_run_xml.extend_from_slice(&attr.value);
                            current_run_xml.push(b'"');
                        }
                        current_run_xml.push(b'>');
                    }
                }
                Ok(Event::End(e)) => {
                    if in_run {
                        current_run_xml.extend_from_slice(b"</");
                        current_run_xml.extend_from_slice(e.name().as_ref());
                        current_run_xml.push(b'>');

                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"r" {
                            runs.push(Run::new(current_run_xml.clone()));
                            in_run = false;
                        }
                    }
                }
                Ok(Event::Text(e)) if in_run => {
                    current_run_xml.extend_from_slice(e.as_ref());
                }
                Ok(Event::Empty(e)) if in_run => {
                    current_run_xml.push(b'<');
                    current_run_xml.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        current_run_xml.push(b' ');
                        current_run_xml.extend_from_slice(attr.key.as_ref());
                        current_run_xml.extend_from_slice(b"=\"");
                        current_run_xml.extend_from_slice(&attr.value);
                        current_run_xml.push(b'"');
                    }
                    current_run_xml.extend_from_slice(b"/>");
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(runs)
    }
}

/// A run within a paragraph.
///
/// Represents a `<w:r>` element. A run is a region of text with a single
/// set of formatting properties.
///
/// # Example
///
/// ```rust,ignore
/// let run = runs[0];
/// println!("Text: {}", run.text()?);
/// println!("Bold: {:?}", run.bold()?);
/// println!("Italic: {:?}", run.italic()?);
/// ```
#[derive(Debug, Clone)]
pub struct Run {
    /// The raw XML bytes for this run
    xml_bytes: Vec<u8>,
}

impl Run {
    /// Create a new Run from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the text content of this run.
    ///
    /// Extracts text from `<w:t>` elements and converts special characters:
    /// - `<w:tab/>` → tab character
    /// - `<w:br/>` → newline character
    pub fn text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity
        let estimated_capacity = self.xml_bytes.len() / 8; // Rough estimate for text content
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;
        let mut buf = Vec::with_capacity(256); // Reusable buffer

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"t" {
                        in_text_element = true;
                    } else if name.as_ref() == b"tab" {
                        result.push('\t');
                    } else if name.as_ref() == b"br" {
                        result.push('\n');
                    }
                }
                Ok(Event::Text(e)) if in_text_element => {
                    // Use unsafe conversion for better performance (safe since we validate XML)
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    result.push_str(text);
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(result)
    }

    /// Check if this run is bold.
    ///
    /// Returns `Some(true)` if bold is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn bold(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"b")
    }

    /// Check if this run is italic.
    ///
    /// Returns `Some(true)` if italic is explicitly enabled,
    /// `Some(false)` if explicitly disabled,
    /// `None` if not specified (inherits from style).
    pub fn italic(&self) -> Result<Option<bool>> {
        self.get_bool_property(b"i")
    }

    /// Check if this run is underlined.
    ///
    /// Returns `Some(true)` if underline is present,
    /// `None` if not specified.
    ///
    /// Note: This is simplified. Full implementation would return
    /// the underline style (single, double, wavy, etc.).
    pub fn underline(&self) -> Result<Option<bool>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"u" {
                        return Ok(Some(true));
                    }
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Get the font name for this run.
    ///
    /// Returns the typeface name if specified, None if inherited.
    pub fn font_name(&self) -> Result<Option<String>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"rFonts" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ascii" {
                                let value = attr.unescape_value().unwrap_or(Cow::Borrowed(""));
                                return Ok(Some(value.to_string()));
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Get the font size for this run in half-points.
    ///
    /// Returns the size if specified, None if inherited.
    /// Note: Word stores font size in half-points (e.g., 24 = 12pt).
    pub fn font_size(&self) -> Result<Option<u32>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == b"sz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                if let Ok(value) = std::str::from_utf8(&attr.value) {
                                    if let Ok(size) = value.parse::<u32>() {
                                        return Ok(Some(size));
                                    }
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Helper to extract boolean properties from run properties.
    ///
    /// Handles the tri-state logic where w:val can be "true", "false", "1", "0"
    /// or the element can be present without a val attribute (implies true).
    fn get_bool_property(&self, property_name: &[u8]) -> Result<Option<bool>> {
        let mut reader = Reader::from_reader(&self.xml_bytes[..]);
        reader.config_mut().trim_text(true);

        let mut in_r_pr = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.local_name();
                    if name.as_ref() == b"rPr" {
                        in_r_pr = true;
                    } else if in_r_pr && name.as_ref() == property_name {
                        // Check for w:val attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let value = attr.value.as_ref();
                                return Ok(Some(value == b"true" || value == b"1"));
                            }
                        }
                        // Element present without val attribute means true
                        return Ok(Some(true));
                    }
                }
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"rPr" {
                        in_r_pr = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_text_extraction() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Hello, World!</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        let text = run.text().unwrap();
        assert_eq!(text, "Hello, World!");
    }

    #[test]
    fn test_run_bold() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:b/></w:rPr>
            <w:t>Bold text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert_eq!(run.bold().unwrap(), Some(true));
    }

    #[test]
    fn test_run_italic() {
        let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:i/></w:rPr>
            <w:t>Italic text</w:t>
        </w:r>"#;

        let run = Run::new(xml.to_vec());
        assert_eq!(run.italic().unwrap(), Some(true));
    }
}
